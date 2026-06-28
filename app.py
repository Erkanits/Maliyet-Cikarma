
import base64
import io
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd
import streamlit as st
import xlwt
from supabase import create_client


st.set_page_config(
    page_title="ITSSystems Cost Calculator",
    page_icon="€",
    layout="wide",
)


# -----------------------------
# GİRİŞ
# -----------------------------
def check_password():
    if st.session_state.get("authenticated"):
        return True

    st.title("ITSSystems Cost Calculator")
    st.caption("Devam etmek için uygulama şifresini girin.")

    password = st.text_input("Şifre", type="password")
    if st.button("Giriş yap", type="primary"):
        if password == st.secrets["APP_PASSWORD"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Şifre hatalı.")

    return False


if not check_password():
    st.stop()


# -----------------------------
# SUPABASE
# -----------------------------
@st.cache_resource
def get_db():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_ROLE_KEY"],
    )


db = get_db()


# -----------------------------
# YARDIMCI FONKSİYONLAR
# -----------------------------
def parse_decimal(value, default=None):
    if value is None:
        return default

    text = str(value).strip()
    if not text:
        return default

    text = (
        text.replace("€", "")
        .replace("₺", "")
        .replace("TL", "")
        .replace(" ", "")
    )

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(Decimal(text))
    except (InvalidOperation, ValueError):
        return default


def format_number(value, digits=2):
    return (
        f"{float(value):,.{digits}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def format_eur(value):
    return f"{format_number(value)} €"


def format_tl(value):
    return f"{format_number(value)} TL"


def convert_price(source_value, source_currency, eur_tl_rate):
    source_value = float(source_value or 0)
    eur_tl_rate = float(eur_tl_rate or 1)

    if source_currency == "TL":
        tl_value = source_value
        eur_value = source_value / eur_tl_rate if eur_tl_rate else 0
    else:
        eur_value = source_value
        tl_value = source_value * eur_tl_rate

    return eur_value, tl_value


def get_settings():
    response = db.table("ayarlar").select("*").eq("id", 1).execute()

    if response.data:
        return response.data[0]

    default_settings = {
        "id": 1,
        "eur_tl_kuru": 50,
        "logo_base64": None,
    }
    db.table("ayarlar").insert(default_settings).execute()
    return default_settings


def update_settings(rate=None, logo_base64="__KEEP__"):
    current = get_settings()

    payload = {
        "id": 1,
        "eur_tl_kuru": (
            float(rate)
            if rate is not None
            else float(current.get("eur_tl_kuru") or 50)
        ),
        "logo_base64": (
            current.get("logo_base64")
            if logo_base64 == "__KEEP__"
            else logo_base64
        ),
    }

    db.table("ayarlar").upsert(payload).execute()


def get_prices():
    response = (
        db.table("fiyat_tanimlari")
        .select("*")
        .order("ad")
        .execute()
    )
    return response.data or []


def get_labors():
    response = (
        db.table("iscilik_tanimlari")
        .select("*")
        .order("ad")
        .execute()
    )
    return response.data or []


def get_parts():
    response = (
        db.table("parcalar")
        .select("*")
        .order("created_at", desc=True)
        .execute()
    )
    return response.data or []


def get_part_items(part_id=None):
    query = db.table("parca_kalemleri").select("*")

    if part_id is not None:
        query = query.eq("parca_id", part_id)

    return query.execute().data or []


def get_part_labors(part_id=None):
    query = db.table("parca_iscilik_kalemleri").select("*")

    if part_id is not None:
        query = query.eq("parca_id", part_id)

    return query.execute().data or []


def get_price_source(item):
    currency = item.get("kaynak_para_birimi") or "EUR"
    source_value = item.get("kaynak_birim_fiyat")

    if source_value is None:
        source_value = item.get("birim_fiyat_eur", 0)

    return currency, float(source_value or 0)


def get_labor_source(item):
    currency = item.get("kaynak_para_birimi") or "EUR"
    source_value = item.get("kaynak_saatlik_ucret") or 0
    return currency, float(source_value)


def set_hours_from_minutes(hours_key, minutes_key):
    st.session_state[hours_key] = (
        float(st.session_state.get(minutes_key, 0)) / 60
    )


# -----------------------------
# BAŞLIK, LOGO VE KUR
# -----------------------------
def render_header(settings):
    left, right = st.columns([4, 1])

    with left:
        st.title("ITSSystems Cost Calculator")
        st.caption(
            "Malzeme, kaplama, ölçüm, ek işlem ve işçilik maliyet hesaplama uygulaması"
        )

    with right:
        logo_value = settings.get("logo_base64")

        if logo_value:
            try:
                logo_bytes = base64.b64decode(logo_value)
                st.image(io.BytesIO(logo_bytes), width=180)
            except Exception:
                st.warning("Logo görüntülenemedi.")

        with st.expander("Logo yükle / değiştir"):
            logo_file = st.file_uploader(
                "PNG veya JPEG seç",
                type=["png", "jpg", "jpeg"],
                key="logo_file",
            )

            if st.button(
                "Logoyu kaydet",
                key="save_logo",
                use_container_width=True,
            ):
                if logo_file is None:
                    st.warning("Önce bir logo dosyası seç.")
                elif logo_file.size > 2_000_000:
                    st.error("Logo dosyası 2 MB'tan küçük olmalı.")
                else:
                    encoded = base64.b64encode(
                        logo_file.getvalue()
                    ).decode("utf-8")
                    update_settings(logo_base64=encoded)
                    st.success("Logo kaydedildi.")
                    st.rerun()

            if logo_value and st.button(
                "Logoyu kaldır",
                key="remove_logo",
                use_container_width=True,
            ):
                update_settings(logo_base64=None)
                st.rerun()


def render_rate(settings):
    current_rate = float(settings.get("eur_tl_kuru") or 50)

    with st.expander("EUR / TL Kur Ayarı", expanded=True):
        col1, col2, col3 = st.columns([1.3, 1, 3])

        with col1:
            rate_text = st.text_input(
                "1 EUR kaç TL?",
                value=format_number(current_rate, 4),
                key="rate_input",
            )

        with col2:
            st.write("")
            st.write("")
            save_rate = st.button(
                "Kuru kaydet",
                type="primary",
                key="save_rate",
                use_container_width=True,
            )

        with col3:
            st.info(
                "TL olarak girilen fiyatın TL değeri sabit kalır, EUR karşılığı kura göre değişir. "
                "EUR olarak girilen fiyatın EUR değeri sabit kalır, TL karşılığı kura göre değişir."
            )

        if save_rate:
            new_rate = parse_decimal(rate_text)

            if new_rate is None or new_rate <= 0:
                st.error("Geçerli ve sıfırdan büyük bir kur gir.")
            else:
                update_settings(rate=new_rate)
                st.success("Kur güncellendi.")
                st.rerun()


# -----------------------------
# XLS DIŞA AKTARMA
# -----------------------------
def build_export_frames(rate):
    prices = get_prices()
    labors = get_labors()
    parts = get_parts()
    part_items = get_part_items()
    part_labors = get_part_labors()

    price_map = {item["id"]: item for item in prices}
    labor_map = {item["id"]: item for item in labors}

    price_rows = []
    for item in prices:
        currency, source_value = get_price_source(item)
        eur_value, tl_value = convert_price(
            source_value,
            currency,
            rate,
        )

        price_rows.append(
            {
                "Ad": item["ad"],
                "Kategori": item["kategori"],
                "Açıklama": item.get("aciklama", ""),
                "Giriş Para Birimi": currency,
                "Giriş Fiyatı": source_value,
                "EUR Karşılığı": eur_value,
                "TL Karşılığı": tl_value,
            }
        )

    labor_rows = []
    for item in labors:
        currency, source_value = get_labor_source(item)
        eur_value, tl_value = convert_price(
            source_value,
            currency,
            rate,
        )

        labor_rows.append(
            {
                "İşçilik Adı": item["ad"],
                "Açıklama": item.get("aciklama", ""),
                "Giriş Para Birimi": currency,
                "Saatlik Giriş Fiyatı": source_value,
                "Saatlik EUR": eur_value,
                "Saatlik TL": tl_value,
            }
        )

    summary_rows = []
    item_detail_rows = []
    labor_detail_rows = []

    for part in parts:
        part_id = part["id"]
        production_qty = int(part["adet"])
        single_eur = 0.0
        single_tl = 0.0

        for row in [
            x for x in part_items
            if x["parca_id"] == part_id
        ]:
            definition = price_map.get(
                row["fiyat_tanimi_id"],
                {},
            )
            currency = row.get("kaynak_para_birimi") or "EUR"
            source_value = row.get("kaynak_birim_fiyat")

            if source_value is None:
                source_value = row.get("birim_fiyat_eur", 0)

            unit_eur, unit_tl = convert_price(
                float(source_value or 0),
                currency,
                rate,
            )
            quantity = int(row.get("miktar") or 0)
            row_eur = unit_eur * quantity
            row_tl = unit_tl * quantity

            single_eur += row_eur
            single_tl += row_tl

            item_detail_rows.append(
                {
                    "Parça Adı": part["parca_adi"],
                    "Üretim Adedi": production_qty,
                    "Kalem Adı": definition.get("ad", ""),
                    "Kategori": definition.get("kategori", ""),
                    "Miktar / Tek Parça": quantity,
                    "Tek Parça Kalem EUR": row_eur,
                    "Tek Parça Kalem TL": row_tl,
                    "Genel Kalem EUR": row_eur * production_qty,
                    "Genel Kalem TL": row_tl * production_qty,
                }
            )

        for row in [
            x for x in part_labors
            if x["parca_id"] == part_id
        ]:
            definition = labor_map.get(
                row["iscilik_tanimi_id"],
                {},
            )
            currency = row.get("kaynak_para_birimi") or "EUR"
            source_value = float(
                row.get("kaynak_saatlik_ucret") or 0
            )
            hourly_eur, hourly_tl = convert_price(
                source_value,
                currency,
                rate,
            )
            hours = float(row.get("saat") or 0)
            row_eur = hourly_eur * hours
            row_tl = hourly_tl * hours

            single_eur += row_eur
            single_tl += row_tl

            labor_detail_rows.append(
                {
                    "Parça Adı": part["parca_adi"],
                    "Üretim Adedi": production_qty,
                    "İşçilik": definition.get("ad", ""),
                    "Saat / Tek Parça": hours,
                    "Tek Parça İşçilik EUR": row_eur,
                    "Tek Parça İşçilik TL": row_tl,
                    "Genel İşçilik EUR": row_eur * production_qty,
                    "Genel İşçilik TL": row_tl * production_qty,
                }
            )

        summary_rows.append(
            {
                "Parça Adı": part["parca_adi"],
                "Adet": production_qty,
                "Tek Parça EUR": single_eur,
                "Tek Parça TL": single_tl,
                "Genel Toplam EUR": single_eur * production_qty,
                "Genel Toplam TL": single_tl * production_qty,
                "Kullanılan EUR/TL Kuru": rate,
            }
        )

    return {
        "Fiyat Tanımları": pd.DataFrame(price_rows),
        "İşçilik Maliyetleri": pd.DataFrame(labor_rows),
        "Parça Özeti": pd.DataFrame(summary_rows),
        "Maliyet Detayları": pd.DataFrame(item_detail_rows),
        "İşçilik Detayları": pd.DataFrame(labor_detail_rows),
    }


def build_xls(rate):
    frames = build_export_frames(rate)
    workbook = xlwt.Workbook(encoding="utf-8")

    header_style = xlwt.easyxf(
        "font: bold on; "
        "pattern: pattern solid, fore_colour gray25;"
    )
    text_style = xlwt.easyxf()
    number_style = xlwt.easyxf(num_format_str="#,##0.0000")
    money_style = xlwt.easyxf(num_format_str="#,##0.00")

    for sheet_name, dataframe in frames.items():
        sheet = workbook.add_sheet(sheet_name[:31])

        if dataframe.empty:
            sheet.write(0, 0, "Kayıt bulunmuyor.", text_style)
            continue

        for column_index, column_name in enumerate(
            dataframe.columns
        ):
            sheet.write(
                0,
                column_index,
                column_name,
                header_style,
            )
            width = min(max(len(column_name) + 4, 15), 45)
            sheet.col(column_index).width = width * 256

        for row_index, row in enumerate(
            dataframe.itertuples(index=False),
            start=1,
        ):
            for column_index, value in enumerate(row):
                if isinstance(value, bool):
                    sheet.write(
                        row_index,
                        column_index,
                        str(value),
                        text_style,
                    )
                elif isinstance(value, int):
                    sheet.write(
                        row_index,
                        column_index,
                        value,
                        number_style,
                    )
                elif isinstance(value, float):
                    sheet.write(
                        row_index,
                        column_index,
                        value,
                        money_style,
                    )
                else:
                    sheet.write(
                        row_index,
                        column_index,
                        "" if value is None else str(value),
                        text_style,
                    )

        sheet.set_panes_frozen(True)
        sheet.set_horz_split_pos(1)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


# -----------------------------
# UYGULAMA BAŞLANGICI
# -----------------------------
settings = get_settings()
exchange_rate = float(settings.get("eur_tl_kuru") or 50)

render_header(settings)
render_rate(settings)

tab_prices, tab_labors, tab_saved, tab_cost = st.tabs(
    [
        "Fiyat Tanımları",
        "İşçilik Maliyetleri",
        "Kayıtlı Parçalar",
        "Parça Maliyeti",
    ]
)


# -----------------------------
# FİYAT TANIMLARI
# -----------------------------
with tab_prices:
    st.subheader("Yeni fiyat tanımı")

    col1, col2 = st.columns(2)

    with col1:
        new_category = st.selectbox(
            "Kategori",
            ["Malzeme", "Kaplama", "Ek İşlem", "Ölçüm"],
            key="new_price_category",
        )
        new_name = st.text_input(
            "Ad",
            key="new_price_name",
        )
        new_description = st.text_input(
            "Açıklama",
            key="new_price_description",
        )

    with col2:
        new_currency = st.selectbox(
            "Fiyat para birimi",
            ["EUR", "TL"],
            key="new_price_currency",
        )
        new_price_text = st.text_input(
            f"Birim fiyat ({new_currency})",
            value="0,00",
            key="new_price_value",
        )

        new_price = parse_decimal(new_price_text, 0)
        new_eur, new_tl = convert_price(
            new_price,
            new_currency,
            exchange_rate,
        )
        st.info(
            f"Karşılık: {format_eur(new_eur)} | "
            f"{format_tl(new_tl)}"
        )

    if st.button(
        "Fiyat tanımını kaydet",
        type="primary",
        key="save_new_price",
    ):
        new_price = parse_decimal(new_price_text)

        if not new_name.strip():
            st.error("Ad alanı boş bırakılamaz.")
        elif new_price is None or new_price < 0:
            st.error("Geçerli bir fiyat gir.")
        else:
            eur_snapshot, _ = convert_price(
                new_price,
                new_currency,
                exchange_rate,
            )

            db.table("fiyat_tanimlari").insert(
                {
                    "kategori": new_category,
                    "ad": new_name.strip(),
                    "aciklama": new_description.strip(),
                    "kaynak_para_birimi": new_currency,
                    "kaynak_birim_fiyat": new_price,
                    "birim_fiyat_eur": eur_snapshot,
                }
            ).execute()

            st.success("Fiyat tanımı kaydedildi.")
            st.rerun()

    st.divider()
    st.subheader("Kayıtlı fiyatlar")

    selected_filter = st.selectbox(
        "Kategoriye göre filtrele",
        ["Tümü", "Malzeme", "Kaplama", "Ek İşlem", "Ölçüm"],
        key="price_filter",
    )

    prices = get_prices()

    if selected_filter != "Tümü":
        prices = [
            item
            for item in prices
            if item["kategori"] == selected_filter
        ]

    if not prices:
        st.info("Bu filtreye uygun fiyat tanımı bulunmuyor.")
    else:
        table_rows = []

        for item in prices:
            currency, source_value = get_price_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            table_rows.append(
                {
                    "Ad": item["ad"],
                    "Kategori": item["kategori"],
                    "Açıklama": item.get("aciklama", ""),
                    "Giriş Fiyatı": (
                        f"{format_number(source_value, 4)} "
                        f"{currency}"
                    ),
                    "EUR Karşılığı": format_eur(eur_value),
                    "TL Karşılığı": format_tl(tl_value),
                }
            )

        st.dataframe(
            pd.DataFrame(table_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Düzenle veya sil")

        for item in prices:
            currency, source_value = get_price_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            expander_title = (
                f'{item["ad"]} | {item["kategori"]} | '
                f'{format_eur(eur_value)} / '
                f'{format_tl(tl_value)}'
            )

            with st.expander(
                expander_title,
                expanded=False,
            ):
                edit_col1, edit_col2 = st.columns(2)

                with edit_col1:
                    categories = [
                        "Malzeme",
                        "Kaplama",
                        "Ek İşlem",
                        "Ölçüm",
                    ]
                    edit_category = st.selectbox(
                        "Kategori",
                        categories,
                        index=categories.index(
                            item["kategori"]
                        ),
                        key=f'edit_category_{item["id"]}',
                    )
                    edit_name = st.text_input(
                        "Ad",
                        value=item["ad"],
                        key=f'edit_name_{item["id"]}',
                    )
                    edit_description = st.text_input(
                        "Açıklama",
                        value=item.get("aciklama", "") or "",
                        key=f'edit_description_{item["id"]}',
                    )

                with edit_col2:
                    edit_currency = st.selectbox(
                        "Fiyat para birimi",
                        ["EUR", "TL"],
                        index=["EUR", "TL"].index(currency),
                        key=f'edit_currency_{item["id"]}',
                    )
                    edit_price_text = st.text_input(
                        f"Birim fiyat ({edit_currency})",
                        value=format_number(
                            source_value,
                            4,
                        ),
                        key=f'edit_price_{item["id"]}',
                    )

                    edit_price = parse_decimal(
                        edit_price_text,
                        0,
                    )
                    edit_eur, edit_tl = convert_price(
                        edit_price,
                        edit_currency,
                        exchange_rate,
                    )
                    st.info(
                        f"Karşılık: {format_eur(edit_eur)} | "
                        f"{format_tl(edit_tl)}"
                    )

                update_col, delete_col = st.columns(2)

                with update_col:
                    if st.button(
                        "Güncelle",
                        type="primary",
                        use_container_width=True,
                        key=f'update_price_{item["id"]}',
                    ):
                        edit_price = parse_decimal(
                            edit_price_text
                        )

                        if not edit_name.strip():
                            st.error(
                                "Ad alanı boş bırakılamaz."
                            )
                        elif (
                            edit_price is None
                            or edit_price < 0
                        ):
                            st.error("Geçerli bir fiyat gir.")
                        else:
                            eur_snapshot, _ = convert_price(
                                edit_price,
                                edit_currency,
                                exchange_rate,
                            )

                            db.table(
                                "fiyat_tanimlari"
                            ).update(
                                {
                                    "kategori": edit_category,
                                    "ad": edit_name.strip(),
                                    "aciklama": (
                                        edit_description.strip()
                                    ),
                                    "kaynak_para_birimi": (
                                        edit_currency
                                    ),
                                    "kaynak_birim_fiyat": (
                                        edit_price
                                    ),
                                    "birim_fiyat_eur": (
                                        eur_snapshot
                                    ),
                                }
                            ).eq(
                                "id",
                                item["id"],
                            ).execute()

                            st.success("Kayıt güncellendi.")
                            st.rerun()

                with delete_col:
                    if st.button(
                        "Sil",
                        use_container_width=True,
                        key=f'delete_price_{item["id"]}',
                    ):
                        try:
                            db.table(
                                "fiyat_tanimlari"
                            ).delete().eq(
                                "id",
                                item["id"],
                            ).execute()

                            st.success("Kayıt silindi.")
                            st.rerun()
                        except Exception:
                            st.error(
                                "Bu fiyat tanımı kayıtlı bir "
                                "parçada kullanıldığı için silinemedi."
                            )


# -----------------------------
# İŞÇİLİK MALİYETLERİ
# -----------------------------
with tab_labors:
    st.subheader("Yeni işçilik maliyeti")

    col1, col2 = st.columns(2)

    with col1:
        labor_name = st.text_input(
            "İşçilik adı",
            placeholder="Örn. Dik İşlem CNC, Torna",
            key="new_labor_name",
        )
        labor_description = st.text_input(
            "Açıklama",
            key="new_labor_description",
        )

    with col2:
        labor_currency = st.selectbox(
            "Saatlik ücret para birimi",
            ["EUR", "TL"],
            key="new_labor_currency",
        )
        labor_price_text = st.text_input(
            f"Saatlik ücret ({labor_currency})",
            value="0,00",
            key="new_labor_price",
        )

        labor_price = parse_decimal(labor_price_text, 0)
        labor_eur, labor_tl = convert_price(
            labor_price,
            labor_currency,
            exchange_rate,
        )
        st.info(
            f"Saatlik karşılık: "
            f"{format_eur(labor_eur)} | "
            f"{format_tl(labor_tl)}"
        )

    if st.button(
        "İşçilik tanımını kaydet",
        type="primary",
        key="save_new_labor",
    ):
        labor_price = parse_decimal(labor_price_text)

        if not labor_name.strip():
            st.error("İşçilik adı boş bırakılamaz.")
        elif labor_price is None or labor_price < 0:
            st.error("Geçerli bir saatlik ücret gir.")
        else:
            db.table("iscilik_tanimlari").insert(
                {
                    "ad": labor_name.strip(),
                    "aciklama": labor_description.strip(),
                    "kaynak_para_birimi": labor_currency,
                    "kaynak_saatlik_ucret": labor_price,
                }
            ).execute()

            st.success("İşçilik tanımı kaydedildi.")
            st.rerun()

    st.divider()
    st.subheader("Kayıtlı işçilikler")

    labors = get_labors()

    if not labors:
        st.info("Henüz işçilik maliyeti tanımlanmadı.")
    else:
        labor_table_rows = []

        for item in labors:
            currency, source_value = get_labor_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            labor_table_rows.append(
                {
                    "İşçilik Adı": item["ad"],
                    "Açıklama": item.get("aciklama", ""),
                    "Giriş Fiyatı": (
                        f"{format_number(source_value, 4)} "
                        f"{currency}/saat"
                    ),
                    "EUR / Saat": format_eur(eur_value),
                    "TL / Saat": format_tl(tl_value),
                }
            )

        st.dataframe(
            pd.DataFrame(labor_table_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Düzenle veya sil")

        for item in labors:
            currency, source_value = get_labor_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            title = (
                f'{item["ad"]} | '
                f'{format_eur(eur_value)}/saat | '
                f'{format_tl(tl_value)}/saat'
            )

            with st.expander(title, expanded=False):
                edit_col1, edit_col2 = st.columns(2)

                with edit_col1:
                    edit_labor_name = st.text_input(
                        "İşçilik adı",
                        value=item["ad"],
                        key=f'edit_labor_name_{item["id"]}',
                    )
                    edit_labor_description = st.text_input(
                        "Açıklama",
                        value=item.get("aciklama", "") or "",
                        key=(
                            f'edit_labor_description_'
                            f'{item["id"]}'
                        ),
                    )

                with edit_col2:
                    edit_labor_currency = st.selectbox(
                        "Saatlik ücret para birimi",
                        ["EUR", "TL"],
                        index=["EUR", "TL"].index(currency),
                        key=(
                            f'edit_labor_currency_'
                            f'{item["id"]}'
                        ),
                    )
                    edit_labor_price_text = st.text_input(
                        (
                            f"Saatlik ücret "
                            f"({edit_labor_currency})"
                        ),
                        value=format_number(
                            source_value,
                            4,
                        ),
                        key=(
                            f'edit_labor_price_'
                            f'{item["id"]}'
                        ),
                    )

                    edit_labor_price = parse_decimal(
                        edit_labor_price_text,
                        0,
                    )
                    edit_labor_eur, edit_labor_tl = (
                        convert_price(
                            edit_labor_price,
                            edit_labor_currency,
                            exchange_rate,
                        )
                    )
                    st.info(
                        f"Saatlik karşılık: "
                        f"{format_eur(edit_labor_eur)} | "
                        f"{format_tl(edit_labor_tl)}"
                    )

                update_col, delete_col = st.columns(2)

                with update_col:
                    if st.button(
                        "Güncelle",
                        type="primary",
                        use_container_width=True,
                        key=f'update_labor_{item["id"]}',
                    ):
                        edit_labor_price = parse_decimal(
                            edit_labor_price_text
                        )

                        if not edit_labor_name.strip():
                            st.error(
                                "İşçilik adı boş bırakılamaz."
                            )
                        elif (
                            edit_labor_price is None
                            or edit_labor_price < 0
                        ):
                            st.error(
                                "Geçerli bir saatlik ücret gir."
                            )
                        else:
                            db.table(
                                "iscilik_tanimlari"
                            ).update(
                                {
                                    "ad": (
                                        edit_labor_name.strip()
                                    ),
                                    "aciklama": (
                                        edit_labor_description.strip()
                                    ),
                                    "kaynak_para_birimi": (
                                        edit_labor_currency
                                    ),
                                    "kaynak_saatlik_ucret": (
                                        edit_labor_price
                                    ),
                                }
                            ).eq(
                                "id",
                                item["id"],
                            ).execute()

                            st.success(
                                "İşçilik güncellendi."
                            )
                            st.rerun()

                with delete_col:
                    if st.button(
                        "Sil",
                        use_container_width=True,
                        key=f'delete_labor_{item["id"]}',
                    ):
                        try:
                            db.table(
                                "iscilik_tanimlari"
                            ).delete().eq(
                                "id",
                                item["id"],
                            ).execute()

                            st.success("İşçilik silindi.")
                            st.rerun()
                        except Exception:
                            st.error(
                                "Bu işçilik kayıtlı bir parçada "
                                "kullanıldığı için silinemedi."
                            )


# -----------------------------
# KAYITLI PARÇALAR
# -----------------------------
with tab_saved:
    st.subheader("Kayıtlı parçalar")

    parts = get_parts()
    prices = get_prices()
    labors = get_labors()

    price_map = {item["id"]: item for item in prices}
    labor_map = {item["id"]: item for item in labors}

    if not parts:
        st.info("Henüz kayıtlı parça bulunmuyor.")
    else:
        part_options = {
            f'{item["parca_adi"]} — '
            f'{item["adet"]} adet': item
            for item in parts
        }

        selected_part_label = st.selectbox(
            "Parça seç",
            list(part_options.keys()),
        )
        selected_part = part_options[selected_part_label]
        part_id = selected_part["id"]
        production_qty = int(selected_part["adet"])

        selected_items = get_part_items(part_id)
        selected_labors = get_part_labors(part_id)

        item_rows = []
        labor_rows = []
        single_eur = 0.0
        single_tl = 0.0

        for row in selected_items:
            definition = price_map.get(
                row["fiyat_tanimi_id"],
                {},
            )
            currency = row.get(
                "kaynak_para_birimi"
            ) or "EUR"
            source_value = row.get(
                "kaynak_birim_fiyat"
            )

            if source_value is None:
                source_value = row.get(
                    "birim_fiyat_eur",
                    0,
                )

            unit_eur, unit_tl = convert_price(
                float(source_value or 0),
                currency,
                exchange_rate,
            )
            quantity = int(row.get("miktar") or 0)
            row_eur = unit_eur * quantity
            row_tl = unit_tl * quantity

            single_eur += row_eur
            single_tl += row_tl

            item_rows.append(
                {
                    "Ad": definition.get("ad", ""),
                    "Kategori": definition.get(
                        "kategori",
                        "",
                    ),
                    "Miktar": quantity,
                    "Tek Parça EUR": format_eur(row_eur),
                    "Tek Parça TL": format_tl(row_tl),
                }
            )

        for row in selected_labors:
            definition = labor_map.get(
                row["iscilik_tanimi_id"],
                {},
            )
            currency = row.get(
                "kaynak_para_birimi"
            ) or "EUR"
            source_value = float(
                row.get("kaynak_saatlik_ucret") or 0
            )
            hourly_eur, hourly_tl = convert_price(
                source_value,
                currency,
                exchange_rate,
            )
            hours = float(row.get("saat") or 0)
            row_eur = hourly_eur * hours
            row_tl = hourly_tl * hours

            single_eur += row_eur
            single_tl += row_tl

            labor_rows.append(
                {
                    "İşçilik": definition.get("ad", ""),
                    "Saat": format_number(hours, 4),
                    "Tek Parça EUR": format_eur(row_eur),
                    "Tek Parça TL": format_tl(row_tl),
                }
            )

        metric1, metric2, metric3, metric4 = st.columns(4)
        metric1.metric(
            "Tek parça EUR",
            format_eur(single_eur),
        )
        metric2.metric(
            "Tek parça TL",
            format_tl(single_tl),
        )
        metric3.metric(
            "Genel EUR",
            format_eur(single_eur * production_qty),
        )
        metric4.metric(
            "Genel TL",
            format_tl(single_tl * production_qty),
        )

        if item_rows:
            st.markdown(
                "#### Malzeme, kaplama, ek işlem ve ölçüm kalemleri"
            )
            st.dataframe(
                pd.DataFrame(item_rows),
                use_container_width=True,
                hide_index=True,
            )

        if labor_rows:
            st.markdown("#### İşçilik kalemleri")
            st.dataframe(
                pd.DataFrame(labor_rows),
                use_container_width=True,
                hide_index=True,
            )

        if st.button(
            "Bu parçayı sil",
            key="delete_saved_part",
        ):
            db.table("parcalar").delete().eq(
                "id",
                part_id,
            ).execute()
            st.success("Parça silindi.")
            st.rerun()


# -----------------------------
# PARÇA MALİYETİ
# -----------------------------
with tab_cost:
    st.subheader("Yeni parça maliyeti")

    prices = get_prices()
    labors = get_labors()

    if "cost_item_count" not in st.session_state:
        st.session_state.cost_item_count = 1

    if "labor_item_count" not in st.session_state:
        st.session_state.labor_item_count = 1

    top_left, top_right = st.columns([3, 1])

    with top_left:
        part_name = st.text_input(
            "Parça adı",
            key="new_part_name",
        )

    with top_right:
        production_qty = st.number_input(
            "Üretilecek adet",
            min_value=1,
            value=1,
            step=1,
            key="production_qty",
        )

    st.markdown("### Maliyet Kalemleri")

    selected_cost_items = []

    if not prices:
        st.warning(
            "Önce Fiyat Tanımları bölümünden kayıt oluştur."
        )
    else:
        price_labels = []
        price_lookup = {}

        for item in prices:
            currency, source_value = get_price_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            label = (
                f'{item["kategori"]} — {item["ad"]} — '
                f'{format_eur(eur_value)} / '
                f'{format_tl(tl_value)}'
            )
            price_labels.append(label)
            price_lookup[label] = item

        for index in range(
            st.session_state.cost_item_count
        ):
            col1, col2 = st.columns([4, 1])

            with col1:
                selected_label = st.selectbox(
                    f"Kalem {index + 1}",
                    price_labels,
                    key=f"part_cost_item_{index}",
                    label_visibility="collapsed",
                )

            with col2:
                quantity = st.number_input(
                    f"Miktar {index + 1}",
                    min_value=1,
                    value=1,
                    step=1,
                    key=f"part_cost_quantity_{index}",
                    label_visibility="collapsed",
                )

            definition = price_lookup[selected_label]
            currency, source_value = get_price_source(
                definition
            )

            selected_cost_items.append(
                {
                    "definition": definition,
                    "quantity": int(quantity),
                    "currency": currency,
                    "source_value": source_value,
                }
            )

        add_col, remove_col = st.columns(2)

        with add_col:
            if st.button(
                "＋ Kalem ekle",
                key="add_cost_item",
                use_container_width=True,
            ):
                st.session_state.cost_item_count += 1
                st.rerun()

        with remove_col:
            if st.button(
                "− Son kalemi kaldır",
                key="remove_cost_item",
                use_container_width=True,
                disabled=(
                    st.session_state.cost_item_count <= 1
                ),
            ):
                st.session_state.cost_item_count -= 1
                st.rerun()

    st.markdown("### İşçilik Kalemleri")

    selected_labor_items = []

    if not labors:
        st.info(
            "İşçilik Maliyetleri bölümünden işçilik "
            "tanımı eklediğinde burada seçebilirsin."
        )
    else:
        labor_labels = []
        labor_lookup = {}

        for item in labors:
            currency, source_value = get_labor_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            label = (
                f'{item["ad"]} — '
                f'{format_eur(eur_value)}/saat — '
                f'{format_tl(tl_value)}/saat'
            )
            labor_labels.append(label)
            labor_lookup[label] = item

        for index in range(
            st.session_state.labor_item_count
        ):
            hours_key = f"labor_hours_{index}"
            minutes_key = f"labor_minutes_{index}"

            if hours_key not in st.session_state:
                st.session_state[hours_key] = 1.0

            if minutes_key not in st.session_state:
                st.session_state[minutes_key] = 30

            col1, col2, col3 = st.columns(
                [4, 1.2, 1.3]
            )

            with col1:
                selected_labor_label = st.selectbox(
                    f"İşçilik {index + 1}",
                    labor_labels,
                    key=f"part_labor_item_{index}",
                    label_visibility="collapsed",
                )

            with col2:
                hours = st.number_input(
                    f"Saat {index + 1}",
                    min_value=0.0,
                    step=0.25,
                    format="%.4f",
                    key=hours_key,
                    label_visibility="collapsed",
                )

            with col3:
                with st.popover("Dakika → Saat"):
                    minutes = st.number_input(
                        "Dakika",
                        min_value=0,
                        step=1,
                        key=minutes_key,
                    )
                    st.write(
                        f"Karşılığı: "
                        f"{format_number(minutes / 60, 4)} saat"
                    )
                    st.button(
                        "Saat alanına uygula",
                        key=f"apply_minutes_{index}",
                        on_click=set_hours_from_minutes,
                        args=(hours_key, minutes_key),
                    )

            definition = labor_lookup[
                selected_labor_label
            ]
            currency, source_value = get_labor_source(
                definition
            )

            selected_labor_items.append(
                {
                    "definition": definition,
                    "hours": float(hours),
                    "currency": currency,
                    "source_value": source_value,
                }
            )

        add_col, remove_col = st.columns(2)

        with add_col:
            if st.button(
                "＋ İşçilik ekle",
                key="add_labor_item",
                use_container_width=True,
            ):
                st.session_state.labor_item_count += 1
                st.rerun()

        with remove_col:
            if st.button(
                "− Son işçiliği kaldır",
                key="remove_labor_item",
                use_container_width=True,
                disabled=(
                    st.session_state.labor_item_count <= 1
                ),
            ):
                st.session_state.labor_item_count -= 1
                st.rerun()

    single_eur = 0.0
    single_tl = 0.0

    for item in selected_cost_items:
        unit_eur, unit_tl = convert_price(
            item["source_value"],
            item["currency"],
            exchange_rate,
        )
        single_eur += (
            unit_eur * item["quantity"]
        )
        single_tl += (
            unit_tl * item["quantity"]
        )

    for item in selected_labor_items:
        hourly_eur, hourly_tl = convert_price(
            item["source_value"],
            item["currency"],
            exchange_rate,
        )
        single_eur += (
            hourly_eur * item["hours"]
        )
        single_tl += (
            hourly_tl * item["hours"]
        )

    st.divider()

    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric(
        "Tek parça EUR",
        format_eur(single_eur),
    )
    metric2.metric(
        "Tek parça TL",
        format_tl(single_tl),
    )
    metric3.metric(
        "Genel toplam EUR",
        format_eur(
            single_eur * int(production_qty)
        ),
    )
    metric4.metric(
        "Genel toplam TL",
        format_tl(
            single_tl * int(production_qty)
        ),
    )

    if st.button(
        "Parçayı kaydet",
        type="primary",
        key="save_part",
        use_container_width=True,
    ):
        if not part_name.strip():
            st.error("Parça adı boş bırakılamaz.")
        elif not selected_cost_items and not selected_labor_items:
            st.error(
                "En az bir maliyet veya işçilik kalemi ekle."
            )
        else:
            part_result = (
                db.table("parcalar")
                .insert(
                    {
                        "parca_adi": part_name.strip(),
                        "adet": int(production_qty),
                    }
                )
                .execute()
            )

            part_id = part_result.data[0]["id"]

            if selected_cost_items:
                rows = []

                for item in selected_cost_items:
                    eur_snapshot, _ = convert_price(
                        item["source_value"],
                        item["currency"],
                        exchange_rate,
                    )

                    rows.append(
                        {
                            "parca_id": part_id,
                            "fiyat_tanimi_id": (
                                item["definition"]["id"]
                            ),
                            "miktar": item["quantity"],
                            "kaynak_para_birimi": (
                                item["currency"]
                            ),
                            "kaynak_birim_fiyat": (
                                item["source_value"]
                            ),
                            "birim_fiyat_eur": (
                                eur_snapshot
                            ),
                        }
                    )

                db.table("parca_kalemleri").insert(
                    rows
                ).execute()

            if selected_labor_items:
                rows = []

                for item in selected_labor_items:
                    rows.append(
                        {
                            "parca_id": part_id,
                            "iscilik_tanimi_id": (
                                item["definition"]["id"]
                            ),
                            "saat": item["hours"],
                            "kaynak_para_birimi": (
                                item["currency"]
                            ),
                            "kaynak_saatlik_ucret": (
                                item["source_value"]
                            ),
                        }
                    )

                db.table(
                    "parca_iscilik_kalemleri"
                ).insert(rows).execute()

            st.session_state.cost_item_count = 1
            st.session_state.labor_item_count = 1
            st.success("Parça maliyeti kaydedildi.")
            st.rerun()


# -----------------------------
# ALTTA XLS İNDİR
# -----------------------------
st.divider()

download_col, _ = st.columns([1, 3])

with download_col:
    st.download_button(
        "Dosyayı indir (.xls)",
        data=build_xls(exchange_rate),
        file_name=(
            f"ITSSystems_Cost_Calculator_"
            f"{datetime.now():%Y-%m-%d}.xls"
        ),
        mime="application/vnd.ms-excel",
        type="primary",
        use_container_width=True,
    )
