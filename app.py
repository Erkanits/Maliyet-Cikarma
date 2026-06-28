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


# =========================================================
# GİRİŞ VE VERİTABANI
# =========================================================
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


@st.cache_resource
def get_db():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_ROLE_KEY"],
    )


db = get_db()


# =========================================================
# GENEL YARDIMCI FONKSİYONLAR
# =========================================================
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

    defaults = {
        "id": 1,
        "eur_tl_kuru": 50,
        "logo_base64": None,
    }
    db.table("ayarlar").insert(defaults).execute()
    return defaults


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


# =========================================================
# BAŞLIK, LOGO VE KUR
# =========================================================
def render_header(settings):
    logo_col, title_col = st.columns([1, 4], vertical_alignment="center")

    with logo_col:
        logo_value = settings.get("logo_base64")

        if logo_value:
            try:
                logo_bytes = base64.b64decode(logo_value)
                st.image(io.BytesIO(logo_bytes), width=210)
            except Exception:
                st.warning("Logo görüntülenemedi.")

        popover_title = "Logo ayarları" if logo_value else "Logo yükle"

        with st.popover(popover_title, use_container_width=True):
            st.caption(
                "Önerilen logo: şeffaf PNG, yaklaşık 800 × 250 px, "
                "3:1 oranında ve 2 MB'tan küçük."
            )

            logo_file = st.file_uploader(
                "Yeni PNG veya JPEG seç",
                type=["png", "jpg", "jpeg"],
                key="logo_file",
            )

            if st.button(
                "Logoyu kaydet / değiştir",
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

    with title_col:
        st.title("ITSSystems Cost Calculator")
        st.caption(
            "Malzeme, kaplama, ölçüm, ek işlem ve işçilik maliyet hesaplama uygulaması"
        )


def render_rate(settings):
    current_rate = float(settings.get("eur_tl_kuru") or 50)

    with st.expander("EUR / TL Kur Ayarı", expanded=False):
        with st.form("exchange_rate_form"):
            col1, col2, col3 = st.columns([1.3, 1, 3])

            with col1:
                rate_text = st.text_input(
                    "1 EUR kaç TL?",
                    value=format_number(current_rate, 4),
                )

            with col2:
                st.write("")
                st.write("")
                save_rate = st.form_submit_button(
                    "Kuru kaydet",
                    type="primary",
                    use_container_width=True,
                )

            with col3:
                st.info(
                    "TL girilen fiyatın TL değeri sabit kalır; EUR karşılığı kura göre değişir. "
                    "EUR girilen fiyatın EUR değeri sabit kalır; TL karşılığı kura göre değişir."
                )

            if save_rate:
                new_rate = parse_decimal(rate_text)

                if new_rate is None or new_rate <= 0:
                    st.error("Geçerli ve sıfırdan büyük bir kur gir.")
                else:
                    update_settings(rate=new_rate)
                    st.success("Kur güncellendi.")
                    st.rerun()


# =========================================================
# XLS DIŞA AKTARMA
# =========================================================
def build_export_frames(rate):
    prices = get_prices()
    labors = get_labors()
    parts = get_parts()
    part_items = get_part_items()
    part_labors = get_part_labors()

    price_map = {item["id"]: item for item in prices}
    labor_map = {item["id"]: item for item in labors}

    summary_rows = []
    detail_rows = []

    for part in parts:
        part_id = part["id"]
        quantity = int(part["adet"])
        single_eur = 0.0
        single_tl = 0.0
        operation_names = []

        for row in [x for x in part_items if x["parca_id"] == part_id]:
            definition = price_map.get(row["fiyat_tanimi_id"], {})
            currency = row.get("kaynak_para_birimi") or "EUR"
            source_value = row.get("kaynak_birim_fiyat")

            if source_value is None:
                source_value = row.get("birim_fiyat_eur", 0)

            unit_eur, unit_tl = convert_price(
                float(source_value or 0),
                currency,
                rate,
            )
            item_quantity = int(row.get("miktar") or 0)
            line_eur = unit_eur * item_quantity
            line_tl = unit_tl * item_quantity

            single_eur += line_eur
            single_tl += line_tl

            operation_names.append(
                f'{definition.get("ad", "")} '
                f'({format_eur(line_eur)} / {format_tl(line_tl)})'
            )

            detail_rows.append(
                {
                    "Parça Adı": part["parca_adi"],
                    "Kategori": definition.get("kategori", ""),
                    "İşlem / Kalem": definition.get("ad", ""),
                    "Miktar / Tek Parça": item_quantity,
                    "Tek Parça Kalem EUR": line_eur,
                    "Tek Parça Kalem TL": line_tl,
                }
            )

        for row in [x for x in part_labors if x["parca_id"] == part_id]:
            definition = labor_map.get(row["iscilik_tanimi_id"], {})
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
            line_eur = hourly_eur * hours
            line_tl = hourly_tl * hours

            single_eur += line_eur
            single_tl += line_tl

            operation_names.append(
                f'{definition.get("ad", "")} '
                f'({format_eur(line_eur)} / {format_tl(line_tl)})'
            )

            detail_rows.append(
                {
                    "Parça Adı": part["parca_adi"],
                    "Kategori": "İşçilik",
                    "İşlem / Kalem": definition.get("ad", ""),
                    "Miktar / Tek Parça": hours,
                    "Tek Parça Kalem EUR": line_eur,
                    "Tek Parça Kalem TL": line_tl,
                }
            )

        summary_rows.append(
            {
                "Parça Adı": part["parca_adi"],
                "Adet": quantity,
                "İşlemler": " + ".join(operation_names),
                "Birim Fiyat EUR": single_eur,
                "Birim Fiyat TL": single_tl,
                "Toplam Fiyat EUR": single_eur * quantity,
                "Toplam Fiyat TL": single_tl * quantity,
                "Kullanılan EUR/TL Kuru": rate,
            }
        )

    return {
        "Liste": pd.DataFrame(summary_rows),
        "Detaylar": pd.DataFrame(detail_rows),
    }


def build_xls(rate):
    frames = build_export_frames(rate)
    workbook = xlwt.Workbook(encoding="utf-8")

    header_style = xlwt.easyxf(
        "font: bold on; pattern: pattern solid, fore_colour gray25;"
    )
    text_style = xlwt.easyxf()
    number_style = xlwt.easyxf(num_format_str="#,##0.0000")
    money_style = xlwt.easyxf(num_format_str="#,##0.00")

    for sheet_name, dataframe in frames.items():
        sheet = workbook.add_sheet(sheet_name[:31])

        if dataframe.empty:
            sheet.write(0, 0, "Kayıt bulunmuyor.", text_style)
            continue

        for column_index, column_name in enumerate(dataframe.columns):
            sheet.write(
                0,
                column_index,
                column_name,
                header_style,
            )
            width = min(max(len(column_name) + 4, 15), 55)
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


# =========================================================
# UYGULAMA VERİLERİ
# =========================================================
settings = get_settings()
exchange_rate = float(settings.get("eur_tl_kuru") or 50)

prices = get_prices()
labors = get_labors()
parts = get_parts()

price_map = {item["id"]: item for item in prices}
labor_map = {item["id"]: item for item in labors}

render_header(settings)
render_rate(settings)

tab_prices, tab_labors, tab_list, tab_cost = st.tabs(
    [
        "Fiyat Tanımları",
        "İşçilik Maliyetleri",
        "Liste",
        "Parça Maliyeti",
    ]
)


# =========================================================
# FİYAT TANIMLARI
# =========================================================
with tab_prices:
    st.subheader("Yeni fiyat tanımı")

    with st.form("new_price_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            new_category = st.selectbox(
                "Kategori",
                ["Malzeme", "Kaplama", "Ek İşlem", "Ölçüm"],
            )
            new_name = st.text_input("Ad")
            new_description = st.text_input("Açıklama")

        with col2:
            new_currency = st.selectbox(
                "Fiyat para birimi",
                ["EUR", "TL"],
            )
            new_price_text = st.text_input(
                f"Birim fiyat ({new_currency})",
                value="0,00",
            )

        save_new_price = st.form_submit_button(
            "Fiyat tanımını kaydet",
            type="primary",
        )

        if save_new_price:
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
    st.subheader("Düzenle veya sil")

    selected_filter = st.selectbox(
        "Kategoriye göre filtrele",
        ["Tümü", "Malzeme", "Kaplama", "Ek İşlem", "Ölçüm"],
        key="price_filter",
    )

    filtered_prices = prices
    if selected_filter != "Tümü":
        filtered_prices = [
            item
            for item in prices
            if item["kategori"] == selected_filter
        ]

    if not filtered_prices:
        st.info("Bu filtreye uygun kayıt bulunmuyor.")
    else:
        for item in filtered_prices:
            currency, source_value = get_price_source(item)
            eur_value, tl_value = convert_price(
                source_value,
                currency,
                exchange_rate,
            )

            title = (
                f'{item["ad"]} | {item["kategori"]} | '
                f'{format_eur(eur_value)} / {format_tl(tl_value)}'
            )

            with st.expander(title, expanded=False):
                with st.form(f'price_edit_form_{item["id"]}'):
                    col1, col2 = st.columns(2)

                    with col1:
                        categories = [
                            "Malzeme",
                            "Kaplama",
                            "Ek İşlem",
                            "Ölçüm",
                        ]
                        edit_category = st.selectbox(
                            "Kategori",
                            categories,
                            index=categories.index(item["kategori"]),
                        )
                        edit_name = st.text_input(
                            "Ad",
                            value=item["ad"],
                        )
                        edit_description = st.text_input(
                            "Açıklama",
                            value=item.get("aciklama", "") or "",
                        )

                    with col2:
                        edit_currency = st.selectbox(
                            "Fiyat para birimi",
                            ["EUR", "TL"],
                            index=["EUR", "TL"].index(currency),
                        )
                        edit_price_text = st.text_input(
                            f"Birim fiyat ({edit_currency})",
                            value=format_number(source_value, 4),
                        )

                    update_col, delete_col = st.columns(2)

                    with update_col:
                        update_price = st.form_submit_button(
                            "Güncelle",
                            type="primary",
                            use_container_width=True,
                        )

                    with delete_col:
                        delete_price = st.form_submit_button(
                            "Sil",
                            use_container_width=True,
                        )

                    if update_price:
                        edit_price = parse_decimal(edit_price_text)

                        if not edit_name.strip():
                            st.error("Ad alanı boş bırakılamaz.")
                        elif edit_price is None or edit_price < 0:
                            st.error("Geçerli bir fiyat gir.")
                        else:
                            eur_snapshot, _ = convert_price(
                                edit_price,
                                edit_currency,
                                exchange_rate,
                            )

                            db.table("fiyat_tanimlari").update(
                                {
                                    "kategori": edit_category,
                                    "ad": edit_name.strip(),
                                    "aciklama": edit_description.strip(),
                                    "kaynak_para_birimi": edit_currency,
                                    "kaynak_birim_fiyat": edit_price,
                                    "birim_fiyat_eur": eur_snapshot,
                                }
                            ).eq("id", item["id"]).execute()

                            st.success("Kayıt güncellendi.")
                            st.rerun()

                    if delete_price:
                        try:
                            db.table("fiyat_tanimlari").delete().eq(
                                "id",
                                item["id"],
                            ).execute()
                            st.success("Kayıt silindi.")
                            st.rerun()
                        except Exception:
                            st.error(
                                "Bu kayıt bir parçada kullanıldığı için silinemedi."
                            )


# =========================================================
# İŞÇİLİK MALİYETLERİ
# =========================================================
with tab_labors:
    st.subheader("Yeni işçilik maliyeti")

    with st.form("new_labor_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            labor_name = st.text_input(
                "İşçilik adı",
                placeholder="Örn. Dik İşlem CNC, Torna",
            )
            labor_description = st.text_input("Açıklama")

        with col2:
            labor_currency = st.selectbox(
                "Saatlik ücret para birimi",
                ["EUR", "TL"],
            )
            labor_price_text = st.text_input(
                f"Saatlik ücret ({labor_currency})",
                value="0,00",
            )

        save_labor = st.form_submit_button(
            "İşçilik tanımını kaydet",
            type="primary",
        )

        if save_labor:
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
    st.subheader("Düzenle veya sil")

    if not labors:
        st.info("Henüz işçilik tanımı bulunmuyor.")
    else:
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
                with st.form(f'labor_edit_form_{item["id"]}'):
                    col1, col2 = st.columns(2)

                    with col1:
                        edit_labor_name = st.text_input(
                            "İşçilik adı",
                            value=item["ad"],
                        )
                        edit_labor_description = st.text_input(
                            "Açıklama",
                            value=item.get("aciklama", "") or "",
                        )

                    with col2:
                        edit_labor_currency = st.selectbox(
                            "Saatlik ücret para birimi",
                            ["EUR", "TL"],
                            index=["EUR", "TL"].index(currency),
                        )
                        edit_labor_price_text = st.text_input(
                            f"Saatlik ücret ({edit_labor_currency})",
                            value=format_number(source_value, 4),
                        )

                    update_col, delete_col = st.columns(2)

                    with update_col:
                        update_labor = st.form_submit_button(
                            "Güncelle",
                            type="primary",
                            use_container_width=True,
                        )

                    with delete_col:
                        delete_labor = st.form_submit_button(
                            "Sil",
                            use_container_width=True,
                        )

                    if update_labor:
                        edit_labor_price = parse_decimal(
                            edit_labor_price_text
                        )

                        if not edit_labor_name.strip():
                            st.error("İşçilik adı boş bırakılamaz.")
                        elif (
                            edit_labor_price is None
                            or edit_labor_price < 0
                        ):
                            st.error("Geçerli bir saatlik ücret gir.")
                        else:
                            db.table("iscilik_tanimlari").update(
                                {
                                    "ad": edit_labor_name.strip(),
                                    "aciklama": edit_labor_description.strip(),
                                    "kaynak_para_birimi": edit_labor_currency,
                                    "kaynak_saatlik_ucret": edit_labor_price,
                                }
                            ).eq("id", item["id"]).execute()

                            st.success("İşçilik güncellendi.")
                            st.rerun()

                    if delete_labor:
                        try:
                            db.table("iscilik_tanimlari").delete().eq(
                                "id",
                                item["id"],
                            ).execute()
                            st.success("İşçilik silindi.")
                            st.rerun()
                        except Exception:
                            st.error(
                                "Bu işçilik bir parçada kullanıldığı için silinemedi."
                            )


# =========================================================
# LİSTE
# =========================================================
with tab_list:
    st.subheader("Parça Listesi")

    all_part_items = get_part_items()
    all_part_labors = get_part_labors()

    if not parts:
        st.info("Henüz kayıtlı parça bulunmuyor.")
    else:
        list_rows = []

        for part in parts:
            part_id = part["id"]
            quantity = int(part["adet"])
            single_eur = 0.0
            single_tl = 0.0
            operations = []

            for row in [
                x for x in all_part_items
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
                    exchange_rate,
                )
                item_quantity = int(row.get("miktar") or 0)
                line_eur = unit_eur * item_quantity
                line_tl = unit_tl * item_quantity

                single_eur += line_eur
                single_tl += line_tl

                operations.append(
                    f'{definition.get("ad", "")} '
                    f'({format_eur(line_eur)} / {format_tl(line_tl)})'
                )

            for row in [
                x for x in all_part_labors
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
                    exchange_rate,
                )
                hours = float(row.get("saat") or 0)
                line_eur = hourly_eur * hours
                line_tl = hourly_tl * hours

                single_eur += line_eur
                single_tl += line_tl

                operations.append(
                    f'{definition.get("ad", "")} '
                    f'({format_eur(line_eur)} / {format_tl(line_tl)})'
                )

            list_rows.append(
                {
                    "Parça Adı": part["parca_adi"],
                    "Adet": quantity,
                    "İşlemler": " + ".join(operations),
                    "Birim Fiyat EUR": format_eur(single_eur),
                    "Birim Fiyat TL": format_tl(single_tl),
                    "Toplam Fiyat EUR": format_eur(
                        single_eur * quantity
                    ),
                    "Toplam Fiyat TL": format_tl(
                        single_tl * quantity
                    ),
                }
            )

        st.dataframe(
            pd.DataFrame(list_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            "Bu bölüm yalnızca listeleme içindir. Parça değişikliklerini "
            "Parça Maliyeti sekmesindeki Güncelle seçeneğinden yapabilirsin."
        )

        st.download_button(
            "Dosyayı indir (.xls)",
            data=build_xls(exchange_rate),
            file_name=(
                f"ITSSystems_Cost_Calculator_"
                f"{datetime.now():%Y-%m-%d}.xls"
            ),
            mime="application/vnd.ms-excel",
            type="primary",
        )


# =========================================================
# PARÇA MALİYETİ: YENİ KAYIT VE GÜNCELLEME
# =========================================================
with tab_cost:
    st.subheader("Parça Maliyeti")

    mode = st.radio(
        "İşlem",
        ["Yeni Parça", "Mevcut Parçayı Güncelle"],
        horizontal=True,
    )

    selected_part = None
    current_items = []
    current_labors = []

    if mode == "Mevcut Parçayı Güncelle":
        if not parts:
            st.warning("Güncellenecek kayıtlı parça bulunmuyor.")
            st.stop()

        part_options = {
            f'{part["parca_adi"]} — {part["adet"]} adet': part
            for part in parts
        }

        selected_part_label = st.selectbox(
            "Güncellenecek parça",
            list(part_options.keys()),
        )
        selected_part = part_options[selected_part_label]
        current_items = get_part_items(selected_part["id"])
        current_labors = get_part_labors(selected_part["id"])

    default_cost_count = (
        len(current_items)
        if selected_part is not None
        else 1
    )
    default_labor_count = (
        len(current_labors)
        if selected_part is not None
        else 0
    )

    count_col1, count_col2 = st.columns(2)

    with count_col1:
        cost_count = st.number_input(
            "Maliyet kalemi sayısı",
            min_value=0,
            max_value=20,
            value=default_cost_count,
            step=1,
            key=(
                f'cost_count_{selected_part["id"]}'
                if selected_part
                else "new_cost_count"
            ),
        )

    with count_col2:
        labor_count = st.number_input(
            "İşçilik kalemi sayısı",
            min_value=0,
            max_value=20,
            value=default_labor_count,
            step=1,
            key=(
                f'labor_count_{selected_part["id"]}'
                if selected_part
                else "new_labor_count"
            ),
        )

    if not prices and cost_count > 0:
        st.warning(
            "Maliyet kalemi eklemek için önce Fiyat Tanımları bölümünden kayıt oluştur."
        )

    if not labors and labor_count > 0:
        st.warning(
            "İşçilik eklemek için önce İşçilik Maliyetleri bölümünden kayıt oluştur."
        )

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
            f'{format_eur(eur_value)} / {format_tl(tl_value)}'
        )
        price_labels.append(label)
        price_lookup[label] = item

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

    with st.form(
        (
            f'part_form_{selected_part["id"]}'
            if selected_part
            else "new_part_form"
        )
    ):
        top_col1, top_col2 = st.columns([3, 1])

        with top_col1:
            part_name = st.text_input(
                "Parça adı",
                value=(
                    selected_part["parca_adi"]
                    if selected_part
                    else ""
                ),
            )

        with top_col2:
            production_quantity = st.number_input(
                "Üretilecek adet",
                min_value=1,
                value=(
                    int(selected_part["adet"])
                    if selected_part
                    else 1
                ),
                step=1,
            )

        selected_cost_rows = []

        if cost_count > 0 and price_labels:
            st.markdown("### Maliyet Kalemleri")

            for index in range(int(cost_count)):
                current_row = (
                    current_items[index]
                    if index < len(current_items)
                    else None
                )

                default_price_index = 0
                default_quantity = 1

                if current_row:
                    default_quantity = int(
                        current_row.get("miktar") or 1
                    )
                    current_definition_id = (
                        current_row["fiyat_tanimi_id"]
                    )

                    for label_index, label in enumerate(price_labels):
                        if (
                            price_lookup[label]["id"]
                            == current_definition_id
                        ):
                            default_price_index = label_index
                            break

                row_col1, row_col2 = st.columns([4, 1])

                with row_col1:
                    selected_label = st.selectbox(
                        f"Kalem {index + 1}",
                        price_labels,
                        index=default_price_index,
                        key=(
                            f'part_price_{selected_part["id"]}_{index}'
                            if selected_part
                            else f'new_part_price_{index}'
                        ),
                    )

                with row_col2:
                    item_quantity = st.number_input(
                        f"Adet {index + 1}",
                        min_value=1,
                        value=default_quantity,
                        step=1,
                        key=(
                            f'part_qty_{selected_part["id"]}_{index}'
                            if selected_part
                            else f'new_part_qty_{index}'
                        ),
                    )

                definition = price_lookup[selected_label]
                currency, source_value = get_price_source(definition)

                selected_cost_rows.append(
                    {
                        "definition": definition,
                        "quantity": int(item_quantity),
                        "currency": currency,
                        "source_value": source_value,
                    }
                )

        selected_labor_rows = []

        if labor_count > 0 and labor_labels:
            st.markdown("### İşçilik Kalemleri")

            for index in range(int(labor_count)):
                current_row = (
                    current_labors[index]
                    if index < len(current_labors)
                    else None
                )

                default_labor_index = 0
                default_duration_type = "Saat"
                default_duration_value = 1.0

                if current_row:
                    current_definition_id = (
                        current_row["iscilik_tanimi_id"]
                    )
                    default_duration_value = float(
                        current_row.get("saat") or 0
                    )

                    for label_index, label in enumerate(labor_labels):
                        if (
                            labor_lookup[label]["id"]
                            == current_definition_id
                        ):
                            default_labor_index = label_index
                            break

                row_col1, row_col2, row_col3 = st.columns(
                    [4, 1.2, 1.4]
                )

                with row_col1:
                    selected_labor_label = st.selectbox(
                        f"İşçilik {index + 1}",
                        labor_labels,
                        index=default_labor_index,
                        key=(
                            f'part_labor_{selected_part["id"]}_{index}'
                            if selected_part
                            else f'new_part_labor_{index}'
                        ),
                    )

                with row_col2:
                    duration_type = st.selectbox(
                        f"Süre birimi {index + 1}",
                        ["Saat", "Dakika"],
                        index=(
                            0
                            if default_duration_type == "Saat"
                            else 1
                        ),
                        key=(
                            f'part_duration_type_'
                            f'{selected_part["id"]}_{index}'
                            if selected_part
                            else f'new_duration_type_{index}'
                        ),
                    )

                with row_col3:
                    duration_value = st.number_input(
                        f"Süre {index + 1}",
                        min_value=0.0,
                        value=default_duration_value,
                        step=1.0 if duration_type == "Dakika" else 0.25,
                        format="%.4f",
                        key=(
                            f'part_duration_'
                            f'{selected_part["id"]}_{index}'
                            if selected_part
                            else f'new_duration_{index}'
                        ),
                    )

                hours = (
                    float(duration_value) / 60
                    if duration_type == "Dakika"
                    else float(duration_value)
                )

                definition = labor_lookup[selected_labor_label]
                currency, source_value = get_labor_source(definition)

                selected_labor_rows.append(
                    {
                        "definition": definition,
                        "hours": hours,
                        "currency": currency,
                        "source_value": source_value,
                    }
                )

        submit_label = (
            "Güncelle"
            if selected_part
            else "Parçayı kaydet"
        )

        submitted = st.form_submit_button(
            submit_label,
            type="primary",
            use_container_width=True,
        )

        if submitted:
            if not part_name.strip():
                st.error("Parça adı boş bırakılamaz.")
            elif not selected_cost_rows and not selected_labor_rows:
                st.error(
                    "En az bir maliyet veya işçilik kalemi ekle."
                )
            else:
                if selected_part:
                    part_id = selected_part["id"]

                    db.table("parcalar").update(
                        {
                            "parca_adi": part_name.strip(),
                            "adet": int(production_quantity),
                        }
                    ).eq("id", part_id).execute()

                    db.table("parca_kalemleri").delete().eq(
                        "parca_id",
                        part_id,
                    ).execute()

                    db.table("parca_iscilik_kalemleri").delete().eq(
                        "parca_id",
                        part_id,
                    ).execute()
                else:
                    result = db.table("parcalar").insert(
                        {
                            "parca_adi": part_name.strip(),
                            "adet": int(production_quantity),
                        }
                    ).execute()
                    part_id = result.data[0]["id"]

                if selected_cost_rows:
                    cost_insert_rows = []

                    for row in selected_cost_rows:
                        eur_snapshot, _ = convert_price(
                            row["source_value"],
                            row["currency"],
                            exchange_rate,
                        )

                        cost_insert_rows.append(
                            {
                                "parca_id": part_id,
                                "fiyat_tanimi_id": (
                                    row["definition"]["id"]
                                ),
                                "miktar": row["quantity"],
                                "kaynak_para_birimi": (
                                    row["currency"]
                                ),
                                "kaynak_birim_fiyat": (
                                    row["source_value"]
                                ),
                                "birim_fiyat_eur": eur_snapshot,
                            }
                        )

                    db.table("parca_kalemleri").insert(
                        cost_insert_rows
                    ).execute()

                if selected_labor_rows:
                    labor_insert_rows = []

                    for row in selected_labor_rows:
                        labor_insert_rows.append(
                            {
                                "parca_id": part_id,
                                "iscilik_tanimi_id": (
                                    row["definition"]["id"]
                                ),
                                "saat": row["hours"],
                                "kaynak_para_birimi": (
                                    row["currency"]
                                ),
                                "kaynak_saatlik_ucret": (
                                    row["source_value"]
                                ),
                            }
                        )

                    db.table("parca_iscilik_kalemleri").insert(
                        labor_insert_rows
                    ).execute()

                st.success(
                    "Parça güncellendi."
                    if selected_part
                    else "Parça kaydedildi."
                )
                st.rerun()
