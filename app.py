
import io
from datetime import datetime

import pandas as pd
import streamlit as st
from supabase import create_client


st.set_page_config(
    page_title="ITSSystems Cost Calculator",
    page_icon="€",
    layout="wide",
)


@st.cache_resource
def get_supabase():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_ROLE_KEY"],
    )


db = get_supabase()


def get_fiyatlar():
    response = (
        db.table("fiyat_tanimlari")
        .select("*")
        .order("kategori")
        .order("ad")
        .execute()
    )
    return response.data or []


def get_parcalar():
    response = (
        db.table("parcalar")
        .select("*")
        .order("created_at", desc=True)
        .execute()
    )
    return response.data or []


def get_kalemler(parca_id=None):
    query = db.table("parca_kalemleri").select("*")
    if parca_id is not None:
        query = query.eq("parca_id", parca_id)
    return query.execute().data or []


def euro(value):
    return f"{float(value):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def excel_dosyasi():
    fiyatlar = get_fiyatlar()
    parcalar = get_parcalar()
    kalemler = get_kalemler()

    fiyat_map = {item["id"]: item for item in fiyatlar}
    parca_map = {item["id"]: item for item in parcalar}

    fiyat_df = pd.DataFrame(
        [
            {
                "ID": item["id"],
                "Kategori": item["kategori"],
                "Ad": item["ad"],
                "Açıklama": item.get("aciklama", ""),
                "Birim Fiyat (€)": float(item["birim_fiyat_eur"]),
                "Oluşturulma Tarihi": item["created_at"],
            }
            for item in fiyatlar
        ]
    )

    parca_df = pd.DataFrame(
        [
            {
                "ID": item["id"],
                "Parça Adı": item["parca_adi"],
                "Adet": item["adet"],
                "Oluşturulma Tarihi": item["created_at"],
            }
            for item in parcalar
        ]
    )

    detay_rows = []
    for kalem in kalemler:
        fiyat = fiyat_map.get(kalem["fiyat_tanimi_id"], {})
        parca = parca_map.get(kalem["parca_id"], {})
        miktar = float(kalem["miktar"])
        birim_fiyat = float(kalem["birim_fiyat_eur"])
        adet = int(parca.get("adet", 0))
        detay_rows.append(
            {
                "Parça ID": kalem["parca_id"],
                "Parça Adı": parca.get("parca_adi", ""),
                "Üretim Adedi": adet,
                "Kategori": fiyat.get("kategori", ""),
                "Kalem": fiyat.get("ad", ""),
                "Açıklama": fiyat.get("aciklama", ""),
                "Miktar / Tek Parça": miktar,
                "Birim Fiyat (€)": birim_fiyat,
                "Tek Parça Kalem Tutarı (€)": miktar * birim_fiyat,
                "Toplam Kalem Tutarı (€)": miktar * birim_fiyat * adet,
            }
        )

    detay_df = pd.DataFrame(detay_rows)

    ozet_rows = []
    for parca in parcalar:
        ilgili = [x for x in kalemler if x["parca_id"] == parca["id"]]
        tek_parca = sum(
            float(x["miktar"]) * float(x["birim_fiyat_eur"]) for x in ilgili
        )
        ozet_rows.append(
            {
                "Parça ID": parca["id"],
                "Parça Adı": parca["parca_adi"],
                "Adet": parca["adet"],
                "Tek Parça Maliyeti (€)": tek_parca,
                "Genel Toplam (€)": tek_parca * int(parca["adet"]),
            }
        )
    ozet_df = pd.DataFrame(ozet_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        fiyat_df.to_excel(writer, index=False, sheet_name="Fiyat Tanımları")
        parca_df.to_excel(writer, index=False, sheet_name="Parçalar")
        detay_df.to_excel(writer, index=False, sheet_name="Maliyet Detayları")
        ozet_df.to_excel(writer, index=False, sheet_name="Maliyet Özeti")

        for sheet_name in writer.book.sheetnames:
            sheet = writer.book[sheet_name]
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                sheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 45
                )

    output.seek(0)
    return output.getvalue()


st.title("ITSSystems Cost Calculator")
st.caption("Malzeme, kaplama ve ek işlem maliyet hesaplama uygulaması")

tab1, tab2, tab3, tab4 = st.tabs(
    ["Fiyat Tanımları", "Parça Maliyeti", "Kayıtlı Parçalar", "Excel"]
)

with tab1:
    st.subheader("Yeni fiyat tanımı")

    with st.form("yeni_fiyat_formu", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            kategori = st.selectbox(
                "Kategori", ["Malzeme", "Kaplama", "Ek İşlem"]
            )
            ad = st.text_input("Ad")
        with col2:
            birim_fiyat = st.number_input(
                "Birim fiyat (€)", min_value=0.0, step=0.01, format="%.4f"
            )
            aciklama = st.text_input("Açıklama")

        kaydet = st.form_submit_button("Fiyat tanımını kaydet", type="primary")

        if kaydet:
            if not ad.strip():
                st.error("Ad alanı boş bırakılamaz.")
            else:
                db.table("fiyat_tanimlari").insert(
                    {
                        "kategori": kategori,
                        "ad": ad.strip(),
                        "aciklama": aciklama.strip(),
                        "birim_fiyat_eur": birim_fiyat,
                    }
                ).execute()
                st.success("Fiyat tanımı kaydedildi.")
                st.rerun()

    st.divider()
    st.subheader("Kayıtlı fiyatlar")

    fiyatlar = get_fiyatlar()

    if not fiyatlar:
        st.info("Henüz fiyat tanımı bulunmuyor.")
    else:
        fiyat_df = pd.DataFrame(
            [
                {
                    "ID": item["id"],
                    "Kategori": item["kategori"],
                    "Ad": item["ad"],
                    "Açıklama": item.get("aciklama", ""),
                    "Birim Fiyat (€)": float(item["birim_fiyat_eur"]),
                }
                for item in fiyatlar
            ]
        )
        st.dataframe(fiyat_df, use_container_width=True, hide_index=True)

        st.markdown("#### Düzenle veya sil")
        secenekler = {
            f'{x["kategori"]} — {x["ad"]} — {euro(x["birim_fiyat_eur"])}': x
            for x in fiyatlar
        }
        secim = st.selectbox("Kayıt seç", list(secenekler.keys()))
        secilen = secenekler[secim]

        with st.form("fiyat_duzenle_formu"):
            col1, col2 = st.columns(2)
            with col1:
                yeni_kategori = st.selectbox(
                    "Kategori",
                    ["Malzeme", "Kaplama", "Ek İşlem"],
                    index=["Malzeme", "Kaplama", "Ek İşlem"].index(
                        secilen["kategori"]
                    ),
                )
                yeni_ad = st.text_input("Ad", value=secilen["ad"])
            with col2:
                yeni_fiyat = st.number_input(
                    "Birim fiyat (€)",
                    min_value=0.0,
                    value=float(secilen["birim_fiyat_eur"]),
                    step=0.01,
                    format="%.4f",
                )
                yeni_aciklama = st.text_input(
                    "Açıklama", value=secilen.get("aciklama", "") or ""
                )

            col_update, col_delete = st.columns(2)
            with col_update:
                guncelle = st.form_submit_button(
                    "Güncelle", type="primary", use_container_width=True
                )
            with col_delete:
                sil = st.form_submit_button("Sil", use_container_width=True)

            if guncelle:
                if not yeni_ad.strip():
                    st.error("Ad alanı boş bırakılamaz.")
                else:
                    db.table("fiyat_tanimlari").update(
                        {
                            "kategori": yeni_kategori,
                            "ad": yeni_ad.strip(),
                            "aciklama": yeni_aciklama.strip(),
                            "birim_fiyat_eur": yeni_fiyat,
                        }
                    ).eq("id", secilen["id"]).execute()
                    st.success("Kayıt güncellendi.")
                    st.rerun()

            if sil:
                try:
                    db.table("fiyat_tanimlari").delete().eq(
                        "id", secilen["id"]
                    ).execute()
                    st.success("Kayıt silindi.")
                    st.rerun()
                except Exception:
                    st.error(
                        "Bu fiyat tanımı kayıtlı bir parçada kullanıldığı için silinemedi."
                    )

with tab2:
    st.subheader("Yeni parça maliyeti")

    fiyatlar = get_fiyatlar()

    if not fiyatlar:
        st.warning("Önce en az bir fiyat tanımı eklemelisin.")
    else:
        if "kalem_sayisi" not in st.session_state:
            st.session_state.kalem_sayisi = 1

        col1, col2 = st.columns([3, 1])
        with col1:
            parca_adi = st.text_input("Parça adı", key="yeni_parca_adi")
        with col2:
            adet = st.number_input(
                "Üretilecek adet", min_value=1, value=1, step=1
            )

        label_map = {
            f'{x["kategori"]} — {x["ad"]} — {euro(x["birim_fiyat_eur"])}': x
            for x in fiyatlar
        }
        labels = list(label_map.keys())

        st.markdown("#### Maliyet kalemleri")
        secilen_kalemler = []

        for i in range(st.session_state.kalem_sayisi):
            col_a, col_b, col_c, col_d = st.columns([4, 1.5, 1.5, 1.5])
            with col_a:
                label = st.selectbox(
                    f"Kalem {i + 1}",
                    labels,
                    key=f"kalem_{i}",
                    label_visibility="collapsed",
                )
            with col_b:
                miktar = st.number_input(
                    f"Miktar {i + 1}",
                    min_value=0.0,
                    value=1.0,
                    step=0.01,
                    format="%.4f",
                    key=f"miktar_{i}",
                    label_visibility="collapsed",
                )

            item = label_map[label]
            satir_toplam = miktar * float(item["birim_fiyat_eur"])

            with col_c:
                st.text_input(
                    f"Birim fiyat {i + 1}",
                    value=euro(item["birim_fiyat_eur"]),
                    disabled=True,
                    key=f"fiyat_goster_{i}",
                    label_visibility="collapsed",
                )
            with col_d:
                st.text_input(
                    f"Tutar {i + 1}",
                    value=euro(satir_toplam),
                    disabled=True,
                    key=f"tutar_goster_{i}",
                    label_visibility="collapsed",
                )

            secilen_kalemler.append(
                {
                    "fiyat_tanimi_id": item["id"],
                    "miktar": miktar,
                    "birim_fiyat_eur": float(item["birim_fiyat_eur"]),
                    "satir_toplam": satir_toplam,
                }
            )

        col_add, col_remove = st.columns(2)
        with col_add:
            if st.button("＋ İşlem / kalem ekle", use_container_width=True):
                st.session_state.kalem_sayisi += 1
                st.rerun()
        with col_remove:
            if st.button(
                "− Son kalemi kaldır",
                use_container_width=True,
                disabled=st.session_state.kalem_sayisi <= 1,
            ):
                st.session_state.kalem_sayisi -= 1
                st.rerun()

        tek_parca_toplam = sum(x["satir_toplam"] for x in secilen_kalemler)
        genel_toplam = tek_parca_toplam * int(adet)

        st.divider()
        col1, col2, col3 = st.columns(3)
        col1.metric("Tek parça maliyeti", euro(tek_parca_toplam))
        col2.metric("Üretilecek adet", int(adet))
        col3.metric("Genel toplam", euro(genel_toplam))

        if st.button("Parçayı kaydet", type="primary", use_container_width=True):
            if not parca_adi.strip():
                st.error("Parça adı boş bırakılamaz.")
            else:
                parca_response = (
                    db.table("parcalar")
                    .insert(
                        {
                            "parca_adi": parca_adi.strip(),
                            "adet": int(adet),
                        }
                    )
                    .execute()
                )

                parca_id = parca_response.data[0]["id"]

                db.table("parca_kalemleri").insert(
                    [
                        {
                            "parca_id": parca_id,
                            "fiyat_tanimi_id": x["fiyat_tanimi_id"],
                            "miktar": x["miktar"],
                            "birim_fiyat_eur": x["birim_fiyat_eur"],
                        }
                        for x in secilen_kalemler
                    ]
                ).execute()

                st.session_state.kalem_sayisi = 1
                st.success("Parça maliyeti kaydedildi.")
                st.rerun()

with tab3:
    st.subheader("Kayıtlı parçalar")

    parcalar = get_parcalar()
    fiyatlar = get_fiyatlar()
    fiyat_map = {item["id"]: item for item in fiyatlar}

    if not parcalar:
        st.info("Henüz kayıtlı parça bulunmuyor.")
    else:
        parca_secimleri = {
            f'{x["parca_adi"]} — {x["adet"]} adet — ID {x["id"]}': x
            for x in parcalar
        }
        secim = st.selectbox("Parça seç", list(parca_secimleri.keys()))
        parca = parca_secimleri[secim]
        kalemler = get_kalemler(parca["id"])

        detay = []
        tek_parca = 0.0
        for kalem in kalemler:
            tanim = fiyat_map.get(kalem["fiyat_tanimi_id"], {})
            miktar = float(kalem["miktar"])
            birim_fiyat = float(kalem["birim_fiyat_eur"])
            tutar = miktar * birim_fiyat
            tek_parca += tutar
            detay.append(
                {
                    "Kategori": tanim.get("kategori", ""),
                    "Kalem": tanim.get("ad", ""),
                    "Açıklama": tanim.get("aciklama", ""),
                    "Miktar": miktar,
                    "Birim Fiyat (€)": birim_fiyat,
                    "Tek Parça Tutarı (€)": tutar,
                }
            )

        col1, col2, col3 = st.columns(3)
        col1.metric("Tek parça maliyeti", euro(tek_parca))
        col2.metric("Adet", parca["adet"])
        col3.metric("Genel toplam", euro(tek_parca * int(parca["adet"])))

        st.dataframe(
            pd.DataFrame(detay),
            use_container_width=True,
            hide_index=True,
        )

        if st.button("Bu parçayı sil", type="secondary"):
            db.table("parcalar").delete().eq("id", parca["id"]).execute()
            st.success("Parça ve maliyet kalemleri silindi.")
            st.rerun()

with tab4:
    st.subheader("Excel dışa aktar")
    st.write(
        "Fiyat tanımları, parçalar, maliyet detayları ve maliyet özeti "
        "ayrı Excel sayfaları olarak hazırlanır."
    )

    st.download_button(
        "Excel dosyasını indir",
        data=excel_dosyasi(),
        file_name=f"ITSSystems_Cost_Calculator_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
