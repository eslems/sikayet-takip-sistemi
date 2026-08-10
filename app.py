"""
Sikayet Takip Sistemi - Web Uygulamasi (Streamlit)
------------------------------------------------------
Bu dosya bilgisayarina kurulmaz -- Streamlit Community Cloud'a (ucretsiz)
yukleyip bir LINK olarak paylasabilirsin. O linki acan herkes (supervision'in
dahil) hicbir kurulum yapmadan, sadece tarayicidan kullanir.

ONEMLI: Bu versiyon her ziyaret icin gecicidir -- yani sayfayi kapatip
tekrar actiginda o oturumdaki kayitlar sifirlanir, herkesin GORDUGU ortak
kalici bir Excel dosyasi degildir. Islenen Excel'i "Indir" butonuyla
indirip kendi bilgisayarina kaydetmen gerekir. Eger BIRDEN FAZLA KISININ
ortak, kalici bir dosya uzerinde calismasini istiyorsan (herkes ekler,
hepsi ayni dosyayi gorur), bunun icin Google Sheets baglantisi eklememiz
gerekir -- bu, ayri bir kurulum adimi, istersen sonra ekleriz.
"""

import re
from pathlib import Path
from io import BytesIO

import streamlit as st
import pytesseract
from PIL import Image
from openpyxl import Workbook, load_workbook

# ---------------------- AYARLAR ----------------------
LANG = "tur+eng"
VALUE_COL_X = 140
RIGHT_COL_LIMIT = 500   # sagdaki ikinci sutuna (orn. "Sikayet") tasmayi engeller
ROW_TOLERANCE = 8
SHEET_HEADERS = ["Gorev No", "Ad Soyad", "Telefon", "Görev Tipi", "Şikayet Metni", "Kaynak Dosya"]
LABEL_KEYWORDS = {"ad_soyad": ["talep eden", "sms"], "telefon": ["telefon"]}
# -------------------------------------------------------


def gorev_no_bul(text):
    m = re.search(r"No\s*-?\s*\[?\s*(\d{4,})\s*\]?", text)
    return m.group(1) if m else ""


def satirlari_grupla(image):
    data = pytesseract.image_to_data(image, lang=LANG, output_type=pytesseract.Output.DICT)
    kelimeler = []
    for i in range(len(data["text"])):
        t = data["text"][i].strip()
        if t:
            kelimeler.append((data["top"][i], data["left"][i], t))
    kelimeler.sort()
    satirlar = []
    for top, left, t in kelimeler:
        eklendi = False
        for satir in satirlar:
            if abs(satir["top"] - top) <= ROW_TOLERANCE:
                satir["items"].append((left, t))
                satir["top"] = (satir["top"] + top) / 2
                eklendi = True
                break
        if not eklendi:
            satirlar.append({"top": top, "items": [(left, t)]})
    satirlar.sort(key=lambda s: s["top"])
    return satirlar


def satirdan_etiket_deger_ayir(satir):
    items = sorted(satir["items"])
    etiket = " ".join(t for left, t in items if left < VALUE_COL_X)
    deger = " ".join(t for left, t in items if VALUE_COL_X <= left < RIGHT_COL_LIMIT)
    return etiket.lower(), deger


def deger_bul(satirlar, anahtar_kelimeler):
    for satir in satirlar:
        etiket, deger = satirdan_etiket_deger_ayir(satir)
        if any(k in etiket for k in anahtar_kelimeler) and deger:
            return deger.strip()
    return ""


def telefon_temizle(ham_deger):
    return re.sub(r"\D", "", ham_deger)


def ad_soyad_temizle(ham_deger):
    return re.sub(r"^[^\wÇĞİÖŞÜçğıöşü]+|[^\wÇĞİÖŞÜçğıöşü]+$", "", ham_deger).strip()


def gorev_tipi_ve_sikayet_metnini_ayikla(satirlar):
    """Gorev Tipi degerini ve Aciklama etiketinden once gelen paragrafi
    AYRI AYRI dondurur: (gorev_tipi, sikayet_metni)."""
    idx_tipi, idx_aciklama = None, None
    for i, satir in enumerate(satirlar):
        etiket, _ = satirdan_etiket_deger_ayir(satir)
        if idx_tipi is None and "tipi" in etiket:
            idx_tipi = i
        if idx_tipi is not None and "iklama" in etiket:
            idx_aciklama = i
            break

    gorev_tipi_degeri = deger_bul(satirlar, ["tipi"])
    sikayet_metni = ""
    if idx_tipi is not None and idx_aciklama is not None and idx_aciklama > idx_tipi:
        parcalar = []
        for satir in satirlar[idx_tipi + 1: idx_aciklama]:
            items = sorted(satir["items"])
            parcalar.append(" ".join(t for _, t in items))
        sikayet_metni = " ".join(p for p in parcalar if p).strip()

    return gorev_tipi_degeri, sikayet_metni


def gorseli_isle(image, dosya_adi):
    tum_metin = pytesseract.image_to_string(image, lang=LANG)
    satirlar = satirlari_grupla(image)
    ad_soyad = ad_soyad_temizle(deger_bul(satirlar, LABEL_KEYWORDS["ad_soyad"]))
    telefon_ham = deger_bul(satirlar, LABEL_KEYWORDS["telefon"])
    gorev_tipi, sikayet_metni = gorev_tipi_ve_sikayet_metnini_ayikla(satirlar)
    return {
        "Gorev No": gorev_no_bul(tum_metin),
        "Ad Soyad": ad_soyad,
        "Telefon": telefon_temizle(telefon_ham),
        "Görev Tipi": gorev_tipi,
        "Şikayet Metni": sikayet_metni,
        "Kaynak Dosya": dosya_adi,
    }


def kayitlari_excele_donustur(kayitlar):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sikayetler"
    ws.append(SHEET_HEADERS)
    for kayit in kayitlar:
        ws.append([kayit["Gorev No"], kayit["Ad Soyad"], kayit["Telefon"], kayit["Görev Tipi"], kayit["Şikayet Metni"], kayit["Kaynak Dosya"]])
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==================== ARAYUZ ====================

st.set_page_config(page_title="Sikayet Takip Sistemi", page_icon="📋")
st.title("📋 Sikayet Takip Sistemi")
st.caption("Ekran goruntusunden isim/telefon cikarip Excel'e aktarir.")

if "kayitlar" not in st.session_state:
    st.session_state.kayitlar = []

yuklenen_dosyalar = st.file_uploader(
    "Sikayet ekran goruntulerini sec (birden fazla secebilirsin)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if st.button("Isle ve Listeye Ekle", type="primary", disabled=not yuklenen_dosyalar):
    ilerleme = st.progress(0)
    for i, dosya in enumerate(yuklenen_dosyalar):
        image = Image.open(dosya)
        kayit = gorseli_isle(image, dosya.name)
        st.session_state.kayitlar.append(kayit)
        if not kayit["Ad Soyad"] or not kayit["Telefon"]:
            st.warning(f"{dosya.name}: Isim veya telefon bos geldi, gorseli manuel kontrol et.")
        ilerleme.progress((i + 1) / len(yuklenen_dosyalar))
    st.success(f"{len(yuklenen_dosyalar)} gorsel islendi.")

if st.session_state.kayitlar:
    st.subheader("Bu Oturumdaki Kayitlar")
    st.dataframe(st.session_state.kayitlar, use_container_width=True)

    excel_buffer = kayitlari_excele_donustur(st.session_state.kayitlar)
    st.download_button(
        "Excel Olarak Indir",
        data=excel_buffer,
        file_name="sikayet_kayitlari.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("Listeyi Temizle"):
        st.session_state.kayitlar = []
        st.rerun()
else:
    st.info("Henuz islenmis kayit yok. Yukaridan gorsel sec ve 'Isle' butonuna bas.")
