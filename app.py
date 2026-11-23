import streamlit as st
import os
import pandas as pd
import io
import re
import json
import urllib.parse
import fitz
from PIL import Image
from rapidfuzz import fuzz
try:
    from rapidocr_onnxruntime import RapidOCR
except Exception:
    RapidOCR = None
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from pydantic import BaseModel

st.set_page_config(page_title="B2B Kurumsal Teklif & Fiyat Portalı", page_icon="🏢", layout="wide")

with st.sidebar:
    st.markdown("### 🎛️ B2B Portal Ayarları")
    tema_rengi = st.color_picker("Kurumsal Tema Rengi", "#0F172A")
    vurgu_rengi = st.color_picker("Vurgu & Buton Rengi", "#2563EB")
    calisma_modu = st.radio(
        "Eşleştirme Motoru:",
        ("🔌 İnternetsiz (Evrensel Dinamik Motor)", "🌐 Yapay Zeka (Gemini API)")
    )
    api_key = ""
    if "Gemini" in calisma_modu:
        api_key = st.text_input("Gemini API Anahtarı", type="password")

    st.divider()
    st.markdown("### 🏢 Firma & Antet Bilgileri")
    logo_dosyasi = st.file_uploader("Firma Logosu (PNG / JPG)", type=("png", "jpg", "jpeg"))
    firma_adi = st.text_input("Satıcı Firma Ünvanı", value="TEKNOLOJİ VE TEDARİK A.Ş.")
    firma_bilgi = st.text_area("Adres & İletişim", value="1. Organize Sanayi Bölgesi 1. Cad. Sakarya\nTel: 0264 000 00 00 | satis@firma.com")
    musteri_adi = st.text_input("Alıcı Bayi / Müşteri Ünvanı", value="Sakarya Elektrik Dağıtım Ltd. Şti.")
    musteri_tel = st.text_input("Bayi WhatsApp No (905xxxxxxxxx)", value="905")

st.markdown(f"""
<style>
    .stApp {{ background-color: #F4F6F9; }}
    .b2b-header {{
        background: linear-gradient(120deg, {tema_rengi} 0%, {vurgu_rengi} 100%);
        padding: 24px 30px; border-radius: 16px; color: white; margin-bottom: 20px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
    }}
    .b2b-header h1 {{ color: white !important; font-size: 26px !important; margin: 0 !important; font-weight: 700 !important; }}
    .b2b-badge {{
        background: rgba(255,255,255,0.18); padding: 5px 12px; border-radius: 20px;
        font-size: 12px; font-weight: 600; display: inline-block; margin-top: 8px;
    }}
    .kpi-card {{
        background: white; border-radius: 14px; padding: 18px 20px;
        border-left: 5px solid {vurgu_rengi}; box-shadow: 0 4px 12px rgba(0,0,0,0.04); margin-bottom: 16px;
    }}
    .kpi-title {{ color: #64748B; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }}
    .kpi-value {{ color: #0F172A; font-size: 24px; font-weight: 800; margin-top: 4px; }}
    .b2b-card {{
        background: white; border: 1px solid #E2E8F0; border-top: 4px solid {vurgu_rengi};
        border-radius: 14px; padding: 20px; margin-bottom: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.04);
    }}
    .b2b-card-title {{ color: {tema_rengi}; font-size: 17px; font-weight: 700; margin-bottom: 6px; }}
    .b2b-card-sub {{ color: #64748B; font-size: 13px; }}
    div.stButton > button {{
        background: linear-gradient(135deg, {vurgu_rengi}, {tema_rengi}) !important;
        color: white !important; border-radius: 10px !important; padding: 12px 18px !important;
        font-weight: 600 !important; border: none !important;
    }}
    div.stDownloadButton > button {{
        background: linear-gradient(135deg, #059669, #047857) !important;
        color: white !important; border-radius: 10px !important; padding: 12px 18px !important; font-weight: 600 !important;
    }}
</style>
""", unsafe_allow_html=True)

@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

def turkce_normalize(metin):
    degisim = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
    return str(metin).translate(degisim).upper()

def talep_satirlarini_cikar(dosya):
    dosya.seek(0)
    ext = dosya.name.split('.')[-1].lower()
    talepler = []
    if ext in ('xlsx', 'xls', 'csv'):
        df = pd.read_excel(dosya, header=None) if ext != 'csv' else pd.read_csv(dosya, header=None)
        for _, row in df.iterrows():
            hucreler = tuple(str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != "")
            if not hucreler:
                continue
            if any(turkce_normalize(h) in ("URUN", "MALZEME", "ADET", "MIKTAR", "ACIKLAMA", "STOK") for h in hucreler):
                continue
            urun_txt = hucreler
            adet = 1
            if len(hucreler) >= 2:
                try:
                    adet = int(float(hucreler[-1].replace(',', '.')))
                except ValueError:
                    adet_m = re.search(r'(\d+)\s*(?:ADET|AD|PCS)', turkce_normalize(urun_txt))
                    if adet_m:
                        adet = int(adet_m.group(1))
            talepler.append((urun_txt, adet))
    elif ext == 'pdf':
        doc = fitz.open(stream=dosya.read(), filetype="pdf")
        for page in doc:
            for line in page.get_text().split('\n'):
                if len(line.strip()) > 3:
                    talepler.append((line.strip(), 1))
    elif ext in ('png', 'jpg', 'jpeg'):
        ocr = get_ocr_engine()
        res, _ = ocr(dosya.read())
        if res:
            for line in res:
                if len(line.strip()) > 3:
                    talepler.append((line.strip(), 1))
    return talepler

def dinamik_katalog_veritabani(dosya):
    dosya.seek(0)
    ext = dosya.name.split('.')[-1].lower()
    urunler = []

    if ext in ('xlsx', 'xls', 'csv'):
        df = pd.read_excel(dosya) if ext != 'csv' else pd.read_csv(dosya)
        for _, row in df.iterrows():
            satir = tuple(str(x).strip() for x in row.values if pd.notna(x))
            fiyat = 0.0
            for val in reversed(satir):
                temiz = re.sub(r'[^\d.,]', '', val).replace(',', '.')
                try:
                    f = float(temiz)
                    if f > 0:
                        fiyat = f
                        break
                except ValueError:
                    pass
            if len(satir) >= 2 and fiyat > 0:
                tam = " ".join(satir[:-1])
                urunler.append({"kod": satir[0], "ad": tam, "arama_metni": turkce_normalize(tam), "fiyat": fiyat, "grup": "", "kutup": "", "amper": ""})
        return urunler

    doc = fitz.open(stream=dosya.read(), filetype="pdf")
    kod_deseni = re.compile(r'\b\d\S{2}-\S{4,6}-\S{3,6}\b')
    fiyat_deseni = re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2})')

    for page in doc:
        lines = tuple(l.strip() for l in page.get_text("text").split('\n') if l.strip())
        if not lines:
            continue
        sayfa_baslik = " ".join(lines[:6])
        norm_sayfa = turkce_normalize(sayfa_baslik)

        grup = "GENEL"
        if "MINYATUR DEVRE KESICI" in norm_sayfa or "HOME SERISI" in norm_sayfa:
            grup = "OTOMATIK SIGORTA"
        elif "KACAK AKIM" in norm_sayfa or "RCCB" in norm_sayfa:
            grup = "KACAK AKIM"
        elif "ARK HATASI" in norm_sayfa or "FAFDD" in norm_sayfa:
            grup = "ARK ALGILAMA"
        elif "DEVRE KESICI" in norm_sayfa and ("SABIT" in norm_sayfa or "AYARLI" in norm_sayfa or "3 KUTUPLU" in norm_sayfa):
            grup = "TERMIK MANYETIK SALTER"
        elif "KONTAKTOR" in norm_sayfa:
            grup = "KONTAKTOR"
        elif "TERMIK ROLE" in norm_sayfa:
            grup = "TERMIK ROLE"

        son_alt_baslik = ""
        son_kutup = ""
        son_ka = ""

        for i, line in enumerate(lines):
            norm_line = turkce_normalize(line)
            if any(k in norm_line for k in ("SERISI", "TIPI", "SABIT", "AYARLI", "KONTAKTOR", "ROLE", "PARAFUDR", "AYIRICI", "TRAFO")):
                if len(line) < 60 and not kod_deseni.search(line):
                    son_alt_baslik = line
            if "1 KUTUP" in norm_line:
                son_kutup = "1P"
            elif "2 KUTUP" in norm_line:
                son_kutup = "2P"
            elif "3 KUTUP" in norm_line:
                son_kutup = "3P"
            elif "4 KUTUP" in norm_line:
                son_kutup = "4P"
            elif "1P+N" in norm_line or "FAZ+NOTR" in norm_line:
                son_kutup = "1P+N"

            ka_m = re.search(r'(\d+(?:[.,]\d+)?)\s*KA', norm_line)
            if ka_m and len(line) < 25:
                son_ka = ka_m.group(1).replace(',', '.') + "KA"

            kod_m = kod_deseni.search(line)
            if kod_m:
                ham_kod = kod_m.group(0)
                fiyat = 0.0
                for p_line in lines[i:min(len(lines), i+5)]:
                    bulunan = fiyat_deseni.findall(p_line)
                    if bulunan:
                        val = float(next(iter(bulunan)).replace('.', '').replace(',', '.'))
                        if val > 0.5:
                            fiyat = val
                            break
                if fiyat <= 0:
                    continue

                amper = ""
                for p_line in reversed(lines[max(0, i-4):i+1]):
                    amp_aralik = re.search(r'(\d+(?:[.,]\d+)?\s*-\s*\d+(?:[.,]\d+)?)', p_line)
                    if amp_aralik and grup == "TERMIK ROLE":
                        amper = amp_aralik.group(1).replace(' ', '').replace(',', '.') + "A"
                        break
                    amp_tek = re.match(r'^(\d+(?:[.,]\d+)?)\s*A?$', p_line.strip())
                    if amp_tek:
                        amper = amp_tek.group(1).replace(',', '.') + "A"
                        break
                if not amper:
                    son_rakam = re.search(r'(\d{2,4})$', ham_kod)
                    if son_rakam:
                        amper = str(int(son_rakam.group(1))) + "A"

                yakin_baglam = " ".join(lines[max(0, i-4):min(len(lines), i+3)])
                genel_baglam = re.sub(r'\s+', ' ', f"{grup} {sayfa_baslik} {son_alt_baslik} {son_kutup} {amper} {son_ka} {yakin_baglam}").strip()

                if "Δ" in ham_kod or "#" in ham_kod:
                    for tip in ("B", "C"):
                        temiz_kod = ham_kod.replace("Δ", tip).replace("#", "E")
                        ad_str = re.sub(r'\s+', ' ', f"{tip} Tipi {son_alt_baslik} {son_kutup} {amper} {son_ka}").strip()
                        urunler.append({
                            "kod": temiz_kod, "ad": ad_str,
                            "arama_metni": turkce_normalize(f"{tip} TIPI {genel_baglam}"),
                            "fiyat": fiyat, "grup": grup, "kutup": son_kutup, "amper": amper, "ka": son_ka, "tip": tip
                        })
                else:
                    tip = "B" if "-B" in ham_kod else ("C" if "-C" in ham_kod else "")
                    ad_str = re.sub(r'\s+', ' ', f"{son_alt_baslik} {son_kutup} {amper} {son_ka} ({yakin_baglam[:35]})").strip()
                    urunler.append({
                        "kod": ham_kod, "ad": ad_str,
                        "arama_metni": turkce_normalize(genel_baglam),
                        "fiyat": fiyat, "grup": grup, "kutup": son_kutup, "amper": amper, "ka": son_ka, "tip": tip
                    })
    return urunler

def evrensel_teknik_eslestir(fiyat_db, talep_listesi, varsayilan_iskonto):
    sonuclar = []
    if not fiyat_db:
        return sonuclar

    for talep_txt, adet in talep_listesi:
        t_norm = turkce_normalize(talep_txt).replace(',', '.')

        hedef_grup = ""
        if "OTOMATIK SIGORTA" in t_norm or "MCB" in t_norm:
            hedef_grup = "OTOMATIK SIGORTA"
        elif "KACAK AKIM" in t_norm or "RCCB" in t_norm:
            hedef_grup = "KACAK AKIM"
        elif "ARK" in t_norm or "AFDD" in t_norm:
            hedef_grup = "ARK ALGILAMA"
        elif "TERMIK MANYETIK" in t_norm or "SALTER" in t_norm:
            hedef_grup = "TERMIK MANYETIK SALTER"
        elif "TERMIK ROLE" in t_norm:
            hedef_grup = "TERMIK ROLE"
        elif "KONTAKTOR" in t_norm:
            hedef_grup = "KONTAKTOR"

        kutup_m = re.search(r'\b(1|2|3|4)\s*(?:X|P)\b', t_norm)
        hedef_kutup = (kutup_m.group(1) + "P") if kutup_m else ""

        hedef_amper = ""
        aralik_m = re.search(r'(\d+(?:\.\d+)?\s*-\s*\d+(?:\.\d+)?)', t_norm)
        if aralik_m and hedef_grup == "TERMIK ROLE":
            hedef_amper = aralik_m.group(1).replace(' ', '') + "A"
        else:
            amp_m = re.search(r'(\d+(?:\.\d+)?)\s*A\b', t_norm)
            if amp_m:
                hedef_amper = amp_m.group(1) + "A"
            elif hedef_grup == "TERMIK ROLE":
                tek_aralik = re.search(r'\b(\d+-\d+)\b', t_norm)
                if tek_aralik:
                    hedef_amper = tek_aralik.group(1) + "A"

        ka_m = re.search(r'(\d+(?:\.\d+)?)\s*KA', t_norm)
        hedef_ka = (ka_m.group(1) + "KA") if ka_m else ""

        hedef_tip = "B" if ("B TIP" in t_norm) else ("C" if ("C TIP" in t_norm) else "")

        en_iyi_urun = None
        en_yuksek_skor = -1.0

        for u in fiyat_db:
            k_metin = u["arama_metni"]
            skor = fuzz.token_set_ratio(t_norm, k_metin) * 0.4

            if hedef_grup and u.get("grup") == hedef_grup:
                skor += 45.0
            if hedef_kutup and (u.get("kutup") == hedef_kutup or (hedef_kutup == "1P" and u.get("kutup") == "1P+N")):
                skor += 30.0
            if hedef_amper and u.get("amper") == hedef_amper:
                skor += 40.0
            if hedef_ka and u.get("ka") == hedef_ka:
                skor += 20.0
            if hedef_tip and u.get("tip") == hedef_tip:
                skor += 20.0

            if skor > en_yuksek_skor:
                en_yuksek_skor = skor
                en_iyi_urun = u

        if en_iyi_urun:
            sonuclar.append({
                "Talep Edilen": talep_txt,
                "Ürün Kodu": en_iyi_urun["kod"],
                "Eşleşen Ürün": en_iyi_urun["ad"],
                "Adet": adet,
                "Liste Birim Fiyat": en_iyi_urun["fiyat"],
                "Özel İskonto (%)": float(varsayilan_iskonto)
            })
    return sonuclar
def pdf_olustur(df, firma_adi, firma_bilgi, musteri_adi, odeme_notu, kdv, para_birimi, logo_file, tema_hex):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    font_yolu = "C:/Windows/Fonts/arial.ttf"
    tr_font = "tr_arial" if os.path.exists(font_yolu) else "helv"
    if os.path.exists(font_yolu):
        page.insert_font(fontname="tr_arial", fontfile=font_yolu)

    h_hex = tema_hex.lstrip('#')
    r = int(h_hex[0:2], 16) / 255.0
    g = int(h_hex[2:4], 16) / 255.0
    b = int(h_hex[4:6], 16) / 255.0
    renk = (r, g, b)

    page.draw_rect(fitz.Rect(30, 30, 565, 95), color=renk, fill=renk)
    page.insert_text((45, 58), firma_adi, fontsize=15, color=(1, 1, 1), fontname=tr_font)
    bilgi_tek_satir = firma_bilgi.replace('\n', ' | ')
    page.insert_text((45, 78), bilgi_tek_satir[:85], fontsize=8.5, color=(0.95, 0.95, 0.95), fontname=tr_font)

    if logo_file is not None:
        page.insert_image(fitz.Rect(475, 36, 555, 89), stream=logo_file.getvalue())

    page.insert_text((35, 120), f"Sayın / Bayi: {musteri_adi}", fontsize=11, color=(0.1, 0.1, 0.1), fontname=tr_font)
    page.insert_text((435, 120), f"Tarih: {pd.Timestamp.now().strftime('%d.%m.%Y')}", fontsize=10, color=(0.2, 0.2, 0.2), fontname=tr_font)
    page.draw_line((35, 130), (560, 130), color=renk, width=1.2)

    # Mavi Tablo Başlık Şeridi
    y = 145
    page.draw_rect(fitz.Rect(35, y, 560, y + 24), color=renk, fill=renk)
    kolonlar = ((38, "Sipariş Kodu"), (118, "Ürün Açıklaması"), (305, "Adet"), (340, f"Liste ({para_birimi})"), (400, "İsk.%"), (440, f"Net ({para_birimi})"), (500, f"Toplam ({para_birimi})"))
    for x_pos, baslik in kolonlar:
        page.insert_text((x_pos, y + 16), baslik, fontsize=8, color=(1, 1, 1), fontname=tr_font)

    # İlk satır ile mavi başlık arasına ferah boşluk
    y += 42
    ara_toplam, kdv_toplam, genel_toplam = 0.0, 0.0, 0.0

    for _, row in df.iterrows():
        if y > 730:
            page = doc.new_page(width=595, height=842)
            if os.path.exists(font_yolu):
                page.insert_font(fontname="tr_arial", fontfile=font_yolu)
            y = 55

        adet = int(row["Adet"])
        birim = float(row["Liste Birim Fiyat"])
        isk = float(row["Özel İskonto (%)"])
        net_birim = birim * (1 - isk / 100.0)
        kdv_tutar = net_birim * (kdv / 100.0)
        satir_top = (net_birim + kdv_tutar) * adet

        ara_toplam += net_birim * adet
        kdv_toplam += kdv_tutar * adet
        genel_toplam += satir_top

        page.insert_text((38, y), str(row["Ürün Kodu"])[:14], fontsize=7.5, fontname=tr_font)
        page.insert_text((118, y), str(row["Eşleşen Ürün"])[:36], fontsize=7.5, fontname=tr_font)
        page.insert_text((308, y), str(adet), fontsize=8, fontname=tr_font)
        page.insert_text((340, y), f"{birim:,.2f}", fontsize=8, fontname=tr_font)
        page.insert_text((402, y), f"%{isk:.0f}", fontsize=8, fontname=tr_font)
        page.insert_text((440, y), f"{net_birim:,.2f}", fontsize=8, fontname=tr_font)
        page.insert_text((500, y), f"{satir_top:,.2f}", fontsize=8, fontname=tr_font)
        page.draw_line((35, y + 6), (560, y + 6), color=(0.85, 0.85, 0.85), width=0.5)
        y += 20

    y += 12
    page.draw_rect(fitz.Rect(350, y, 560, y + 64), color=renk, width=1)
    page.insert_text((360, y + 18), f"Ara Toplam: {ara_toplam:,.2f} {para_birimi}", fontsize=9, fontname=tr_font)
    page.insert_text((360, y + 36), f"KDV (%{kdv}): {kdv_toplam:,.2f} {para_birimi}", fontsize=9, fontname=tr_font)
    page.insert_text((360, y + 54), f"GENEL TOPLAM: {genel_toplam:,.2f} {para_birimi}", fontsize=9.5, color=renk, fontname=tr_font)

    page.insert_text((35, y + 18), "ÖDEME YÖNTEMİ VE KOŞULLAR:", fontsize=9, color=renk, fontname=tr_font)
    kosul = odeme_notu if odeme_notu else "Peşin / Havale."
    page.insert_text((35, y + 36), kosul[:65], fontsize=8.5, fontname=tr_font)

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes
def excel_olustur(df, firma_adi, firma_bilgi, musteri_adi, odeme_notu, kdv, para_birimi, logo_file, tema_hex):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "B2B Fiyat Teklifi"
    renk_kodu = tema_hex.lstrip('#')

    baslik_font = Font(name="Calibri", size=16, bold=True, color=renk_kodu)
    tablo_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    tablo_fill = PatternFill(start_color=renk_kodu, end_color=renk_kodu, fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    if logo_file is not None:
        logo_file.seek(0)
        img = XLImage(logo_file)
        img.width = 110
        img.height = 55
        ws.add_image(img, "G1")

    ws["A1"] = firma_adi
    ws["A1"].font = baslik_font
    ws["A2"] = firma_bilgi
    ws["A4"] = f"Sayın / Bayi: {musteri_adi}"
    ws["A4"].font = Font(bold=True)
    ws["F4"] = f"Tarih: {pd.Timestamp.now().strftime('%d.%m.%Y')}"

    headers = ("Sipariş Kodu", "Ürün Açıklaması", "Adet", f"Liste Fiyat ({para_birimi})", "İskonto (%)", f"Net Birim ({para_birimi})", f"KDV (%{kdv})", f"Toplam ({para_birimi})")
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=c_idx, value=h)
        cell.font = tablo_font
        cell.fill = tablo_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    r_idx = 7
    ara_toplam, kdv_toplam, genel_toplam = 0.0, 0.0, 0.0
    for _, row in df.iterrows():
        adet = row["Adet"]
        birim = row["Liste Birim Fiyat"]
        isk = row["Özel İskonto (%)"]
        net_birim = birim * (1 - isk / 100)
        kdv_tutar = net_birim * (kdv / 100)
        satir_top = (net_birim + kdv_tutar) * adet

        ara_toplam += net_birim * adet
        kdv_toplam += kdv_tutar * adet
        genel_toplam += satir_top

        vals = (row["Ürün Kodu"], row["Eşleşen Ürün"], adet, round(birim, 2), isk, round(net_birim, 2), round(kdv_tutar, 2), round(satir_top, 2))
        for c_idx, v in enumerate(vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.border = border
            if c_idx >= 4:
                cell.number_format = '#,##0.00'
        r_idx += 1

    r_idx += 1
    for lbl, val in (("Ara Toplam:", ara_toplam), (f"Toplam KDV (%{kdv}):", kdv_toplam), ("GENEL TOPLAM:", genel_toplam)):
        ws.cell(row=r_idx, column=7, value=lbl).font = Font(bold=True)
        c = ws.cell(row=r_idx, column=8, value=round(val, 2))
        c.font = Font(bold=True)
        c.number_format = '#,##0.00'
        r_idx += 1

    r_idx += 1
    ws.cell(row=r_idx, column=1, value="ÖDEME YÖNTEMİ VE KOŞULLAR:").font = Font(bold=True, color=renk_kodu)
    ws.cell(row=r_idx + 1, column=1, value=odeme_notu if odeme_notu else "Peşin / Havale")

    widths = (18, 46, 10, 18, 14, 18, 16, 20)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
banner_sol, banner_sag = st.columns((5, 1))
with banner_sol:
    st.markdown(f"""
    <div class="b2b-header">
        <div>
            <h1>🏢 B2B Kurumsal Teklif & Sipariş Portalı</h1>
            <div style="opacity:0.9; font-size:14px; margin-top:4px;">Satıcı: <b>{firma_adi}</b> &nbsp;|&nbsp; Bayi: <b>{musteri_adi}</b></div>
            <div class="b2b-badge">● Sistem Aktif &nbsp;•&nbsp; {calisma_modu}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
with banner_sag:
    if logo_dosyasi:
        st.image(logo_dosyasi, width=140)
    else:
        st.markdown('<div class="b2b-card" style="text-align:center; padding:14px;"><b>LOGO</b><br><span style="font-size:11px; color:#64748B;">Sol menüden yükleyin</span></div>', unsafe_allow_html=True)

df_aktif = st.session_state.get("teklif_df", pd.DataFrame())
aktif_kur = st.session_state.get("aktif_kur", "USD")
kalem_sayisi = len(df_aktif) if not df_aktif.empty else 0
toplam_adet = int(df_aktif["Adet"].sum()) if not df_aktif.empty else 0

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">📦 Eşleşen Kalem</div><div class="kpi-value">{kalem_sayisi} Kalem</div></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">🔢 Toplam Sipariş Adedi</div><div class="kpi-value">{toplam_adet} Adet</div></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">💱 Aktif Teklif Dövizi</div><div class="kpi-value">{aktif_kur}</div></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">📅 Teklif Tarihi</div><div class="kpi-value">{pd.Timestamp.now().strftime("%d.%m.%Y")}</div></div>', unsafe_allow_html=True)

col_sol, col_sag = st.columns(2)

with col_sol:
    st.markdown('<div class="b2b-card"><div class="b2b-card-title">📑 1. Katalog & Ana Fiyat Listesi</div><div class="b2b-card-sub">Herhangi bir PDF veya Excel fiyat listenizi sürükleyip bırakın ya da masaüstünden seçin.</div></div>', unsafe_allow_html=True)
    ana_fiyat_dosyasi = st.file_uploader("Fiyat Listesi Yükle", type=("pdf", "xlsx", "xls", "csv"), label_visibility="collapsed", key="ana_liste")

with col_sag:
    st.markdown('<div class="b2b-card"><div class="b2b-card-title">🛒 2. Bayi / Müşteri Talep Sepeti & Sohbet</div><div class="b2b-card-sub">PDF, Ekran Görüntüsü, Excel yükleyin veya alttaki B2B talep kutusuna ürün listesini yazın.</div></div>', unsafe_allow_html=True)
    musteri_dosyalari = st.file_uploader("Müşteri Belgeleri (PDF, PNG, JPG, Excel)", type=("pdf", "png", "jpg", "jpeg", "xlsx", "csv"), accept_multiple_files=True, label_visibility="collapsed", key="musteri_doc")
    sohbet_prompt = st.chat_input("Hızlı B2B Sipariş Girişi (İstediğiniz ürün, amper, kutup veya kodu yazın)...")

col_isk_on, col_kdv_on = st.columns(2)
with col_isk_on:
    baslangic_iskonto = st.number_input("🏷️ Teklife Uygulanacak İskonto Oranı (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
with col_kdv_on:
    baslangic_kdv = st.number_input("🧾 KDV Oranı (%)", min_value=0, max_value=100, value=20)

baslat_butonu = st.button("⚡ B2B Kataloğu Tara ve İskontolu Teklifi Oluştur", use_container_width=True)

if baslat_butonu or sohbet_prompt:
    if not ana_fiyat_dosyasi:
        st.error("⚠️ Lütfen önce 1. Katalog kartına ana fiyat listenizi yükleyin!")
    elif not musteri_dosyalari and not sohbet_prompt:
        st.warning("⚠️ Lütfen 2. Talep kartına müşteri belgesi yükleyin veya sipariş kutusuna talep yazın.")
    else:
        with st.spinner("Tüm katalog taranıyor ve teknik parametreler dinamik eşleştiriliyor..."):
            if "İnternetsiz" in calisma_modu or not api_key:
                fiyat_db = dinamik_katalog_veritabani(ana_fiyat_dosyasi)
                tum_talepler = []
                if sohbet_prompt:
                    for s in sohbet_prompt.split('\n'):
                        if s.strip():
                            tum_talepler.append((s.strip(), 1))
                for d in musteri_dosyalari:
                    tum_talepler.extend(talep_satirlarini_cikar(d))

                eslesenler = evrensel_teknik_eslestir(fiyat_db, tum_talepler, baslangic_iskonto)
                if eslesenler:
                    st.session_state["teklif_df"] = pd.DataFrame(eslesenler)
                    st.session_state["aktif_kur"] = "USD"
                    st.rerun()
                else:
                    st.warning("Eşleşen ürün bulunamadı.")
            else:
                from google import genai
                from google.genai import types

                class UrunItem(BaseModel):
                    talep: str
                    kod: str
                    ad: str
                    adet: int
                    fiyat: float

                class UrunListe(BaseModel):
                    items: list[UrunItem]

                client = genai.Client(api_key=api_key)
                ana_fiyat_dosyasi.seek(0)
                pdf_part = types.Part.from_bytes(data=ana_fiyat_dosyasi.read(), mime_type="application/pdf")
                talepler_str = "\n".join([f"{u} -> Adet: {a}" for d in musteri_dosyalari for u, a in talep_satirlarini_cikar(d)])
                if sohbet_prompt:
                    talepler_str += f"\n{sohbet_prompt}"

                prompt_txt = f"Ekteki fiyat listesi PDF'inden aşağıdaki müşteri taleplerini en ayrıştırıcı teknik kriterlere göre bul ve sipariş kodu ile birim fiyatını çıkar:\n{talepler_str}"
                resp = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=(pdf_part, prompt_txt),
                    config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=UrunListe)
                )
                veri = json.loads(resp.text).get("items", [])
                st.session_state["teklif_df"] = pd.DataFrame([{
                    "Talep Edilen": x["talep"], "Ürün Kodu": x["kod"], "Eşleşen Ürün": x["ad"],
                    "Adet": x["adet"], "Liste Birim Fiyat": x["fiyat"], "Özel İskonto (%)": float(baslangic_iskonto)
                } for x in veri])
                st.session_state["aktif_kur"] = "USD"
                st.rerun()

if "teklif_df" in st.session_state and not st.session_state["teklif_df"].empty:
    st.markdown('<div class="b2b-card"><div class="b2b-card-title">📊 B2B Teklif Sepeti, İskonto & Döviz Yönetimi</div><div class="b2b-card-sub">Toplu veya ürün bazlı özel iskonto tanımlayın, USD/TL kurunu çevirin ve teklifinizi oluşturun.</div></div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns((1.2, 1.2, 1.3))
    with c1:
        usd_kuru = st.number_input("💱 Güncel USD/TL Kuru", min_value=1.0, value=38.50, step=0.10)
    with c2:
        st.write("**Döviz Dönüşümü**")
        if st.button("🔄 USD ⇄ TL Çevir", use_container_width=True):
            if st.session_state.get("aktif_kur", "USD") == "USD":
                st.session_state["teklif_df"]["Liste Birim Fiyat"] *= usd_kuru
                st.session_state["aktif_kur"] = "TL"
            else:
                st.session_state["teklif_df"]["Liste Birim Fiyat"] /= usd_kuru
                st.session_state["aktif_kur"] = "USD"
            st.rerun()
    with c3:
        toplu_iskonto = st.number_input("🏷️ Toplu İskontoyu Güncelle (%)", min_value=0.0, max_value=100.0, value=baslangic_iskonto, step=1.0)
        if st.button("📉 Tüm Kalemlere Uygula", use_container_width=True):
            st.session_state["teklif_df"]["Özel İskonto (%)"] = toplu_iskonto
            st.rerun()

    with st.expander("🎯 Ürün Üzerinden Özel İskonto Belirle (Tekil Kalem Seçimi)"):
        secilen_urun = st.selectbox("İskonto uygulanacak ürünü seçin:", st.session_state["teklif_df"]["Eşleşen Ürün"].tolist())
        yeni_ozel_isk = st.number_input(f"'{secilen_urun}' için Özel İskonto (%)", min_value=0.0, max_value=100.0, value=10.0)
        if st.button("✅ Seçili Kaleme Özel İskontoyu İşle"):
            st.session_state["teklif_df"].loc[st.session_state["teklif_df"]["Eşleşen Ürün"] == secilen_urun, "Özel İskonto (%)"] = yeni_ozel_isk
            st.rerun()

    duzenli_df = st.data_editor(st.session_state["teklif_df"], num_rows="dynamic", use_container_width=True)
    st.session_state["teklif_df"] = duzenli_df

    hesap = duzenli_df.copy()
    hesap["Net Birim Fiyat"] = hesap["Liste Birim Fiyat"] * (1 - hesap["Özel İskonto (%)"] / 100)
    hesap[f"KDV Dahil Birim (%{baslangic_kdv})"] = hesap["Net Birim Fiyat"] * (1 + baslangic_kdv / 100)
    hesap["Satır Toplam (KDV Dahil)"] = hesap[f"KDV Dahil Birim (%{baslangic_kdv})"] * hesap["Adet"]

    st.dataframe(hesap[["Ürün Kodu", "Eşleşen Ürün", "Adet", "Liste Birim Fiyat", "Özel İskonto (%)", "Net Birim Fiyat", f"KDV Dahil Birim (%{baslangic_kdv})", "Satır Toplam (KDV Dahil)"]].style.format({
        "Liste Birim Fiyat": "{:,.2f}",
        "Özel İskonto (%)": "%{:.1f}",
        "Net Birim Fiyat": "{:,.2f}",
        f"KDV Dahil Birim (%{baslangic_kdv})": "{:,.2f}",
        "Satır Toplam (KDV Dahil)": "{:,.2f}"
    }), use_container_width=True)

    ara_top = (hesap["Net Birim Fiyat"] * hesap["Adet"]).sum()
    kdv_top = (hesap["Net Birim Fiyat"] * (baslangic_kdv / 100) * hesap["Adet"]).sum()
    genel_toplam = hesap["Satır Toplam (KDV Dahil)"].sum()

    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Ara Toplam (KDV Hariç)</div><div class="kpi-value">{ara_top:,.2f} {aktif_kur}</div></div>', unsafe_allow_html=True)
    with t2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Toplam KDV (%{baslangic_kdv})</div><div class="kpi-value">{kdv_top:,.2f} {aktif_kur}</div></div>', unsafe_allow_html=True)
    with t3:
        st.markdown(f'<div class="kpi-card" style="border-left-color:#059669;"><div class="kpi-title">💰 KDV Dahil Genel Toplam</div><div class="kpi-value" style="color:#059669;">{genel_toplam:,.2f} {aktif_kur}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="b2b-card"><div class="b2b-card-title">💳 Ödeme Yöntemi ve B2B Sevkiyat Koşulları</div><div class="b2b-card-sub">Antetli teklif formunun altında görünecek ödeme şartlarını el ile yazın.</div></div>', unsafe_allow_html=True)
    odeme_yontemi = st.text_area("Ödeme Yöntemi:", placeholder="Örn: %50 Siparişte peşin, %50 teslimatta Havale/EFT. Teklif 5 iş günü geçerlidir.", label_visibility="collapsed")

    logo_excel = io.BytesIO(logo_dosyasi.getvalue()) if logo_dosyasi else None
    logo_pdf = io.BytesIO(logo_dosyasi.getvalue()) if logo_dosyasi else None

    excel_veri = excel_olustur(duzenli_df, firma_adi, firma_bilgi, musteri_adi, odeme_yontemi, baslangic_kdv, aktif_kur, logo_excel, vurgu_rengi)
    pdf_veri = pdf_olustur(duzenli_df, firma_adi, firma_bilgi, musteri_adi, odeme_yontemi, baslangic_kdv, aktif_kur, logo_pdf, vurgu_rengi)

    btn_col1, btn_col2, btn_col3 = st.columns(3)
    with btn_col1:
        st.download_button(
            label="📥 Teklifi Excel İndir",
            data=excel_veri,
            file_name=f"B2B_Teklif_{musteri_adi.replace(' ', '_')}_{aktif_kur}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with btn_col2:
        st.download_button(
            label="📄 Antetli Teklifi PDF İndir",
            data=pdf_veri,
            file_name=f"B2B_Teklif_{musteri_adi.replace(' ', '_')}_{aktif_kur}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    with btn_col3:
        wp_satirlar = [f"*{firma_adi} - B2B FİYAT TEKLİFİ*", f"Sayın {musteri_adi},", ""]
        for _, r in hesap.iterrows():
            wp_satirlar.append(f"• [{r['Ürün Kodu']}] {r['Eşleşen Ürün'][:35]} ({r['Adet']} Ad.) -> {r['Net Birim Fiyat']:,.2f} {aktif_kur} + KDV")
        wp_satirlar.append("")
        wp_satirlar.append(f"*KDV Dahil Genel Toplam: {genel_toplam:,.2f} {aktif_kur}*")
        if odeme_yontemi:
            wp_satirlar.append(f"Ödeme Koşulu: {odeme_yontemi}")

        wp_mesaj = urllib.parse.quote("\n".join(wp_satirlar))
        wp_link = f"https://wa.me{musteri_tel.strip()}?text={wp_mesaj}"

        st.link_button("📲 WhatsApp ile Gönder", url=wp_link, use_container_width=True)
