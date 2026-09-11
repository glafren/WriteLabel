import re
import pandas as pd
import fitz  # PyMuPDF
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

# -------------------------------------------------
# AYARLANABİLİR SABİTLER
# -------------------------------------------------
EXTRA_BOTTOM_PT = 25

TEXT_X = 10
TEXT_Y_IN_PADDING = 10

FONT_SIZE = 8
FONT_COLOR = (0, 0, 0)

# Sürüm bağımsız Base-14 font adları
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

SHIPMENT_REGEX = re.compile(r"\b(\d{4}\s*\d{4}\s*-\s*\d{4}\s*-\s*\d+)\b")


def oku_siparis_dosyasi_yolundan(yol: str) -> pd.DataFrame:
    yol_lc = yol.lower()
    if yol_lc.endswith(".csv"):
        try:
            return pd.read_csv(yol, header=0, encoding="utf-8-sig", sep=None, engine="python")
        except Exception:
            return pd.read_csv(yol, header=0, encoding="utf-8-sig")
    return pd.read_excel(yol, header=0)


def kisa_urun_kodu_getir(urun_kodu, kisa_urun_df_listesi):
    for kisa_urun_df in kisa_urun_df_listesi:
        if 'Article Code' not in kisa_urun_df.columns or 'Kısa Ürün Kodu' not in kisa_urun_df.columns:
            continue
        kisa_urun_kodu_dict = pd.Series(
            kisa_urun_df['Kısa Ürün Kodu'].values,
            index=kisa_urun_df['Article Code'].values
        ).to_dict()
        if urun_kodu in kisa_urun_kodu_dict:
            return kisa_urun_kodu_dict[urun_kodu]
    return urun_kodu


def normalize_shipment(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")
    s = re.sub(r"\s+", "", s)
    return s


def extract_shipment_candidates_from_text(page_text: str) -> set[str]:
    if not page_text:
        return set()
    page_text = page_text.replace("–", "-").replace("—", "-").replace("−", "-")
    found = set()
    for m in SHIPMENT_REGEX.finditer(page_text):
        found.add(normalize_shipment(m.group(1)))
    return found


def draw_segments(page, x, y, segments, fontsize, color):
    """
    segments: [(text, fontname), ...]
    Aynı satıra farklı fontlarla yan yana yazar.
    """
    cur_x = x
    for text, fontname in segments:
        if not text:
            continue
        page.insert_text((cur_x, y), text, fontsize=fontsize, color=color, fontname=fontname)
        cur_x += fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)


def main():
    root = Tk()
    root.withdraw()

    siparis_dosyasi = filedialog.askopenfilename(
        title="Sipariş Listesini Seç",
        filetypes=[
            ("Excel or CSV", "*.xlsx;*.csv"),
            ("Excel Files", "*.xlsx"),
            ("CSV Files", "*.csv"),
        ]
    )
    if not siparis_dosyasi:
        messagebox.showwarning("Uyarı", "Sipariş dosyası seçilmedi.")
        return

    etiket_dosyasi = filedialog.askopenfilename(
        title="Etiket PDF Dosyasını Seç",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not etiket_dosyasi:
        messagebox.showwarning("Uyarı", "PDF dosyası seçilmedi.")
        return

    try:
        df = oku_siparis_dosyasi_yolundan(siparis_dosyasi)
    except Exception as e:
        messagebox.showerror("Hata", f"Sipariş dosyası okunamadı:\n{e}")
        return

    # Kısa ürün kodu eşlemesi (opsiyonel)
    kisa_urun_df_listesi = []
    try:
        kisa_df = pd.read_excel('omega_kisaurunkodu.xlsx', header=0)
        kisa_urun_df_listesi.append(kisa_df)
    except Exception as e:
        messagebox.showwarning(
            "Uyarı",
            f"'omega_kisaurunkodu.xlsx' okunamadı. Kısa kodlar kullanılmayacak.\n{e}"
        )

    df.columns = df.columns.str.strip()
    for kdf in kisa_urun_df_listesi:
        kdf.columns = kdf.columns.str.strip()

    gerekli = {'Shipment number', 'Article code', 'Quantity'}
    if not gerekli.issubset(df.columns):
        messagebox.showerror(
            "Hata",
            "Sipariş dosyasında gerekli sütunlar bulunamadı.\n"
            "Gerekli sütunlar: 'Shipment number', 'Article code', 'Quantity'"
        )
        return

    df['Shipment number'] = df['Shipment number'].astype(str).apply(normalize_shipment)
    df['Article code'] = df['Article code'].astype(str).str.strip()

    siparis_gruplari = {sn: grp for sn, grp in df.groupby('Shipment number')}
    rapor_dict = {}

    try:
        src_pdf = fitz.open(etiket_dosyasi)
    except Exception as e:
        messagebox.showerror("Hata", f"PDF açılamadı:\n{e}")
        return

    dst_pdf = fitz.open()

    try:
        for sayfa_num in range(src_pdf.page_count):
            src_page = src_pdf[sayfa_num]
            rect = src_page.rect
            w, h = rect.width, rect.height

            dst_page = dst_pdf.new_page(width=w, height=h + EXTRA_BOTTOM_PT)
            dst_page.show_pdf_page(fitz.Rect(0, 0, w, h), src_pdf, sayfa_num)

            page_text = src_page.get_text() or ""

            candidates = extract_shipment_candidates_from_text(page_text)

            matched_sn = None
            if candidates:
                for c in candidates:
                    if c in siparis_gruplari:
                        matched_sn = c
                        break
            else:
                norm_page_text = normalize_shipment(page_text)
                for sn in siparis_gruplari.keys():
                    if sn and sn in norm_page_text:
                        matched_sn = sn
                        break

            if not matched_sn:
                continue

            urun_sirasi_df = siparis_gruplari.get(matched_sn, pd.DataFrame())
            if urun_sirasi_df.empty:
                continue

            segments = []
            first = True

            for _, satir in urun_sirasi_df.iterrows():
                orijinal_urun_kodu = str(satir.get('Article code', '')).strip()
                adet = satir.get('Quantity', 0)

                kisa_kod = kisa_urun_kodu_getir(orijinal_urun_kodu, kisa_urun_df_listesi)

                try:
                    adet_int = int(adet)
                except Exception:
                    adet_int = 0

                if not first:
                    segments.append((" + ", FONT_REGULAR))
                first = False

                segments.append((str(kisa_kod), FONT_REGULAR))

                # SADECE adet kısmı kalın
                if adet_int > 1:
                    segments.append((f" {adet_int}x", FONT_BOLD))

                rapor_dict[orijinal_urun_kodu] = rapor_dict.get(orijinal_urun_kodu, 0) + max(adet_int, 0)

            x_koord = TEXT_X
            y_koord = h + TEXT_Y_IN_PADDING

            draw_segments(dst_page, x_koord, y_koord, segments, fontsize=FONT_SIZE, color=FONT_COLOR)

        bugun = datetime.today().strftime("%d.%m.%Y")
        varsayilan_ad = f"{bugun} Yazılı Etiketler.pdf"

        pdf_save_path = filedialog.asksaveasfilename(
            title="Etiketli PDF Dosyasını Kaydet",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=varsayilan_ad
        )

        if pdf_save_path:
            dst_pdf.save(pdf_save_path)
        else:
            messagebox.showwarning(
                "Uyarı",
                f"PDF kaydedilmedi. Varsayılan olarak '{varsayilan_ad}' ismiyle kaydediliyor."
            )
            dst_pdf.save(varsayilan_ad)

        messagebox.showinfo("İşlem Tamamlandı", "Etiketler başarıyla güncellendi ve kaydedildi!")

    except Exception as e:
        messagebox.showerror("Hata", f"İşlem sırasında hata oluştu:\n{e}")

    finally:
        try:
            src_pdf.close()
        except Exception:
            pass
        try:
            dst_pdf.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
