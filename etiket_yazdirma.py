import re
import os
import html
import json
import sys
import pandas as pd
import fitz  # PyMuPDF
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

# -------------------------------------------------
# AYARLANABİLİR SABİTLER
# -------------------------------------------------
DEFAULT_SETTINGS = {
    "extra_bottom_pt": 25,
    "text_x": 10,
    "text_top_padding": 4,
    "text_bottom_padding": 4,
    "font_size": 8,
    "line_height_factor": 1.25,
    "sort_labels_by_first_product": True,
}

EXTRA_BOTTOM_PT = DEFAULT_SETTINGS["extra_bottom_pt"]

TEXT_X = DEFAULT_SETTINGS["text_x"]
TEXT_TOP_PADDING = DEFAULT_SETTINGS["text_top_padding"]
TEXT_BOTTOM_PADDING = DEFAULT_SETTINGS["text_bottom_padding"]

FONT_SIZE = DEFAULT_SETTINGS["font_size"]
LINE_HEIGHT_FACTOR = DEFAULT_SETTINGS["line_height_factor"]
FONT_COLOR = (0, 0, 0)

FONT_REGULAR = "regular"
FONT_BOLD = "bold"
WINDOWS_FONT_DIR = r"C:\Windows\Fonts"
FONT_REGULAR_FILE = os.path.join(WINDOWS_FONT_DIR, "arial.ttf")
FONT_BOLD_FILE = os.path.join(WINDOWS_FONT_DIR, "arialbd.ttf")

SHIPMENT_REGEX = re.compile(r"\b(\d{4}\s*\d{4}\s*-\s*\d{4}\s*-\s*\d+)\b")

FONT_OBJECTS = {
    FONT_REGULAR: fitz.Font(fontfile=FONT_REGULAR_FILE),
    FONT_BOLD: fitz.Font(fontfile=FONT_BOLD_FILE),
}
FONT_ARCHIVE = fitz.Archive(WINDOWS_FONT_DIR)


def uygulama_klasoru():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ayar_dosyasi_yolu():
    return os.path.join(uygulama_klasoru(), "ayarlar.json")


def ayarlari_oku():
    yol = ayar_dosyasi_yolu()
    if not os.path.exists(yol):
        with open(yol, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_SETTINGS, f, ensure_ascii=False, indent=4)
        return DEFAULT_SETTINGS.copy()

    with open(yol, "r", encoding="utf-8") as f:
        kullanici_ayarlari = json.load(f)

    ayarlar = DEFAULT_SETTINGS.copy()
    if isinstance(kullanici_ayarlari, dict):
        ayarlar.update(kullanici_ayarlari)

    return ayarlar


def pozitif_sayi(value, default):
    try:
        value = float(value)
        if value > 0:
            return value
    except Exception:
        pass
    return default


def ayarlari_uygula(ayarlar):
    global EXTRA_BOTTOM_PT, TEXT_X, TEXT_TOP_PADDING, TEXT_BOTTOM_PADDING
    global FONT_SIZE, LINE_HEIGHT_FACTOR

    EXTRA_BOTTOM_PT = pozitif_sayi(
        ayarlar.get("extra_bottom_pt"),
        DEFAULT_SETTINGS["extra_bottom_pt"],
    )
    TEXT_X = pozitif_sayi(
        ayarlar.get("text_x"),
        DEFAULT_SETTINGS["text_x"],
    )
    TEXT_TOP_PADDING = pozitif_sayi(
        ayarlar.get("text_top_padding"),
        DEFAULT_SETTINGS["text_top_padding"],
    )
    TEXT_BOTTOM_PADDING = pozitif_sayi(
        ayarlar.get("text_bottom_padding"),
        DEFAULT_SETTINGS["text_bottom_padding"],
    )
    FONT_SIZE = pozitif_sayi(
        ayarlar.get("font_size"),
        DEFAULT_SETTINGS["font_size"],
    )
    LINE_HEIGHT_FACTOR = pozitif_sayi(
        ayarlar.get("line_height_factor"),
        DEFAULT_SETTINGS["line_height_factor"],
    )


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


def segment_width(segment, fontsize):
    text, font_style = segment
    return FONT_OBJECTS[font_style].text_length(str(text), fontsize=fontsize)


def segments_width(segments, fontsize):
    return sum(segment_width(segment, fontsize) for segment in segments)


def line_height(fontsize):
    return fontsize * LINE_HEIGHT_FACTOR


def split_text_for_width(text, font_style, max_width, fontsize):
    words = str(text).split(" ")
    lines = []
    current = ""

    for word in words:
        candidate = word if not current else f"{current} {word}"
        if current and segment_width((candidate, font_style), fontsize) > max_width:
            lines.append([(current, font_style)])
            current = word
        else:
            current = candidate

    if current:
        lines.append([(current, font_style)])

    return lines or [[("", font_style)]]


def wrap_segments_to_width(segments, max_width, fontsize):
    if segments_width(segments, fontsize) <= max_width:
        return [segments]

    lines = []
    current = []

    for text, font_style in segments:
        text = str(text)
        candidate = current + [(text, font_style)]
        if current and segments_width(candidate, fontsize) > max_width:
            lines.append(current)
            current = [(text.lstrip(), font_style)]
        else:
            current = candidate

    if current:
        lines.append(current)

    wrapped_lines = []
    for line in lines:
        if segments_width(line, fontsize) <= max_width:
            wrapped_lines.append(line)
            continue

        for text, font_style in line:
            wrapped_lines.extend(split_text_for_width(text, font_style, max_width, fontsize))

    return wrapped_lines


def build_label_lines(items, max_width):
    lines = []
    for item_segments in items:
        lines.extend(wrap_segments_to_width(item_segments, max_width, FONT_SIZE))
    return lines, FONT_SIZE


def required_bottom_height(line_count, fontsize):
    return TEXT_TOP_PADDING + TEXT_BOTTOM_PADDING + (line_count * line_height(fontsize))


def sort_text(value):
    return str(value).strip().casefold()


def html_color(color):
    r, g, b = (int(max(0, min(1, c)) * 255) for c in color)
    return f"#{r:02x}{g:02x}{b:02x}"


def lines_to_html(lines):
    html_lines = []
    for line in lines:
        parts = []
        for text, font_style in line:
            safe_text = html.escape(str(text))
            if font_style == FONT_BOLD:
                parts.append(f"<strong>{safe_text}</strong>")
            else:
                parts.append(safe_text)
        html_lines.append(f"<div>{''.join(parts)}</div>")
    return "".join(html_lines)


def draw_label_lines(page, x, top, width, lines, fontsize, color):
    css = f"""
    @font-face {{
        font-family: ArialLocal;
        src: url(arial.ttf);
    }}
    @font-face {{
        font-family: ArialLocal;
        font-weight: bold;
        src: url(arialbd.ttf);
    }}
    * {{
        font-family: ArialLocal;
        font-size: {fontsize}pt;
        line-height: {LINE_HEIGHT_FACTOR};
        color: {html_color(color)};
        margin: 0;
        padding: 0;
    }}
    div {{
        white-space: nowrap;
    }}
    """
    rect = fitz.Rect(
        x,
        top,
        x + width,
        top + required_bottom_height(len(lines), fontsize),
    )
    page.insert_htmlbox(rect, lines_to_html(lines), css=css, archive=FONT_ARCHIVE)


def main():
    root = Tk()
    root.withdraw()

    try:
        ayarlar = ayarlari_oku()
        ayarlari_uygula(ayarlar)
    except Exception as e:
        messagebox.showerror("Hata", f"Ayar dosyası okunamadı:\n{e}")
        return

    siralama_aktif = bool(ayarlar.get(
        "sort_labels_by_first_product",
        DEFAULT_SETTINGS["sort_labels_by_first_product"],
    ))

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
        page_records = []

        for sayfa_num in range(src_pdf.page_count):
            src_page = src_pdf[sayfa_num]
            rect = src_page.rect
            w, h = rect.width, rect.height

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

            items = []
            label_lines = []
            label_font_size = FONT_SIZE

            if not matched_sn:
                page_records.append({
                    "page_index": sayfa_num,
                    "width": w,
                    "height": h,
                    "bottom_height": EXTRA_BOTTOM_PT,
                    "label_lines": [],
                    "label_font_size": FONT_SIZE,
                    "sort_key": (1, sayfa_num) if siralama_aktif else (sayfa_num,),
                })
                continue

            urun_sirasi_df = siparis_gruplari.get(matched_sn, pd.DataFrame())
            if urun_sirasi_df.empty:
                page_records.append({
                    "page_index": sayfa_num,
                    "width": w,
                    "height": h,
                    "bottom_height": EXTRA_BOTTOM_PT,
                    "label_lines": [],
                    "label_font_size": FONT_SIZE,
                    "sort_key": (1, sayfa_num) if siralama_aktif else (sayfa_num,),
                })
                continue

            first_label_text = ""
            first_article_code = ""

            for _, satir in urun_sirasi_df.iterrows():
                orijinal_urun_kodu = str(satir.get('Article code', '')).strip()
                adet = satir.get('Quantity', 0)

                kisa_kod = kisa_urun_kodu_getir(orijinal_urun_kodu, kisa_urun_df_listesi)

                try:
                    adet_int = int(adet)
                except Exception:
                    adet_int = 0

                item_segments = [(str(kisa_kod), FONT_REGULAR)]

                if adet_int > 1:
                    item_segments.append((f" {adet_int}x", FONT_BOLD))

                if not first_label_text:
                    first_label_text = str(kisa_kod)
                    first_article_code = orijinal_urun_kodu

                items.append(item_segments)

                rapor_dict[orijinal_urun_kodu] = rapor_dict.get(orijinal_urun_kodu, 0) + max(adet_int, 0)

            max_text_width = max(1, w - (TEXT_X * 2))
            label_lines, label_font_size = build_label_lines(items, max_text_width)
            bottom_height = max(EXTRA_BOTTOM_PT, required_bottom_height(len(label_lines), label_font_size))

            page_records.append({
                "page_index": sayfa_num,
                "width": w,
                "height": h,
                "bottom_height": bottom_height,
                "label_lines": label_lines,
                "label_font_size": label_font_size,
                "sort_key": (sayfa_num,) if not siralama_aktif else (
                    0,
                    sort_text(first_label_text),
                    sort_text(first_article_code),
                    sort_text(matched_sn),
                    sayfa_num,
                ),
            })

        for record in sorted(page_records, key=lambda item: item["sort_key"]):
            sayfa_num = record["page_index"]
            w = record["width"]
            h = record["height"]

            dst_page = dst_pdf.new_page(width=w, height=h + record["bottom_height"])
            dst_page.show_pdf_page(fitz.Rect(0, 0, w, h), src_pdf, sayfa_num)

            if record["label_lines"]:
                draw_label_lines(
                    dst_page,
                    TEXT_X,
                    h + TEXT_TOP_PADDING,
                    max(1, w - (TEXT_X * 2)),
                    record["label_lines"],
                    fontsize=record["label_font_size"],
                    color=FONT_COLOR,
                )

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
