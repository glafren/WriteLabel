import pandas as pd
import fitz  # PyMuPDF
from tkinter import Tk, filedialog, messagebox

# Kısa ürün kodunu almak için güncellenmiş fonksiyon
def kisa_urun_kodu_getir(urun_kodu, kisa_urun_df_listesi):
    for kisa_urun_df in kisa_urun_df_listesi:
        # Kısa kod tablosundan, siparişin ürün kodunu bulalım
        kisa_urun_kodu_dict = pd.Series(
            kisa_urun_df['Kısa Ürün Kodu'].values,
            index=kisa_urun_df['Article Code'].values
        ).to_dict()
        
        # Ürün kodunu kısaltılmış ürün kodu ile eşleştirme
        if urun_kodu in kisa_urun_kodu_dict:
            return kisa_urun_kodu_dict[urun_kodu]
    return urun_kodu  # Eğer hiçbir eşleşme bulunamazsa orijinal ürün kodunu döndür

# Dosya seçme penceresi aç
root = Tk()
root.withdraw()  # Ana pencereyi gizle
siparis_dosyasi = filedialog.askopenfilename(title="Sipariş Listesini Seç", filetypes=[("Excel Files", "*.xlsx")])
etiket_dosyasi = filedialog.askopenfilename(title="Etiket PDF Dosyasını Seç", filetypes=[("PDF Files", "*.pdf")])

# Excel dosyasını oku, başlıkların 0. satırda olduğunu belirtiyoruz
df = pd.read_excel(siparis_dosyasi, header=0)

# Kısa Ürün Kodu verilerini yükle (birden fazla dosya)
kisa_urun_df_listesi = [
    pd.read_excel('ogulcan_kisaurunkodu.xlsx', header=0),
    pd.read_excel('dogcan_kisaurunkodu.xlsx', header=0)
]

# Başlık satırındaki boşlukları temizleyelim
df.columns = df.columns.str.strip()  # Sipariş listesinde başlıkları düzenle

# Kısa Ürün Kodu listelerindeki başlıkları temizle
for kisa_urun_df in kisa_urun_df_listesi:
    kisa_urun_df.columns = kisa_urun_df.columns.str.strip()

# PDF dosyasını aç
pdf = fitz.open(etiket_dosyasi)

# Her bir sayfa üzerinde işlem yapalım
for sayfa_num in range(pdf.page_count):
    sayfa = pdf[sayfa_num]
    
    # PDF sayfasındaki sipariş numarasını tespit et
    for siparis_numarasi in df['Shipment number']:
        # Sayfa üzerindeki metni arayıp, sipariş numarasını bulalım
        if siparis_numarasi in sayfa.get_text():
            # Siparişin ilgili satırını bul
            urun_sirasi = df[df['Shipment number'] == siparis_numarasi]
            urun_kodu = urun_sirasi['Article code'].values[0]  # Ürün Kodu
            adet = urun_sirasi['Quantity'].values[0]  # Adet bilgisi
            shipment_date = urun_sirasi['Shipment date'].values[0]  # Shipment date

            # Shipment date'i string formatına dönüştür
            shipment_date_str = pd.to_datetime(shipment_date).strftime('%d.%m')

            # Kısa Ürün Kodu'nu al
            kisa_urun_kodu = kisa_urun_kodu_getir(urun_kodu, kisa_urun_df_listesi)

            # Eğer adet 1'den fazlaysa, [Quantity]x ekleyelim
            if adet > 1:
                kisa_urun_kodu += f" {adet}x"
            
            # Ürün kodunu ve adet bilgisini PDF'ye ekleyelim
            sayfa.insert_text((40, 265), kisa_urun_kodu, fontsize=10, color=(0, 0, 0))  # Siyah yazı, 11 punto
            # Shipment date'i yazdıralım
            sayfa.insert_text((165, 265), shipment_date_str, fontsize=10, color=(0, 0, 0))  # Shipment date yazdır

# PDF dosyasını kaydedelim
pdf.save("Ogulcan_guncellenmis_etiketler.pdf")

# İşlem tamamlandığında kullanıcıya mesaj göster
messagebox.showinfo("İşlem Tamamlandı", "Etiketler başarıyla güncellendi ve kaydedildi!")
