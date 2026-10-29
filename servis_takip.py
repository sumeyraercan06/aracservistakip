from datetime import datetime
import pandas as pd

# 1. Excel dosyasını oku
df = pd.read_excel("servis_kayitlari.xlsx")

# 2. Bugünden tam 11 ay önceki yıl ve ayı belirle (Manuel Sabit Filtre)
# 2026 yılındayız; 11 ay öncesi yaklaşık Temmuz 2025 yapar.
# Kodun anlaması için bunu string (metin) olarak yazıyoruz.
filtre_tarihi = "2025-07-01"

print("Müşteri listesi taranıyor...")
print(f"Kriter: {filtre_tarihi} tarihinden önce gelenler bakıma çağrılacak.\n")

# 3. Bakım zamanı gelenleri toplamak için boş bir liste oluştur
hatirlatma_listesi = []

# 4. Excel'deki tüm satırları tek tek döngüyle dönerek kontrol et
for indeks, satir in df.iterrows():
    # Satırdaki tarihi metne çevir ve ilk 10 karakterini (YIL-AY-GÜN) al
    servis_tarihi = str(satir["Son_Servis_Tarihi"])[:10]

    # Eğer müşterinin servis tarihi, bizim belirlediğimiz filtreden eski veya eşitse
    if servis_tarihi <= filtre_tarihi:
        # Müşterinin bilgilerini listemize ekle
        hatirlatma_listesi.append(
            {
                "Musteri_Adi": satir["Musteri_Adi"],
                "Telefon": str(satir["Telefon"]),
                "Plaka": satir["Plaka"],
                "Son_Servis_Tarihi": servis_tarihi,
            }
        )

# 5. Sonuçları yeni bir Excel dosyasına kaydet
if len(hatirlatma_listesi) > 0:
    # Listeyi Excel formatına çevir
    yeni_df = pd.DataFrame(hatirlatma_listesi)
    # Excel dosyası olarak bilgisayara kaydet
    yeni_df.to_excel("hatirlatilacak_musteriler.xlsx", index=False)

    print(f"İşlem Başarılı! {len(hatirlatma_listesi)} kişi bulundu.")
    print("Liste 'hatirlatilacak_musteriler.xlsx' olarak kaydedildi.")
else:
    print("Bakım zamanı gelmiş hiçbir müşteri bulunamadı.")