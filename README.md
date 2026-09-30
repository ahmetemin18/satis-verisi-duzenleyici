# Satış Verisi Temizleyici ve Rapor Üretici

Dağınık bir satış dosyasını (CSV veya Excel) otomatik olarak temizleyen ve tek tıkla okunabilir bir Excel raporu üreten Python scripti.

## Hangi sorunu çözüyor?

Küçük işletmelerin satış tabloları genelde elle doldurulur ve zamanla dağılır:

| Sorun | Örnek | Script ne yapıyor? |
|---|---|---|
| Aynı şehir farklı yazılmış | `istanbul`, `ISTANBUL`, `Istanbul` | Tek doğru ada indiriyor: `İstanbul` |
| Karışık tarih formatları | `2024-03-15`, `15.03.2024`, `15/03/2024` | Hepsini gerçek tarihe çeviriyor |
| Tutarlar metin olarak girilmiş | `1.250,50 TL`, `₺1.250,50`, `1250.50` | Sayıya çeviriyor: `1250.5` |
| Çift girilmiş siparişler | Aynı sipariş numarası iki kez | Tekrarları siliyor |
| Fazla boşluklar | `" Elif Koç  "` | Temizliyor |

## Sonuç

Örnek veride:

- **208 satırlık** dağınık dosya → **198 temiz satır** (10 çift kayıt silindi)
- **14 farklı şehir yazımı** → **5 şehir**
- Toplam ciro, *adet × birim fiyat* ile ayrıca hesaplanıp doğrulanıyor

Üretilen `temiz_rapor.xlsx` iki sayfadan oluşuyor:

- **Özet:** genel özet, en çok sipariş veren 3 müşteri, şehirlere göre, ürünlere göre ve aylara göre ciro
- **Temiz Veri:** düzeltilmiş tüm siparişler

## Kurulum ve kullanım

```bash
pip install pandas openpyxl
python ornek_veri_uret.py            # deneme için dağınık örnek veri üretir
python temizle.py ornek_satislar.csv # temiz_rapor.xlsx dosyasını üretir
```

Kendi dosyanızla:

```bash
python temizle.py satislar.xlsx rapor.xlsx
```

## Dosyalar

- `temizle.py`: temizleme ve raporlama scripti
- `ornek_veri_uret.py`: kasıtlı olarak dağınık örnek veri üretir
- `temiz_rapor.xlsx`: örnek çıktı

## Kullanılan teknolojiler

Python, pandas, openpyxl
