"""
Deneme için kasıtlı olarak DAĞINIK bir satış dosyası üretir: ornek_satislar.csv

Gerçek hayatta müşteriden gelen dosyalar böyle olur:
- Aynı şehir farklı yazılmış (istanbul, ISTANBUL, İstanbul...)
- Tarihler farklı formatlarda (2024-03-15, 15.03.2024, 15/03/2024)
- Tutarlar metin olarak ve farklı biçimlerde (1.250,50 TL, 1250.50, ₺1.250,50)
- Bazı satırlar yanlışlıkla iki kez girilmiş
"""

import csv
import random
from datetime import date, timedelta

random.seed(42)  # Her çalıştırmada aynı veri çıksın diye

SIPARIS_SAYISI = 198
TEKRAR_SAYISI = 10  # Kasıtlı olarak kopyalanacak satır sayısı

# Her şehrin dosyada karşımıza çıkabilecek farklı yazımları
SEHIR_YAZIMLARI = {
    "İstanbul": ["İstanbul", "istanbul", "ISTANBUL", "Istanbul"],
    "Ankara": ["Ankara", "ankara", "ANKARA"],
    "İzmir": ["İzmir", "izmir", "IZMIR"],
    "Bursa": ["Bursa", "bursa"],
    "Antalya": ["Antalya", "antalya"],
}

URUNLER = {
    "Kablosuz Kulaklık": 850,
    "Telefon Kılıfı": 150,
    "Şarj Aleti": 250,
    "Akıllı Saat": 2400,
    "Powerbank": 600,
    "Bluetooth Hoparlör": 1250,
}

MUSTERILER = [
    "Ayşe Yılmaz", "Mehmet Kaya", "Zeynep Demir", "Ali Çelik", "Fatma Şahin",
    "Mustafa Arslan", "Elif Koç", "Hasan Aydın", "Emine Öztürk", "Burak Kurt",
    "Selin Aksoy", "Can Polat",
]


def tarih_yaz(t):
    """Tarihi rastgele bir formatta metne çevirir."""
    format_ = random.choice(["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"])
    return t.strftime(format_)


def turk_formati(sayi):
    """1250.5 -> '1.250,50'"""
    metin = f"{sayi:,.2f}"  # '1,250.50'
    return metin.replace(",", "X").replace(".", ",").replace("X", ".")


def tutar_yaz(sayi):
    """Tutarı rastgele bir biçimde metne çevirir."""
    secim = random.choice(["tl_sonda", "duz", "tl_isareti", "kurussuz"])
    if secim == "tl_sonda":
        return f"{turk_formati(sayi)} TL"      # 1.250,50 TL
    if secim == "duz":
        return f"{sayi:.2f}"                   # 1250.50
    if secim == "tl_isareti":
        return f"₺{turk_formati(sayi)}"        # ₺1.250,50
    # kurussuz: 1.250 TL (sadece tam sayılarda anlamlı)
    if sayi == int(sayi):
        return f"{turk_formati(sayi)[:-3]} TL"
    return f"{turk_formati(sayi)} TL"


def main():
    baslangic = date(2024, 1, 1)
    satirlar = []

    for i in range(1, SIPARIS_SAYISI + 1):
        sehir = random.choice(list(SEHIR_YAZIMLARI))
        urun = random.choice(list(URUNLER))
        adet = random.randint(1, 5)
        fiyat = URUNLER[urun]
        tarih = baslangic + timedelta(days=random.randint(0, 180))
        musteri = random.choice(MUSTERILER)

        # Bazı müşteri adlarının başına/sonuna yanlışlıkla boşluk girilmiş
        if random.random() < 0.15:
            musteri = f" {musteri}  "

        satirlar.append({
            "Sipariş No": f"S{i:04d}",
            "Tarih": tarih_yaz(tarih),
            "Müşteri": musteri,
            "Şehir": random.choice(SEHIR_YAZIMLARI[sehir]),
            "Ürün": urun,
            "Adet": adet,
            "Birim Fiyat": fiyat,
            "Tutar": tutar_yaz(adet * fiyat),
        })

    # Bazı satırları kopyala (çift girilmiş kayıtlar)
    kopyalar = random.sample(satirlar, TEKRAR_SAYISI)
    satirlar.extend(dict(s) for s in kopyalar)
    random.shuffle(satirlar)

    with open("ornek_satislar.csv", "w", newline="", encoding="utf-8-sig") as f:
        yazici = csv.DictWriter(f, fieldnames=satirlar[0].keys(), delimiter=";")
        yazici.writeheader()
        yazici.writerows(satirlar)

    print(f"ornek_satislar.csv oluşturuldu: {len(satirlar)} satır "
          f"({TEKRAR_SAYISI} tanesi kasıtlı tekrar)")


if __name__ == "__main__":
    main()
