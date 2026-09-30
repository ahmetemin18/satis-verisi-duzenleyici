"""
Dağınık bir satış dosyasını (CSV veya Excel) temizler ve Excel raporu üretir.

Kullanım:
    python temizle.py ornek_satislar.csv
    python temizle.py satislar.xlsx cikti.xlsx
"""

import sys
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Dosyada görülen yazım -> doğru şehir adı.
# Anahtarlar küçük harfle yazılır, çünkü karşılaştırmadan önce metni küçültüyoruz.
SEHIR_ESLEME = {
    "istanbul": "İstanbul",
    "i̇stanbul": "İstanbul",  # "İstanbul".lower() Python'da böyle çıkar (i + nokta işareti)
    "ankara": "Ankara",
    "izmir": "İzmir",
    "i̇zmir": "İzmir",
    "bursa": "Bursa",
    "antalya": "Antalya"
}


TARIH_FORMATLARI = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]


def dosya_oku(yol):
    """CSV veya Excel dosyasını okur."""
    if yol.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(yol)
    return pd.read_csv(yol, sep=None, engine="python", encoding="utf-8-sig")


def tutar_temizle(deger):
    """'1.250,50 TL', '₺1.250,50', '1250.50', '1.250 TL' -> 1250.5"""
    if pd.isna(deger):
        return None
    if isinstance(deger, (int, float)):
        return float(deger)

    kucuk_metin = str(deger).strip().lower()
    metin = kucuk_metin.replace("tl", "").replace("₺", "").replace(" ", "").strip()

    if "," in metin:
        # Türk biçimi: nokta binlik ayırıcı, virgül ondalık
        metin = metin.replace(".", "").replace(",", ".")
    elif metin.count(".") == 1 and len(metin.split(".")[1]) == 3:
        # '1.250' gibi: nokta binlik ayırıcı, ondalık kısım yok
        metin = metin.replace(".", "")

    return float(metin)


def tarih_temizle(deger):
    """Farklı formatlardaki tarihleri tek tip tarihe çevirir."""
    if pd.isna(deger):
        return None
    metin = str(deger).strip()
    for format_ in TARIH_FORMATLARI:
        try:
            return datetime.strptime(metin, format_).date()
        except ValueError:
            continue
    print(f"  Uyarı: tarih okunamadı -> {metin!r}")
    return None


def sehir_temizle(deger):
    """Şehir yazımlarını tek bir doğru ada indirir."""
    if pd.isna(deger):
        return None
    anahtar = str(deger).strip().lower()
    if anahtar not in SEHIR_ESLEME:
        print(f"  Uyarı: bilinmeyen şehir -> {deger!r}")
        return str(deger).strip()
    return SEHIR_ESLEME[anahtar]


def temizle(df):
    df = df.copy()
    df["Müşteri"] = df["Müşteri"].astype(str).str.strip()
    df["Şehir"] = df["Şehir"].apply(sehir_temizle)
    df["Tarih"] = pd.to_datetime(df["Tarih"].apply(tarih_temizle))
    df["Tutar"] = df["Tutar"].apply(tutar_temizle)
    df["Adet"] = pd.to_numeric(df["Adet"])
    df["Birim Fiyat"] = pd.to_numeric(df["Birim Fiyat"])

    # Aynı sipariş numarası birden fazla kez girilmişse ilkini tut
    df = df.drop_duplicates(subset="Sipariş No", keep="first")
    return df.sort_values("Tarih").reset_index(drop=True)


def ozet_tablolari(df):
    """Rapordaki özet tablolarını hazırlar. Her tablo (başlık, DataFrame) çiftidir."""
    genel = pd.DataFrame({
        "Bilgi": ["Toplam ciro (TL)", "Sipariş sayısı", "Ortalama sipariş (TL)",
                  "İlk tarih", "Son tarih"],
        "Değer": [df["Tutar"].sum(), len(df), round(df["Tutar"].mean(), 2),
                  df["Tarih"].min().date(), df["Tarih"].max().date()],
    })

    musteri = (df.groupby("Müşteri")
                .agg(Sipariş_Sayısı=("Müşteri", "count"), Ciro = ("Tutar", "sum"))  
                .sort_values("Sipariş_Sayısı", ascending = False)
                .reset_index())

    sehir = (df.groupby("Şehir")
               .agg(Sipariş=("Sipariş No", "count"), Ciro=("Tutar", "sum"))
               .sort_values("Ciro", ascending=False)
               .reset_index())

    urun = (df.groupby("Ürün")
              .agg(Adet=("Adet", "sum"), Ciro=("Tutar", "sum"))
              .sort_values("Ciro", ascending=False)
              .reset_index())

    aylik = (df.groupby(df["Tarih"].dt.strftime("%Y-%m"))
               .agg(Sipariş=("Sipariş No", "count"), Ciro=("Tutar", "sum"))
               .reset_index()
               .rename(columns={"Tarih": "Ay"}))

    return [
        ("Genel Özet", genel),
        ("En çok sipariş veren ilk 3 müşteri", musteri.head(3)),
        ("Şehirlere Göre", sehir),
        ("Ürünlere Göre", urun),
        ("Aylara Göre", aylik),
    ]


def rapor_yaz(df, tablolar, cikti_yolu):
    baslik_font = Font(bold=True, color="FFFFFF")
    baslik_dolgu = PatternFill("solid", fgColor="305496")
    bolum_font = Font(bold=True, size=13)

    with pd.ExcelWriter(cikti_yolu, engine="openpyxl") as yazici:
        # 1) Özet sayfası: tabloları alt alta yaz
        satir = 0
        konumlar = []
        for baslik, tablo in tablolar:
            tablo.to_excel(yazici, sheet_name="Özet", startrow=satir + 1, index=False)
            konumlar.append((satir, baslik, tablo))
            satir += len(tablo) + 4

        # 2) Temiz veri sayfası
        df_yaz = df.copy()
        df_yaz["Tarih"] = df_yaz["Tarih"].dt.date
        df_yaz.to_excel(yazici, sheet_name="Temiz Veri", index=False)

        # Biçimlendirme
        ozet = yazici.sheets["Özet"]
        for satir, baslik, tablo in konumlar:
            ozet.cell(row=satir + 1, column=1, value=baslik).font = bolum_font
            for sutun in range(1, len(tablo.columns) + 1):
                hucre = ozet.cell(row=satir + 2, column=sutun)
                hucre.font = baslik_font
                hucre.fill = baslik_dolgu

        veri = yazici.sheets["Temiz Veri"]
        for hucre in veri[1]:
            hucre.font = baslik_font
            hucre.fill = baslik_dolgu
        veri.freeze_panes = "A2"

        for sayfa in (ozet, veri):
            for sutun in sayfa.columns:
                genislik = max(len(str(h.value)) if h.value is not None else 0 for h in sutun)
                sayfa.column_dimensions[get_column_letter(sutun[0].column)].width = genislik + 3
            for satir_hucreleri in sayfa.iter_rows(min_row=2):
                for hucre in satir_hucreleri:
                    if isinstance(hucre.value, float):
                        hucre.number_format = "#,##0.00"
                    elif hucre.is_date:
                        hucre.number_format = "DD.MM.YYYY"


def main():
    if len(sys.argv) < 2:
        print("Kullanım: python temizle.py <girdi.csv|xlsx> [cikti.xlsx]")
        sys.exit(1)

    girdi = sys.argv[1]
    cikti = sys.argv[2] if len(sys.argv) > 2 else "temiz_rapor.xlsx"

    ham = dosya_oku(girdi)
    print(f"Okundu: {len(ham)} satır, {ham['Şehir'].nunique()} farklı şehir yazımı")

    df = temizle(ham)
    print(f"Temizlendi: {len(df)} satır, {df['Şehir'].nunique()} şehir")

    # Kontrol: yazılan tutar, adet x birim fiyat ile tutuyor mu?
    hesaplanan = (df["Adet"] * df["Birim Fiyat"]).sum()
    toplam = df["Tutar"].sum()
    durum = "TUTUYOR" if abs(hesaplanan - toplam) < 0.01 else "TUTMUYOR!"
    print(f"Toplam ciro: {toplam:,.2f} TL (adet x fiyat: {hesaplanan:,.2f} TL) -> {durum}")

    rapor_yaz(df, ozet_tablolari(df), cikti)
    print(f"Rapor yazıldı: {cikti}")


if __name__ == "__main__":
    main()
