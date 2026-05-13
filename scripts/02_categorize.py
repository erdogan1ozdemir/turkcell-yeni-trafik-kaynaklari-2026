"""
Turkcell Content Gap - Kategori Tabanli Firsat Analizi
Hedef: Turkcell yok, rakip(ler) var olan yuksek hacimli keyword'leri evergreen kategorilere ayir
"""
import pandas as pd
import re
import json

CSV_PATH = "/Users/Erdo/Desktop/Claude Projects/Turkcel/www.turkcell.com.tr-content-gap-subdomains-t_2026-05-13_14-14-31.csv"
OUTPUT_DIR = "/Users/Erdo/Desktop/Claude Projects/Turkcel/output"

df = pd.read_csv(CSV_PATH, low_memory=False)

competitors = ["turktelekom.com.tr", "vodafone.com.tr", "mediamarkt.com.tr",
               "vatanbilgisayar.com", "turk.net", "pttcell.com.tr"]

# Turkcell ranking yok
df = df[df["www.turkcell.com.tr/: URL"].isna()].copy()

# En az bir rakip rank ediyor + en iyi rakip pozisyonunu bul
def best_competitor(row):
    best_pos = 999
    best_comp = None
    best_traffic = 0
    for comp in competitors:
        pos = row.get(f"{comp}/: Organic Position", None)
        traf = row.get(f"{comp}/: Organic Traffic", 0)
        if pd.notna(pos) and pos < best_pos:
            best_pos = pos
            best_comp = comp
            best_traffic = traf if pd.notna(traf) else 0
    return pd.Series([best_comp, best_pos if best_pos < 999 else None, best_traffic])

df[["best_competitor", "best_position", "best_traffic"]] = df.apply(best_competitor, axis=1)
df = df[df["best_competitor"].notna()].copy()
print(f"Rakipli firsat sayisi: {len(df)}")

# Anahtar kelime tabanli kategori atama
# Evergreen kategoriler:
CATEGORIES = {
    "Hesaplama_Calculator": [
        r"\bhesapla\b", r"\bhesaplama\b", r"\bhesab[ıi]\b", r"\bhesap kitap\b",
    ],
    "Donusturucu_Converter": [
        r"\bdonustur", r"\bçevir", r"\bcevir\b", r"\bkac\s+(saat|dakika|kilogram|kg|gram|metre|cm|inc|gb|mb|tl|dolar|euro)",
        r"\b(saat|dakika|kg|gram|metre|cm|gb|mb|tl|dolar|euro)\s+kac\b",
    ],
    "Ne_Zaman_When": [
        r"\bne zaman\b",
    ],
    "Nedir_WhatIs": [
        r"\bnedir\b", r"\bne demek\b", r"\bne anlama gelir\b",
    ],
    "Nasil_HowTo": [
        r"\bnas[ıi]l\b",
    ],
    "Kac_HowMany": [
        r"\bkaç\b", r"\bkac\b",
    ],
    "Ozel_Gunler_SpecialDays": [
        r"\banneler g[uü]n[uü]\b", r"\bbabalar g[uü]n[uü]\b", r"\bsevgililer g[uü]n[uü]\b",
        r"\böğretmenler g[uü]n[uü]\b", r"\boğretmenler g[uü]n[uü]\b",
        r"\bçocuk(lar)? g[uü]n[uü]\b", r"\bcocuk(lar)? g[uü]n[uü]\b",
        r"\bcumhuriyet bayram[ıi]\b", r"\b29 ekim\b", r"\b23 nisan\b",
        r"\b19 may[ıi]s\b", r"\b30 ağustos\b", r"\b30 agustos\b",
        r"\b14 şubat\b", r"\b14 subat\b", r"\b8 mart\b",
        r"\bkad[ıi]nlar g[uü]n[uü]\b", r"\bkadinlar g[uü]n[uü]\b",
        r"\bdoğum g[uü]n[uü]\b", r"\bdogum g[uü]n[uü]\b",
        r"\byeni y[ıi]l\b", r"\by[ıi]lbaş[ıi]\b", r"\byilbasi\b",
        r"\bnewroz\b", r"\bnevruz\b", r"\b30 nisan\b",
    ],
    "Tatiller_Bayramlar": [
        r"\btatil\b", r"\bbayram\b", r"\bramazan\b", r"\boruç\b", r"\boruc\b",
        r"\bkurban\b", r"\bramazan bayram[ıi]\b", r"\bkurban bayram[ıi]\b",
        r"\bmevlit\b", r"\bkandil\b", r"\barefe\b",
    ],
    "Mesaj_Sozler_Kart": [
        r"\bmesaj(lar[ıi])?\b", r"\bsöz(ler[ıi])?\b", r"\bsozler[ıi]?\b",
        r"\bkart(lar[ıi])?\b", r"\bşiir\b", r"\bsiir\b",
    ],
    "Saat_Tarih_Time": [
        r"\bbug[uü]n\b", r"\bbugun\b", r"\byar[ıi]n\b", r"\bsaat kaç\b",
        r"\bsaat fark[ıi]\b", r"\bzaman fark[ıi]\b", r"\bnamaz vakit\b",
        r"\biftar\b", r"\bsahur\b", r"\bezan\b", r"\bgüneş\b",
    ],
    "Burc_Astroloji": [
        r"\bbur(c|ç)\b", r"\bastroloji\b", r"\by[ıi]ld[ıi]z\b",
        r"\bkoc bur(c|ç)\b", r"\bboğa bur(c|ç)\b", r"\biki(z|zler) bur(c|ç)\b",
        r"\byengec bur(c|ç)\b", r"\baslan bur(c|ç)\b", r"\bbaşak bur(c|ç)\b",
        r"\bterazi bur(c|ç)\b", r"\bakrep bur(c|ç)\b", r"\byay bur(c|ç)\b",
        r"\boğlak bur(c|ç)\b", r"\bkova bur(c|ç)\b", r"\bbal[ıi]k bur(c|ç)\b",
        r"\byükselen burç\b", r"\byükselen burc\b",
    ],
    "Saglik_Health": [
        r"\bbelirti(ler)?\b", r"\bhastal[ıi]k\b", r"\btedavi\b", r"\bilac\b", r"\bilaç\b",
        r"\bhamile\b", r"\bgebelik\b", r"\bdoğum\b", r"\bbebek\b",
        r"\bagri\b", r"\bağr[ıi]\b", r"\battak\b",
    ],
    "Yemek_Tarif": [
        r"\btarif(i)?\b", r"\byemek\b", r"\bnas[ıi]l yap[ıi]l[ıi]r\b", r"\bnasil yapilir\b",
        r"\b(corba|çorba|salata|tatl[ıi]|kek|pasta|börek|borek|pilav|et|tavuk|bal[ıi]k|sebze)\b",
    ],
    "Ruya_Tabiri": [
        r"\bruya\b", r"\brüya\b", r"\btabir\b", r"\bgörmek\b", r"\bgormek\b",
    ],
    "Kelime_Sozluk": [
        r"\bne demek\b", r"\bes anlaml[ıi]\b", r"\besş anlaml[ıi]\b", r"\bzit anlamlisi\b",
        r"\bnoktal[ıi] virgül\b", r"\bbüyük harf\b", r"\bküçük harf\b",
        r"\bingilizce(ce|si)?\b", r"\bturkçe(ce|si)?\b",
    ],
    "Maas_Maddi_Hesap": [
        r"\bmaaş\b", r"\bmaas\b", r"\bmaaş hesapla\b", r"\bvergi\b", r"\bkdv\b", r"\bötv\b", r"\botv\b",
        r"\bsgk\b", r"\bemekli\b", r"\bborc sorgula\b",
        r"\baidat\b", r"\bdamga\b", r"\bgelir vergisi\b",
    ],
    "Kredi_Bankacilik": [
        r"\bkredi\b", r"\btaksit\b", r"\bfaiz\b", r"\biban\b", r"\bdöviz\b", r"\bdoviz\b",
        r"\baltin\b", r"\balt[ıi]n\b", r"\bborsa\b",
    ],
    "Telekom_Tarife": [
        r"\btarife\b", r"\bpaket\b", r"\bsanal hat\b", r"\befatur\b", r"\bnumara taş[ıi]\b",
        r"\binternet h[ıi]z[ıi] testi\b", r"\bh[ıi]z testi\b", r"\bsim\b", r"\bcep telefon\b",
        r"\bfatura sorgu\b", r"\bbakiye\b", r"\btl yükle\b",
    ],
    "Egitim_Sinav": [
        r"\bsinav\b", r"\bs[ıi]nav\b", r"\byks\b", r"\btyt\b", r"\bayt\b", r"\blgs\b",
        r"\bkpss\b", r"\baleus\b", r"\bösym\b", r"\bosym\b", r"\bders\b",
        r"\bsoru\b", r"\bçözüm\b", r"\bcozum\b", r"\bkonu anlat[ıi]m[ıi]\b",
    ],
    "Spor": [
        r"\bmaç\b", r"\bmac\b", r"\bfutbol\b", r"\bbasketbol\b", r"\bvoleybol\b",
        r"\bfenerbahce\b", r"\bgalatasaray\b", r"\bbeşiktaş\b", r"\bbesiktas\b",
        r"\bskor\b", r"\bfik(s)?tür\b", r"\bpuan(durumu)?\b",
    ],
    "Otomotiv": [
        r"\baraba\b", r"\botomobil\b", r"\baraç\b", r"\barac\b", r"\bmuayene\b",
        r"\btrafik\b", r"\bsürücü\b", r"\bsurucu\b", r"\bplaka\b",
        r"\bbenzin\b", r"\bmotorin\b", r"\botv\b",
    ],
}

def categorize(keyword):
    kw_low = str(keyword).lower()
    matches = []
    for cat, patterns in CATEGORIES.items():
        for p in patterns:
            if re.search(p, kw_low):
                matches.append(cat)
                break
    return matches

df["categories"] = df["Keyword"].apply(categorize)
df["primary_category"] = df["categories"].apply(lambda x: x[0] if x else "Other")
df["has_category"] = df["categories"].apply(lambda x: len(x) > 0)

# Yuksek hacim filtresi
mid = df[df["Volume"] >= 500].copy()
print(f"\nVolume >= 500 firsat: {len(mid)}")
print(f"Kategorize edilen: {mid['has_category'].sum()}")
print(f"Other (sinif disi): {(~mid['has_category']).sum()}")

# Kategori ozeti
cat_summary = []
for cat in CATEGORIES.keys():
    sub = mid[mid["primary_category"] == cat]
    if len(sub) > 0:
        cat_summary.append({
            "Kategori": cat,
            "Keyword_Sayisi": len(sub),
            "Toplam_Aylik_Hacim": int(sub["Volume"].sum()),
            "Ortalama_Hacim": int(sub["Volume"].mean()),
            "Ortalama_KD": round(sub["KD"].mean(), 1) if sub["KD"].notna().any() else None,
            "Max_Hacim": int(sub["Volume"].max()),
            "Ornek_KW_1": sub.iloc[0]["Keyword"] if len(sub) > 0 else "",
            "Ornek_KW_2": sub.iloc[1]["Keyword"] if len(sub) > 1 else "",
            "Ornek_KW_3": sub.iloc[2]["Keyword"] if len(sub) > 2 else "",
        })

cat_df = pd.DataFrame(cat_summary).sort_values("Toplam_Aylik_Hacim", ascending=False)
print("\n=== KATEGORI OZETI (Volume >= 500) ===")
print(cat_df.to_string(index=False))

cat_df.to_csv(f"{OUTPUT_DIR}/kategori_ozeti.csv", index=False)

# Her kategori icin top 20 keyword'u kaydet
mid_sorted = mid.sort_values("Volume", ascending=False)
for cat in CATEGORIES.keys():
    sub = mid_sorted[mid_sorted["primary_category"] == cat].head(30)
    if len(sub) > 0:
        sub[["Keyword", "Volume", "KD", "CPC", "SERP features",
             "best_competitor", "best_position", "best_traffic"]].to_csv(
            f"{OUTPUT_DIR}/cat_{cat}.csv", index=False)

# Other kategori - yuksek hacim ama siniflandirilmamis
other = mid_sorted[mid_sorted["primary_category"] == "Other"].head(100)
other[["Keyword", "Volume", "KD", "best_competitor", "best_position", "best_traffic"]].to_csv(
    f"{OUTPUT_DIR}/cat_Other_top100.csv", index=False)

print(f"\nCikti dosyalari kaydedildi: {OUTPUT_DIR}/")
