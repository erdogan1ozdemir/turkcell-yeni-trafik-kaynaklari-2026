"""
Daha rafine cluster'lar - 'Other' kategorisindeki gizli firsatlari bul
Telekom firmasi icin uygun: hesaplama, tatil, dijital servis, bilgi sayfalari
"""
import pandas as pd
import re

CSV_PATH = "/Users/Erdo/Desktop/Claude Projects/Turkcel/www.turkcell.com.tr-content-gap-subdomains-t_2026-05-13_14-14-31.csv"
OUTPUT_DIR = "/Users/Erdo/Desktop/Claude Projects/Turkcel/output"

df = pd.read_csv(CSV_PATH, low_memory=False)
competitors = ["turktelekom.com.tr", "vodafone.com.tr", "mediamarkt.com.tr",
               "vatanbilgisayar.com", "turk.net", "pttcell.com.tr"]

df = df[df["www.turkcell.com.tr/: URL"].isna()].copy()

def best_comp_info(row):
    best_pos = 999
    best_comp = None
    best_url = None
    best_traf = 0
    for comp in competitors:
        pos = row.get(f"{comp}/: Organic Position", None)
        if pd.notna(pos) and pos < best_pos:
            best_pos = pos
            best_comp = comp
            best_url = row.get(f"{comp}/: URL", None)
            best_traf = row.get(f"{comp}/: Organic Traffic", 0)
    return pd.Series([best_comp, best_pos if best_pos < 999 else None, best_url, best_traf if pd.notna(best_traf) else 0])

df[["best_competitor", "best_position", "best_url", "best_traffic"]] = df.apply(best_comp_info, axis=1)
df = df[df["best_competitor"].notna()].copy()

# Genis hacim havuzu - Volume 200+ (cluster'lar icin)
pool = df[df["Volume"] >= 200].copy()
print(f"Volume>=200 firsat havuzu: {len(pool)}")

# Rafine cluster'lar - amac telekom firmasinin acabilecegi sayfa tiplerini bulmak
CLUSTERS = {
    "01_MAAS_HESAPLAMA": {
        "include": [r"\b(brüt|brut|net|maaş|maas|asgari ücret|asgari ucret)\b",
                    r"\b(yıllık|aylık|yillik|aylik) maaş",
                    r"\bnet brüt\b", r"\bbrüt net\b",
                    r"\bbordro\b", r"\bsgk\b", r"\bemekli\b", r"\bişe giriş\b"],
        "exclude": [r"\bbayrami\b"],
    },
    "02_HESAPLAMA_ARACLARI": {
        "include": [r"\bhesapla(ma)?\b", r"\bhesap makinesi\b", r"\bhesab[ıi]\b",
                    r"\byüzde\b", r"\byuzde\b", r"\byaş hesap", r"\byas hesap",
                    r"\bkdv\b", r"\bötv\b", r"\botv\b", r"\bvergi\b",
                    r"\bnot ortalama\b", r"\bkredi hesap", r"\btaksit hesap",
                    r"\biban\b", r"\bbenzin hesap", r"\bmotorin hesap",
                    r"\bgöğüs hesap", r"\byükselen burç hesap"],
        "exclude": [r"\binstagram hesab", r"\bpubg hesab", r"\bgmail hesab",
                    r"\bgoogle hesab", r"\bsteam hesab", r"\btwitter hesab"],
    },
    "03_DONUSTURUCU_BIRIM": {
        "include": [r"\b(inc|inç)\b", r"\b(metre|cm|mm|km|gram|kg|miligram|ton|litre|ml|gallon|inch)\b",
                    r"\bkac (saat|dakika|saniye|kg|gram|metre|cm|gb|mb|tl|dolar|euro|km)\b",
                    r"\bçevirme\b", r"\bcevirme\b", r"\bdönüştürücü\b", r"\bdonusturucu\b",
                    r"\bbirim\b"],
        "exclude": [r"\bcumhuriyet bayram\b"],
    },
    "04_DOVIZ_FINANS": {
        "include": [r"\bdolar\b", r"\beuro\b", r"\bsterlin\b", r"\bdöviz\b", r"\bdoviz\b",
                    r"\baltın\b", r"\baltin\b", r"\bborsa\b", r"\bfaiz oran",
                    r"\bbist\b", r"\bbankamatik\b"],
        "exclude": [],
    },
    "05_RESMI_TATIL_BAYRAM": {
        "include": [r"\bresmî tatil\b", r"\bresmi tatil\b", r"\btatil mi\b",
                    r"\bbayram\b", r"\bkurban\b", r"\bramazan\b", r"\bayrefe\b", r"\barefe\b",
                    r"\byılbaş[ıi]\b", r"\byilbasi\b",
                    r"\bcumhuriyet bayram\b", r"\b29 ekim\b", r"\b23 nisan\b",
                    r"\b19 may[ıi]s\b", r"\b30 ağustos\b", r"\b30 agustos\b",
                    r"\bsömestr\b", r"\bsomestr\b", r"\byarı yıl\b", r"\byari yil\b",
                    r"\b15 tatil\b"],
        "exclude": [r"\byunan adaları\b"],
    },
    "06_OZEL_GUNLER": {
        "include": [r"\banneler g[uü]n[uü]\b", r"\bbabalar g[uü]n[uü]\b",
                    r"\bsevgililer g[uü]n[uü]\b", r"\b14 şubat\b", r"\b14 subat\b",
                    r"\böğretmenler g[uü]n[uü]\b", r"\bogretmenler g[uü]n[uü]\b",
                    r"\bçocuklar g[uü]n[uü]\b", r"\bcocuklar g[uü]n[uü]\b",
                    r"\bkad[ıi]nlar g[uü]n[uü]\b", r"\bkadinlar g[uü]n[uü]\b",
                    r"\bdoğum g[uü]n[uü]\b", r"\bdogum g[uü]n[uü]\b",
                    r"\bnewroz\b", r"\bnevruz\b",
                    r"\b8 mart\b", r"\bdoğum tarihi\b", r"\bdogum tarihi\b"],
        "exclude": [],
    },
    "07_MESAJ_SOZ_SIIR_KART": {
        "include": [r"\b(mesaj|söz|sozler|şiir|siir|yazi|notu|kart) ?(ı|i|ları|leri)?\b",
                    r"\b(en güzel|guzel|duygusal|romantik|kısa|uzun|anlamlı|anlamli)\b.*\b(mesaj|söz|sözleri|şiir|sözler)\b"],
        "exclude": [r"\bcarrefour\b", r"\bsd kart\b", r"\bbankkart\b", r"\bkart okuyucu\b",
                    r"\bparaf kart\b", r"\bsim kart bloke\b", r"\bwhatsapp.*mesaj\b",
                    r"\binstagram.*mesaj\b", r"\bsilinen mesaj\b", r"\bmesaj gelmiyor\b",
                    r"\bbiletix\b"],
    },
    "08_NE_ZAMAN_TARIH": {
        "include": [r"\bne zaman\b", r"\bhangi g[uü]n\b", r"\bhangi tarih\b",
                    r"\bhangi ay\b", r"\bkaçında\b", r"\bkacinda\b",
                    r"\bne zamand[ıi]\b", r"\bgeri sayım\b", r"\bgeri sayim\b",
                    r"\bkaç gün kaldı\b", r"\bkac gun kaldi\b"],
        "exclude": [],
    },
    "09_NEDIR_TANIM": {
        "include": [r"\bnedir\b", r"\bne demek\b", r"\bne anlama gelir\b",
                    r"\bne işe yarar\b", r"\bne ise yarar\b"],
        "exclude": [],
    },
    "10_NASIL_YAPILIR": {
        "include": [r"\bnas[ıi]l\b"],
        "exclude": [r"\bnas[ıi]l biri\b", r"\bnas[ıi]l adam\b"],
    },
    "11_KAC_INC_EKRAN": {
        "include": [r"\bkaç (inç|cm|ekran|gb|mb)\b", r"\b(inç|cm|ekran) kaç\b",
                    r"\b(ekran|televizyon|tv) (boyutu|olcusu)\b"],
        "exclude": [],
    },
    "12_WHATSAPP_INSTAGRAM_REHBER": {
        "include": [r"\bwhatsapp\b", r"\binstagram\b", r"\btiktok\b", r"\bfacebook\b",
                    r"\bsnapchat\b", r"\btwitter\b", r"\b\bx giriş\b", r"\btelegram\b"],
        "exclude": [r"\bnedir\b"],
    },
    "13_TELEFON_AYAR_SORUN": {
        "include": [r"\biphone (\d+|\w+) ne|kaç|nasıl\b",
                    r"\bgüvenli arama\b", r"\bnfc\b", r"\bairdrop\b",
                    r"\bairtag\b", r"\bekran goruntusu\b", r"\bekran görüntüsü\b",
                    r"\bekran kayd[ıi]\b", r"\btelefon\b", r"\bsim kart\b",
                    r"\besim\b", r"\bgsm\b", r"\bnumara öğrenme\b", r"\bnumara ogrenme\b"],
        "exclude": [r"\bmasaj salonu\b", r"\bsofa\b"],
    },
    "14_DIZI_FILM_REHBER": {
        "include": [r"\b(dizi|film|sezon|bolum|bölüm) (konusu|oyuncuları|oyuncular)\b",
                    r"\b(oyuncular[ıi]|oyuncular)\b",
                    r"\bnetflix\b", r"\bdisney\+ \b", r"\bexxen\b", r"\btabii\b", r"\bblutv\b",
                    r"\bmaç (hangi|nerede|saat)\b"],
        "exclude": [r"\bmaç tipi\b", r"\bkahve makinesi\b"],
    },
    "15_OYUN_REHBER": {
        "include": [r"\b(gta|pubg|lol|valorant|fortnite|minecraft|fifa|playstation|ps5|ps6|xbox|nintendo|steam|monopoly|wordle)\b"],
        "exclude": [],
    },
    "16_DEPREM_AFET_GUVENLIK": {
        "include": [r"\bdeprem\b", r"\bafad\b", r"\btsunami\b", r"\bafet\b",
                    r"\bguvenlik\b", r"\bgüvenlik\b"],
        "exclude": [r"\bgüvenli arama\b"],
    },
    "17_KIMLIK_EDEVLET_NUMARA": {
        "include": [r"\be-devlet\b", r"\bedevlet\b", r"\btc kimlik\b",
                    r"\bkimlik no\b", r"\bpasaport\b", r"\bvergi numaras\b",
                    r"\bbarkod sorgulama\b", r"\biban sorgu\b", r"\bborç sorgu\b",
                    r"\bplaka sorgu\b", r"\bmuayene\b", r"\btrafik cezas\b",
                    r"\bsgk\b", r"\bn[uü]fus\b"],
        "exclude": [],
    },
    "18_INTERNET_HIZ_MODEM": {
        "include": [r"\binternet h[ıi]z\b", r"\bh[ıi]z testi\b", r"\bspeed test\b",
                    r"\bmodem\b", r"\bwifi\b", r"\brouter\b", r"\bping\b",
                    r"\b192\.168\.\b"],
        "exclude": [],
    },
    "19_POSTA_KODU_ALAN": {
        "include": [r"\bposta kodu\b", r"\balan kodu\b", r"\bil kodu\b",
                    r"\bplaka kodu\b", r"\btelefon kodu\b"],
        "exclude": [],
    },
    "20_SAGLIK_VUCUT": {
        "include": [r"\bbmi\b", r"\bvücut\b", r"\bvucut\b", r"\bkalori\b",
                    r"\bhamilelik\b", r"\bgebelik\b", r"\bdoğum tarihi hesap\b",
                    r"\bregl\b", r"\başı\b", r"\basi takvim\b", r"\bilac\b",
                    r"\bilaç\b", r"\bsemptom\b", r"\bbelirti\b", r"\bhastalik\b",
                    r"\bhastalık\b"],
        "exclude": [],
    },
    "21_KELIME_DIL_YAZIM": {
        "include": [r"\bçeviri\b", r"\bceviri\b", r"\bdil\b", r"\bingilizce\b", r"\btürkçe\b",
                    r"\bnoktal[ıi] virgül\b", r"\bbüyük harf\b", r"\beşanlamlı\b", r"\bzit anlamli\b",
                    r"\bkelime say\b", r"\bkarakter say\b", r"\byaz[ıi]m\b"],
        "exclude": [r"\bturkce dublaj\b"],
    },
    "22_OTOMOTIV_TRAFIK": {
        "include": [r"\bplaka\b", r"\bmuayene\b", r"\btrafik\b", r"\baraç vergisi\b",
                    r"\barac vergisi\b", r"\botv\b", r"\bbenzin (fiyat|hesap)\b",
                    r"\bmotorin (fiyat|hesap)\b", r"\bsürücü kurs\b"],
        "exclude": [],
    },
}

def cluster_match(keyword):
    kw = str(keyword).lower()
    for cname, rules in CLUSTERS.items():
        # Exclude check
        if any(re.search(p, kw) for p in rules["exclude"]):
            continue
        # Include check
        if any(re.search(p, kw) for p in rules["include"]):
            return cname
    return "ZZ_DIGER"

pool["cluster"] = pool["Keyword"].apply(cluster_match)

# Cluster ozeti
print("\n=== CLUSTER OZETI (Volume >= 200) ===")
summary = pool.groupby("cluster").agg(
    keyword_sayisi=("Keyword", "count"),
    toplam_hacim=("Volume", "sum"),
    ort_hacim=("Volume", "mean"),
    ort_kd=("KD", "mean"),
    max_hacim=("Volume", "max"),
).sort_values("toplam_hacim", ascending=False)
summary["ort_hacim"] = summary["ort_hacim"].astype(int)
summary["toplam_hacim"] = summary["toplam_hacim"].astype(int)
summary["max_hacim"] = summary["max_hacim"].astype(int)
print(summary.to_string())

# Cluster basina top 50 keyword (Diger haric)
pool_sorted = pool.sort_values("Volume", ascending=False)
for cname in CLUSTERS.keys():
    sub = pool_sorted[pool_sorted["cluster"] == cname].head(50)
    if len(sub) > 0:
        sub[["Keyword", "Volume", "KD", "SERP features",
             "best_competitor", "best_position", "best_url", "best_traffic"]].to_csv(
            f"{OUTPUT_DIR}/cluster_{cname}.csv", index=False)

# Diger - top 200
diger = pool_sorted[pool_sorted["cluster"] == "ZZ_DIGER"].head(200)
diger[["Keyword", "Volume", "KD", "best_competitor", "best_position", "best_url", "best_traffic"]].to_csv(
    f"{OUTPUT_DIR}/cluster_ZZ_DIGER_top200.csv", index=False)

# Tum havuzu da kaydet (Excel icin)
pool[["Keyword", "Volume", "KD", "CPC", "SERP features", "cluster",
      "best_competitor", "best_position", "best_url", "best_traffic"]].to_csv(
    f"{OUTPUT_DIR}/all_opportunities_v200.csv", index=False)

print(f"\nToplam cluster sayisi (Diger haric): {len([c for c in pool['cluster'].unique() if c != 'ZZ_DIGER'])}")
print(f"Diger: {(pool['cluster'] == 'ZZ_DIGER').sum()} keyword")
