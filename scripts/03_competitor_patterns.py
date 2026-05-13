"""
Rakip URL pattern analizi
Vodafone, Turk Telekom ve diger rakiplerin hangi URL pattern'leri ile trafik aldigini cikar
"""
import pandas as pd
import re
from collections import Counter

CSV_PATH = "/Users/Erdo/Desktop/Claude Projects/Turkcel/www.turkcell.com.tr-content-gap-subdomains-t_2026-05-13_14-14-31.csv"
df = pd.read_csv(CSV_PATH, low_memory=False)

competitors = ["turktelekom.com.tr", "vodafone.com.tr", "mediamarkt.com.tr",
               "vatanbilgisayar.com", "turk.net", "pttcell.com.tr"]

# Turkcell siralamiyor
df = df[df["www.turkcell.com.tr/: URL"].isna()].copy()

def extract_path_prefix(url, depth=2):
    """ /a/b/c/d -> /a/b """
    if pd.isna(url):
        return None
    m = re.match(r"https?://[^/]+(/[^?#]*)", str(url))
    if not m:
        return None
    path = m.group(1)
    parts = [p for p in path.split("/") if p]
    if not parts:
        return "/"
    return "/" + "/".join(parts[:depth])

# Tum rakipler icin URL prefix'lerini ve toplam trafigi cikar
for comp in competitors:
    url_col = f"{comp}/: URL"
    traf_col = f"{comp}/: Organic Traffic"
    pos_col = f"{comp}/: Organic Position"

    sub = df[df[url_col].notna() & df[traf_col].notna()].copy()
    sub = sub[sub["Volume"] >= 200]  # En azindan kayda deger hacim
    sub["path_prefix"] = sub[url_col].apply(lambda x: extract_path_prefix(x, 2))

    prefix_stats = sub.groupby("path_prefix").agg(
        keyword_sayisi=("Keyword", "count"),
        toplam_trafik=(traf_col, "sum"),
        toplam_hacim=("Volume", "sum"),
    ).sort_values("toplam_trafik", ascending=False).head(20)

    print(f"\n{'='*80}")
    print(f"## {comp} - En cok trafik aldigi URL path prefix'leri")
    print(f"{'='*80}")
    print(prefix_stats.to_string())
