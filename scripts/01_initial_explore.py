"""
Turkcell Content Gap CSV - Initial Exploration
Goal: Understand the structure, find keywords where Turkcell has no rank
"""
import pandas as pd
import re

CSV_PATH = "/Users/Erdo/Desktop/Claude Projects/Turkcel/www.turkcell.com.tr-content-gap-subdomains-t_2026-05-13_14-14-31.csv"

# Load with focus on key columns
df = pd.read_csv(CSV_PATH, low_memory=False)
print(f"Toplam satir: {len(df)}")
print(f"Toplam kolon: {len(df.columns)}")
print("\nKolon adlari:")
for c in df.columns:
    print(f"  - {c}")

# Identify the target and competitor columns
target_url_col = "www.turkcell.com.tr/: URL"
target_pos_col = "www.turkcell.com.tr/: Organic Position"
competitors = ["turktelekom.com.tr", "vodafone.com.tr", "mediamarkt.com.tr",
               "vatanbilgisayar.com", "turk.net", "pttcell.com.tr"]

# Filter: Turkcell has no organic position (NaN)
no_rank_mask = df[target_pos_col].isna()
turkcell_no_rank = df[no_rank_mask].copy()
print(f"\nTurkcell'in siralamasi olmayan keyword sayisi: {len(turkcell_no_rank)}")

# Of these, how many have at least one competitor ranking?
def has_competitor_rank(row):
    for comp in competitors:
        if pd.notna(row.get(f"{comp}/: Organic Position", None)):
            return True
    return False

turkcell_no_rank["has_competitor"] = turkcell_no_rank.apply(has_competitor_rank, axis=1)
opportunities = turkcell_no_rank[turkcell_no_rank["has_competitor"]].copy()
print(f"Turkcell yok ama rakip(ler) var: {len(opportunities)}")

# Volume distribution
print(f"\nHacim dagilimi (opportunities):")
print(f"  Volume >= 10000: {len(opportunities[opportunities['Volume'] >= 10000])}")
print(f"  Volume >= 5000:  {len(opportunities[opportunities['Volume'] >= 5000])}")
print(f"  Volume >= 1000:  {len(opportunities[opportunities['Volume'] >= 1000])}")
print(f"  Volume >= 500:   {len(opportunities[opportunities['Volume'] >= 500])}")
print(f"  Volume >= 100:   {len(opportunities[opportunities['Volume'] >= 100])}")

# KD (Keyword Difficulty) distribution for opportunities with Volume >= 500
mid = opportunities[opportunities['Volume'] >= 500].copy()
print(f"\nKD dagilimi (Volume >= 500):")
print(f"  KD <= 10: {len(mid[mid['KD'] <= 10])}")
print(f"  KD <= 20: {len(mid[mid['KD'] <= 20])}")
print(f"  KD <= 30: {len(mid[mid['KD'] <= 30])}")
print(f"  KD <= 50: {len(mid[mid['KD'] <= 50])}")
