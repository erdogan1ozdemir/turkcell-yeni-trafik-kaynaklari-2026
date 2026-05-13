"""
xlsx -> base64 -> HTML embed
HTML icindeki onceki xlsxData script tag'lerini temizler ve yenisini ekler.
Idempotent: kac kez calistirilirsa calistirilsin tek bir embed kalir.
"""
import base64
import re
from pathlib import Path

ROOT = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel")
XLSX = ROOT / "TURKCELL_TRAFIK_FIRSATLARI.xlsx"
HTML = ROOT / "index.html"

# 1) Read and base64 encode
with open(XLSX, "rb") as f:
    raw = f.read()

b64 = base64.b64encode(raw).decode("ascii")
print(f"xlsx size: {len(raw):,} bytes ({len(raw)/1024/1024:.2f} MB)")
print(f"base64 size: {len(b64):,} chars ({len(b64)/1024/1024:.2f} MB)")

# 2) Read HTML
with open(HTML, "r", encoding="utf-8") as f:
    html = f.read()

# 3) Onceki xlsxData script tag'lerini sil (idempotent)
pattern = re.compile(r'\n?<script id="xlsxData">window\.__XLSX_B64__ = "[^"]*";</script>\n?', re.DOTALL)
matches = pattern.findall(html)
print(f"Removing {len(matches)} existing xlsxData script(s)")
html = pattern.sub("", html)

# 4) Yenisini </body>'den once enjekte et
inject = f'\n<script id="xlsxData">window.__XLSX_B64__ = "{b64}";</script>\n</body>'
if "</body>" not in html:
    print("ERROR: </body> not found")
    raise SystemExit(1)
html = html.replace("</body>", inject, 1)
print("Injected fresh xlsxData script.")

# 5) Write back
with open(HTML, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nUpdated: {HTML}")
print(f"New HTML size: {len(html):,} chars ({len(html)/1024/1024:.2f} MB)")
