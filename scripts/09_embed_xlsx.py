"""
xlsx -> base64 -> HTML embed
HTML icindeki XLSX_B64_PLACEHOLDER stringini gercek base64 verisi ile degistirir.
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

# 3) Build replacement marker - if marker exists, replace it
marker = "/*__XLSX_B64_PLACEHOLDER__*/"
new_script = f'window.__XLSX_B64__ = "{b64}";'

if marker in html:
    html = html.replace(marker, new_script)
    print("Replaced existing placeholder.")
else:
    # If first run, inject as a script just before closing </body>
    inject = f'\n<script id="xlsxData">window.__XLSX_B64__ = "{b64}";</script>\n</body>'
    if "</body>" in html:
        html = html.replace("</body>", inject, 1)
        print("Injected xlsxData script before </body>.")
    else:
        print("No </body> found. ERROR.")
        exit(1)

# 4) Write back
with open(HTML, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nUpdated: {HTML}")
print(f"New HTML size: {len(html):,} chars ({len(html)/1024/1024:.2f} MB)")
