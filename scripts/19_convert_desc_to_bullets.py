"""
Tum <p class="idea-desc">...</p> paragraflarini cumle bazli bullet list'e cevir.
2+ cumle olanlar bullet, tek cumle olanlar paragraf kalir.
"""
import re
from pathlib import Path

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")
orig_len = len(text)

# Pattern: <p class="idea-desc">CONTENT</p> (multiline)
pattern = re.compile(r'<p class="idea-desc">(.+?)</p>', re.DOTALL)

# Turkce buyuk harf seti
SENT_SPLIT = re.compile(r'(?<=[.!?])\s+(?=[A-ZÇĞİÖŞÜ])')

def to_bullets(match):
    content = match.group(1).strip()
    # Tum whitespace'leri normalize et
    normalized = re.sub(r'\s+', ' ', content)
    # Cumlelere bol
    parts = SENT_SPLIT.split(normalized)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) < 2:
        # Tek cumle, paragraf olarak kalsin
        return '<p class="idea-desc">' + normalized + '</p>'

    # Cok kisa "cumle" parcalarini bir oncekine birlestir (yanlis bolme)
    merged = []
    for p in parts:
        # Sadece kisaltma gibi gorunen tek harf vs.
        if merged and len(p.split()) < 3:
            merged[-1] = merged[-1] + ' ' + p
        else:
            merged.append(p)

    if len(merged) < 2:
        return '<p class="idea-desc">' + normalized + '</p>'

    lis = '\n'.join('          <li>' + p + '</li>' for p in merged)
    return '<ul class="idea-bullets">\n' + lis + '\n        </ul>'

new_text, count = pattern.subn(to_bullets, text)

# Sayim
ul_count = new_text.count('<ul class="idea-bullets">') - text.count('<ul class="idea-bullets">')
p_remain = new_text.count('<p class="idea-desc">')

HTML.write_text(new_text, encoding="utf-8")
print(f"Toplam islem: {count}")
print(f"Bullet list'e cevrilen: {ul_count}")
print(f"Paragraf kalan (tek cumle): {p_remain}")
print(f"HTML size delta: {len(new_text)-orig_len:+,} chars")
