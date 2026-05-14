"""
Her <ul class="idea-bullets">'in onune bir toggle butonu enjekte et.
Sonra JS click handler ekle.
"""
import re
from pathlib import Path

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")

# Her <ul class="idea-bullets"> ondan onceki whitespace ile birlikte yakalanir
# Toggle butonu enjekte et
pattern = re.compile(r'(\s*)(<ul class="idea-bullets">)', re.MULTILINE)

def inject_toggle(m):
    indent = m.group(1)
    ul = m.group(2)
    # idea-toggle butonu
    toggle = '\n' + indent + '<button type="button" class="idea-toggle" aria-expanded="false">'
    toggle += '\n' + indent + '  <span>Detayları göster <span class="summary">(açıklama ve mantık)</span></span>'
    toggle += '\n' + indent + '  <span class="chev">▾</span>'
    toggle += '\n' + indent + '</button>'
    return toggle + indent + ul

new_text, count = pattern.subn(inject_toggle, text)
print(f"Toggle butonu eklendi: {count} adet")

# JS click handler ekle - mevcut DOMContentLoaded sonuna
js_handler = """
  // =========================================================
  // BULLET LIST ACCORDION - Her kartin bullet'lari toggle olabilir
  // =========================================================
  document.querySelectorAll('.idea-card .idea-toggle').forEach(toggle => {
    toggle.addEventListener('click', (e) => {
      e.stopPropagation();
      const card = toggle.closest('.idea-card');
      if (!card) return;
      const isOpen = card.classList.toggle('bullets-open');
      toggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
      const label = toggle.querySelector('span:first-child');
      if (label) {
        label.innerHTML = isOpen
          ? 'Detayları gizle <span class="summary">(tıklayıp kapatabilirsin)</span>'
          : 'Detayları göster <span class="summary">(açıklama ve mantık)</span>';
      }
    });
  });
"""

# JS'i scrollHandler'dan once eklenmeli ki erkenden bind olsin
marker = "  // =========================================================\n  // EXCEL APP MODE"
new_text = new_text.replace(marker, js_handler + "\n" + marker, 1)
print("JS handler eklendi.")

HTML.write_text(new_text, encoding="utf-8")
print(f"Total size: {len(new_text):,} chars")
