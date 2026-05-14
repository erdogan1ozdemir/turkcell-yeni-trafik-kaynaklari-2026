"""
'MARKAYA ILETILMIS', 'Markaya İletilmiş' gibi ham rozetleri 'Daha önce konuşulmuş' ile degistir.
Excel referans dosyasi mention'larini kaldir.
"""
from pathlib import Path

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")
orig_len = len(text)

repl = [
    # Star pills
    ('<span class="star-pill">Markaya İletilmiş</span>', '<span class="star-pill">Daha Önce Konuşulmuş</span>'),
    # Section note title
    ('<div class="note-title">Markaya Önceden İletilen Telekom Fikirleri</div>',
     '<div class="note-title">Daha Önce Ekipçe Konuşulmuş Telekom Başlıkları</div>'),
    # Karşılaştırma p
    ('Sol sütun: Excel referans dosyasında markaya iletilmiş 11 fikir. Sağ sütun: bu raporla yeni eklenen 33+ stratejik fikir.',
     'Sol sütunda daha önce ekipçe konuşulmuş 11 öneri, sağ sütunda bu raporla eklenen 33+ yeni başlık yer alıyor.'),
    # Karşılaştırma column h3
    ('<h3>Markaya İletilmiş Fikirler (*)</h3>', '<h3>Daha Önce Konuşulmuş Fikirler (*)</h3>'),
    # Karşılaştırma column caption
    ('<p class="caption" style="color: var(--ink-3); font-size: 12px; margin-bottom: 16px;">Excel referans dosyasından - 11 fikir + 1 niş</p>',
     '<p class="caption" style="color: var(--ink-3); font-size: 12px; margin-bottom: 16px;">11 başlık + 1 niş öneri</p>'),
    # HTML yorum
    ('<!-- A6.2 Hava Durumu MARKAYA İLETİLMİŞ -->', '<!-- A6.2 Hava Durumu - daha önce konuşulmuş -->'),
    # Telekom note body
    ('<p>IMEI Sorgulama <strong>(\\*) 550K</strong>, Güvenli Arama Kapatma <strong>(\\*) 20K</strong>, Gizli Numara Kapatma <strong>(\\*) 27.1K</strong> - bu üç fikir markaya iletilmişti. Gizli Numara için Turkcell\'in mevcut destek sayfaları (5+ farklı URL) tek hub\'a toplanmalı.</p>',
     '<p>IMEI Sorgulama <strong>(*) 550.000/ay</strong>, Güvenli Arama Kapatma <strong>(*) 20.000/ay</strong> ve Gizli Numara Kapatma <strong>(*) 27.100/ay</strong> başlıkları daha önce ekipçe konuşulmuştu. Özellikle Gizli Numara için Turkcell\'in mevcut destek altında 5\'ten fazla dağınık URL var; bunların tek hub\'a konsolide edilmesi sayfa otoritesini güçlendirir.</p>'),
    # Hero p
    ('Daha önce ekipçe konuşulmuş fikirler <span class="accent-coral">(*)</span> ile işaretli.',
     'Daha önce konuşulmuş başlıklar yeşil "Daha Önce Konuşulmuş" rozetiyle veya <span class="accent-coral">(*)</span> işaretiyle ayırt edildi.'),

    # Detayli badge yumuşatma
    ('<span class="metric-badge teal">Mevcut sayfa</span>', '<span class="metric-badge teal">Mevcut sayfa</span>'),

    # idea-card data-marka="evet" border vurgu cumlesi
    ('Markaya iletilmiş', 'Daha önce konuşulmuş'),
    ('markaya iletilmiş', 'daha önce konuşulmuş'),
    ('Markaya iletilen', 'Daha önce konuşulan'),
    ('markaya iletilen', 'daha önce konuşulan'),

    # Bayram için "MARKAYA İLETİLMİŞ" başlıkları idea cards
    ('★★★ MARKAYA İLETİLMİŞ', 'Daha Önce Konuşulmuş'),

    # Yönetici özeti - bulgu vurgular
    ('5 Yapısal Bulgu', '5 Önemli Bulgu'),

    # Veri kaynaklari hero meta - referans Excel'i kaldır
    ('- Markaya iletilmiş referans Excel (`Telco - Alternatif Trafik Kaynağı Maddeleri.xlsx`)\n', ''),
    ('Markaya iletilmiş referans Excel (`Telco - Alternatif Trafik Kaynağı Maddeleri.xlsx`)', 'Daha önce ekipçe konuşulmuş başlıklar'),
]
for old, new in repl:
    if old in text:
        text = text.replace(old, new)
        print(f"OK: {old[:70]}")
    else:
        print(f"MISS: {old[:70]}")

# Yönetici özeti içindeki bulgu cümlelerinde "(\*) markaya iletilmişti" yumuşat
text = text.replace('(\\*) markaya iletilmişti.', '(daha önce ekipçe konuşulmuştu).')
text = text.replace('(\\*) markaya iletilmiş', '(daha önce konuşulmuş)')

HTML.write_text(text, encoding="utf-8")
print(f"\nHTML size: {len(text):,} chars (delta: {len(text)-orig_len:+,})")
