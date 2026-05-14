"""
MD raporunda ham wording'leri yumusat (A1-A6 etiketleri, hedef cumleleri, markaya iletilmis)
"""
from pathlib import Path

MD = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/TURKCELL_TRAFIK_FIRSATLARI_RAPORU.md")
text = MD.read_text(encoding="utf-8")
orig_len = len(text)

repl = [
    # Veri kaynaklari - markaya iletilmis referans Excel kaldir
    ('- Markaya iletilmiş referans Excel (`Telco - Alternatif Trafik Kaynağı Maddeleri.xlsx`)\n', ''),

    # Bolum baslik yumusatma (NASIL OKUMALI?)
    ('- **BÖLÜM A1 (Hesaplama Hub)** — Calculator fikirleri ve kelime listeleri',
     '- **Bölüm: Hesaplama Araçları** — Hesap makinesi tarzı yardımcı araç önerileri ve kelime kümeleri'),
    ('- **BÖLÜM A2 (Tatil & Bayram Hub)** — Mevcut `/kampanya/resmi-tatil-gunleri/` sayfasını genişletme (yeni sayfa açılmıyor)',
     '- **Bölüm: Resmi Tatil Sayfasını Genişletme** — Mevcut sayfanın kapsamını genişletme önerisi; yeni URL açılmıyor'),
    ('- **BÖLÜM A3 (Özel Gün & Mesaj)** — Özel günler + mesaj kartları',
     '- **Bölüm: Özel Gün ve Mesaj İçerikleri** — Anneler/Babalar Günü, mesaj galerileri, doğum günü mesajları'),
    ('- **BÖLÜM A4 (Telekom Yardım & Sorgu)** — Brand-uyumu en yüksek cluster',
     '- **Bölüm: Telekom Yardım ve Sorgu** — Brand güveniyle en kolay yakalanabilecek başlık grubu'),
    ('- **BÖLÜM A5 (AI, Siber, Diğer)** — Tech blog ve niş fikirler',
     '- **Bölüm: Teknoloji ve Tech Blog Konuları** — Yapay zeka, siber güvenlik, dizi/film, CV aracı, sosyal medya'),
    ('- **BÖLÜM A6 (2026 Plan Perspektifleri)** — **YENİ**: Hava durumu, Data Calculator (Vodafone UK), Coverage Map (Verizon), Sınırsız Mecra Paketleri, Persona Tarifeleri, Teknoloji Sözlüğü, Tarife Karşılaştırma, Paket Listeleme Kategori İçerikleri',
     '- **Bölüm: 2026 İçin Yeni Perspektifler** — Hava Durumu, Veri Kullanım Hesaplayıcı (Vodafone UK örneği), Kapsama Haritası (Verizon örneği), Sınırsız Mecra Paketleri, Persona Tarifeleri, Teknoloji Sözlüğü, Tarife Karşılaştırma, Paket Listeleme İçerikleri'),
    ('- **BÖLÜM B** — Ahrefs CSV\'sinden çıkarılan 35 niche cluster + tüm 60.715 keyword',
     '- **Veri ekleri** — 35 konu başlığı özeti ve toplam 60.715 anahtar kelimenin tam listesi (Excel App\'te)'),

    # Notasyon
    ('- **(\\*)** işareti — daha önce markaya iletilmiş ya da üstüne konuşulmuş fikirler',
     '- **(\\*)** işareti — bu başlığın ekipçe daha önce konuşulmuş olduğunu gösterir'),
    ('- ★★★ Acil — yüksek hacim + düşük KD + brand uyumu yüksek',
     '- ★★★ Öncelikli — yüksek arama hacmi, düşük zorluk skoru ve marka uyumu güçlü'),
    ('- ★★ Hızlı kazanım — orta hacim + orta KD + internal link mümkün',
     '- ★★ Hızlı kazanım — orta hacim, orta zorluk, dahili bağlantı olanağı yüksek'),
    ('- ★ Stratejik — düşük hacim ama yıllık spike veya niş güven sayfası',
     '- ★ Stratejik — düşük yıllık hacim ama mevsimsel zirve veya niş güven değeri taşıyan başlık'),

    # YONETICI OZETI baslik
    ('### 2026 Plan İle Hizalama\n\nMarkanın 2026 ana hedefi: **+5M click**. Aşağıdaki maddeler bu raporun fikirleriyle eşleşiyor:',
     '### Konu Başlıklarının Rapordaki Yerleri\n\nAşağıdaki tablo, çalışma planı sırasında ekipçe konuşulan başlıkları bu raporun ilgili bölümleriyle eşliyor:'),

    # 2026 Plan tablo basliklari
    ('| 2026 Planı | Bu Raporun Karşılığı | Hedef |',
     '| Konu Başlığı | Rapor Bölümü | Notlar |'),
    ('| Hesaplama sayfaları | A1 ve A6.1 (Data Calculator) | 5.5M hacim → 500K+ click |',
     '| Genel hesaplama / yardımcı araç sayfaları | Hesaplama Araçları ve Yeni Perspektifler / Veri Kullanım Hesaplayıcı | Toplam aylık 5,5 milyon arama |'),
    ('| Alternatif mass trafik LP\'ler | A6.2 (Hava durumu yeni LP), A3 (yeni LP), A2 (mevcut /kampanya/resmi-tatil-gunleri/ sayfasını genişletme) | 44M hacim → 500K+ click |',
     '| Tek seferde yüksek trafik çekebilecek konu sayfaları | Yeni Perspektifler / Hava Durumu, Özel Gün ve Mesaj İçerikleri, Resmi Tatil Sayfası Genişletme | Toplam aylık 44 milyon arama; hava durumu en güçlü lokomotif |'),
    ('| Paket listeleme sayfalarına içerik | A6.8 | %7-10 click artış |',
     '| Paket listeleme sayfalarına kategori içeriği | Yeni Perspektifler / Paket Listeleme | Mevcut sayfa trafiği üzerinde içerik genişletme |'),
    ('| Aylık düzenli blog | Tüm A4/A5 cluster\'ları besler | %5-7 click artış |',
     '| Aylık düzenli blog içerikleri | Telekom Yardım ve Tech Blog Konuları başlıklarını besler | Konu havuzu hazır |'),
    ('| Schema markup | Her sayfa tipinde önerildi | %6-10 click artış |',
     '| Schema markup uygulaması | Her sayfa tipinde uygulanacak | FAQPage, Event, BreadcrumbList önceliği |'),
    ('| Persona odaklı tarifeler | A6.5 | 50K+ click |',
     '| Persona odaklı tarife sayfaları | Yeni Perspektifler / Persona Tarifeleri | Hacim düşük, kullanıcı niyeti yüksek |'),
    ('| Mecra bazlı (sınırsız YouTube/TikTok) paketler | A6.4 | %8 click artış |',
     '| Sınırsız mecra paketleri (YouTube, TikTok vb.) | Yeni Perspektifler / Sınırsız Mecra | Listelemede facet olarak da açılabilir |'),
    ('| Teknoloji sözlüğü | A6.6 | (Plan içinde önerildiği gibi) |',
     '| Teknoloji sözlüğü | Yeni Perspektifler / Teknoloji Sözlüğü | Aylık 200 binin üzerinde doğrudan arama |'),
    ('| 4.5G karşılaştırma "fark" detayı | A6.7 | UX/CR odaklı |',
     '| 4.5G - 5G karşılaştırma detayları | Yeni Perspektifler / Tarife Karşılaştırma | Mevcut pop-up genişletme; UX odaklı |'),
    ('| Coverage Map (Verizon modeli) | A6.3 | Trend +%1614 quarterly |',
     '| 5G Kapsama Haritası | Yeni Perspektifler / Kapsama Haritası | Verizon örneği; çeyreklik +%2.500 trend |'),

    # 5 yapisal bulgu
    ('### En Büyük 5 Yapısal Bulgu', '### Öne Çıkan 5 Bulgu'),
    ('1. **Hava durumu — 41M+ aylık toplam (mass-traffic LP\'lerin lokomotifi).** "Hava durumu" tek başına **55.6M/ay**',
     '1. **Hava Durumu — Aylık 95 milyonu aşan toplam arama hacmi.** "Hava durumu" tek başına aylık 55,6 milyon arama'),
    ('(\\*) markaya iletilmişti.', '(daha önce ekipçe konuşulmuştu).'),

    # Toplam adreslenebilir hacim
    ('### Toplam Adreslenebilir Hacim', '### Konu Başlıklarının Aylık Toplam Arama Hacmi'),
    ('| Hava durumu (A6.2) | ~95M (tüm varyantlar dahil) |',
     '| Hava Durumu | ~95 milyon (tüm varyasyonlar dahil) |'),
    ('| Hesaplama hub (A1) | ~5.5M |', '| Hesaplama Araçları | ~5,5 milyon |'),
    ('| Resmi tatil & bayram (A2 - mevcut sayfa genişletme) | ~10M |',
     '| Resmi Tatil Sayfası (mevcut sayfa genişletme) | ~10 milyon |'),
    ('| Özel gün & mesaj (A3) | ~5M |', '| Özel Gün ve Mesaj İçerikleri | ~5 milyon |'),
    ('| Telekom yardım (A4) | ~3M |', '| Telekom Yardım ve Sorgu | ~3 milyon |'),
    ('| AI / Siber / Diğer (A5) | ~3M |', '| Teknoloji ve Tech Blog Konuları | ~3 milyon |'),
    ('| Sınırsız paket + persona + sözlük (A6) | ~250K (direkt) + dolaylı CR etkisi |',
     '| Yeni Perspektifler (sınırsız paket, persona, sözlük vb.) | ~250 bin doğrudan + dolaylı dönüşüm etkisi |'),
    ('| **TOPLAM** | **~120M+ adreslenebilir** |',
     '| **Toplam** | **Aylık 120 milyonu aşan toplam arama hacmi** |'),

    ('**Tahmini realize edilebilir trafik (12-18 ayda, 3-10 pozisyon hedefiyle):** 500K-1M+/ay.\n\n---',
     '\n---'),

    # Bolum baslik degisimleri
    ('# BÖLÜM A — KENDİ ARAŞTIRMA + DataForSEO Doğrulama',
     '# Konu Başlıkları + Anahtar Kelime Verisi'),
    ('## A1. HESAPLAMA ARAÇLARI HUB', '## Hesaplama Araçları'),
    ('## A2. RESMİ TATİL & BAYRAM HUB ★★★ — Mevcut Sayfayı Genişletme',
     '## Resmi Tatil Sayfasını Genişletme (Mevcut Sayfa)'),
    ('## A3. ÖZEL GÜNLER + MESAJ / SÖZ / KART ★★★', '## Özel Gün ve Mesaj İçerikleri'),
    ('## A4. TELEKOM-ÖZEL BİLGİ & YARDIM ★★★', '## Telekom Yardım ve Sorgu'),
    ('## A5. AI, SİBER, DİĞER ★★', '## Teknoloji ve Tech Blog Konuları'),
    ('# A6. 2026 PLAN PERSPEKTİFLERİ (YENİ EKLENEN)', '# 2026 İçin Yeni Perspektifler'),
    ('# BÖLÜM B — AHREFS CSV RAKİP GAP ANALİZİ (TÜM HACİM SEVİYELERİ)', '# Veri Ekleri — Konu Başlıkları ve Rakip Modeller'),
]
ok = miss = 0
for old, new in repl:
    if old in text:
        text = text.replace(old, new)
        ok += 1
    else:
        miss += 1
        print(f"MISS: {old[:70]}")

# Genel cleanup
text = text.replace('Markaya iletilmiş', 'Daha önce konuşulmuş')
text = text.replace('markaya iletilmiş', 'daha önce konuşulmuş')
text = text.replace('Markaya iletilen', 'Daha önce konuşulan')
text = text.replace('markaya iletilen', 'daha önce konuşulan')
text = text.replace('MARKAYA İLETİLMİŞ', 'DAHA ÖNCE KONUŞULMUŞ')
text = text.replace('Adreslenebilir Hacim', 'Toplam Arama Hacmi')
text = text.replace('adreslenebilir hacim', 'aylık toplam arama hacmi')

MD.write_text(text, encoding="utf-8")
print(f"OK {ok} / MISS {miss}, MD size: {len(text):,} chars (delta {len(text)-orig_len:+,})")
