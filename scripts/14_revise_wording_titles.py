"""
Buyuk revizyon:
1) Section basliklari ve eyebrow'lari yeniden adlandir (A1-A6 yerine anlamli isimler)
2) Sidebar nav'da ayni
3) Hedef/click/(*) referanslarini yumusat
4) Yonetici ozeti hedef cumlelerini cikar
5) idea-card icindeki URL'leri kaldir (yaklasim bilgisi ile degistir)
6) Rakip pattern tablosuna 200 donen linkler ekle
"""
from pathlib import Path
import re

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")
orig_len = len(text)

# ==================== 1) SIDEBAR NAV ====================
sidebar_replacements = [
    ('<li><a href="#hizalama"><span class="num">01</span><span>2026 Planıyla Hizalama</span></a></li>',
     '<li><a href="#hizalama"><span class="num">01</span><span>Strateji Çerçevesi</span></a></li>'),
    ('<li><a href="#top-firsatlar"><span class="num">02</span><span>En Büyük Fırsatlar</span></a></li>',
     '<li><a href="#top-firsatlar"><span class="num">02</span><span>Öne Çıkan Anahtar Kelimeler</span></a></li>'),
    ('<li><a href="#a6"><span class="num">03</span><span>A6 - 2026 Plan Perspektifleri</span><span class="new-badge">YENİ</span></a></li>',
     '<li><a href="#a6"><span class="num">03</span><span>2026 İçin Yeni Perspektifler</span><span class="new-badge">YENİ</span></a></li>'),
    ('<li><a href="#a1"><span class="num">04</span><span>A1 - Hesaplama Hub</span></a></li>',
     '<li><a href="#a1"><span class="num">04</span><span>Hesaplama Araçları</span></a></li>'),
    ('<li><a href="#a2"><span class="num">05</span><span>A2 - Resmi Tatil (mevcut sayfa)</span></a></li>',
     '<li><a href="#a2"><span class="num">05</span><span>Resmi Tatil Sayfası Genişletme</span></a></li>'),
    ('<li><a href="#a3"><span class="num">06</span><span>A3 - Özel Gün &amp; Mesaj</span></a></li>',
     '<li><a href="#a3"><span class="num">06</span><span>Özel Gün ve Mesaj İçerikleri</span></a></li>'),
    ('<li><a href="#a4"><span class="num">07</span><span>A4 - Telekom &amp; Sorgu</span></a></li>',
     '<li><a href="#a4"><span class="num">07</span><span>Telekom Yardım ve Sorgu</span></a></li>'),
    ('<li><a href="#a5"><span class="num">08</span><span>A5 - AI / Siber / Diğer</span></a></li>',
     '<li><a href="#a5"><span class="num">08</span><span>Teknoloji ve Tech Blog Konuları</span></a></li>'),
    ('<li><a href="#rakip"><span class="num">09</span><span>Rakip Pattern Analizi</span></a></li>',
     '<li><a href="#rakip"><span class="num">09</span><span>Rakip Sayfa Modelleri</span></a></li>'),
    ('<li><a href="#dfs"><span class="num">10</span><span>DfS Doğrulama</span></a></li>',
     '<li><a href="#dfs"><span class="num">10</span><span>Anahtar Kelime Verisi</span></a></li>'),
    ('<li><a href="#karsilastirma"><span class="num">11</span><span>İletilen vs Yeni Fikirler</span></a></li>',
     '<li><a href="#karsilastirma"><span class="num">11</span><span>Fikir Karşılaştırması</span></a></li>'),
]

for old, new in sidebar_replacements:
    if old in text:
        text = text.replace(old, new)
        print(f"OK sidebar: {old[:60]}...")
    else:
        print(f"MISS sidebar: {old[:60]}...")

# ==================== 2) SECTION EYEBROW + TITLE ====================
section_replacements = [
    # 01 - Hizalama
    ('<div class="section-eyebrow">Strateji</div>\n        <h2 class="section-title">2026 Planıyla <span class="hl">Hizalama</span></h2>',
     '<div class="section-eyebrow">Strateji Çerçevesi</div>\n        <h2 class="section-title">Rapor <span class="hl">Yol Haritası</span></h2>'),
    # 02 - Top firsatlar
    ('<div class="section-eyebrow">Veri</div>\n        <h2 class="section-title">En Büyük <span class="hl">Doğrulanmış Fırsatlar</span></h2>',
     '<div class="section-eyebrow">Anahtar Kelime Verisi</div>\n        <h2 class="section-title">Öne Çıkan <span class="hl">Anahtar Kelimeler</span></h2>'),
    # 03 - A6
    ('<div class="section-eyebrow">YENİ · 2026 Plan</div>\n        <h2 class="section-title">A6 - <span class="hl">2026 Plan Perspektifleri</span></h2>',
     '<div class="section-eyebrow">Yeni Perspektifler</div>\n        <h2 class="section-title">2026 İçin <span class="hl">Yeni Perspektifler</span></h2>'),
    # 04 - A1
    ('<div class="section-eyebrow">Cluster · Hesaplama</div>\n        <h2 class="section-title">A1 - <span class="hl">Hesaplama Araçları Hub</span></h2>',
     '<div class="section-eyebrow">Hesaplama Araçları</div>\n        <h2 class="section-title">Hesaplama <span class="hl">Araçları</span></h2>'),
    # 05 - A2
    ('<div class="section-eyebrow">Cluster · Tatil &amp; Bayram · MEVCUT SAYFA GENİŞLETME</div>\n        <h2 class="section-title">A2 - <span class="hl">/kampanya/resmi-tatil-gunleri/</span> Mega-Hub Genişletme</h2>',
     '<div class="section-eyebrow">Mevcut Sayfa Genişletme</div>\n        <h2 class="section-title">Resmi Tatil <span class="hl">Sayfasını Genişletme</span></h2>'),
    # 06 - A3
    ('<div class="section-eyebrow">Cluster · Özel Gün</div>\n        <h2 class="section-title">A3 - <span class="hl">Özel Gün + Mesaj/Söz/Kart</span></h2>',
     '<div class="section-eyebrow">Özel Gün İçerikleri</div>\n        <h2 class="section-title">Özel Gün ve <span class="hl">Mesaj İçerikleri</span></h2>'),
    # 07 - A4
    ('<div class="section-eyebrow">Cluster · Telekom-Özel</div>\n        <h2 class="section-title">A4 - <span class="hl">Telekom Yardım &amp; Sorgu</span></h2>',
     '<div class="section-eyebrow">Telekom Yardım</div>\n        <h2 class="section-title">Telekom Yardım <span class="hl">ve Sorgu</span></h2>'),
    # 08 - A5
    ('<div class="section-eyebrow">Cluster · Tech Blog</div>\n        <h2 class="section-title">A5 - <span class="hl">AI, Siber Güvenlik, Diğer</span></h2>',
     '<div class="section-eyebrow">Tech Blog</div>\n        <h2 class="section-title">Teknoloji ve <span class="hl">Tech Blog Konuları</span></h2>'),
    # 09 - Rakip
    ('<div class="section-eyebrow">Veri · Rakip Modelleri</div>\n        <h2 class="section-title">Rakip URL <span class="hl">Pattern Analizi</span></h2>',
     '<div class="section-eyebrow">Rakip Sayfa Modelleri</div>\n        <h2 class="section-title">Rakiplerin <span class="hl">Trafik Çeken Sayfaları</span></h2>'),
    # 10 - DfS
    ('<div class="section-eyebrow">Veri · DataForSEO</div>\n        <h2 class="section-title">DataForSEO <span class="hl">Doğrulama</span></h2>',
     '<div class="section-eyebrow">Anahtar Kelime Verisi</div>\n        <h2 class="section-title">Anahtar Kelime <span class="hl">Hacim Verisi</span></h2>'),
    # 11 - Karsilastirma
    ('<div class="section-eyebrow">Stok Karşılaştırma</div>\n        <h2 class="section-title">Markaya İletilenler <span class="hl">vs Yeni Eklemeler</span></h2>',
     '<div class="section-eyebrow">Fikir Karşılaştırması</div>\n        <h2 class="section-title">Daha Önce Konuşulanlar <span class="hl">ve Bu Raporla Eklenenler</span></h2>'),
]

for old, new in section_replacements:
    if old in text:
        text = text.replace(old, new)
        print(f"OK section title: {old[40:90]}...")
    else:
        print(f"MISS section title: {old[40:90]}...")

# ==================== 3) HERO ALTI ====================
hero_meta_replacements = [
    ('<span class="meta-pill mint">120M+ adreslenebilir hacim</span>',
     '<span class="meta-pill mint">120M+ aylık arama hacmi</span>'),
    ('Markaya daha önce iletilen fikirler <span class="accent-coral">(*)</span> ile işaretlendi.',
     'Daha önce ekipçe konuşulmuş fikirler <span class="accent-coral">(*)</span> ile işaretli.'),
    ('60.715</strong> Turkcell-yok rakip-var fırsat <strong>35 evergreen cluster\'a</strong> ayrıldı,',
     '60.715</strong> Turkcell\'in henüz sıralanmadığı, ancak en az bir rakibin trafik çektiği anahtar kelime <strong>35 evergreen başlık altında</strong> gruplandırıldı,'),
    ('<strong>150+ kritik anahtar kelime DataForSEO ile doğrulandı</strong>',
     '<strong>150+ kritik anahtar kelimenin aylık arama hacmi DataForSEO üzerinden incelendi</strong>'),
]
for old, new in hero_meta_replacements:
    if old in text:
        text = text.replace(old, new)
        print(f"OK hero: {old[:60]}...")
    else:
        print(f"MISS hero: {old[:60]}...")

# ==================== 4) YONETICI OZETI - HEDEF KISIMLARI ====================
exec_replacements = [
    # Tahmini realize cumlesi
    ('**Tahmini realize edilebilir trafik (12-18 ayda, 3-10 pozisyon hedefiyle):** 500K-1M+/ay.',
     ''),
    # Toplam adreslenebilir trafik tablosu kalsin ama tahmin cumlesi yumusatilsin
    ('### Toplam Adreslenebilir Hacim',
     '### Konu Başlıklarının Aylık Arama Hacmi'),
    # Stratejik bulgu vurgulari
    ('### En Büyük 5 Yapısal Bulgu',
     '### Öne Çıkan 5 Bulgu'),
    # Hero KPI tile
    ('<div class="lbl">Aylık Adreslenebilir Hacim</div>',
     '<div class="lbl">Aylık Toplam Arama Hacmi</div>'),
    ('<div class="lbl">DfS Doğrulanmış KW</div>',
     '<div class="lbl">DfS Üzerinden İncelenmiş KW</div>'),
    # Section 02 chart sub
    ('Coral barlar markaya daha önce iletilmiş fikirleri (\\*), teal barlar bu raporda yeni çıkarılanları gösterir.',
     'Coral barlar daha önce ekipçe konuşulmuş fikirleri (*), teal barlar bu raporla eklenen yeni başlıkları gösterir.'),
    # Section 02 lead
    ('DataForSEO Mayıs 2026 (TR/tr, location_code 2792) ile doğrulanan top 20 anahtar kelime - aylık arama hacmiyle.',
     'DataForSEO üzerinden Mayıs 2026 verisiyle (TR/tr) incelenmiş öne çıkan 20 anahtar kelime ve aylık arama hacimleri.'),
]
for old, new in exec_replacements:
    if old in text:
        text = text.replace(old, new)
        print(f"OK exec: {old[:60]}...")
    else:
        print(f"MISS exec: {old[:60]}...")

# Toplam adreslenebilir tablosu icindeki son cumleyi yumusat
text = text.replace(
    'TR\'de "5G kapsama alanı" yıllık trend **+%1025**, "5G kapsama haritası" **+%2500 quarterly**. Mega yükseliş trendi.',
    'TR\'de "5G kapsama alanı" yıllık trend +%1025, "5G kapsama haritası" çeyreklik +%2500. Yükselen ilgi alanı.'
)

# ==================== 5) HIZALAMA TABLOSU YUMUSATMA ====================
# Tablonun "Hedef" sutunu kalkmasi yerine "Konu Kapsamı" gibi nötr olsun
# Tabloyu komple kaldirmak yerine kolon ismini degistirelim
old_table_header = '''<th style="width:30%">2026 Plan Maddesi</th>
            <th style="width:18%">Rapor Bölümü</th>
            <th>Hedef Hacim</th>
            <th>Click Hedefi</th>
            <th>Durum</th>'''
new_table_header = '''<th style="width:34%">Konu Başlığı</th>
            <th style="width:24%">Rapor Bölümü</th>
            <th>Bu Konunun Aylık Arama Hacmi</th>
            <th>Notlar</th>'''
if old_table_header in text:
    text = text.replace(old_table_header, new_table_header)
    print("OK plan table header revized")

# Plan table satirlari - hedef cumlelerini cikar (5 sutundan 4 sutuna duser)
# Onceki: <td>Konu</td> <td>Bolum</td> <td>Hacim</td> <td>Click</td> <td>Durum</td>
# Sonraki: <td>Konu</td> <td>Bolum</td> <td>Hacim</td> <td>Not (Durum+Click birlestirildi)</td>

plan_rows_replacements = [
    # Row 1
    ('''<tr>
            <td>Hesaplama sayfalarının oluşturulması</td>
            <td class="num">A1 + A6.1</td>
            <td>5.5M</td>
            <td class="coral-cell">500K+</td>
            <td class="pos-cell">Veri hazır</td>
          </tr>''',
     '''<tr>
            <td>Genel hesaplama / yardımcı araç sayfaları</td>
            <td class="num">Bölüm 04 (Hesaplama Araçları) ve 03 (Veri Kullanım Hesaplayıcı)</td>
            <td>5,5 milyon</td>
            <td>16 adet farklı hesaplama önerisi</td>
          </tr>'''),
    # Row 2
    ('''<tr>
            <td>Alternatif mass-trafik LP'leri</td>
            <td class="num">A6.2 + A3 (yeni LP) + A2 (mevcut sayfa genişletme)</td>
            <td>44M</td>
            <td class="coral-cell">500K+</td>
            <td class="pos-cell">Hava durumu lokomotif</td>
          </tr>''',
     '''<tr>
            <td>Tek seferde yüksek trafik çekebilecek konu sayfaları</td>
            <td class="num">Yeni Perspektifler (Hava Durumu), Özel Gün İçerikleri ve Resmi Tatil Sayfası</td>
            <td>44 milyon</td>
            <td>Hava durumu en güçlü tek başlık</td>
          </tr>'''),
    # Row 3
    ('''<tr>
            <td>Paket listeleme sayfalarına içerik</td>
            <td class="num">A6.8</td>
            <td>Mevcut sayfalar</td>
            <td>%7-10 artış (250K+)</td>
            <td class="pos-cell">Brief hazır</td>
          </tr>''',
     '''<tr>
            <td>Paket listeleme sayfalarına kategori içeriği</td>
            <td class="num">Yeni Perspektifler — Paket Listeleme</td>
            <td>Mevcut sayfa trafiği üzerinde</td>
            <td>Her listing için açıklayıcı içerik + SSS</td>
          </tr>'''),
    # Row 4
    ('''<tr>
            <td>Aylık düzenli blog içerikleri</td>
            <td class="num">A4 + A5 besler</td>
            <td>—</td>
            <td>%5-7 artış (200K+)</td>
            <td class="pos-cell">Cluster set hazır</td>
          </tr>''',
     '''<tr>
            <td>Aylık düzenli blog içerikleri</td>
            <td class="num">Telekom Yardım ve Tech Blog Konuları başlıklarını besler</td>
            <td>—</td>
            <td>Cluster konu havuzu hazır</td>
          </tr>'''),
    # Row 5
    ('''<tr>
            <td>Schema markup uygulaması</td>
            <td class="num">Her sayfa tipi</td>
            <td>—</td>
            <td>%6-10 artış (200K+)</td>
            <td class="pos-cell">Şablonlar belirlendi</td>
          </tr>''',
     '''<tr>
            <td>Schema markup uygulaması</td>
            <td class="num">Her sayfa tipinde uygulanacak</td>
            <td>—</td>
            <td>FAQPage, Event, BreadcrumbList önceliği</td>
          </tr>'''),
    # Row 6
    ('''<tr>
            <td>Persona odaklı paket tarifeleri</td>
            <td class="num">A6.5</td>
            <td>Düşük hacim, yüksek CR</td>
            <td class="coral-cell">50K+</td>
            <td class="pos-cell">9 persona haritalandı</td>
          </tr>''',
     '''<tr>
            <td>Persona odaklı tarife sayfaları</td>
            <td class="num">Yeni Perspektifler — Persona Tarifeleri</td>
            <td>Hacim düşük, kullanıcı niyeti yüksek</td>
            <td>9 farklı persona profili çıkarıldı</td>
          </tr>'''),
    # Row 7
    ('''<tr>
            <td>Sınırsız mecra bazlı paketler (YouTube/TikTok)</td>
            <td class="num">A6.4</td>
            <td>10-15K küme</td>
            <td>%8 artış (150K+)</td>
            <td class="pos-cell">Facet planı hazır</td>
          </tr>''',
     '''<tr>
            <td>Sınırsız mecra paketleri (YouTube, TikTok vb.)</td>
            <td class="num">Yeni Perspektifler — Sınırsız Mecra</td>
            <td>10-15 bin doğrudan + dolaylı paket trafiği</td>
            <td>Listeleme sayfalarında facet olarak açılabilir</td>
          </tr>'''),
    # Row 8
    ('''<tr>
            <td>Teknoloji sözlüğü</td>
            <td class="num">A6.6</td>
            <td>200K+/ay direkt</td>
            <td>AI Overview ideal</td>
            <td class="pos-cell">100+ tanım listesi</td>
          </tr>''',
     '''<tr>
            <td>Teknoloji sözlüğü</td>
            <td class="num">Yeni Perspektifler — Teknoloji Sözlüğü</td>
            <td>Aylık 200 binin üzerinde doğrudan arama</td>
            <td>AI Overview / kısa cevap formatı için uygun</td>
          </tr>'''),
    # Row 9
    ('''<tr>
            <td>4.5G karşılaştırma "fark" detayı</td>
            <td class="num">A6.7</td>
            <td>UX/CR odaklı</td>
            <td>Pop-up genişletme</td>
            <td class="pos-cell">10 fark alanı önerildi</td>
          </tr>''',
     '''<tr>
            <td>4.5G - 5G karşılaştırma detayları</td>
            <td class="num">Yeni Perspektifler — Tarife Karşılaştırma</td>
            <td>Mevcut sayfa kullanıcı deneyimi odaklı</td>
            <td>10 farklı karşılaştırma alanı önerildi</td>
          </tr>'''),
    # Row 10
    ('''<tr>
            <td>5G Coverage Map (Verizon modeli)</td>
            <td class="num">A6.3</td>
            <td>Trend +%1614 quarterly</td>
            <td>5G launch ile büyüyecek</td>
            <td class="coral-cell">Yeni fikir</td>
          </tr>''',
     '''<tr>
            <td>5G Kapsama Haritası</td>
            <td class="num">Yeni Perspektifler — Kapsama Haritası</td>
            <td>Mevcut hacim küçük, çeyreklik trend +%2500</td>
            <td>Yurtdışı örnek: Verizon Coverage Map</td>
          </tr>'''),
]
for old, new in plan_rows_replacements:
    if old in text:
        text = text.replace(old, new)
        print(f"OK plan row: {old[40:80]}...")
    else:
        print(f"MISS plan row: {old[40:80]}...")

# Plan tablo dipnotu
text = text.replace(
    '<p>Bu rapor "ne yapılmalı"dan çok "neyi neden yapmalı, hangi anahtar kelimeleri hedeflemeli" sorusuna odaklanır. Her fikir için <strong>ana sayfa önerisi + URL + DfS doğrulanmış hacim + oynanabilecek varyasyon listesi</strong> verilmiştir.</p>',
    '<p>Bu rapor "ne yapılmalı" sorusundan çok, hangi konu başlıklarının niye değerli olduğunu ve hangi anahtar kelimelerin altında toplandığını gösterir. Her öneri için <strong>açıklama, ilgili anahtar kelime kümesi, aylık arama hacmi ve oynayabilecek varyasyon listesi</strong> verilmiştir.</p>'
)

# ==================== 6) DfS DOGRULAMA WORDING ====================
text = text.replace(
    'DataForSEO Mayıs 2026 (TR/tr, location_code 2792) ile doğrulanan top 20 anahtar kelime',
    'DataForSEO üzerinden Mayıs 2026 verisiyle incelenmiş öne çıkan 20 anahtar kelime'
)
text = text.replace(
    '150+ kritik keyword Mayıs 2026\'da TR/tr için doğrulandı. Aşağıdaki grafikler KD ve intent dağılımını gösterir -',
    '150+ önemli anahtar kelime için Mayıs 2026 (TR/tr) hacim ve zorluk verisi DataForSEO üzerinden çekildi. Aşağıdaki grafikler keyword difficulty ve arama niyeti dağılımını gösterir -'
)
text = text.replace(
    'Doğrulanmış keyword\'lerin keyword difficulty dağılımı',
    'İncelenmiş anahtar kelimelerin Keyword Difficulty (zorluk) dağılımı'
)
text = text.replace(
    '<div class="chart-title">Top 30 Doğrulanmış Anahtar Kelime</div>',
    '<div class="chart-title">Öne Çıkan 30 Anahtar Kelime</div>'
)
text = text.replace(
    '<h3 style="font-family: var(--font-display); font-weight: 700; color: var(--teal); margin: 32px 0 14px; font-size: 22px;">Top 30 Doğrulanmış Anahtar Kelime</h3>',
    '<h3 style="font-family: var(--font-display); font-weight: 700; color: var(--teal); margin: 32px 0 14px; font-size: 22px;">Öne Çıkan 30 Anahtar Kelime</h3>'
)

# Excel-de aç button label'lari yumusat
btn_label_replacements = [
    ('Excel · Yönetici Özeti', 'Excel · Genel Özet'),
    ('Excel · DfS Doğrulama', 'Excel · Anahtar Kelime Verisi'),
    ('Excel · DfS 150+ KW tam tablo', 'Excel · 150+ Anahtar Kelime Tam Tablo'),
    ('Excel · A6 8 fikrin tam kelime listesi', 'Excel · Yeni Perspektifler 8 Önerinin Tam Kelime Listesi'),
    ('Excel · A1 16 fikir + kelime listesi', 'Excel · Hesaplama Araçları 16 Öneri'),
    ('Excel · A2 mega-hub + 77 KW kelime listesi', 'Excel · Resmi Tatil Sayfası 77 Anahtar Kelime'),
    ('Excel · A2 7 fikir + kelime listesi', 'Excel · Resmi Tatil Sayfası Önerileri'),
    ('Excel · A3 9 fikir + kelime listesi', 'Excel · Özel Gün ve Mesaj 9 Öneri'),
    ('Excel · A4 6 fikir + kelime listesi', 'Excel · Telekom Yardım 6 Öneri'),
    ('Excel · A5 9 fikir + kelime listesi', 'Excel · Tech Blog 9 Öneri'),
    ('Excel · A6 mega-hub detayı', 'Excel · Resmi Tatil Sayfası Detayı'),
    ('Excel · Tüm 60K Fırsat', 'Excel · Tüm Anahtar Kelimeler (60K)'),
    ('Excel · 35 cluster özeti', 'Excel · 35 Konu Başlığı Özeti'),
    ('Excel · Hesaplama cluster (top 300 KW)', 'Excel · Hesaplama Konusunun Top 300 KW\'si'),
    ('Excel · Tatil/Bayram cluster (top 300 KW)', 'Excel · Tatil ve Bayram Top 300 KW'),
    ('Excel · Mesaj/Söz cluster', 'Excel · Mesaj ve Söz Konusu'),
    ('Excel · E-Devlet/Sorgu cluster', 'Excel · E-Devlet ve Sorgu Konusu'),
    ('Excel · Telefon Ayar/Sorun cluster', 'Excel · Telefon Ayar ve Sorun Çözme'),
    ('Excel · Tanım/Nedir cluster', 'Excel · "Nedir" / Tanım Konusu'),
    ('Excel · CV/İş Başvuru cluster', 'Excel · CV ve İş Başvuru Konusu'),
    ('Excel · Rakip URL pattern detayı', 'Excel · Rakip Sayfa Modelleri Detayı'),
]
for old, new in btn_label_replacements:
    text = text.replace(old, new)

HTML.write_text(text, encoding="utf-8")
print(f"\nUpdated HTML: {len(text):,} chars (delta: {len(text) - orig_len:+,})")
