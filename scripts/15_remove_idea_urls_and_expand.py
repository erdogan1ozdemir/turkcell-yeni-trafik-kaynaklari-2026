"""
1) Tum <span class="idea-url">/...</span> satirlarini kaldir
2) idea-card'larin .idea-desc cumlelerini genislet (kim icin, ne icin, nasil entegre)
"""
from pathlib import Path
import re

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")
orig_len = len(text)

# 1) URL'leri kaldir
url_pattern = re.compile(r'\s*<span class="idea-url"[^>]*>[^<]*</span>\n?', re.DOTALL)
matches = url_pattern.findall(text)
print(f"Removing {len(matches)} idea-url spans")
text = url_pattern.sub("\n        ", text)

# 2) Aciklamalari genislet - her fikir icin
# Format: ('<p class="idea-desc">OLD</p>', '<p class="idea-desc">NEW</p>')
desc_replacements = [
    # ===== A6 - 2026 Plan Perspektifleri =====
    # A6.1 Data Calculator
    (
        '<p class="idea-desc">Slider\'lı kullanım girişi + kişiselleştirilmiş paket önerisi. Vodafone UK örneğinde 469 KW / 56.6K trafik, tüm informational.</p>',
        '<p class="idea-desc">Kullanıcı kaydırma çubuklarıyla aylık YouTube, Netflix, Instagram ve TikTok kullanımını girer; sayfa bu girdiyi GB tahminine çevirip mevcut Turkcell tarifelerinden uygun olanları önerir. Vodafone UK örneğindeki tek sayfa 469 farklı sorguda sıralanıp ayda 56.600 ziyaret çekiyor. Tüm sorgular bilgi-amaçlı; bu da AI Overview alıntılarına çıkma ihtimalini yüksek tutuyor. Hem dönüşüm hem AI görünürlüğüne aynı anda fayda sağlayan nadir bir yapı.</p>'
    ),
    # A6.2 Hava Durumu
    (
        '<p class="idea-desc">Markaya iletilmiş ana mass-trafik LP fikri. MGM API + ICT desteği gerekli. WeatherForecast + Place schema.</p>',
        '<p class="idea-desc">81 il ve ilçe için hava durumu hub\'ı; ana sayfa, "yarınki", "5 günlük", "10 günlük" ve "saatlik" alt görünümler. Tek başına aylık 95 milyonun üzerinde aramayı kapsıyor — bu kümede İstanbul, Ankara ve İzmir başlıkları milyonluk hacimler. Veri MGM veya lisanslı bir hava durumu sağlayıcısından çekilir; WeatherForecast ve Place schema eklenir. ICT ekibinin entegrasyon desteği gerekiyor. Daha önce ekipçe konuşulmuş bir başlık; bu raporla doğrulanmış arama hacimleri ve URL yapısı önerisi netleşti.</p>'
    ),
    # A6.3 Coverage Map
    (
        '<p class="idea-desc">Verizon: 7041 KW / 109K trafik / toplamın %1\'i tek sayfada. TR 5G launch ile patlayacak - adres girişi + interaktif harita + 5G FAQ.</p>',
        '<p class="idea-desc">Kullanıcı adres veya konum bilgisi girerek bulunduğu yerde 5G / 4.5G / LTE kapsamasını görür. Verizon\'un tek bir kapsama haritası sayfası 7.000\'den fazla sorguda sıralanarak aylık 109.000 ziyaret çekiyor — bu, sitenin toplam trafiğinin yaklaşık %1\'i. Türkiye\'de "5G kapsama alanı" sorgusu çeyrek bazda +%2.500 yükseliş gösterdi; 5G yaygınlaşmasıyla ilginin daha da artması bekleniyor. Sayfanın altına 5G hızı, hangi telefonların desteklediği gibi SSS\'ler eklenerek uzun kuyruk sorgular da kapsanabilir.</p>'
    ),
    # A6.4 Sınırsız Mecra
    (
        '<p class="idea-desc">Plan listeleme sayfalarında "sınırsız [mecra]" facet\'leri açılır. Her facet bir landing page\'e çıkar.</p>',
        '<p class="idea-desc">Mevcut paket listeleme sayfalarına "sınırsız YouTube", "sınırsız TikTok", "sınırsız WhatsApp", "sınırsız Instagram" gibi filtreler eklenir; her filtre kendi açıklayıcı sayfasına yönlenir. Bu sayfalar hem ticari sorgu (sınırsız YouTube paketi 590/ay) hem de dolaylı paket trafiği çeker. TikTok başlığı yıllık +%250 ile en hızlı büyüyen. Hem listeleme sayfasının dönüşümünü yükseltir hem ayrı landing\'lerle SEO trafiği kazandırır.</p>'
    ),
    # A6.5 Persona Tarifeler
    (
        '<p class="idea-desc">9 persona: öğrenci, asker, emekli, KOBİ, engelli, aile/çift hat, gençlik, expat, oyuncu. Her persona için tipik kullanım + tarife öneri + FAQ.</p>',
        '<p class="idea-desc">Her persona için ayrı sayfa: tipik kullanım profili (GB, dakika, sosyal medya), önerilen 2-3 tarife, "neden bu persona için" kısa açıklama ve sık sorulan sorular. 9 profil çıkarıldı (öğrenci, asker, emekli, KOBİ, engelli, aile/çift hat, gençlik, yurt dışından gelen abone, oyuncu). Arama hacimleri tek tek küçük ama kullanıcı niyeti çok yüksek; dolayısıyla satış dönüşümüne ciddi katkı sağlar. "Emekli tarifesi" gibi sorgularda KD çok düşük (1).</p>'
    ),
    # A6.6 Teknoloji Sözlüğü
    (
        '<p class="idea-desc">100+ teknik terim "nedir" sayfası. Schema: DefinedTerm + FAQPage. <strong>Muud nedir / Muud Premium TT\'ye gidiyor</strong> (15K+11K) - Turkcell kendi açıklaması yok!</p>',
        '<p class="idea-desc">100\'den fazla teknik terim için kısa, net "nedir" sayfaları (GSM, RAM, eSIM, NFC, DNS, modem, fiber internet, 5G vb.). Her sayfa DefinedTerm + FAQPage schema ile işaretlenir; bu format AI Overview ve Featured Snippet alıntılarına çok uygun. Önemli bir tespit: "Muud nedir" ve "Muud Premium nedir" — ki bunlar Turkcell\'in kendi markası — Türk Telekom\'un blog sayfalarına gidiyor (15K + 11K aylık arama Turkcell\'in eline geçmiyor). Bu hub açıldığında brand sorgular da kazanılmış olur.</p>'
    ),
    # A6.7 Tarife Fark
    (
        '<p class="idea-desc">Mevcut karşılaştırma pop-up\'ında sadece görsel var. "Fark" alanı detaylandırılmalı: GB, dakika, hotspot, 5G, roaming, dijital servisler.</p>',
        '<p class="idea-desc">Mevcut /paket-ve-tarifeler/4-5-g-hizinda sayfasının "karşılaştır" pop-up\'ında şu an yalnızca görsel var; tarifeler arasındaki fark belirsiz. Bu pop-up\'a 10 farklı karşılaştırma alanı eklenir: GB miktarı ve sınırlandırma, dakika her yöne, SMS, hotspot izni ve limiti, 5G desteği ve hız sınırı, roaming ülke listesi, dahili dijital servisler (BiP / fizy / TV+ / lifebox), eSIM uyumluluğu, taahhüt süresi ve cayma bedeli, ek avantajlar (genç indirimi, mağaza puanı). ComparisonTable schema eklenmesi bu pop-up\'ın görünürlüğünü artırır.</p>'
    ),
    # A6.8 Paket Listeleme
    (
        '<p class="idea-desc">Her listing sayfasının üstüne 300-500 kelime kategori içeriği: tanım + karşılaştırma tablosu + 5-10 FAQ + internal link + schema (BreadcrumbList, FAQPage, ItemList).</p>',
        '<p class="idea-desc">Mevcut paket listeleme sayfalarının üst kısmına 300-500 kelimelik kategori açıklaması, ardından özellik bazlı karşılaştırma tablosu, 5-10 sık sorulan soru ve ilgili sayfalara dahili link eklenir. Sayfa türüne göre uyarlanır: ana paket listeleme, 4.5G/5G karşılaştırma, fiber paketler, genç tarifeleri, aile paketleri vb. Schema markup ile FAQPage, BreadcrumbList ve ItemList işaretlemesi yapılır. Bu yaklaşım hem mevcut sayfaların organik görünürlüğünü artırır hem AI Overview\'da alıntılanma ihtimalini yükseltir.</p>'
    ),
    # ===== A1 - Hesaplama =====
    # Fikir 1 - Hesap Makinesi
    (
        '<p class="idea-desc">Vatan modeli; basit + bilimsel + fonksiyonlu modlar.</p>',
        '<p class="idea-desc">Tek bir hesap makinesi hub\'ı altında basit, bilimsel ve fonksiyonlu modlar — bunlara ek olarak yüzde, KDV, döviz, BMI gibi hızlı geçişler de eklenebilir. Vatan Bilgisayar\'ın bu pattern\'le açtığı tek sayfa 48 sorguda sıralanıp aylık 33.991 ziyaret çekiyor; "hesap makinesi" tek başına aylık 2,17 milyon arama alıyor. Hem üst başlık trafik magnetli hem de mevcut yüzde/yaş hesaplama sayfalarını besleyen iç bağlantı mimarisi kuruyor.</p>'
    ),
    # Fikir 2 - Yas Hesaplama (mevcut)
    (
        '<p class="idea-desc">Mevcut sayfa; kedi/köpek yaşı + iki tarih arası gün widget ekle.</p>',
        '<p class="idea-desc">Halihazırda canlı olan yaş hesaplama sayfasının kapsamı genişletilir: doğum tarihinden yaş hesaplama (8.100/ay), kedi yaşı hesaplama (12.000/ay), köpek yaşı hesaplama ve "iki tarih arası gün" widget\'ı (22.200/ay) gibi ek aletler aynı sayfanın alt başlıkları olarak konumlanır. Vodafone\'un kedi yaşı sayfası 1. sırada, oysa Turkcell brand otoritesiyle bu sorguyu kolayca kapabilir. Mevcut URL korunur, yalnızca sayfa içeriği zenginleşir.</p>'
    ),
    # Fikir 3 - Yuzde Hesaplama (mevcut)
    (
        '<p class="idea-desc">Mevcut sayfayı senaryo bazlı genişlet: indirim, KDV, maaş zammı.</p>',
        '<p class="idea-desc">Mevcut yüzde hesaplama sayfasına kullanım senaryoları eklenir: alışveriş indirim oranı, KDV dahil/hariç hesaplama, maaş zammı, yüzdelik dilim, yüzde artış ve azalış. "İndirim hesaplama" sorgusu tek başına aylık 49.500 arama alıyor; bu da daha önce ekipçe konuşulmuş bir başlık. Ana sayfa korunur; alt başlıklar accordion veya sekme yapısıyla eklenebilir.</p>'
    ),
    # Fikir 4 - Kac gun kaldi
    (
        '<p class="idea-desc">Tek hub + dinamik geri sayım. Bayramlar/tatiller dahil.</p>',
        '<p class="idea-desc">Önemli tarihler için dinamik geri sayım hub\'ı: Kurban Bayramı\'na kaç gün, Ramazan\'a kaç gün, yılbaşına kaç gün, doğum gününe kaç gün — hepsi tek hub altında. URL parametresiyle ("?tarih=2026-10-29") herhangi bir tarihe özel geri sayım paylaşılabilir, sosyal medya paylaşımına uygundur. "İki tarih arası gün" (22.200/ay) ve "kaç gün kaldı" (33.100/ay) gibi ana sorgular küme olarak aylık 1,4 milyonu aşıyor.</p>'
    ),
    # Fikir 5 - Brüt-Net Maaş
    (
        '<p class="idea-desc">Vodafone modeli; Asgari Ücret 2026 (368K, +%210K trend!) canlı tablo bileşeniyle.</p>',
        '<p class="idea-desc">Brüt-net maaş hesaplama aracı + güncel asgari ücret tablosu birlikte sunulur. Vodafone\'un benzer sayfası 110 sorguda sıralanıp aylık 7.365 ziyaret çekiyor. "Asgari ücret 2026" sorgusu yıllık +%210.212 yükselişle aylık 368.000 aramaya ulaşmış durumda; bu rakam canlı bir tablo bileşeniyle sayfada her yıl güncel tutulduğunda kalıcı trafik kaynağı oluşturur. Hem bireysel çalışan hem işveren niyeti hedefleniyor.</p>'
    ),
    # Fikir 6 - KDV Hesaplama
    (
        '<p class="idea-desc">İşletme/freelancer kitle. Paycell köprüsü.</p>',
        '<p class="idea-desc">KDV dahil/hariç hesaplama, oran çevirici (%1, %8, %10, %20), tevkifat ve farklı senaryolar için tek sayfa. Aylık 201.000 arama alan ana sorgunun KD\'si 19 — orta seviye zorluk, dolayısıyla brand otoritesiyle yakalanabilir. Hedef kitle: KOBİ, freelancer, e-ticaret yapan küçük işletme sahipleri. Paycell ürünleriyle dahili bağlantı kurulabilir.</p>'
    ),
    # Fikir 7 - Kidem Ihbar
    (
        '<p class="idea-desc">Tek sayfada çift araç. HR/iş hayatı kitle.</p>',
        '<p class="idea-desc">Tek bir sayfada hem kıdem tazminatı hem ihbar tazminatı hesaplama aracı; ek olarak kıdem tazminatı tavanı 2026 canlı tablosu, hak ediş şartları ve sık sorulanlar. İki ana sorgu birlikte aylık 228.000 arama alıyor (kıdem 201.000, ihbar 27.100). Kullanıcı niyeti çoğunlukla bilgi-amaçlı; HR yöneticileri, işverenler ve iş arayanlar hedef.</p>'
    ),
    # Fikir 8 - Calisan Haklari
    (
        '<p class="idea-desc">Yıllık izin, fazla mesai, SGK, emeklilik yaşı tek hub\'tan dallar.</p>',
        '<p class="idea-desc">Çalışan hakları başlığı altında bir mini hub: yıllık izin hesaplama (9.900/ay, KD 4), fazla mesai hesaplama (9.900/ay), gece zammı, bayram mesaisi, SGK hizmet dökümü erişimi (49.500/ay), SGK işe giriş (135.000/ay) ve emeklilik yaşı hesaplama (3.600/ay). Bu sorguların ortak özelliği düşük KD ve yüksek kullanıcı niyeti. Tek hub\'tan dallandığında her alt başlık ana sayfaya iç bağlantı sağlar.</p>'
    ),
    # Fikir 9 - Yukselen Burc
    (
        '<p class="idea-desc">Doğum tarihi + saat girişli widget. Eğlence niş, mega hacim.</p>',
        '<p class="idea-desc">Doğum tarihi, saati ve doğum yeri girilerek yükselen burç hesaplanır; sonuçla birlikte burcun özelliklerine dair kısa açıklama da gösterilir. Aylık 201.000 arama ve KD 10 — yüksek hacim, düşük zorluk kombinasyonu. Eğlence içerikli bir başlık olduğu için brand tonu kararı işletme tercihine bağlı; ancak fizy/Turkcell TV+\'ın "günlük burç yorumu" özellikleriyle entegre edilebilir.</p>'
    ),
    # Fikir 10 - Saglik
    (
        '<p class="idea-desc">BMI, kalori, hamilelik, ovulasyon, regl. Kadın+sağlık kitle.</p>',
        '<p class="idea-desc">Sağlık alanında bir mini hub: BMI/ideal kilo hesaplama (8.100/ay), kalori hesaplama (74.000/ay), hamilelik haftası (2.400/ay, KD 8), ovulasyon hesaplama (6.600/ay, KD 7), regl takvimi (18.100/ay). Bu sorguların büyük çoğunluğu kadın kullanıcılardan; kullanıcılar pipeline mantığında birden fazla aracı sırayla kullanıyor (regl → ovulasyon → gebelik). Lifebox ve BiP gibi servislerle "kişisel sağlık verisi saklama" temalı kampanyalara köprü kurulabilir.</p>'
    ),
    # Fikir 11 - Yazma Araclari
    (
        '<p class="idea-desc">EN DÜŞÜK KD! Sosyal medya+öğrenci kitle.</p>',
        '<p class="idea-desc">Sosyal medya ve öğrenci kitlesine yönelik bir yardımcı araç hub\'ı: karakter sayacı (14.800/ay, KD 3 — en düşük zorluk), kelime sayacı (18.100/ay, KD 12), büyük harf - küçük harf çevirici, "@" gibi klavye sembol rehberleri (72.000/ay), QR kod oluşturucu, şifre oluşturucu. Tüm sorgular düşük KD\'li; mobil-first widget formatında planlandığında Twitter/X kullanıcılarının karakter limiti, öğrencilerin ödev kelime sayımı, sosyal medya yöneticilerinin günlük kullanım ihtiyacı tek noktadan karşılanır.</p>'
    ),
    # Fikir 12 - Dil Sozluk
    (
        '<p class="idea-desc">Eğitim cluster düşük KD. BTK Akademi entegrasyonu mantıklı.</p>',
        '<p class="idea-desc">Eğitim ve dil alanında bir yardımcı hub: noktalama işaretleri kuralları (40.500/ay, KD 4), iyelik eki (14.800/ay, KD 5), TDK sözlük entegrasyonu (368.000/ay), yazım kılavuzu, isim anlamı sözlüğü (6.600/ay, KD 3), eşanlamlı/zıt anlamlı sözcükler, atasözleri-deyimler, ki bağlacı yazımı. Eğitim kategorisi genellikle düşük KD\'li; brand güveni ve içerik kalitesiyle yakalanabilir. BTK Akademi içerikleriyle çapraz tanıtım yapılabilir.</p>'
    ),
    # Fikir 13 - Birim Cevirici
    (
        '<p class="idea-desc">İnç-cm, kg-libre, GB-MB (telekom-özel!), döviz, sıcaklık.</p>',
        '<p class="idea-desc">Birim çevirici mini hub: inç-cm (TV ekran ölçüleri için 100K+ küme), kg-libre, metre-km, litre-ml, celsius-fahrenheit, döviz çevirici (74.000/ay), 1 GB kaç MB (5.400/ay) gibi telekom\'a özel veri birimi çevirileri. Özellikle GB-MB çeviricisi "data paketim ne kadar yeter" sorusuna doğrudan cevap veren bir araç olarak Turkcell\'in kendi paket kullanım eğitimine bağlanabilir. TV inç sorgularında MediaMarkt dominant; ancak Turkcell TV+ köprüsüyle bu trafik içeriğe çevrilebilir.</p>'
    ),
    # Fikir 14 - Zaman Araclari
    (
        '<p class="idea-desc">Bugün/saat farkı/geri sayım widget\'ları.</p>',
        '<p class="idea-desc">Tarih ve zaman alanında ufak yardımcılar: "bugün ayın kaçı" (90.500/ay), il bazlı saat (2.900/ay), saat farkı (ABD, Avrupa, Asya), yarın hava nasıl olacak, hicri-miladi takvim çevirici, "bu yılın kaçıncı günü" tarzı sayaçlar. Tek tek hacimleri küçük ama "araç sayfası" pipeline\'ında düşük çıkış oranı sağlıyor — kullanıcı bir aracı kullanıp diğerine geçiyor.</p>'
    ),
    # Fikir 15 - Kredi Taksit
    (
        '<p class="idea-desc">Ana zor, ama niş alt sayfalar açık: cayma bedeli, vade farkı, YouTube gelir.</p>',
        '<p class="idea-desc">Kredi hesaplama ana sorgusu (aylık 1,22 milyon, KD 45) bankaların hakimiyetinde; bu yüzden niş alt sayfalara odaklanmak daha verimli. Telekom için en doğal başlık: cayma bedeli hesaplama (3.000/ay, KD 0) — turk.net bu sayfayla 1.427 trafik alıyor. Diğer niş hesaplayıcılar: vade farkı (2.400/ay), iskonto (1.000/ay, KD 17), YouTube gelir hesaplama (2.100/ay), kira artış oranı (6.600/ay, yıllık trend +%83). Paycell ürünlerine internal link kapısı açık.</p>'
    ),
    # Fikir 16 - Vergi
    (
        '<p class="idea-desc">KDV/MTV/ÖTV/gelir vergisi/damga. KOBİ blog.</p>',
        '<p class="idea-desc">Vergi alanında bir hub: KDV hesaplama (201.000/ay, KD 19 — Fikir 6 ile ortak), gelir vergisi dilimleri 2026, damga vergisi, stopaj, kurumlar vergisi, motorlu taşıtlar vergisi sorgulama (22.200/ay), ÖTV oranları, vergi mükellefliği ve vergi levhası süreçleri. Hedef kitle: KOBİ\'ler, freelancer\'lar, küçük işletme sahipleri. Turkcell\'in iş dünyasına yönelik konumlanmasını güçlendirir.</p>'
    ),
    # ===== A4 - Telekom & Sorgu (en sığ olanları zenginleştir) =====
    (
        '<p class="idea-desc">30+ how-to sayfa. Vodafone/turk.net dominant, Turkcell brand güveni daha yüksek.</p>',
        '<p class="idea-desc">Mobil cihazlarla ilgili "nasıl yapılır" tarzı 30\'dan fazla yardım sayfası: güvenli arama kapatma (20.000/ay, KD 0), NFC açma (23.000/ay küme), konum paylaşma (10.000/ay), kendi numarasını öğrenme (3.600/ay), eSIM kurulumu, 5G destekleyen telefonlar, mobil/e-imza süreci, numara taşıma rehberi, IMEI sorgulama (550.000/ay) ve gizli numara kapatma (27.100/ay). Bu sorgularda Vodafone ve turk.net dominant; Turkcell\'in brand güveni daha yüksek olduğundan içerikler doğru yapıldığında pozisyon kazanması hızlı.</p>'
    ),
    (
        '<p class="idea-desc">Superonline ürün otoritesi. Modem markaları için alt sayfalar.</p>',
        '<p class="idea-desc">İnternet ve modem konularında yardım hub\'ı: 192.168.1.1 modem arayüzü (106.000/ay), Wi-Fi şifresi değiştirme (18.100/ay), internet hız testi (1.000.000/ay), DNS ayarlama ve değiştirme (5.400/ay), VPN nedir/kullanımı, Wi-Fi sinyal güçlendirme. Superonline ürün portföyü bu hub\'a doğrudan bağlanır — fiber paket sayfalarına internal link köprüsü kurulur. Modem markalarına özel alt sayfalar (Zyxel, Huawei, ZTE) açılabilir.</p>'
    ),
    (
        '<p class="idea-desc">Turkcell brand güveni = bu cluster\'ın doğal sahibi.</p>',
        '<p class="idea-desc">Resmi sorgu hub\'ı: IBAN sorgulama (33.100/ay), HGS bakiye sorgulama (110.000/ay), MTV sorgulama (22.200/ay), plaka sorgu (74.000/ay, KD 4 — çok düşük zorluk), trafik cezası sorgulama, posta kodu sorgulama (14.800/ay), TC kimlik no sorgulama, pasaport ücretleri 2026 (18.100/ay, KD 3), yeşil pasaport ücreti (27.100/ay, KD 3), bordo pasaport (12.100/ay), il bazlı alan kodları. Turkcell\'in "iletişim ve güven" markasıyla en uyumlu cluster; e-Devlet sürecine destek konumlanması mümkün.</p>'
    ),
    (
        '<p class="idea-desc">HR/işveren kitle; e-Devlet bridge.</p>',
        '<p class="idea-desc">SGK ile ilgili konularda rehber hub\'ı: SGK işe giriş (135.000/ay, KD 20), SGK işe giriş bildirgesi nasıl alınır (11.000/ay), SGK hizmet dökümü (49.500/ay, KD 28), SGK başlangıç tarihi öğrenme, 4A/4B/4C farkları, Emekli Sandığı ve Bağ-Kur sorgulamaları, emeklilik yaşı hesaplama (3.600/ay). Hedef kitle: HR çalışanları, işveren KOBİ\'ler, yeni işe giren çalışanlar. Her sayfa e-Devlet süreçlerine yönlendiren açık bilgi sağlar.</p>'
    ),
    (
        '<p class="idea-desc">KOBİ blog için niş.</p>',
        '<p class="idea-desc">Kira ve konut süreçleri için yardımcı içerikler: kira artış oranı hesaplama (6.600/ay, yıllık trend +%83), TÜFE ile kira artışı, kira kontratı örneği (8.100/ay), dilekçe örneği (74.000/ay), tahliye dilekçesi, mazeret dilekçesi, taahhütname örneği, ev sahibine ihtar metni. KOBİ ve bireysel kullanıcı kesişiminde duran niş bir alan; Turkcell\'in iş dünyasına yönelik içerik blogunda bu başlıklar yer bulabilir.</p>'
    ),
    (
        '<p class="idea-desc">Turkcell EN DOĞAL otorite; mega trend +%9329 yıllık.</p>',
        '<p class="idea-desc">Cezaevi iletişim süreçleri Turkcell\'in en doğal otorite alanlarından biri. Ana sorgu olan "cezaevi kontör" (6.600/ay) yıllık +%9.329 büyüme gösterdi. Sayfa kapsamı: cezaevi kontör yükleme, mahkum görüşme süreçleri, telefondan mahkum görüşme, dakika alma yöntemleri, cezaevine para yatırma, e-mektup gönderme, cezaevi adres listesi. Hassas konu olduğu için ton dikkatli, bilgilendirici ve resmi kaynaklara atıflı olmalı.</p>'
    ),
    # ===== A5 - Tech Blog =====
    (
        '<p class="idea-desc">Turkcell + McAfee + BiP güvenlik servisleriyle direkt köprü.</p>',
        '<p class="idea-desc">Siber güvenlik rehberi olarak konumlanan bir hub: siber güvenlik nedir (9.900/ay, KD 10), phishing nedir ve nasıl korunulur (368.000/ay — özellikle Mart\'ta 4 milyona çıkıyor), e-posta dolandırıcılığı, oltalama saldırıları, siber zorbalık, deepfake nedir (8.100/ay), iki faktörlü kimlik doğrulama (1.900/ay), McAfee nedir (11.000/ay), AirTag takip cihazları, açık Wi-Fi güvenliği, çocuklar için internet güvenliği. Turkcell\'in McAfee ortaklığı ve BiP\'in güvenli mesajlaşma konumlanması doğrudan bu içeriklere bağlanır.</p>'
    ),
    (
        '<p class="idea-desc">turk.net 1. (2.571 trafik); Turkcell brand güvenilirliği = bu cluster\'ın doğal sahibi.</p>',
        '<p class="idea-desc">Türkiye için kritik bir başlık. Ana sorgu "deprem güvenlik bilgileri" aylık 158.000 arama alıyor (KD 3); şu an turk.net 1. sırada. Sayfa kapsamı: deprem çantası içeriği, deprem sırasında yapılacaklar, artçı deprem hazırlığı, Android ve iOS için deprem uyarı uygulamaları (4.300/ay), AFAD uyarı sistemi kaydı, Kandilli Rasathanesi son depremler, fay hatları haritası, yangın söndürme tüpü, 112 ne zaman aranır. Turkcell\'in deprem uyarı ürünleri ve sosyal sorumluluk projeleriyle uyumlu içerik konumlanması.</p>'
    ),
    (
        '<p class="idea-desc">Yüksek KD ama trend mega; long-tail kapısı açık.</p>',
        '<p class="idea-desc">Yapay zeka araçlarının kullanımına dair tanıtım ve rehber içerikleri: yapay zeka nedir (823.000/ay, KD 50 zor), ChatGPT Türkçe nasıl kullanılır (135.000/ay, KD 16), Gemini AI (450.000/ay), Claude AI (74.000/ay), DeepSeek (368.000/ay), Grok, CapCut, Canva, MidJourney, Google Lens, Google Form oluşturma (12.100/ay, KD 14). Ana KW\'lerde KD yüksek; ancak uzun kuyruk sorgular ("chatgpt türkçe ücretsiz", "deepseek nasıl kullanılır") düşük zorlukla yakalanabilir.</p>'
    ),
    (
        '<p class="idea-desc">turk.net dominant; BiP köprü içerik.</p>',
        '<p class="idea-desc">Sosyal medya uygulamaları için yardım ve rehber içerikleri: WhatsApp\'ta silinen mesajları geri getirme (16.000/ay), WhatsApp Web giriş (201.000/ay), 1 sene önceki WhatsApp mesajları, Instagram hesap dondurma/silme (20.000+ küme), Telegram Web kullanımı (170.000/ay), X (Twitter) giriş sorunları, Threads nedir (22.000/ay), TikTok rehberi, Snapchat hesap silme. Şu an turk.net dominant. BiP\'in "veri korumalı mesajlaşma" konumlanmasına köprü içerikler.</p>'
    ),
    (
        '<p class="idea-desc">Her popüler dizi için sayfa; Turkcell TV+ köprü.</p>',
        '<p class="idea-desc">Popüler diziler için "konusu, oyuncuları, yayın platformu" şablonunda Turkcell TV+ köprü içerikleri. Ahrefs verisinde turk.net her büyük dizi için tek sayfa açıp 5-20 bin trafik çekiyor: Kızılcık Şerbeti oyuncuları (105.000/ay), Eşref Rüya (79.000/ay), Taşacak Bu Deniz (80.000/ay), Wednesday sezonları (51K+9K), GTA 6 ne zaman çıkacak (16.000/ay), iPhone 16 ne zaman, Samsung S26 Ultra. Turkcell TV+ kataloğunda olan içeriklere doğrudan platforma yönlendirme yapılır.</p>'
    ),
    (
        '<p class="idea-desc">Tarife önerilerine bridge (büyük dosya = büyük paket).</p>',
        '<p class="idea-desc">Oyun severlere yönelik içerik: GTA 6 ne zaman çıkacak (73.000/ay), PS5 sistem gereksinimleri ve fiyat (88.000+/ay), Nintendo Switch (88.000/ay), PUBG kaç GB (2.400/ay), LOL kaç GB (5.700/ay), Fortnite kaç GB (4.000/ay), Valorant Mobile, sistem gereksinimleri rehberleri. Bu sorguların büyük çoğunluğu indirme boyutu/data kullanımıyla ilgili — Turkcell\'in geniş data paketlerine doğal bir bridge oluşturur.</p>'
    ),
    (
        '<p class="idea-desc">Vodafone tam bu modeli kullanıyor (/cv-hazirlama-araci, 6.295 trafik). Online CV oluşturucu + PDF çıktı.</p>',
        '<p class="idea-desc">Online CV oluşturucu — kullanıcı bilgilerini girer, sayfada yapılandırılmış CV görür, PDF olarak indirir. Vodafone\'un /cv-hazirlama-araci sayfası aylık 6.295 trafik alıyor. Ana sorgu "CV hazırlama" aylık 186.000 arama; "CV hazırlama ücretsiz" 27.100 + 20.000 varyasyonlarla 50.000\'i aşıyor. Aracın yanına özgeçmiş şablonları (3.000/ay), motivasyon mektubu örnekleri, mülakat soruları rehberi eklenir. Turkcell İK kampanyaları veya gençlik tarifeleriyle uyumlu konumlanma.</p>'
    ),
    (
        '<p class="idea-desc">Turkcell Prime ile köprü.</p>',
        '<p class="idea-desc">Çiçek, hediye ve etkinlik bilet süreçlerinde Turkcell Prime ortaklıklarına bağlayan içerikler: taze çiçek (155.000/ay), Biletix etkinlik sorguları (227.000/ay), sinema bilet fiyatları (18.100/ay), Vialand bilet (11.000/ay), uçak bileti kampanyaları, Müzekart öğrenci süreci. Doğrudan ticari trafik değil; ancak Prime ayrıcalıklarına yönlendirme noktası olarak değer üretir.</p>'
    ),
    (
        '<p class="idea-desc">Eğlence niche, lifestyle blog için.</p>',
        '<p class="idea-desc">Eğlence ağırlıklı içerik kümesi: yükselen burç hesaplama (201.000/ay, KD 10), burç uyumu, uyumlu burçlar, kahve falı (18.100/ay, KD 12), tarot falı, el falı, yıldız falı, isim anlamı (6.600/ay, KD 3), bebek ismi önerileri. Brand tonu kararı işletme tercihine bağlı. Lifestyle bloğu kapsamında konumlandırılabilir veya fizy\'nin "günlük şarkı önerisi" gibi kişiselleştirme özellikleriyle entegre edilebilir.</p>'
    ),
    # ===== A3 - Özel Gün & Mesaj =====
    (
        '<p class="idea-desc">Mayıs ayında 5M; hub + mesaj galerisi + hediye fikirleri.</p>',
        '<p class="idea-desc">"Anneler günü ne zaman" sorgusu Mayıs ayında 5 milyona çıkıyor (yıllık ortalama 673.000/ay, KD 8). Sayfa üç katmanlı: (1) "Ne zaman" bilgi başlığı, (2) Mesaj/söz/şiir galerisi (anneler günü mesajları 74.000/ay; Mayıs\'ta 1 milyon!), (3) Hediye fikirleri — Turkcell mağaza ve dijital servislere köprü. Tek hub üç farklı niyet tipini birden karşılıyor.</p>'
    ),
    (
        '<p class="idea-desc">Haziran ayında 4M. Hub + mesaj galerisi.</p>',
        '<p class="idea-desc">"Babalar günü ne zaman" sorgusu Haziran ayında 4 milyona çıkıyor (yıllık ortalama 450.000/ay). Anneler günü sayfası şablonuyla aynı yapı: ne zaman bilgisi + mesaj/söz/şiir galerileri (babalar günü mesajları 4.300/ay, babalar günü sözleri 15.000/ay) + hediye önerileri. Eşten babaya mesaj, duygusal mesaj varyasyonları gibi alt galeriler trafik çeşitliliği sağlıyor.</p>'
    ),
    (
        '<p class="idea-desc">Şubat 368K spike.</p>',
        '<p class="idea-desc">Sevgililer günü mesajları sorgusu Şubat ayında 368.000\'e çıkıyor (yıllık ortalama 33.100/ay). Sayfa: 14 Şubat ne zaman bilgisi, sevgililer günü mesajları/sözleri, sevgiliye şiir (6.600/ay), aşk şiirleri, kısa romantik mesajlar, sevgililer günü hediyesi önerileri. Turkcell\'in fizy "aşk şarkıları çalma listesi" gibi servisleriyle çapraz tanıtım fırsatı.</p>'
    ),
    (
        '<p class="idea-desc">Kasım spike, KD çok düşük.</p>',
        '<p class="idea-desc">24 Kasım Öğretmenler Günü için içerik: "öğretmenler günü için güzel sözler" (5.600/ay, KD 1 — çok düşük zorluk), "unutulmaz öğretmen sözleri" (6.600/ay, KD 0), öğretmenler günü notu (2.400/ay), kısa öğretmen sözleri, öğretmen şiirleri. Kasım ayında spike yaşar. Düşük KD ve yüksek bilgi-amaçlı niyet sebebiyle hızlı kazanım fırsatı.</p>'
    ),
    (
        '<p class="idea-desc">Mart spike.</p>',
        '<p class="idea-desc">8 Mart Dünya Kadınlar Günü için içerik: 8 mart ne zaman, kadınlar günü hangi gün, dünya kadınlar günü tarihçesi, 8 mart mesajları, kadınlar günü sözleri ve şiirleri, kadınlar günü hediye önerileri. Mart ayında spike yaşar; sayfa yıl boyu açık tutulup içeriği her yıl yenilendiğinde kalıcı trafik kaynağı olur.</p>'
    ),
    (
        '<p class="idea-desc">Nisan spike.</p>',
        '<p class="idea-desc">23 Nisan Ulusal Egemenlik ve Çocuk Bayramı için içerik: 23 nisan ne zaman, mesajları, şiirleri, ulusal egemenlik bayramı tarihçesi, çocuk bayramı mesajları, çocuk şiirleri, 23 nisan hediye önerileri. Nisan ayında spike yaşar; ebeveynler ve eğitimciler ana hedef kitle.</p>'
    ),
    (
        '<p class="idea-desc">Yıl boyu sabit hacim. BiP/mesajlaşma servisleriyle natürel köprü.</p>',
        '<p class="idea-desc">"Doğum günü mesajları" sorgusu aylık 201.000 arama ile yıl boyu sabit bir hacme sahip — özel gün başlığı değil evergreen bir başlık. Alt galeriler: eşe, anneye, babaya, arkadaşa, sevgiliye, kardeşe, çocuğa doğum günü mesajları; komik/duygusal varyantlar; iyi ki doğdun mesajları; doğum günü hediyesi (26.000/ay). BiP\'in mesajlaşma konumlanmasıyla doğal köprü kurulur.</p>'
    ),
    (
        '<p class="idea-desc">EN DÜŞÜK KD! 135K+40K+165K+110K = 450K+/ay; sosyal medya paylaşım kitlesi.</p>',
        '<p class="idea-desc">Günlük paylaşılan mesaj türlerinin tamamı için tek bir galeri hub\'ı: günaydın mesajları (135.000/ay, KD 4 — en düşük zorluk), iyi geceler mesajları (40.500/ay), güzel sözler (165.000/ay, KD 11), anlamlı sözler (110.000/ay, KD 4), duygusal sözler, motivasyon sözleri, taziye/başsağlığı/geçmiş olsun/şifa mesajları. Toplam aylık 450.000\'i aşan hacim ve çok düşük zorluk kombinasyonu — hızlı kazanım için en güçlü kümelerden biri.</p>'
    ),
    (
        '<p class="idea-desc">Ramazan bayramı mesajları 74K (Mart 823K!). Spike aylar Mart/Ekim/Aralık.</p>',
        '<p class="idea-desc">Bayram mesajları hub\'ı: bayram mesajları (14.800/ay; Mart\'ta 110.000), ramazan bayramı mesajları (74.000/ay; Mart\'ta 823.000!), kurban bayramı mesajları, kandil mesajları, yeni yıl ve yılbaşı mesajları, 29 Ekim mesajları (Ekim\'de 12.000). Spike aylar Mart, Ekim ve Aralık. Bu sayfa A2 (Resmi Tatil) sayfasının "mesaj" katmanı olarak konumlandırılabilir veya bağımsız bir galeri olarak açılabilir.</p>'
    ),
]

count_replaced = 0
for old, new in desc_replacements:
    if old in text:
        text = text.replace(old, new)
        count_replaced += 1
    else:
        print(f"MISS desc: {old[20:80]}")

print(f"\nReplaced {count_replaced}/{len(desc_replacements)} idea descriptions")
HTML.write_text(text, encoding="utf-8")
print(f"HTML size: {len(text):,} chars (delta: {len(text)-orig_len:+,})")
