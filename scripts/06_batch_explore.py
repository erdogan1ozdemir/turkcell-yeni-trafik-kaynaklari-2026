"""
Tum 60.715 Turkcell-yok rakip-var fursati batch batch incele.
Hacim filtresi yok - dusuk hacimli niche firsatlar dahil.
Amac: Eksik kategorileri ve niche evergreen pattern'leri yakalamak.
"""
import pandas as pd
import re
from collections import Counter

CSV_PATH = "/Users/Erdo/Desktop/Claude Projects/Turkcel/www.turkcell.com.tr-content-gap-subdomains-t_2026-05-13_14-14-31.csv"

df = pd.read_csv(CSV_PATH, low_memory=False)
competitors = ["turktelekom.com.tr", "vodafone.com.tr", "mediamarkt.com.tr",
               "vatanbilgisayar.com", "turk.net", "pttcell.com.tr"]
df = df[df["www.turkcell.com.tr/: URL"].isna()].copy()

# En az bir rakip rank
def has_comp(row):
    for comp in competitors:
        if pd.notna(row.get(f"{comp}/: Organic Position", None)):
            return True
    return False

df = df[df.apply(has_comp, axis=1)].copy()
print(f"Toplam Turkcell-yok rakip-var: {len(df)}")
print(f"Volume >= 10: {len(df[df['Volume'] >= 10])}")
print(f"Volume >= 50: {len(df[df['Volume'] >= 50])}")
print(f"Volume >= 100: {len(df[df['Volume'] >= 100])}")

# Niche pattern aramasi - genis ag at
NICHE_PATTERNS = {
    # Hesaplama dunyasi
    "HESAPLAMA_DUNYA": [
        r"\bsayac\b", r"\bsay[ıi]c[ıi]\b", r"\bhesap\w+", r"\bhesaplay\w+",
        r"\b(yüzde|yuzde|oran|fark|toplam|farkı) hesap",
        r"\b(kdv|otv|ötv|stopaj|damga vergisi|gelir vergisi|kurumlar vergisi)\b",
        r"\b(brüt|net|maaş|maas|asgari) ?(ücret|hesap)",
        r"\b(kıdem|kidem|ihbar) tazmin",
        r"\b(emekli|emeklilik) (yaşı|hesap|maaş)",
        r"\b(yıllık izin|yillik izin|izin g[üu]nü|izin g[üu]nleri)",
        r"\b(fazla mesai|gece mesai|hafta sonu mesai)",
        r"\b(yıl sonu notu|not ortalama|gpa|aort|üni\.? ort)",
        r"\b(bmi|vücut kitle|ideal kilo|metabolizma|kalori|protein|su tüketim)",
        r"\b(hamilelik|gebelik|ovulasyon|regl|aşı takvim|menapoz)",
        r"\b(doğum tarihi|dogum tarihi|kaç haftalık|kac haftalik)\b",
        r"\b(burç|burc|yükselen|astroloji|yıldız haritası|doğum haritası)",
        r"\b(kredi|taksit|faiz|vade) hesap",
        r"\b(iban|swift) (hesap|sorgu|olu[şs]tur)",
        r"\b(döviz|altın|gümüş|euro|dolar) (kuru|fiyat|hesap|çevir)",
        r"\b(benzin|motorin|lpg|elektrik|do[ğg]algaz) (fiyat|hesap|tüketim)",
        r"\b(yol|mesafe|km|rota) hesap",
        r"\b(saat fark|zaman fark|gmt|utc)",
        r"\b(yıl|gün|hafta|ay) (sayac|hesap)",
        r"\b(geri sayım|geri sayim|countdown)",
        r"\b(rastgele|şans|piyango|loto)",
        r"\b(faktöriyel|yuvarla|yüzdelik dilim)",
        r"\b(kira|aidat|tapu|emlak vergi) hesap",
        r"\bdaire hesap", r"\bnoter (ücret|hesap)",
        r"\btoki (başvuru|hesap)",
    ],

    # Birim donusum
    "BIRIM_DONUSUM": [
        r"\b(inç|inc|inch) (kaç|kac) ?(cm|ekran)",
        r"\b(cm|metre|km|mm|mil) (kaç|kac)",
        r"\b(gb|mb|tb|kb) (kaç|kac)",
        r"\b(kg|gram|libre|pound|ton) (kaç|kac)",
        r"\b(litre|ml|gallon|fincan|bardak|kase) (kaç|kac)",
        r"\b(\d+) ?(inç|cm|kg|gram|metre|km|ml|litre|tl|usd|eur|gb|mb)",
        r"\b(celsius|fahrenheit|kelvin|santigrat)",
        r"\b(saat|dakika|saniye) (çevir|hesap|kaç)",
        r"\bderece (çevir|kaç)",
    ],

    # Resmi tatil ve bayram
    "TATIL_BAYRAM_OZEL_GUN": [
        r"\bresmî tatil\b", r"\bresmi tatil\b", r"\btatil mi\b",
        r"\b(kurban|ramazan|şeker|berat|mevlit|miraç|regaip) (bayram|kandil)",
        r"\barefe\b", r"\bayrefe\b", r"\barife\b",
        r"\b(yılbaş[ıi]|yilbasi|yeni yıl|yeni yil)",
        r"\b(cumhuriyet|zafer|gençlik|egemenlik|barış|çocuk) bayram",
        r"\b(29 ekim|23 nisan|19 mayıs|30 ağustos|10 kasım|18 mart)",
        r"\b(yarı yıl|yariyil|yari yil|sömestr|somestr|yaz tatili|kış tatili)",
        r"\b(black friday|kara cuma|cyber monday|valentine)",
        r"\b(anneler|babalar|sevgililer|öğretmenler|çocuklar|kadınlar) g[uü]n[uü]",
        r"\b(8 mart|14 şubat|24 kasım|1 mayıs)",
        r"\b(nevruz|hıdırellez|hidirellez)",
        r"\b(geri sayım|kaç gün kaldı) (ramazan|kurban|bayram|yılbaş|tatil)",
    ],

    # Mesaj soz siir kart
    "MESAJ_SOZ_KART": [
        r"\b(mesaj|söz|sozler|şiir|siir|söy|yazi|notu) ?(ı|i|ları|leri|larım)?\b.*\b(günü|bayram|sevgili|anne|baba|eş|kardeş|arkadaş|öğretmen|patron)",
        r"\b(en güzel|guzel|duygusal|romantik|kısa|uzun|anlamlı|anlamli|komik|esprili) (mesaj|söz|sözler|şiir)",
        r"\b(günaydın|iyi geceler|tatlı uykular|hayırlı sabahlar) (mesaj|söz)",
        r"\b(taziye|geçmiş olsun|şifa|başsağlığı) (mesaj|söz)",
        r"\b(düğün|nişan|kına|söz|bebek|bilek) (kart|davetiye|mesaj)",
        r"\b(doğum günü|yas günü) (hediyesi|mesaj|kart)",
    ],

    # Tanim ve ne demek
    "TANIM_NEDIR": [
        r"\bnedir\b", r"\bne demek\b", r"\bne anlama gelir\b", r"\bne işe yarar\b",
        r"\bne ise yarar\b", r"\baçılımı\b", r"\bacilimi\b", r"\bkısaltma\b",
        r"\bne anlama gelir\b",
    ],

    # Nasil
    "NASIL_YAPILIR": [r"\bnas[ıi]l\b"],

    # Ne zaman
    "NE_ZAMAN": [r"\bne zaman\b", r"\bhangi g[uü]n\b", r"\bhangi tarih\b", r"\bhangi ay\b"],

    # Kac (how many)
    "KAC": [r"\bkaç\b", r"\bkac\b"],

    # Sosyal medya / dijital uygulamalar
    "SOSYAL_MEDYA_APP": [
        r"\b(whatsapp|wp|whatsapp web)\b",
        r"\b(instagram|insta|ig)\b",
        r"\b(facebook|fb)\b",
        r"\b(twitter|x|x giriş)\b",
        r"\b(telegram|telegram web)\b",
        r"\b(snapchat|snap)\b",
        r"\b(tiktok|tik tok)\b",
        r"\b(linkedin|reddit|pinterest|discord)\b",
        r"\b(youtube|yt)\b", r"\b(netflix|disney\+|exxen|tabii|blutv|gain)\b",
        r"\b(spotify|apple music|fizy|tidal)\b",
        r"\b(bip|whatsapp|signal|wechat)\b",
        r"\b(zoom|teams|skype|meet)\b",
        r"\b(canva|figma|photoshop|illustrator)\b",
        r"\b(chatgpt|gemini|claude|grok|copilot|deepseek)\b",
    ],

    # Telefon / cihaz
    "TELEFON_CIHAZ": [
        r"\biphone\b", r"\bsamsung\b", r"\bxiaomi\b", r"\bhuawei\b", r"\boppo\b",
        r"\bvivo\b", r"\bredmi\b", r"\bhonor\b", r"\brealme\b", r"\boneplus\b",
        r"\b(akıllı saat|akilli saat|smart watch)",
        r"\b(kulaklık|kulaklik|airpods|earbuds)",
        r"\b(şarj|sarj) (cihaz|aleti|adaptör)",
        r"\b(power bank|powerbank|taşınabilir şarj)",
        r"\b(kılıf|kilif|ekran koruyucu|cam)",
        r"\b(sim|esim|nano sim|micro sim)",
        r"\b(imei|seri no|model no)",
    ],

    # Telefon ayar / sorun
    "TELEFON_AYAR_SORUN": [
        r"\b(güvenli arama|safe search|family link)",
        r"\b(nfc|airdrop|airtag|airpods)",
        r"\b(ekran görüntüsü|screenshot|ekran kaydı|screen record)",
        r"\b(face id|touch id|parmak izi|şifre|sifre)",
        r"\b(yedek|backup|icloud|google drive|onedrive)",
        r"\b(format|sıfırla|fabrika ayar)",
        r"\b(güncelleme|update|ios|android)",
        r"\b(virüs|virus|malware|antivirus)",
        r"\b(konum|location|gps)",
    ],

    # Internet / modem / WiFi
    "INTERNET_MODEM_WIFI": [
        r"\b(internet|wi[- ]?fi|modem|router)",
        r"\b(hız testi|hiz testi|speed test|ping)",
        r"\b(192\.168\.|10\.0\.|172\.16\.)",
        r"\bdns\b", r"\bip adres", r"\bping\b", r"\bvpn\b",
        r"\b(fiber|adsl|vdsl|kablo internet)",
    ],

    # Saglik
    "SAGLIK_VUCUT": [
        r"\b(belirti|semptom|hastalık|hastalik|teşhis|teshis)",
        r"\b(ağrı|agri|sancı|sanci)",
        r"\b(soğuk algınlığı|grip|nezle|sinüzit|alerji)",
        r"\b(diyabet|hipertansiyon|kolesterol|tansiyon)",
        r"\b(hamilelik|gebelik|loğusalık|emzirme|bebek)",
        r"\b(diş|dis|kulak|göz|gozluk) (ağrı|temizlik|implant|muayene)",
        r"\b(astım|astim|akciğer|kalp|böbrek|karaciğer|safra)",
        r"\b(panik atak|anksiyete|depresyon|stres|uyku)",
        r"\b(vitamin|mineral|takviye|d3|b12|c vitamini)",
        r"\b(diyet|kilo verme|kilo alma|protein)",
    ],

    # Yemek tarifi
    "YEMEK_TARIF": [
        r"\btarif(i)?\b", r"\bnasıl yap[ıi]l[ıi]r\b",
        r"\b(corba|çorba|salata|tatlı|tatli|kek|pasta|börek|borek|pilav|et yemeği|tavuk yemeği)",
        r"\b(makarna|mantı|köfte|kebap|kavurma|musakka|dolma|sarma)",
        r"\b(çay|cay|kahve|smoothie|kokteyl|içecek)",
    ],

    # E-devlet / sorgulama / resmi islem
    "EDEVLET_SORGULAMA": [
        r"\be-devlet\b", r"\bedevlet\b", r"\btc kimlik\b", r"\bnüfus\b",
        r"\b(pasaport|ehliyet|sgk|emekli sandığı)",
        r"\b(borç sorgu|borc sorgu|trafik cezası|hgs|ogs|mtv)",
        r"\b(plaka sorgu|araç sorgu|arac sorgu|muayene tarihi)",
        r"\b(barkod|qr|karekod) sorgu",
        r"\b(adli sicil|sabıka kayıt|ikametgâh)",
        r"\b(askerlik|bedelli|sevk|tecil)",
        r"\b(başvuru|basvuru|form|dilekçe|dilekce)",
        r"\b(noter|tapu|gemlik|nüfus müdürlüğü)",
        r"\b(vergi numarası|gelir vergisi|kurumlar vergisi)",
        r"\b(posta kodu|alan kodu|plaka kodu|telefon kodu)",
    ],

    # Otomotiv
    "OTOMOTIV": [
        r"\b(araba|otomobil|araç|arac|otomotiv|otomotİv)",
        r"\b(motor|motorsiklet|scooter|bisiklet)",
        r"\b(plaka|trafik|muayene|sigorta|kasko)",
        r"\b(sürücü|surucu|ehliyet) (kurs|sınav|belge)",
        r"\b(benzinli|dizel|hibrit|elektrikli) (araba|otomobil)",
        r"\b(yedek parça|yedek parca|fren|debriyaj|amortisör)",
    ],

    # Egitim / sinav
    "EGITIM_SINAV": [
        r"\b(yks|tyt|ayt|lgs|kpss|ales|yds|yökdil)",
        r"\b(ösym|osym|meb|yök|yok)",
        r"\b(sınav takvim|sinav takvim|sınav tarih|tercih kılavuz)",
        r"\b(üniversite|universite|lise|ortaokul|ilkokul) (puan|tercih|sınav)",
        r"\b(ders çalışma|ders calisma|konu anlatım|özet|ozet|soru çözüm)",
        r"\b(eba|btk akademi|udemy|coursera)",
        r"\b(diploma|sertifika|burs)",
    ],

    # Burç astroloji
    "BURC_ASTROLOJI": [
        r"\b(burç|burc|astroloji|yıldız|zodyak|horoskop)",
        r"\b(koç|boğa|ikizler|yengeç|aslan|başak|terazi|akrep|yay|oğlak|kova|balık) burc",
        r"\b(yükselen|yukselen|ay|güneş|gunes) burc",
        r"\b(doğum haritası|dogum haritasi)",
        r"\b(uyumlu burç|uyumsuz burç|burç eşleşmesi)",
    ],

    # Spor
    "SPOR": [
        r"\b(maç|mac|skor|fikstür|puan durumu|lig|şampiyona)",
        r"\b(fenerbahçe|fenerbahce|galatasaray|beşiktaş|besiktas|trabzonspor)",
        r"\b(milli takım|euro|dünya kupası|şampiyonlar ligi|avrupa ligi)",
        r"\b(futbol|basketbol|voleybol|hentbol|tenis|yüzme)",
        r"\b(formula|f1|nascar)",
        r"\b(antrenman|kondisyon|fitness|yoga|pilates)",
    ],

    # Oyun
    "OYUN_REHBER": [
        r"\b(gta|pubg|lol|valorant|fortnite|minecraft|fifa|efootball)",
        r"\b(playstation|ps[345]|xbox|nintendo|switch)",
        r"\b(steam|epic games|origin|battle\.net)",
        r"\b(roblox|among us|free fire|call of duty|cod)",
        r"\bkonsol\b", r"\boyun (kolu|seti|simülasyon)",
    ],

    # Dizi/film
    "DIZI_FILM": [
        r"\b(dizi|film|sezon|bölüm|bolum) (konusu|oyuncu|özet|özet|yayın)",
        r"\b(yapımcı|yönetmen|senaryo|kanal)",
        r"\bbilim kurgu\b", r"\bromantik komedi\b", r"\bgerilim\b",
    ],

    # Klavye sembol
    "KLAVYE_SEMBOL": [
        r"\b(et işareti|et isareti|@|#|&|%|\\$)",
        r"\b(noktalama|virgül|nokta|ünlem|soru işareti|tırnak|ayraç)",
        r"\b(klavye|tuş|tus|combo|kısayol|kisayol)",
        r"\b(emoji|emojikler|ifade)",
        r"\b(büyük harf|kucuk harf|caps lock)",
    ],

    # Kelime / yazim / dil
    "KELIME_YAZIM_DIL": [
        r"\b(eşanlamlı|esanlamli|zıt anlamlı|zit anlamli|terim sözlük)",
        r"\b(noktalı virgül|ki bağlacı|de da bağlacı|herşey her şey)",
        r"\b(ingilizce|almanca|fransızca|arapça|rusça|çince) (öğren|ogren|cümle)",
        r"\b(çeviri|ceviri|sözlük|sozluk|kelime anlam)",
        r"\b(yazım|yazim) (kuralları|kontrol|kılavuzu|kilavuzu)",
    ],

    # Hayvanlar / evcil
    "HAYVANLAR": [
        r"\b(köpek|kopek|kedi|kuş|kus|balık|balik|kuş bakımı)",
        r"\b(mama|yem|kafes|tasma)",
        r"\b(yavru|yetişkin|yaşlı) (kedi|köpek)",
    ],

    # Bitkiler / bahce
    "BITKILER_BAHCE": [
        r"\b(çiçek|cicek|bitki|fidan|tohum|sera|saksı)",
        r"\b(orkide|kaktüs|gül|menekşe|begonya|sakız|aloe)",
        r"\b(toprak|gübre|sulama|budama|yetiştirme)",
    ],

    # Ruya tabiri
    "RUYA_TABIRI": [
        r"\b(rüya|ruya) ?(tabir|görmek|gormek|yorum)",
        r"\b(rüyada|ruyada)",
    ],

    # Astronomi
    "ASTRONOMI": [
        r"\b(ay tutulması|güneş tutulması|gunes tutulmasi|kuyruklu yıldız)",
        r"\b(gezegen|asteroid|takım yıldız|samanyolu)",
        r"\b(uzay|nasa|spacex|mars|venüs|jüpiter|satürn|uranüs|neptün)",
        r"\b(meteor|şahap|göktaşı|göktaşı yağmuru)",
    ],

    # Konser / etkinlik / mekan
    "KONSER_ETKINLIK": [
        r"\b(konser|festival|etkinlik|tiyatro|opera|bale)",
        r"\b(bilet|ticket|biletix|passo)",
        r"\b(sergi|müze|fuar)",
    ],

    # Seyahat / ulasim
    "SEYAHAT_ULASIM": [
        r"\b(uçak|ucak) (bileti|firma|kalkış|iniş)",
        r"\b(otobüs|otobus) (bileti|firma|sefer)",
        r"\b(tren|hızlı tren|yht) (bileti|sefer)",
        r"\b(otel|tatil köyü|all inclusive|kamp)",
        r"\b(vize|pasaport|gümrük|seyahat sigortası)",
        r"\b(yurt dışı|yurtdışı|yurtiçi) (seyahat|tatil|gezi)",
        r"\b(navlun|kargo|posta) (takip|fiyat)",
    ],

    # Iş başvuru / cv
    "IS_BASVURU_CV": [
        r"\b(cv|özgeçmiş|ozgecmis) (hazırlama|hazirlama|şablon|sablon|örnek|ornek)",
        r"\b(iş başvuru|is basvuru|işe alım|mülakat|mulakat) (sorular|tüyolar)",
        r"\b(staj|kariyer|iş ilanları|is ilanlari)",
        r"\b(linkedin|linkedlin) (profil|optimizasyon)",
        r"\b(motivasyon mektup|kapak yazısı|niyet mektup)",
    ],

    # Finans / Bankacilik / Yatirim
    "FINANS_BANKA_YATIRIM": [
        r"\b(banka|bankacılık|bankacilik|atm|kart)",
        r"\b(kredi notu|findeks|sicil|borç sorgu|borc sorgu)",
        r"\b(yatırım|yatirim|borsa|hisse senedi|bist|tahvil|fon|portföy)",
        r"\b(kripto|bitcoin|ethereum|nft|blockchain|metaverse)",
        r"\b(altın fiyatı|gümüş fiyatı|döviz kuru|dolar tl|euro tl)",
        r"\b(faiz oran|enflasyon|tüfe|enflasyon hesap|reeskont)",
        r"\b(emekli|individual emeklilik|bes)",
    ],

    # Kanun / Hukuk
    "HUKUK_KANUN": [
        r"\b(kanun|yasa|tüzük|yönetmelik)",
        r"\b(boşanma|nafaka|velayet|alacak)",
        r"\b(iş hukuku|sözleşme|fesih)",
        r"\b(miras|tereke|veraset)",
        r"\b(ceza|tazminat|şikayet|tüketici)",
    ],

    # Ev / dekor / mobilya
    "EV_DEKOR_MOBILYA": [
        r"\b(mobilya|koltuk|yatak|masa|sandalye|gardırop)",
        r"\b(perde|halı|hali|kilim)",
        r"\b(boya|duvar kağıdı|duvar kagidi|tadilat|kalıcı)",
        r"\b(banyo|mutfak|salon|yatak odası) (dekor|tasarım)",
    ],

    # Beauty / kozmetik
    "BEAUTY_KOZMETIK": [
        r"\b(makyaj|fondoten|ruj|maskara|allık|göz kalemi)",
        r"\b(cilt bakım|nemlendir|krem|serum|ton|spf)",
        r"\b(saç bakım|saç boya|saç şekillendir|saç tipi)",
        r"\b(tırnak|manikür|pedikür|ojesi)",
        r"\b(parfüm|deodorant|koku)",
    ],

    # Cocuk / bebek
    "COCUK_BEBEK": [
        r"\bbebek bakım", r"\bemzir", r"\bemzirme",
        r"\b(çocuk gelişimi|cocuk gelisimi|montessori|ana sınıfı)",
        r"\b(beslenme önerileri|mama tarif|pure)",
    ],
}

def categorize_multi(keyword):
    kw = str(keyword).lower()
    matches = []
    for cname, patterns in NICHE_PATTERNS.items():
        if any(re.search(p, kw) for p in patterns):
            matches.append(cname)
    return matches

df["niche_cats"] = df["Keyword"].apply(categorize_multi)
df["primary_niche"] = df["niche_cats"].apply(lambda x: x[0] if x else "ZZ_DIGER")
df["matched_cats_count"] = df["niche_cats"].apply(len)

# Tum hacim seviyelerinde cluster ozeti
print("\n=== NICHE CLUSTER OZETI (TUM HACIM) ===")
ozet = df.groupby("primary_niche").agg(
    kw=("Keyword", "count"),
    toplam_hacim=("Volume", "sum"),
    ort_hacim=("Volume", "mean"),
    max_hacim=("Volume", "max"),
).sort_values("toplam_hacim", ascending=False)
print(ozet.head(50).to_string())

# Volume seviyelerine gore breakdown
print("\n=== HACIM SEVIYESI BAZINDA ===")
for thr in [0, 50, 100, 200, 500, 1000, 5000]:
    print(f"  Volume >= {thr}: {len(df[df['Volume'] >= thr])} keyword")

# Cikti kaydet - tum 60K firsat clustering ile
df[["Keyword", "Volume", "KD", "CPC", "SERP features", "primary_niche", "niche_cats"]].to_csv(
    "/Users/Erdo/Desktop/Claude Projects/Turkcel/output/niche_all_opportunities.csv", index=False)

# Her niche icin top 200 keyword (sort by volume)
sorted_df = df.sort_values("Volume", ascending=False)
for cname in NICHE_PATTERNS.keys():
    sub = sorted_df[sorted_df["primary_niche"] == cname].head(200)
    if len(sub) > 0:
        sub[["Keyword", "Volume", "KD"]].to_csv(
            f"/Users/Erdo/Desktop/Claude Projects/Turkcel/output/niche_{cname}.csv", index=False)
print(f"\n{len(NICHE_PATTERNS)} niche cluster icin top 200 keyword kaydedildi.")

# DIGER kategorisi - low volume da dahil
diger = sorted_df[sorted_df["primary_niche"] == "ZZ_DIGER"].copy()
print(f"\nZZ_DIGER (siniflandirilmadi): {len(diger)} keyword, toplam hacim: {diger['Volume'].sum():,}")
# Diger'in en yuksek 500'unu kaydet
diger.head(500)[["Keyword", "Volume", "KD"]].to_csv(
    "/Users/Erdo/Desktop/Claude Projects/Turkcel/output/niche_ZZ_DIGER_top500.csv", index=False)
