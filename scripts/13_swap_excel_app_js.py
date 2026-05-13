"""
Eski 'INLINE EXCEL VIEWER' JS blogunu yeni Excel App JS'i ile degistir.
- Open/close overlay
- Sheet kategorize listele (sidebar)
- Sheet -> rapor mapping (geri linkler)
- Tum data-xls-sheet ve data-xls-open butonlarini yeni open fonksiyonuna bagla
"""
from pathlib import Path

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")

js_start_marker = "  // =========================================================\n  // INLINE EXCEL VIEWER (parses embedded base64 xlsx with SheetJS)\n  // =========================================================\n"
js_end_marker = "  if (xlsSection) xlsObserver.observe(xlsSection);"

s = text.find(js_start_marker)
e = text.find(js_end_marker)
if s < 0 or e < 0:
    print("Markers not found", s, e)
    raise SystemExit(1)

e_end = e + len(js_end_marker)
print(f"Replacing JS block of {e_end - s} chars")

NEW_JS = """  // =========================================================
  // EXCEL APP MODE - Full-screen separate UI (parses embedded base64 xlsx)
  // =========================================================
  const xaOverlay = document.getElementById('excelApp');
  const xaBack = document.getElementById('xaBack');
  const xaClose = document.getElementById('xaClose');
  const xaSidebar = document.getElementById('xaSidebar');
  const xaSidebarToggle = document.getElementById('xaSidebarToggle');
  const xaSearch = document.getElementById('xaSearch');
  const xaData = document.getElementById('xaData');
  const xaSheetTitle = document.getElementById('xaSheetTitle');
  const xaMetaRows = document.getElementById('xaMetaRows');
  const xaMetaCols = document.getElementById('xaMetaCols');
  const xaMetaFilter = document.getElementById('xaMetaFilter');
  const xaPrev = document.getElementById('xaPrev');
  const xaNext = document.getElementById('xaNext');
  const xaPageInfo = document.getElementById('xaPageInfo');
  const xaRelated = document.getElementById('xaRelated');
  const xaDownload = document.getElementById('xaDownload');

  // Sheet alias -> sheet name mapping (rapordaki butonlar icin)
  const SHEET_ALIASES = {
    'ozet': '00_Yonetici_Ozeti',
    'a1': '01_A1_Hesaplama_Fikirleri',
    'a2': '02_A2_Tatil_Bayram_Fikirleri',
    'a3': '03_A3_OzelGun_Mesaj_Fikirleri',
    'a4': '04_A4_Telekom_Sorgu_Fikirleri',
    'a5': '05_A5_AI_Siber_Diger',
    'a6': '06_A6_2026_Plan_Perspektif',
    'a6_full': '06_A6_2026_Plan_Perspektif',
    'cluster': '07_B1_Cluster_Ozeti',
    'rakip': '08_B2_Rakip_URL_Patternleri',
    'dfs': '09_DfS_Dogrulama_TR',
    'hesaplama': '09_n_HESAPLAMA_DUNYA',
    'tatil': '10_n_TATIL_BAYRAM_OZEL_GUN',
    'telefon_ayar': '11_n_TELEFON_AYAR_SORUN',
    'internet': '12_n_INTERNET_MODEM_WIFI',
    'edevlet': '13_n_EDEVLET_SORGULAMA',
    'mesaj': '14_n_MESAJ_SOZ_KART',
    'nasil': '15_n_NASIL_YAPILIR',
    'ne_zaman': '16_n_NE_ZAMAN',
    'nedir': '17_n_TANIM_NEDIR',
    'birim': '18_n_BIRIM_DONUSUM',
    'kac': '19_n_KAC',
    'klavye': '20_n_KLAVYE_SEMBOL',
    'kelime': '21_n_KELIME_YAZIM_DIL',
    'cv': '22_n_IS_BASVURU_CV',
    'hukuk': '23_n_HUKUK_KANUN',
    'saglik': '24_n_SAGLIK_VUCUT',
    'konser': '25_n_KONSER_ETKINLIK',
    'finans': '26_n_FINANS_BANKA_YATIRIM',
    'burc': '27_n_BURC_ASTROLOJI',
    'egitim': '28_n_EGITIM_SINAV',
    'otomotiv': '29_n_OTOMOTIV',
    'spor': '30_n_SPOR',
    'oyun': '31_n_OYUN_REHBER',
    'dizi': '32_n_DIZI_FILM',
    'seyahat': '33_n_SEYAHAT_ULASIM',
    'cocuk': '34_n_COCUK_BEBEK',
    'hayvan': '35_n_HAYVANLAR',
    'bitki': '36_n_BITKILER_BAHCE',
    'beauty': '37_n_BEAUTY_KOZMETIK',
    'ev': '38_n_EV_DEKOR_MOBILYA',
    'yemek': '39_n_YEMEK_TARIF',
    'sosyal': '40_n_SOSYAL_MEDYA_APP',
    'cihaz': '41_n_TELEFON_CIHAZ',
    'astronomi': '42_n_ASTRONOMI',
    'diger': '43_n_DIGER_top500',
    'tum': '44_TUM_60K_FIRSATLAR'
  };

  // Sidebar kategorileri
  const SIDEBAR_GROUPS = [
    { title: 'Yönetici Özeti', sheets: ['00_Yonetici_Ozeti'] },
    { title: 'A1-A6 Fikir Kartları', sheets: [
      '01_A1_Hesaplama_Fikirleri',
      '02_A2_Tatil_Bayram_Fikirleri',
      '03_A3_OzelGun_Mesaj_Fikirleri',
      '04_A4_Telekom_Sorgu_Fikirleri',
      '05_A5_AI_Siber_Diger',
      '06_A6_2026_Plan_Perspektif'
    ]},
    { title: 'Veri & Doğrulama', sheets: [
      '07_B1_Cluster_Ozeti',
      '08_B2_Rakip_URL_Patternleri',
      '09_DfS_Dogrulama_TR'
    ]},
    { title: 'Niche Cluster\\'lar (35)', sheets: 'niche' },  // dinamik
    { title: 'Diğer & Tüm Fırsatlar', sheets: [
      '43_n_DIGER_top500',
      '44_TUM_60K_FIRSATLAR'
    ]}
  ];

  // Sheet -> rapor bolumu geri-link mapping
  // Her sheet hangi rapor section(lar)inde mentions oluyor
  const SHEET_TO_REPORT = {
    '00_Yonetici_Ozeti': [
      { id: 'top', label: 'Yönetici Özeti' },
      { id: 'hizalama', label: '2026 Plan Hizalama' }
    ],
    '01_A1_Hesaplama_Fikirleri': [
      { id: 'a1', label: 'Bölüm 04 · A1 Hesaplama' }
    ],
    '02_A2_Tatil_Bayram_Fikirleri': [
      { id: 'a2', label: 'Bölüm 05 · A2 Resmi Tatil (mevcut sayfa)' }
    ],
    '03_A3_OzelGun_Mesaj_Fikirleri': [
      { id: 'a3', label: 'Bölüm 06 · A3 Özel Gün & Mesaj' }
    ],
    '04_A4_Telekom_Sorgu_Fikirleri': [
      { id: 'a4', label: 'Bölüm 07 · A4 Telekom & Sorgu' }
    ],
    '05_A5_AI_Siber_Diger': [
      { id: 'a5', label: 'Bölüm 08 · A5 AI/Siber/Diğer' }
    ],
    '06_A6_2026_Plan_Perspektif': [
      { id: 'a6', label: 'Bölüm 03 · A6 2026 Plan Perspektifleri' }
    ],
    '07_B1_Cluster_Ozeti': [
      { id: 'rakip', label: 'Bölüm 09 · Rakip Pattern' },
      { id: 'top-firsatlar', label: 'Bölüm 02 · Top Fırsatlar' }
    ],
    '08_B2_Rakip_URL_Patternleri': [
      { id: 'rakip', label: 'Bölüm 09 · Rakip Pattern Analizi' }
    ],
    '09_DfS_Dogrulama_TR': [
      { id: 'dfs', label: 'Bölüm 10 · DfS Doğrulama' },
      { id: 'top-firsatlar', label: 'Bölüm 02 · Top Fırsatlar' }
    ],
    '09_n_HESAPLAMA_DUNYA': [
      { id: 'a1', label: 'A1 · Hesaplama Hub' }
    ],
    '10_n_TATIL_BAYRAM_OZEL_GUN': [
      { id: 'a2', label: 'A2 · Resmi Tatil Hub' },
      { id: 'a3', label: 'A3 · Özel Gün & Mesaj' }
    ],
    '11_n_TELEFON_AYAR_SORUN': [
      { id: 'a4', label: 'A4 · Telekom Yardım' }
    ],
    '12_n_INTERNET_MODEM_WIFI': [
      { id: 'a4', label: 'A4 · Telekom Yardım' }
    ],
    '13_n_EDEVLET_SORGULAMA': [
      { id: 'a4', label: 'A4 · Sorgu Rehberi' }
    ],
    '14_n_MESAJ_SOZ_KART': [
      { id: 'a3', label: 'A3 · Mesaj/Söz' }
    ],
    '15_n_NASIL_YAPILIR': [
      { id: 'a4', label: 'A4 · Telekom How-to' },
      { id: 'a5', label: 'A5 · Tech Blog' }
    ],
    '16_n_NE_ZAMAN': [
      { id: 'a2', label: 'A2 · Tatil/Bayram' },
      { id: 'a3', label: 'A3 · Özel Gün' }
    ],
    '17_n_TANIM_NEDIR': [
      { id: 'a6', label: 'A6.6 · Teknoloji Sözlüğü' },
      { id: 'a5', label: 'A5 · Tech Blog' }
    ],
    '18_n_BIRIM_DONUSUM': [
      { id: 'a1', label: 'A1.6 · Birim Çevirici' }
    ],
    '19_n_KAC': [
      { id: 'a1', label: 'A1 · Hesaplama' }
    ],
    '20_n_KLAVYE_SEMBOL': [
      { id: 'a1', label: 'A1.5 · Yazma Araçları' }
    ],
    '21_n_KELIME_YAZIM_DIL': [
      { id: 'a1', label: 'A1.5 · Dil/Sözlük' }
    ],
    '22_n_IS_BASVURU_CV': [
      { id: 'a5', label: 'A5.7 · CV Aracı' }
    ],
    '23_n_HUKUK_KANUN': [
      { id: 'a4', label: 'A4 · Cezaevi/Hukuk' }
    ],
    '24_n_SAGLIK_VUCUT': [
      { id: 'a1', label: 'A1.3 · Sağlık Hesap' }
    ],
    '25_n_KONSER_ETKINLIK': [
      { id: 'a5', label: 'A5 · Etkinlik' }
    ],
    '26_n_FINANS_BANKA_YATIRIM': [
      { id: 'a1', label: 'A1.2 · Finans' }
    ],
    '27_n_BURC_ASTROLOJI': [
      { id: 'a1', label: 'A1.4 · Astroloji' },
      { id: 'a5', label: 'A5.9 · Burç/Fal' }
    ],
    '28_n_EGITIM_SINAV': [
      { id: 'a1', label: 'A1.5 · Eğitim' }
    ],
    '29_n_OTOMOTIV': [
      { id: 'a4', label: 'A4 · Plaka/MTV' }
    ],
    '30_n_SPOR': [
      { id: 'a5', label: 'A5 · Genel' }
    ],
    '31_n_OYUN_REHBER': [
      { id: 'a5.6', label: 'A5 · Oyun Rehberi', section: 'a5' }
    ],
    '32_n_DIZI_FILM': [
      { id: 'a5', label: 'A5.5 · Dizi (TV+ Köprü)' }
    ],
    '33_n_SEYAHAT_ULASIM': [
      { id: 'a4', label: 'A4 · Pasaport' }
    ],
    '34_n_COCUK_BEBEK': [
      { id: 'a1', label: 'A1.3 · Hamilelik/Bebek' }
    ],
    '35_n_HAYVANLAR': [
      { id: 'a1', label: 'A1.1 · Kedi/Köpek Yaşı' }
    ],
    '36_n_BITKILER_BAHCE': [],
    '37_n_BEAUTY_KOZMETIK': [],
    '38_n_EV_DEKOR_MOBILYA': [],
    '39_n_YEMEK_TARIF': [],
    '40_n_SOSYAL_MEDYA_APP': [
      { id: 'a5', label: 'A5.4 · Sosyal Medya' }
    ],
    '41_n_TELEFON_CIHAZ': [
      { id: 'a4', label: 'A4 · Telefon/Cihaz' }
    ],
    '42_n_ASTRONOMI': [],
    '43_n_DIGER_top500': [
      { id: 'top-firsatlar', label: 'Bölüm 02 · Top Fırsatlar' }
    ],
    '44_TUM_60K_FIRSATLAR': [
      { id: 'top-firsatlar', label: 'Bölüm 02 · Top Fırsatlar' },
      { id: 'karsilastirma', label: 'Bölüm 11 · Karşılaştırma' }
    ]
  };

  let workbook = null;
  let sheetCache = {};
  let currentSheet = null;
  let currentRows = [];
  let currentHeader = [];
  let currentPage = 0;
  let searchTerm = '';
  const PAGE_SIZE = 500;
  let downloadUrlSet = false;

  function base64ToUint8Array(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function ensureWorkbook() {
    if (workbook) return true;
    if (!window.__XLSX_B64__) return false;
    try {
      const bytes = base64ToUint8Array(window.__XLSX_B64__);
      workbook = XLSX.read(bytes, { type: 'array' });
      if (!downloadUrlSet) {
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        xaDownload.href = url;
        downloadUrlSet = true;
      }
      return true;
    } catch (err) {
      console.error('xlsx parse error:', err);
      return false;
    }
  }

  function buildSidebar() {
    if (!workbook) return;
    xaSidebar.innerHTML = '';
    const allSheets = workbook.SheetNames;
    // Niche dinamik liste
    const nicheGroup = SIDEBAR_GROUPS.find(g => g.sheets === 'niche');
    if (nicheGroup) {
      nicheGroup.sheets = allSheets.filter(n => n.match(/^(09|1\\d|2\\d|3\\d|4[012])_n_/));
    }

    SIDEBAR_GROUPS.forEach(group => {
      const groupEl = document.createElement('div');
      groupEl.className = 'xa-sidebar-group';
      const titleEl = document.createElement('div');
      titleEl.className = 'xa-sidebar-group-title';
      titleEl.textContent = group.title;
      groupEl.appendChild(titleEl);
      group.sheets.forEach(name => {
        if (!allSheets.includes(name)) return;
        const ws = workbook.Sheets[name];
        const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
        const rows = range ? range.e.r : 0;
        const btn = document.createElement('button');
        btn.className = 'xa-sheet-item';
        btn.dataset.sheet = name;
        const numMatch = name.match(/^(\\d+)/);
        const num = numMatch ? numMatch[1] : '';
        const shortName = name.replace(/^\\d+_n?_?/, '').replace(/_/g, ' ');
        btn.innerHTML = '<span class="num">' + num + '</span>' +
                        '<span class="name">' + shortName + '</span>' +
                        '<span class="rows-pill">' + rows + '</span>';
        btn.addEventListener('click', () => selectSheet(name));
        groupEl.appendChild(btn);
      });
      xaSidebar.appendChild(groupEl);
    });
  }

  function getSheetData(name) {
    if (sheetCache[name]) return sheetCache[name];
    const ws = workbook.Sheets[name];
    if (!ws) return [];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    sheetCache[name] = aoa;
    return aoa;
  }

  function selectSheet(name) {
    if (!ensureWorkbook()) {
      xaData.innerHTML = '<div class="xa-empty">Excel verisi yüklenemedi.</div>';
      return;
    }
    currentSheet = name;
    const aoa = getSheetData(name);

    // Sheet title
    const numMatch = name.match(/^(\\d+)/);
    const num = numMatch ? numMatch[1] : '';
    const shortName = name.replace(/^\\d+_n?_?/, '').replace(/_/g, ' ');
    xaSheetTitle.innerHTML = (num ? '<span class="num-badge">' + num + '</span>' : '') + shortName;

    // Active sidebar item
    document.querySelectorAll('.xa-sheet-item').forEach(b => {
      b.classList.toggle('active', b.dataset.sheet === name);
    });
    const activeBtn = xaSidebar.querySelector('.xa-sheet-item.active');
    if (activeBtn) activeBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    if (!aoa.length) {
      currentHeader = [];
      currentRows = [];
      xaData.innerHTML = '<div class="xa-empty">Bu sheet boş.</div>';
      updateMeta();
      updateRelated(name);
      return;
    }

    currentHeader = aoa[0] || [];
    const allRows = aoa.slice(1);
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      currentRows = allRows.filter(r => r.some(c => String(c).toLowerCase().includes(term)));
    } else {
      currentRows = allRows;
    }
    currentPage = 0;
    renderTable();
    updateMeta();
    updateRelated(name);
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function highlightTerm(s, term) {
    if (!term) return escapeHtml(s);
    const safe = escapeHtml(s);
    const idx = safe.toLowerCase().indexOf(term.toLowerCase());
    if (idx < 0) return safe;
    return safe.slice(0, idx)
      + '<span class="xa-highlight">' + safe.slice(idx, idx + term.length) + '</span>'
      + safe.slice(idx + term.length);
  }

  function renderTable() {
    const totalRows = currentRows.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / PAGE_SIZE));
    if (currentPage >= totalPages) currentPage = totalPages - 1;
    if (currentPage < 0) currentPage = 0;
    const start = currentPage * PAGE_SIZE;
    const end = Math.min(start + PAGE_SIZE, totalRows);
    const slice = currentRows.slice(start, end);

    if (totalRows === 0) {
      xaData.innerHTML = '<div class="xa-empty">' + (searchTerm ? 'Arama sonucu yok.' : 'Veri yok.') + '</div>';
      xaPageInfo.textContent = '0 / 0';
      xaPrev.disabled = true;
      xaNext.disabled = true;
      return;
    }

    let html = '<table class="xa-table"><thead><tr><th>#</th>';
    currentHeader.forEach(h => {
      html += '<th>' + escapeHtml(h) + '</th>';
    });
    html += '</tr></thead><tbody>';
    slice.forEach((row, idx) => {
      const rowNum = start + idx + 1;
      html += '<tr><td>' + rowNum + '</td>';
      currentHeader.forEach((_, c) => {
        const cell = row[c] === undefined ? '' : row[c];
        html += '<td>' + highlightTerm(String(cell), searchTerm) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    xaData.innerHTML = html;
    xaData.scrollTop = 0;

    xaPageInfo.textContent = (currentPage + 1) + ' / ' + totalPages;
    xaPrev.disabled = currentPage === 0;
    xaNext.disabled = currentPage >= totalPages - 1;
  }

  function updateMeta() {
    xaMetaRows.textContent = currentRows.length.toLocaleString('tr-TR');
    xaMetaCols.textContent = currentHeader.length;
    xaMetaFilter.textContent = searchTerm ? '"' + searchTerm + '"' : 'Yok';
  }

  function updateRelated(sheetName) {
    const related = SHEET_TO_REPORT[sheetName] || [];
    if (related.length === 0) {
      xaRelated.innerHTML = '<span class="xa-related-label">Bu sheet rapora doğrudan referans verilmiyor.</span>';
      return;
    }
    let html = '<span class="xa-related-label">Rapor\\'da geçtiği yerler:</span>';
    related.forEach(r => {
      html += '<a class="xa-related-link" href="#' + r.id + '" data-xa-goto="' + r.id + '">' + r.label + '</a>';
    });
    xaRelated.innerHTML = html;
    // Bind click handlers
    xaRelated.querySelectorAll('[data-xa-goto]').forEach(a => {
      a.addEventListener('click', (e) => {
        e.preventDefault();
        const targetId = a.dataset.xaGoto;
        closeExcelApp();
        setTimeout(() => {
          const target = document.getElementById(targetId);
          if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
      });
    });
  }

  function openExcelApp(sheetAliasOrName) {
    if (!ensureWorkbook()) {
      console.error('Workbook not loaded');
      return;
    }
    if (xaSidebar.children.length === 0) buildSidebar();

    const sheetName = SHEET_ALIASES[sheetAliasOrName] || sheetAliasOrName || workbook.SheetNames[0];
    xaOverlay.classList.add('active');
    document.body.classList.add('excel-app-open');
    selectSheet(sheetName);
    // Sidebar mobile'da kapali baslat
    xaSidebar.classList.remove('open');
  }

  function closeExcelApp() {
    xaOverlay.classList.remove('active');
    document.body.classList.remove('excel-app-open');
  }

  // Top bar event listeners
  xaBack.addEventListener('click', closeExcelApp);
  xaClose.addEventListener('click', closeExcelApp);
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && xaOverlay.classList.contains('active')) closeExcelApp();
  });
  xaSidebarToggle.addEventListener('click', (e) => {
    e.stopPropagation();
    xaSidebar.classList.toggle('open');
  });

  // Search
  let searchTimer = null;
  xaSearch.addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      searchTerm = e.target.value.trim();
      if (currentSheet) selectSheet(currentSheet);
    }, 220);
  });

  // Pagination
  xaPrev.addEventListener('click', () => { currentPage--; renderTable(); });
  xaNext.addEventListener('click', () => { currentPage++; renderTable(); });

  // Tüm rapor butonlari Excel App'i acsin
  document.querySelectorAll('[data-xls-sheet], [data-xls-open]').forEach(el => {
    el.addEventListener('click', (e) => {
      e.preventDefault();
      const alias = el.dataset.xlsSheet || el.dataset.xlsOpen;
      openExcelApp(alias);
    });
  });"""

text = text[:s] + NEW_JS + text[e_end:]
HTML.write_text(text, encoding="utf-8")
print(f"Updated HTML: {len(text):,} chars")
