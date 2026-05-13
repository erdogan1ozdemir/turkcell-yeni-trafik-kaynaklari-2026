"""
1) Eski Excel modal JS blogunu yeni inline Excel viewer JS'i ile degistir.
2) Footer'dan once 'Excel Veri' section'i ekle.
3) Sidebar nav'a 'Excel Veri' linki ekle.
4) data-xls-sheet alias mantigini koru.
"""
from pathlib import Path
import re

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")

# ---------------------------------------------------------------
# 1) Eski JS blogu degistir
# ---------------------------------------------------------------
js_start = text.find("  // =========================================================\n  // EXCEL VIEWER (SheetJS)")
js_end_marker = """  // Tüm "Excel'de göster" linkleri/butonları
  document.querySelectorAll('[data-xls-sheet]').forEach(el => {
    el.addEventListener('click', (e) => {
      e.preventDefault();
      const alias = el.dataset.xlsSheet;
      const sheetName = SHEET_ALIASES[alias] || alias;
      openXlsModal(sheetName);
    });
  });"""

js_end = text.find(js_end_marker)
if js_start < 0 or js_end < 0:
    print("JS markers not found.", js_start, js_end)
    raise SystemExit(1)
js_end_full = js_end + len(js_end_marker)
print(f"Replacing JS block of {js_end_full - js_start} chars.")

NEW_JS = """  // =========================================================
  // INLINE EXCEL VIEWER (parses embedded base64 xlsx with SheetJS)
  // =========================================================
  const xlsSection = document.getElementById('excelSection');
  const xlsTabsBar = document.getElementById('xlsTabsBar');
  const xlsDataWrap = document.getElementById('xlsDataWrap');
  const xlsSearch = document.getElementById('xlsSearch');
  const xlsMetaSheet = document.getElementById('xlsMetaSheet');
  const xlsMetaRows = document.getElementById('xlsMetaRows');
  const xlsMetaCols = document.getElementById('xlsMetaCols');
  const xlsFooterStats = document.getElementById('xlsFooterStats');
  const xlsPrev = document.getElementById('xlsPrev');
  const xlsNext = document.getElementById('xlsNext');
  const xlsPageInfo = document.getElementById('xlsPageInfo');
  const xlsDownload = document.getElementById('xlsDownload');

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

  let workbook = null;
  let sheetCache = {};       // sheetName -> [[cell,cell,...]] (AoA)
  let currentSheet = null;
  let currentRows = [];      // filtrelenmis tum satirlar (header haric)
  let currentHeader = [];
  let currentPage = 0;
  const PAGE_SIZE = 500;
  let searchTerm = '';

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
      // Download blob URL'i hazirla
      const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      xlsDownload.href = url;
      return true;
    } catch (err) {
      console.error('xlsx parse error:', err);
      return false;
    }
  }

  function buildTabs() {
    xlsTabsBar.innerHTML = '';
    workbook.SheetNames.forEach(name => {
      const tab = document.createElement('button');
      tab.className = 'xls-tab';
      tab.dataset.sheet = name;
      // Daha kisa label
      const short = name.replace(/^\\d+_n?_?/, '').replace(/_/g, ' ');
      const numPrefix = (name.match(/^\\d+/) || [''])[0];
      tab.innerHTML = numPrefix
        ? '<strong style="color:var(--coral);margin-right:6px;">' + numPrefix + '</strong>' + short
        : short;
      tab.title = name;
      tab.addEventListener('click', () => selectSheet(name));
      xlsTabsBar.appendChild(tab);
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

  function selectSheet(name, opts = {}) {
    if (!ensureWorkbook()) {
      xlsDataWrap.innerHTML = '<div class="xls-loading">Excel verisi yüklenemedi.</div>';
      return;
    }
    currentSheet = name;
    const aoa = getSheetData(name);
    if (!aoa.length) {
      xlsDataWrap.innerHTML = '<div class="xls-empty">Boş sheet.</div>';
      currentHeader = [];
      currentRows = [];
      updateMeta();
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
    // Aktif tab vurgu + viewport'a kaydir
    document.querySelectorAll('.xls-tab').forEach(t => {
      t.classList.toggle('active', t.dataset.sheet === name);
    });
    const activeTab = document.querySelector('.xls-tab.active');
    if (activeTab) activeTab.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    if (opts.scrollIntoView) {
      xlsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
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
      + '<span class="xls-highlight">' + safe.slice(idx, idx + term.length) + '</span>'
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
      xlsDataWrap.innerHTML = '<div class="xls-empty">' + (searchTerm ? 'Arama sonucu yok.' : 'Veri yok.') + '</div>';
      xlsPageInfo.textContent = '0 / 0';
      xlsPrev.disabled = true;
      xlsNext.disabled = true;
      return;
    }

    let html = '<table class="xls-data-table"><thead><tr>';
    html += '<th style="width:50px;">#</th>';
    currentHeader.forEach(h => {
      html += '<th>' + escapeHtml(h) + '</th>';
    });
    html += '</tr></thead><tbody>';
    slice.forEach((row, idx) => {
      html += '<tr><td><span class="xls-row-num"></span></td>';
      const rowNum = start + idx + 1;
      html = html.replace('<span class="xls-row-num"></span>', '<span class="xls-row-num">' + rowNum + '</span>');
      currentHeader.forEach((_, c) => {
        const cell = row[c] === undefined ? '' : row[c];
        html += '<td>' + highlightTerm(String(cell), searchTerm) + '</td>';
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    xlsDataWrap.innerHTML = html;
    xlsDataWrap.scrollTop = 0;

    xlsPageInfo.textContent = (currentPage + 1) + ' / ' + totalPages;
    xlsPrev.disabled = currentPage === 0;
    xlsNext.disabled = currentPage >= totalPages - 1;
  }

  function updateMeta() {
    xlsMetaSheet.textContent = currentSheet || '-';
    xlsMetaRows.textContent = currentRows.length.toLocaleString('tr-TR');
    xlsMetaCols.textContent = currentHeader.length;
    if (currentRows.length === 0) {
      xlsFooterStats.innerHTML = '<strong>' + (currentSheet || '-') + '</strong> · veri yok';
    } else {
      xlsFooterStats.innerHTML = '<strong>' + (currentSheet || '-') + '</strong> · ' +
        currentRows.length.toLocaleString('tr-TR') + ' satır · ' + currentHeader.length + ' sütun' +
        (searchTerm ? ' · filtre: "' + escapeHtml(searchTerm) + '"' : '');
    }
  }

  // Search
  let searchTimer = null;
  xlsSearch.addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      searchTerm = e.target.value.trim();
      if (currentSheet) selectSheet(currentSheet);
    }, 220);
  });

  // Pagination
  xlsPrev.addEventListener('click', () => { currentPage--; renderTable(); });
  xlsNext.addEventListener('click', () => { currentPage++; renderTable(); });

  // Tüm 'Excel'de göster' butonları
  document.querySelectorAll('[data-xls-sheet]').forEach(el => {
    el.addEventListener('click', (e) => {
      e.preventDefault();
      const alias = el.dataset.xlsSheet;
      const sheetName = SHEET_ALIASES[alias] || alias;
      if (!ensureWorkbook()) return;
      if (!xlsTabsBar.hasChildNodes()) buildTabs();
      selectSheet(sheetName, { scrollIntoView: true });
    });
  });

  // Sayfa Excel section'a ilk geldiginde otomatik yukle (lazy)
  const xlsObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting && !workbook) {
        if (ensureWorkbook()) {
          buildTabs();
          selectSheet(workbook.SheetNames[0]);
        } else {
          xlsDataWrap.innerHTML = '<div class="xls-empty">Excel verisi gömülmemiş. Build scriptini çalıştırın.</div>';
        }
        xlsObserver.disconnect();
      }
    });
  }, { rootMargin: '300px' });
  if (xlsSection) xlsObserver.observe(xlsSection);"""

text = text[:js_start] + NEW_JS + text[js_end_full:]

# ---------------------------------------------------------------
# 2) Sidebar nav'a Excel Veri linki ekle
# ---------------------------------------------------------------
nav_marker = '<li><a href="#sonraki"><span class="num">12</span><span>Sonraki Adımlar</span></a></li>'
nav_replace = nav_marker + '\n      <li><a href="#excelSection"><span class="num">13</span><span>Excel Veri (46 sheet)</span><span class="new-badge">CANLI</span></a></li>'
if nav_marker in text:
    text = text.replace(nav_marker, nav_replace, 1)
    print("Sidebar nav updated.")
else:
    print("WARN: nav marker not found")

# ---------------------------------------------------------------
# 3) Footer'dan once Excel Veri section'i ekle
# ---------------------------------------------------------------
section_marker = '  <!-- FOOTER -->'
section_html = '''  <!-- ============================================================
       13. EXCEL VERİ - Inline (tüm 46 sheet)
       ============================================================ -->
  <section class="section" id="excelSection">
    <div class="section-header">
      <div class="section-num">13</div>
      <div class="section-title-block">
        <div class="section-eyebrow">Canlı Veri</div>
        <h2 class="section-title">Excel <span class="hl">Veri Tarayıcısı</span></h2>
        <p class="section-lead">
          Tüm <strong>46 sheet</strong> ve 60.715 anahtar kelime bu HTML'in içine gömülü. Hiçbir dosya yüklemeye gerek yok.
          Üstteki sekmelerden sheet seç, arama ile filtrele, sayfa sayfa gez. Rapordaki her "Excel · ..." butonu doğrudan ilgili sheet'i açar.
        </p>
      </div>
    </div>

    <div class="excel-section">
      <!-- Header bar -->
      <div class="xls-header">
        <div class="xls-header-left">
          <div class="xls-header-icon">⊞</div>
          <div>
            <div class="xls-header-title">Turkcell Trafik Fırsatları - Excel Veri</div>
            <div class="xls-header-sub">46 sheet · 60.715 KW · DataForSEO + Ahrefs · Mayıs 2026</div>
          </div>
        </div>
        <div class="xls-header-right">
          <div class="xls-search-wrap">
            <input type="text" class="xls-search" id="xlsSearch" placeholder="Sheet içinde ara..." aria-label="Excel sheet içinde ara">
          </div>
          <a href="TURKCELL_TRAFIK_FIRSATLARI.xlsx" download class="xls-download" id="xlsDownload">⬇ XLSX İndir</a>
        </div>
      </div>

      <!-- Sheet tabs -->
      <div class="xls-tabs-bar" id="xlsTabsBar"></div>

      <!-- Meta bar (active sheet info + pagination) -->
      <div class="xls-tab-meta">
        <div>
          <span class="meta-label">Aktif sheet:</span><span class="meta-value" id="xlsMetaSheet">-</span>
          <span style="margin-left:14px;">
            <span class="meta-label">Satır:</span><span class="meta-value" id="xlsMetaRows">-</span>
          </span>
          <span style="margin-left:14px;">
            <span class="meta-label">Sütun:</span><span class="meta-value" id="xlsMetaCols">-</span>
          </span>
        </div>
        <div class="xls-pagination">
          <button class="xls-page-btn" id="xlsPrev" aria-label="Önceki sayfa">‹ Önceki</button>
          <span class="xls-page-info" id="xlsPageInfo">1 / 1</span>
          <button class="xls-page-btn" id="xlsNext" aria-label="Sonraki sayfa">Sonraki ›</button>
        </div>
      </div>

      <!-- Data area -->
      <div class="xls-data-wrap" id="xlsDataWrap">
        <div class="xls-loading">Excel verisi hazırlanıyor...</div>
      </div>

      <!-- Footer stats -->
      <div class="xls-footer">
        <span class="footer-stats" id="xlsFooterStats">Hazırlanıyor...</span>
        <span style="font-family: var(--font-body); color: var(--ink-3);">
          Her sheet sayfa başına 500 satır gösterir. Tam veriyi <a href="TURKCELL_TRAFIK_FIRSATLARI.xlsx" download style="color:var(--coral);font-weight:600;">XLSX olarak indir</a>.
        </span>
      </div>
    </div>
  </section>

  <!-- FOOTER -->'''
text = text.replace(section_marker, section_html, 1)
print("Section inserted.")

# ---------------------------------------------------------------
# 4) Footer copy guncelle (46 sheet baglantili)
# ---------------------------------------------------------------
# rapor-footer kismi: zaten genel ama dokunmuyorum

# Kaydet
HTML.write_text(text, encoding="utf-8")
print(f"Updated HTML: {len(text):,} chars ({len(text)/1024:.1f} KB)")
