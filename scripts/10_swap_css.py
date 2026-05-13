"""
Eski modal CSS blogunu (/* Modal */ ... } @media (max-width: 768px) ile biten)
yeni inline Excel section CSS'i ile degistir.
"""
from pathlib import Path

HTML = Path("/Users/Erdo/Desktop/Claude Projects/Turkcel/index.html")
text = HTML.read_text(encoding="utf-8")

# Eski blogu bul
start_marker = "/* Modal */\n.xls-modal {"
end_marker = "@media (max-width: 768px) {\n  .xls-modal { padding: 0; }\n  .xls-modal-box { height: 100vh; max-width: 100%; border-radius: 0; }\n  .xls-modal-title { font-size: 14px; }\n}"

s = text.find(start_marker)
e = text.find(end_marker)
if s < 0 or e < 0:
    print("MARKERS NOT FOUND")
    print("start:", s, "end:", e)
    raise SystemExit(1)

# end marker'ın sonuna kadar al
e_end = e + len(end_marker)
old_block = text[s:e_end]
print(f"Will replace block of {len(old_block)} chars (lines around {text[:s].count(chr(10))+1})")

NEW_CSS = """/* ============================================================
   INLINE EXCEL DATA SECTION - Google Sheets-like Layout
   ============================================================ */
.excel-section {
  background: var(--off-white);
  border-radius: var(--r-lg);
  border: 1px solid var(--line);
  margin: 32px 0 28px;
  overflow: hidden;
  box-shadow: var(--shadow-card);
}

/* Header strip */
.xls-header {
  background: var(--teal);
  color: var(--white);
  padding: 16px 22px;
  display: flex; align-items: center; justify-content: space-between;
  gap: 18px;
  flex-wrap: wrap;
}
.xls-header-left { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.xls-header-icon {
  width: 38px; height: 38px;
  border-radius: 8px;
  background: var(--coral);
  display: inline-flex; align-items: center; justify-content: center;
  font-family: var(--font-display); font-weight: 800; font-size: 18px;
}
.xls-header-title {
  font-family: var(--font-display); font-weight: 800; font-size: 17px;
  letter-spacing: -0.01em; line-height: 1.2;
}
.xls-header-sub {
  font-family: var(--font-body); font-size: 11.5px;
  color: rgba(255,255,255,0.7); margin-top: 2px; letter-spacing: 0.02em;
}
.xls-header-right { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }

.xls-search-wrap { position: relative; }
.xls-search {
  background: rgba(255,255,255,0.1); color: var(--white);
  border: 1px solid rgba(255,255,255,0.18);
  padding: 8px 14px 8px 32px;
  border-radius: var(--r-pill);
  font-family: var(--font-body); font-size: 12.5px;
  width: 220px; outline: none; transition: all 0.18s;
}
.xls-search::placeholder { color: rgba(255,255,255,0.5); }
.xls-search:focus {
  background: rgba(255,255,255,0.18);
  border-color: var(--coral);
  width: 260px;
}
.xls-search-wrap::before {
  content: '\\2315'; position: absolute; left: 12px; top: 50%;
  transform: translateY(-50%); font-size: 14px;
  color: rgba(255,255,255,0.6); pointer-events: none;
}
.xls-download {
  display: inline-flex; align-items: center; gap: 6px;
  background: var(--coral); color: var(--white);
  font-family: var(--font-display); font-weight: 700;
  font-size: 12px; padding: 8px 14px;
  border-radius: var(--r-pill);
  text-decoration: none; border: none;
  cursor: pointer; transition: background 0.18s;
  letter-spacing: 0.02em;
}
.xls-download:hover { background: var(--coral-deep); }

/* Sheet tabs */
.xls-tabs-bar {
  background: var(--teal-soft);
  border-bottom: 2px solid var(--coral);
  padding: 0 8px;
  display: flex; align-items: stretch;
  overflow-x: auto;
  scrollbar-width: thin;
  scrollbar-color: var(--coral) transparent;
}
.xls-tabs-bar::-webkit-scrollbar { height: 6px; }
.xls-tabs-bar::-webkit-scrollbar-track { background: transparent; }
.xls-tabs-bar::-webkit-scrollbar-thumb { background: var(--coral); border-radius: 3px; }

.xls-tab {
  background: transparent;
  color: rgba(255,255,255,0.62);
  border: none;
  padding: 11px 13px;
  font-family: var(--font-body); font-size: 11.5px;
  font-weight: 500;
  cursor: pointer; white-space: nowrap;
  border-bottom: 3px solid transparent;
  transition: all 0.16s;
  flex-shrink: 0;
  letter-spacing: 0.01em;
}
.xls-tab:hover { color: var(--white); background: rgba(255,255,255,0.06); }
.xls-tab.active {
  color: var(--white);
  border-bottom-color: var(--coral);
  background: rgba(255,123,82,0.18);
  font-weight: 700;
  font-family: var(--font-display);
  font-size: 12px;
}

/* Meta bar */
.xls-tab-meta {
  background: var(--white);
  padding: 10px 18px;
  display: flex; align-items: center; justify-content: space-between;
  font-size: 12px;
  border-bottom: 1px solid var(--line);
  flex-wrap: wrap; gap: 12px;
}
.xls-tab-meta .meta-label { color: var(--ink-3); font-family: var(--font-body); }
.xls-tab-meta .meta-value {
  color: var(--teal); font-family: var(--font-display);
  font-weight: 700; margin-left: 4px;
}
.xls-pagination { display: inline-flex; align-items: center; gap: 6px; }
.xls-page-btn {
  background: var(--white); border: 1px solid var(--line);
  color: var(--teal); padding: 5px 10px; border-radius: 6px;
  font-family: var(--font-display); font-weight: 600;
  font-size: 11px; cursor: pointer;
  transition: all 0.16s;
}
.xls-page-btn:hover:not(:disabled) { border-color: var(--coral); color: var(--coral); }
.xls-page-btn:disabled { opacity: 0.45; cursor: not-allowed; }
.xls-page-info {
  font-family: var(--font-display); font-weight: 600;
  color: var(--teal); font-size: 11px; padding: 0 4px;
}

/* Data area */
.xls-data-wrap {
  background: var(--white);
  max-height: 720px;
  overflow: auto;
  position: relative;
}
.xls-data-wrap::-webkit-scrollbar { width: 10px; height: 10px; }
.xls-data-wrap::-webkit-scrollbar-track { background: var(--off-white); }
.xls-data-wrap::-webkit-scrollbar-thumb { background: #c0c0c0; border-radius: 5px; }
.xls-data-wrap::-webkit-scrollbar-thumb:hover { background: var(--coral); }

.xls-loading {
  padding: 60px 24px;
  text-align: center;
  color: var(--ink-2);
  font-size: 14px;
  font-family: var(--font-body);
}
.xls-loading::before {
  content: '\\27f3';
  display: inline-block;
  font-size: 28px; color: var(--coral);
  margin-right: 12px;
  animation: xlsSpin 1.2s linear infinite;
  vertical-align: middle;
}
@keyframes xlsSpin { to { transform: rotate(360deg); } }

.xls-data-table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-family: var(--font-body);
  font-size: 12.5px;
  font-variant-numeric: tabular-nums;
}
.xls-data-table thead th {
  background: var(--teal); color: var(--white);
  font-family: var(--font-display); font-weight: 700;
  font-size: 11px; letter-spacing: 0.04em;
  text-align: left;
  padding: 11px 14px;
  position: sticky; top: 0; z-index: 3;
  border-right: 1px solid var(--teal-soft);
  border-bottom: 2px solid var(--coral);
  white-space: nowrap;
}
.xls-data-table thead th:first-child {
  position: sticky; left: 0; z-index: 4;
  background: var(--teal);
  border-right: 2px solid var(--coral);
}
.xls-data-table tbody td {
  padding: 8px 14px;
  border-bottom: 1px solid var(--line-soft);
  border-right: 1px solid var(--line-soft);
  color: var(--teal);
  vertical-align: top;
  max-width: 420px;
  word-wrap: break-word;
  line-height: 1.45;
}
.xls-data-table tbody td:first-child {
  position: sticky; left: 0;
  background: var(--off-white);
  z-index: 2;
  font-family: var(--font-display); font-weight: 600;
  border-right: 2px solid var(--line);
}
.xls-data-table tbody tr:nth-child(even) td { background: rgba(232,245,233,0.4); }
.xls-data-table tbody tr:nth-child(even) td:first-child { background: rgba(232,245,233,0.85); }
.xls-data-table tbody tr:hover td { background: rgba(255,123,82,0.08); }
.xls-data-table tbody tr:hover td:first-child { background: rgba(255,123,82,0.18); }

.xls-highlight {
  background: var(--coral); color: var(--white);
  padding: 1px 4px; border-radius: 3px; font-weight: 700;
}

.xls-empty {
  padding: 60px 24px; text-align: center;
  color: var(--ink-3);
  font-family: var(--font-body); font-size: 14px;
}

.xls-footer {
  background: var(--off-white);
  border-top: 1px solid var(--line);
  padding: 10px 22px;
  display: flex; justify-content: space-between; align-items: center;
  font-size: 11.5px; color: var(--ink-3);
  flex-wrap: wrap; gap: 12px;
}
.xls-footer .footer-stats {
  font-family: var(--font-display);
  color: var(--teal); font-weight: 600;
}
.xls-footer .footer-stats strong { color: var(--coral); }

@media (max-width: 768px) {
  .xls-header { padding: 12px 16px; }
  .xls-header-title { font-size: 14px; }
  .xls-search { width: 140px; font-size: 11.5px; }
  .xls-search:focus { width: 160px; }
  .xls-data-wrap { max-height: 540px; }
  .xls-data-table thead th, .xls-data-table tbody td { padding: 7px 10px; font-size: 11.5px; }
}"""

new_text = text[:s] + NEW_CSS + text[e_end:]
HTML.write_text(new_text, encoding="utf-8")
print(f"Replaced. New file size: {len(new_text):,} chars")
