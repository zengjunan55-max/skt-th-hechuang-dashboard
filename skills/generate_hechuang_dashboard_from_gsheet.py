from __future__ import annotations

import json
import re
import time
from datetime import datetime
from io import StringIO
from pathlib import Path

import pandas as pd
import requests

import generate_kol_dashboard as base


SHEET_ID = "142eo8ASRHj0utna0cHxiWRnMY8s7r0ZFKAc-7s_jdgs"
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = BASE_DIR / "SKT-鍚堝垱鍐呭鏁版嵁鐪嬫澘_鑷姩鐗?html"
ALL_DATA_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&sheet=ALL_data"
WORKBOOK_XLSX_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
DMS_GVIZ_URL = (
    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq"
    "?tqx=out:json&sheet=DMS&range=A1:CV10000"
)


def enhance_html_for_live_refresh(html_text: str) -> str:
    button_html = (
        '<button id="refreshNow" type="button">绔嬪嵆鍒锋柊</button>'
    )
    status_html = (
        '<span id="refreshStatus" class="refresh-status refresh-status-idle">'
        '<span class="refresh-status-dot"></span>'
        '<span class="refresh-status-text">鏁版嵁姝ｅ父</span>'
        '</span>'
    )
    html_text = html_text.replace(
        '<button id="applyDates" type="button">搴旂敤鏃堕棿娈?/button>',
        f'<button id="applyDates" type="button">搴旂敤鏃堕棿娈?/button>{button_html}',
        1,
    )
    html_text = html_text.replace(
        '<div class="controls-meta">\n        <span class="update-time">',
        f'<div class="controls-meta">\n        {status_html}\n        <span class="update-time">',
        1,
    )
    html_text = html_text.replace(
        "</style>",
        """
    .refresh-status {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      font-weight: 700;
      color: var(--muted);
      vertical-align: middle;
      white-space: nowrap;
    }
    .refresh-status-dot {
      width: 10px;
      height: 10px;
      border-radius: 999px;
      background: #30a46c;
      box-shadow: 0 0 0 2px rgba(48, 164, 108, 0.12);
      flex: 0 0 10px;
    }
    .refresh-status-refreshing {
      color: #b26a00;
    }
    .refresh-status-refreshing .refresh-status-dot {
      background: #f59e0b;
      box-shadow: 0 0 0 2px rgba(245, 158, 11, 0.18);
    }
    .refresh-status-success,
    .refresh-status-idle {
      color: #1f7a4d;
    }
    .refresh-status-success .refresh-status-dot,
    .refresh-status-idle .refresh-status-dot {
      background: #22c55e;
      box-shadow: 0 0 0 2px rgba(34, 197, 94, 0.16);
    }
    .refresh-status-error {
      color: #c0362c;
    }
    .refresh-status-error .refresh-status-dot {
      background: #ef4444;
      box-shadow: 0 0 0 2px rgba(239, 68, 68, 0.16);
    }
    .controls-filter-panel {
      justify-content: flex-start;
    }
"""
        + "</style>",
        1,
    )
    inject_script = f"""
<script>
(function() {{
  const SHEET_ID = "{SHEET_ID}";
  const ALL_DATA_GVIZ_URL = "https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:json;responseHandler:__allDataRefreshHandler&sheet=ALL_data&range=A1:AL300000";
  const DMS_GVIZ_URL = "https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:json;responseHandler:__dmsRefreshHandler&sheet=DMS&range=A1:CV10000";
  // Source and dashboard monetary metrics are both in USD.
  const refreshBtn = document.getElementById('refreshNow');
  const refreshStatus = document.getElementById('refreshStatus');
  const updateTimeNode = document.querySelector('.update-time');
  const periodSummaryNode = document.querySelector('header > p');
  const REFRESH_CACHE_VERSION = 3;
  const REFRESH_CACHE_NAME = 'hechuang-dashboard-refresh-v1';
  const REFRESH_CACHE_KEY = 'hechuang-dashboard-refresh-v1';
  const REFRESH_CACHE_URL = window.location.origin && window.location.origin !== 'null'
    ? window.location.origin + '/__hechuang_cache__/latest.json'
    : '';

  function setRefreshStatus(state) {{
    if (!refreshStatus) return;
    const config = {{
      idle: {{ text: '鏁版嵁姝ｅ父', cls: 'refresh-status-idle' }},
      refreshing: {{ text: '鍒锋柊涓?, cls: 'refresh-status-refreshing' }},
      success: {{ text: '鏁版嵁姝ｅ父', cls: 'refresh-status-success' }},
      error: {{ text: '鏁版嵁寮傚父', cls: 'refresh-status-error' }},
    }}[state] || {{ text: state || '', cls: 'refresh-status-idle' }};
    refreshStatus.className = 'refresh-status ' + config.cls;
    const textNode = refreshStatus.querySelector('.refresh-status-text');
    if (textNode) textNode.textContent = config.text;
  }}

  function localDayKey(value) {{
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    return `${{date.getFullYear()}}-${{String(date.getMonth() + 1).padStart(2, '0')}}-${{String(date.getDate()).padStart(2, '0')}}`;
  }}

  function formatRefreshStamp(value) {{
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    return `${{date.getFullYear()}}/${{String(date.getMonth() + 1).padStart(2, '0')}}/${{String(date.getDate()).padStart(2, '0')}} ${{String(date.getHours()).padStart(2, '0')}}:${{String(date.getMinutes()).padStart(2, '0')}}:${{String(date.getSeconds()).padStart(2, '0')}}`;
  }}

  function isValidCachedRefresh(snapshot) {{
    return Boolean(
      snapshot &&
      snapshot.version === REFRESH_CACHE_VERSION &&
      snapshot.refreshedAt &&
      Array.isArray(snapshot.queryRows) && snapshot.queryRows.length &&
      Array.isArray(snapshot.compareRows) && snapshot.compareRows.length &&
      validPeriod(snapshot.queryPeriod) &&
      validPeriod(snapshot.comparePeriod)
    );
  }}

  function isCachedRefreshToday(snapshot) {{
    return isValidCachedRefresh(snapshot) && localDayKey(snapshot.refreshedAt) === localDayKey(new Date());
  }}

  async function readCachedRefresh() {{
    let snapshot = null;
    if (REFRESH_CACHE_URL && 'caches' in window) {{
      try {{
        const cache = await caches.open(REFRESH_CACHE_NAME);
        const response = await cache.match(REFRESH_CACHE_URL);
        if (response) snapshot = await response.json();
      }} catch (err) {{
        console.warn('读取看板缓存失败', err);
      }}
    }}
    if (!snapshot) {{
      try {{
        const stored = window.localStorage.getItem(REFRESH_CACHE_KEY);
        if (stored) snapshot = JSON.parse(stored);
      }} catch (err) {{
        console.warn('读取本地看板缓存失败', err);
      }}
    }}
    return isValidCachedRefresh(snapshot) ? snapshot : null;
  }}

  async function writeCachedRefresh(snapshot) {{
    const serialized = JSON.stringify(snapshot);
    let stored = false;
    if (REFRESH_CACHE_URL && 'caches' in window) {{
      try {{
        const cache = await caches.open(REFRESH_CACHE_NAME);
        const response = new Response(serialized, {{ headers: {{ 'Content-Type': 'application/json' }} }});
        await cache.put(REFRESH_CACHE_URL, response);
        stored = true;
      }} catch (err) {{
        console.warn('保存看板缓存失败', err);
      }}
    }}
    if (!stored) {{
      try {{
        window.localStorage.setItem(REFRESH_CACHE_KEY, serialized);
        stored = true;
      }} catch (err) {{
        console.warn('保存本地看板缓存失败', err);
      }}
    }}
    return stored;
  }}

  function normalizeText(value) {{
    return value === undefined || value === null ? '' : String(value).trim();
  }}

  function gvizRowsToObjects(payload) {{
    const cols = payload.table.cols.map(col => normalizeText(col.label));
    return (payload.table.rows || []).map(item => {{
      const cells = item.c || [];
      const row = {{}};
      cols.forEach((col, idx) => {{
        const cell = idx < cells.length ? cells[idx] : null;
        row[col] = cell ? cell.v : null;
      }});
      return row;
    }});
  }}

  function fetchAllDataRows() {{
    return new Promise((resolve, reject) => {{
      const old = document.getElementById('all-data-refresh-script');
      if (old) old.remove();
      const timer = setTimeout(() => reject(new Error('璇诲彇 ALL_data 瓒呮椂')), 120000);
      window.__allDataRefreshHandler = payload => {{
        clearTimeout(timer);
        try {{
          resolve(gvizRowsToObjects(payload));
        }} catch (err) {{
          reject(err);
        }}
      }};
      const script = document.createElement('script');
      script.id = 'all-data-refresh-script';
      script.src = ALL_DATA_GVIZ_URL + '&_=' + Date.now();
      script.onerror = () => {{
        clearTimeout(timer);
        reject(new Error('鍔犺浇 ALL_data 鑴氭湰澶辫触'));
      }};
      document.head.appendChild(script);
    }});
  }}

  function fetchDmsRows() {{
    return new Promise((resolve, reject) => {{
      const old = document.getElementById('dms-refresh-script');
      if (old) old.remove();
      const timer = setTimeout(() => reject(new Error('璇诲彇 DMS 瓒呮椂')), 120000);
      window.__dmsRefreshHandler = payload => {{
        clearTimeout(timer);
        try {{
          resolve(gvizRowsToObjects(payload));
        }} catch (err) {{
          reject(err);
        }}
      }};
      const script = document.createElement('script');
      script.id = 'dms-refresh-script';
      script.src = DMS_GVIZ_URL + '&_=' + Date.now();
      script.onerror = () => {{
        clearTimeout(timer);
        reject(new Error('鍔犺浇 DMS 鑴氭湰澶辫触'));
      }};
      document.head.appendChild(script);
    }});
  }}

  function toNumber(value) {{
    const n = Number(value);
    return Number.isFinite(n) ? n : 0;
  }}

  function formatDateOnly(value) {{
    const text = normalizeText(value);
    if (!text) return '';
    const match = text.match(/^(\\d{{4}})-(\\d{{2}})-(\\d{{2}})/);
    if (match) return `${{match[1]}}-${{match[2]}}-${{match[3]}}`;
    const gvizMatch = text.match(/^Date\\((\\d{{4}}),(\\d{{1,2}}),(\\d{{1,2}})(?:,\\d{{1,2}},\\d{{1,2}},\\d{{1,2}})?\\)$/);
    if (gvizMatch) {{
      const year = Number(gvizMatch[1]);
      const month = Number(gvizMatch[2]) + 1;
      const day = Number(gvizMatch[3]);
      return `${{year}}-${{String(month).padStart(2, '0')}}-${{String(day).padStart(2, '0')}}`;
    }}
    if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {{
      return `${{value.getFullYear()}}-${{String(value.getMonth() + 1).padStart(2, '0')}}-${{String(value.getDate()).padStart(2, '0')}}`;
    }}
    const d = new Date(text);
    if (Number.isNaN(d.getTime())) return '';
    return `${{d.getFullYear()}}-${{String(d.getMonth() + 1).padStart(2, '0')}}-${{String(d.getDate()).padStart(2, '0')}}`;
  }}

  function normalizeScaleBySystem(systemType) {{
    const text = normalizeText(systemType).toLowerCase().replace(/\\s+/g, '');
    if (text === 'bigkol绯荤粺' || text === '澶ol绯荤粺' || text === '澶ol绯荤当') return 'Big KOL绯荤粺';
    if (text === 'regularkol绯荤粺') return 'Regular KOL绯荤粺';
    return '鍏朵粬';
  }}

  function buildDmsLookup(rows) {{
    const map = new Map();
    rows.forEach(row => {{
      const id = normalizeText(row['绱犳潗缂栧彿']);
      if (!id) return;
      const current = map.get(id) || {{
        绱犳潗缂栧彿: id,
        绮変笣鏁? 0,
        杈句汉閲忕骇: '',
        username: '',
        url: '',
        绯荤粺绫诲瀷: '',
        鍐呭骞冲彴: '',
        鍙戝竷鏃堕棿: ''
      }};
      current.绮変笣鏁?= Math.max(current.绮変笣鏁?|| 0, toNumber(row['绮変笣鏁?]));
      current.杈句汉閲忕骇 = current.杈句汉閲忕骇 || normalizeText(row['杈句汉閲忕骇']);
      current.username = current.username || normalizeText(row['username']);
      current.url = current.url || normalizeText(row['url']);
      current.绯荤粺绫诲瀷 = current.绯荤粺绫诲瀷 || normalizeText(row['绯荤粺绫诲瀷']);
      current.鍐呭骞冲彴 = current.鍐呭骞冲彴 || normalizeText(row['鍐呭骞冲彴']);
      current.鍙戝竷鏃堕棿 = current.鍙戝竷鏃堕棿 || normalizeText(row['鍙戝竷鏃堕棿']);
      map.set(id, current);
    }});
    return map;
  }}

  function defaultPeriods(filteredRows) {{
    const now = new Date();
    const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    const queryStart = `${{yesterday.getFullYear()}}-${{String(yesterday.getMonth() + 1).padStart(2, '0')}}-01`;
    const queryEnd = `${{yesterday.getFullYear()}}-${{String(yesterday.getMonth() + 1).padStart(2, '0')}}-${{String(yesterday.getDate()).padStart(2, '0')}}`;
    const prevMonthEnd = new Date(yesterday.getFullYear(), yesterday.getMonth(), 0);
    const compareStart = `${{prevMonthEnd.getFullYear()}}-${{String(prevMonthEnd.getMonth() + 1).padStart(2, '0')}}-01`;
    const compareEndDay = Math.min(yesterday.getDate(), prevMonthEnd.getDate());
    const compareEnd = `${{prevMonthEnd.getFullYear()}}-${{String(prevMonthEnd.getMonth() + 1).padStart(2, '0')}}-${{String(compareEndDay).padStart(2, '0')}}`;
    return {{
      query: {{ start: queryStart, end: queryEnd }},
      compare: {{ start: compareStart, end: compareEnd }},
    }};
  }}

  function applyDefaultPeriodsToInputs() {{
    const periods = defaultPeriods([]);
    const qStartInput = document.getElementById('qStart');
    const qEndInput = document.getElementById('qEnd');
    const cStartInput = document.getElementById('cStart');
    const cEndInput = document.getElementById('cEnd');
    if (qStartInput) qStartInput.value = periods.query.start;
    if (qEndInput) qEndInput.value = periods.query.end;
    if (cStartInput) cStartInput.value = periods.compare.start;
    if (cEndInput) cEndInput.value = periods.compare.end;
    updatePeriodSummary(periods);
    return periods;
  }}

  function selectedPeriodsFromInputs() {{
    const qStart = document.getElementById('qStart')?.value || '';
    const qEnd = document.getElementById('qEnd')?.value || qStart;
    const cStart = document.getElementById('cStart')?.value || '';
    const cEnd = document.getElementById('cEnd')?.value || cStart;
    return {{
      query: {{ start: qStart, end: qEnd }},
      compare: {{ start: cStart, end: cEnd }},
    }};
  }}

  function validPeriod(period) {{
    return normalizeText(period?.start) && normalizeText(period?.end) && period.start <= period.end;
  }}

  function updatePeriodSummary(periods) {{
    if (!periodSummaryNode || !periods) return;
    periodSummaryNode.textContent = `鏌ヨ鍛ㄦ湡锛?{{periods.query.start}} 鑷?${{periods.query.end}}锛涚幆姣斿懆鏈燂細${{periods.compare.start}} 鑷?${{periods.compare.end}}銆俙;
  }}

  function buildSourceRows(allRows, dmsMap) {{
    const filtered = allRows.filter(row => normalizeText(row['骞垮憡绫诲瀷']).includes('鍚堝垱骞垮憡'));
    return filtered.map(row => {{
      const postId = normalizeText(row['甯栧瓙ID']);
      const materialNo = normalizeText(row['绱犳潗缂栧彿']);
      const matchId = postId || materialNo;
      const dms = dmsMap.get(matchId) || {{}};
      const materialId = postId || materialNo || normalizeText(row['URL']) || normalizeText(row['骞垮憡鍚嶇О']);
      const publishDate = formatDateOnly(dms['鍙戝竷鏃堕棿']) || formatDateOnly(row['绱犳潗鍒涘缓鏃堕棿']) || formatDateOnly(row['鍒涘缓鏃堕棿']);
      const systemType = normalizeText(row['绯荤粺绫诲瀷']) || normalizeText(dms['绯荤粺绫诲瀷']) || '鍏朵粬';
      return {{
        date: formatDateOnly(row['鏃ユ湡']),
        post_id: postId,
        ad_type: normalizeText(row['骞垮憡绫诲瀷']),
        '广告类型': normalizeText(row['骞垮憡绫诲瀷']),
        material_id: materialId,
        material_code: normalizeText(dms['绱犳潗缂栧彿']) || materialNo || materialId,
        material_name: normalizeText(row['鍓ф湰鍚嶇О']) || normalizeText(row['骞垮憡鍚嶇О']) || materialId,
        material_product: normalizeText(row['浜у搧']),
        product: normalizeText(row['浜у搧']),
        platform: normalizeText(row['骞冲彴']) || normalizeText(dms['鍐呭骞冲彴']) || '-',
        username: normalizeText(dms['username']) || normalizeText(row['KOL']) || '-',
        system_type: systemType,
        scale_label: normalizeScaleBySystem(systemType),
        scale_raw: normalizeText(dms['杈句汉閲忕骇']),
        fans: toNumber(dms['绮変笣鏁?]),
        url: normalizeText(row['URL']) || normalizeText(dms['url']) || '-',
        source_url: normalizeText(row['URL']) || normalizeText(dms['url']) || '-',
        publish_date: publishDate,
        ad_form2: normalizeText(row['骞垮憡褰㈠紡2']),
        spend: Math.max(0, toNumber(row['鑺辫垂閲戦锛圲SD锛?])),
        gmv: Math.max(0, toNumber(row['璐墿杞寲浠峰€?])),
        orders: Math.max(0, toNumber(row['璁㈠崟鏁?])),
        impressions: Math.max(0, toNumber(row['灞曠ず娆℃暟'])),
        clicks: Math.max(0, toNumber(row['鐐瑰嚮閲?])),
        reach: Math.max(0, toNumber(row['灞曠ず娆℃暟'])),
      }};
    }}).filter(row => row.date && row.material_id);
  }}

  function updateNewMaterialIds(rows, queryPeriod) {{
    fixedNewMaterialIds.clear();
    rows.forEach(row => {{
      if (row.publish_date && row.publish_date >= queryPeriod.start && row.publish_date <= queryPeriod.end) {{
        fixedNewMaterialIds.add(String(row.material_id));
      }}
    }});
  }}

  function applyCachedRefresh(snapshot) {{
    if (!isValidCachedRefresh(snapshot)) return null;
    embedded.queryRaw = snapshot.queryRows;
    embedded.compareRaw = snapshot.compareRows;
    embedded.queryPeriod = snapshot.queryPeriod;
    embedded.comparePeriod = snapshot.comparePeriod;
    embedded.cpmMedian = Number(snapshot.cpmMedian || 0);
    fixedNewMaterialIds.clear();
    (snapshot.newMaterialIds || []).forEach(id => fixedNewMaterialIds.add(String(id)));
    embedded.newMaterialIds = Array.from(fixedNewMaterialIds);
    document.getElementById('qStart').value = snapshot.queryPeriod.start;
    document.getElementById('qEnd').value = snapshot.queryPeriod.end;
    document.getElementById('cStart').value = snapshot.comparePeriod.start;
    document.getElementById('cEnd').value = snapshot.comparePeriod.end;
    updatePeriodSummary({{ query: snapshot.queryPeriod, compare: snapshot.comparePeriod }});
    refreshAllBoards(snapshot.queryRows, snapshot.compareRows);
    const stamp = formatRefreshStamp(snapshot.refreshedAt);
    if (updateTimeNode && stamp) updateTimeNode.textContent = '最近更新：' + stamp;
    setRefreshStatus('success');
    return snapshot;
  }}

  async function restoreCachedRefresh() {{
    const snapshot = await readCachedRefresh();
    return snapshot ? applyCachedRefresh(snapshot) : null;
  }}

  async function refreshFromGoogleSheet(periodsOverride = null, options = {{}}) {{
    if (!refreshBtn) return;
    refreshBtn.disabled = true;
    setRefreshStatus('refreshing');
    try {{
      const [allRows, dmsRows] = await Promise.all([fetchAllDataRows(), fetchDmsRows()]);
      const sourceRows = buildSourceRows(allRows, buildDmsLookup(dmsRows));
      const periods = periodsOverride || defaultPeriods(sourceRows.map(row => ({{ 鏃ユ湡: row.date }})));
      const queryRows = sourceRows.filter(row => row.date >= periods.query.start && row.date <= periods.query.end);
      const compareRows = sourceRows.filter(row => row.date >= periods.compare.start && row.date <= periods.compare.end);
      if (!queryRows.length) throw new Error(`鏌ヨ鏃堕棿娈?${{periods.query.start}} 鑷?${{periods.query.end}} 鏃犳暟鎹甡);
      if (!compareRows.length) throw new Error(`鐜瘮鏃堕棿娈?${{periods.compare.start}} 鑷?${{periods.compare.end}} 鏃犳暟鎹甡);
      updateNewMaterialIds(sourceRows, periods.query);
      embedded.queryRaw = queryRows;
      embedded.compareRaw = compareRows;
      embedded.queryPeriod = periods.query;
      embedded.comparePeriod = periods.compare;
      const mats = materialRows(queryRows);
      embedded.cpmMedian = mats.length ? mats.map(item => Number(item.cpm || 0)).sort((a,b)=>a-b)[Math.floor(mats.length/2)] : 0;
      document.getElementById('qStart').value = periods.query.start;
      document.getElementById('qEnd').value = periods.query.end;
      document.getElementById('cStart').value = periods.compare.start;
      document.getElementById('cEnd').value = periods.compare.end;
      updatePeriodSummary(periods);
      refreshAllBoards(queryRows, compareRows);
      const now = new Date();
      const stamp = formatRefreshStamp(now);
      if (updateTimeNode) updateTimeNode.textContent = '鏈€杩戞洿鏂帮細' + stamp;
      embedded.newMaterialIds = Array.from(fixedNewMaterialIds);
      await writeCachedRefresh({{
        version: REFRESH_CACHE_VERSION,
        refreshedAt: now.toISOString(),
        queryPeriod: {{ ...periods.query }},
        comparePeriod: {{ ...periods.compare }},
        queryRows,
        compareRows,
        cpmMedian: embedded.cpmMedian,
        newMaterialIds: embedded.newMaterialIds,
      }});
      setRefreshStatus('success');
    }} catch (err) {{
      console.error(err);
      setRefreshStatus('error');
      if (!options.silent) {{
        alert('绔嬪嵆鍒锋柊澶辫触锛? + (err && err.message ? err.message : err));
      }}
    }} finally {{
      refreshBtn.disabled = false;
    }}
  }}

  if (refreshBtn) {{
    setRefreshStatus('idle');
    refreshBtn.addEventListener('click', () => refreshFromGoogleSheet());
  }}
  applyDefaultPeriodsToInputs();
  window.refreshFromGoogleSheet = refreshFromGoogleSheet;
  const cachedRefreshPromise = restoreCachedRefresh();
  const autoRefreshOnPageLoad = async () => {{
    if (window.__hechuangAutoRefreshStarted) return;
    window.__hechuangAutoRefreshStarted = true;
    const cachedSnapshot = await cachedRefreshPromise;
    if (isCachedRefreshToday(cachedSnapshot)) return;
    await refreshFromGoogleSheet(null, {{ silent: true }});
  }};
  if (document.readyState === 'complete') {{
    window.setTimeout(autoRefreshOnPageLoad, 0);
  }} else {{
    window.addEventListener('load', autoRefreshOnPageLoad, {{ once: true }});
  }}
}})();
</script>
"""
    return html_text.replace("</body>", inject_script + "\n</body>", 1)


def enhance_html_with_overview_filter(html_text: str) -> str:
    overview_html = """
    <section id="hechuangOverviewSection">
      <h2>鍚堝垱鍩虹鏁版嵁</h2>
      <p class="small">浠呭睍绀哄悎鍒涜浆鍖栦笌鍚堝垱鏇濆厜涓よ锛岀偣鍑诲彲瑙﹀彂鍏ㄥ眬骞垮憡绫诲瀷鑱斿姩銆?/p>
      <div class="table-wrap" id="hechuangOverviewTable"></div>
    </section>
    """
    html_text = html_text.replace(
        "</header>",
        overview_html + "\n  </header>",
        1,
    )
    html_text = html_text.replace(
        "</style>",
        """
    .overview-chip {
      border: 1px solid var(--line);
      background: #f5fbfd;
      color: var(--ink);
      border-radius: 999px;
      padding: 8px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .active-overview-row .overview-chip {
      background: var(--brand);
      color: #fff;
      border-color: var(--brand);
    }
    .linked-filter-row {
      cursor: default;
    }
    .compare-filter-chip {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border: 1px solid var(--line);
      background: #f5fbfd;
      color: var(--ink);
      border-radius: 999px;
      padding: 8px 16px;
      max-width: 100%;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      box-sizing: border-box;
    }
    .compare-filter-chip.active {
      background: var(--brand);
      color: #fff;
      border-color: var(--brand);
    }
    .linked-filter-row.active-filter-row td {
      background: #f8fbfd;
    }
"""+ "</style>",
        1,
    )
    inject_script = """
<script>
(function () {
  let activeOverviewKind = 'all';
  let activeCompareSystemType = 'all';
  let activeComparePlatform = 'all';
  let activeCompareProducts = [];
  window.__activeCompareSystemType = activeCompareSystemType;
  window.__activeComparePlatform = activeComparePlatform;
  window.__activeCompareProducts = activeCompareProducts;
  const activeFilterSummary = document.getElementById('activeFilterSummary');
  const overviewCompareKeys = new Set(['spend', 'gmv', 'roi', 'impressions', 'cpm', 'total_materials', 'ctr', 'test_new_spend', 'test_new_posts', 'test_new_spend_ratio', 'store_gmv_thb', 'store_visitors']);

  function normalizeText(value) {
    return value === undefined || value === null ? '' : String(value).trim();
  }

  function deltaMeta(currentValue, compareValue) {
    const current = Number(currentValue);
    const compare = Number(compareValue);
    if (!Number.isFinite(current) || !Number.isFinite(compare) || compare === 0) return null;
    const pct = (current - compare) / compare * 100;
    return {
      text: `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`,
      cls: pct > 0 ? 'delta-up' : 'delta-down',
    };
  }

  function fmtPct(value) {
    return `${(Number(value || 0) * 100).toFixed(1)}%`;
  }

  function fmtCtrPct(value) {
    return `${(Number(value || 0) * 100).toFixed(2)}%`;
  }

  function cellValue(row, key) {
    if (['spend', 'gmv', 'cpm', 'test_new_spend', 'store_gmv_thb'].includes(key)) return fmtMoney(row[key]);
    if (['impressions', 'total_materials', 'test_new_posts', 'store_visitors'].includes(key)) return fmtInt(row[key]);
    if (key === 'roi') return Number(row[key] || 0).toFixed(2);
    if (key === 'ctr') return fmtCtrPct(row[key]);
    if (key === 'test_new_spend_ratio') return fmtPct(row[key]);
    return row[key] === undefined || row[key] === null ? '' : String(row[key]);
  }

  function cellWithCompare(row, key) {
    const primary = cellValue(row, key);
    if (!row.__compare || !overviewCompareKeys.has(key)) return primary;
    const meta = deltaMeta(row[key], row.__compare[key]);
    if (!meta) return primary;
    return `<div class="cell-main">${primary}</div><div class="cell-sub ${meta.cls}">${meta.text}</div>`;
  }

  function rowsByOverviewKind(rows, kind) {
    if (kind === 'conversion') return rows.filter(row => normalizeText(row.ad_form2) === '\u6536\u5272\u5e7f\u544a');
    if (kind === 'exposure') return rows.filter(row => normalizeText(row.ad_form2) === '\u79cd\u8349\u5e7f\u544a');
    return rows;
  }

  function rowsBySystemType(rows, systemType) {
    if (!systemType || systemType === 'all') return rows;
    return rows.filter(row => normalizeText(row.system_type) === systemType);
  }

  function rowsByPlatform(rows, platform) {
    if (!platform || platform === 'all') return rows;
    return rows.filter(row => normalizeText(row.platform) === platform);
  }

  function rowsByProduct(rows, products) {
    if (!Array.isArray(products) || !products.length) return rows;
    const selected = new Set(products.map(normalizeText).filter(Boolean));
    return rows.filter(row => selected.has(normalizeText(row.product)));
  }

  function rowsByProductPicker(rows, products) {
    if (!Array.isArray(products) || !products.length) return rows;
    const selected = new Set(products.map(normalizeText).filter(Boolean));
    return rows.filter(row => selected.has(normalizeText(row.product || row.material_product)));
  }

  function rowsByProductCategory(rows, categories) {
    if (!Array.isArray(categories) || !categories.length) return rows;
    const selected = new Set(categories.map(normalizeText).filter(Boolean));
    const categoryOf = row => {
      if (typeof window.productCategory === 'function') {
        return normalizeText(window.productCategory(row.product || row.material_product || ''));
      }
      return '';
    };
    return rows.filter(row => selected.has(categoryOf(row)));
  }

  function rowsByMaterialPublishDate(rows, start, end) {
    const normalizedStart = normalizeText(start);
    const normalizedEnd = normalizeText(end);
    if (!normalizedStart && !normalizedEnd) return rows;
    return rows.filter(row => {
      const publishDate = normalizeText(row.publish_date);
      if (!publishDate) return false;
      if (normalizedStart && publishDate < normalizedStart) return false;
      if (normalizedEnd && publishDate > normalizedEnd) return false;
      return true;
    });
  }

  function rowsForProductCompare(rows) {
    return (Array.isArray(rows) ? rows : []).filter(row =>
      normalizeText(row.ad_type || row['骞垮憡绫诲瀷']).includes('鍚堝垱骞垮憡')
    );
  }

  function filterChip(label, value, key) {
    return `<span class="active-filter-chip">${label}锛?{escapeHtml(value)}<button type="button" data-clear-filter="${key}" aria-label="娓呴櫎绛涢€?>脳</button></span>`;
  }

  function renderActiveFilterSummary() {
    if (!activeFilterSummary) return;
    const chips = [];
    if (activeOverviewKind === 'conversion') chips.push(filterChip('骞垮憡绫诲瀷', '鍚堝垱杞寲', 'overview'));
    if (activeOverviewKind === 'exposure') chips.push(filterChip('骞垮憡绫诲瀷', '鍚堝垱鏇濆厜', 'overview'));
    if (activeCompareSystemType !== 'all') chips.push(filterChip('绯荤粺绫诲瀷', activeCompareSystemType, 'system_type'));
    if (activeComparePlatform !== 'all') chips.push(filterChip('骞冲彴', activeComparePlatform, 'platform'));
    activeCompareProducts.forEach(product => chips.push(filterChip('浜у搧', product, 'product:' + product)));
    (window.__activeProductSelections || []).forEach(product => chips.push(filterChip('浜у搧绛涢€?, product, 'product_picker:' + product)));
    (window.__activeProductCategorySelections || []).forEach(category => chips.push(filterChip('鍝佺被绛涢€?, category, 'product_category:' + category)));
    if (normalizeText(window.__activeMaterialPublishStart) || normalizeText(window.__activeMaterialPublishEnd)) {
      chips.push(filterChip('绱犳潗鍙戝竷鏃堕棿', `${normalizeText(window.__activeMaterialPublishStart) || '涓嶉檺'} ~ ${normalizeText(window.__activeMaterialPublishEnd) || '涓嶉檺'}`, 'material_publish'));
    }
    const clearButton = chips.length ? '<button type="button" class="clear-filters-btn" id="clearAllFilters">涓€閿竻闄ょ瓫閫?/button>' : '';
    activeFilterSummary.innerHTML = chips.length ? chips.join('') + clearButton : '<span class="active-filter-empty">褰撳墠鏈瓫閫?/span>';
    activeFilterSummary.querySelectorAll('[data-clear-filter]').forEach(btn => {
      btn.addEventListener('click', event => {
        event.stopPropagation();
        const key = btn.dataset.clearFilter || '';
        if (key === 'overview') activeOverviewKind = 'all';
        if (key === 'system_type') {
          activeCompareSystemType = 'all';
          window.__activeCompareSystemType = activeCompareSystemType;
        }
        if (key === 'platform') {
          activeComparePlatform = 'all';
          window.__activeComparePlatform = activeComparePlatform;
        }
        if (key.startsWith('product:')) {
          const value = key.slice('product:'.length);
          activeCompareProducts = activeCompareProducts.filter(item => item !== value);
          window.__activeCompareProducts = activeCompareProducts;
        }
        if (key.startsWith('product_picker:')) {
          const value = key.slice('product_picker:'.length);
          const next = (window.__activeProductSelections || []).filter(item => item !== value);
          if (typeof window.__setProductListFilterState === 'function') window.__setProductListFilterState({ products: next });
        }
        if (key.startsWith('product_category:')) {
          const value = key.slice('product_category:'.length);
          const next = (window.__activeProductCategorySelections || []).filter(item => item !== value);
          if (typeof window.__setProductListFilterState === 'function') window.__setProductListFilterState({ categories: next });
        }
        if (key === 'material_publish') {
          if (typeof window.__setProductListFilterState === 'function') window.__setProductListFilterState({ materialPublishStart: '', materialPublishEnd: '' });
        }
        rerenderLinkedBoards();
      });
    });
    const clearAllBtn = document.getElementById('clearAllFilters');
    if (clearAllBtn) {
      clearAllBtn.addEventListener('click', () => {
        activeOverviewKind = 'all';
        activeCompareSystemType = 'all';
        activeComparePlatform = 'all';
        activeCompareProducts = [];
        window.__activeCompareSystemType = activeCompareSystemType;
        window.__activeComparePlatform = activeComparePlatform;
        window.__activeCompareProducts = activeCompareProducts;
        if (typeof window.__setProductListFilterState === 'function') {
          window.__setProductListFilterState({ products: [], categories: [], postId: '', materialPublishStart: '', materialPublishEnd: '' });
        }
        rerenderLinkedBoards();
      });
    }
  }

  function applyGlobalFilters(rows, options = {}) {
    const includeOverview = options.includeOverview !== false;
    const includeSystemType = options.includeSystemType !== false;
    const includePlatform = options.includePlatform !== false;
    const includeProduct = options.includeProduct !== false;
    const includeProductPicker = options.includeProductPicker !== false;
    const includeProductCategory = options.includeProductCategory !== false;
    const includeMaterialPublish = options.includeMaterialPublish !== false;
    let scoped = Array.isArray(rows) ? rows.slice() : [];
    if (includeSystemType) scoped = rowsBySystemType(scoped, activeCompareSystemType);
    if (includePlatform) scoped = rowsByPlatform(scoped, activeComparePlatform);
    if (includeProduct) scoped = rowsByProduct(scoped, activeCompareProducts);
    if (includeProductPicker) scoped = rowsByProductPicker(scoped, window.__activeProductSelections || []);
    if (includeProductCategory) scoped = rowsByProductCategory(scoped, window.__activeProductCategorySelections || []);
    if (includeMaterialPublish) scoped = rowsByMaterialPublishDate(scoped, window.__activeMaterialPublishStart, window.__activeMaterialPublishEnd);
    if (includeOverview) scoped = rowsByOverviewKind(scoped, activeOverviewKind);
    return scoped;
  }

  function inDateRange(row, start, end) {
    const date = normalizeText(row.date);
    if (!date) return false;
    if (start && date < start) return false;
    if (end && date > end) return false;
    return true;
  }

  function currentVisibleRows() {
    const qStart = document.getElementById('qStart')?.value || '';
    const qEnd = document.getElementById('qEnd')?.value || qStart;
    const cStart = document.getElementById('cStart')?.value || '';
    const cEnd = document.getElementById('cEnd')?.value || cStart;
    return {
      queryRows: (embedded.queryRaw || []).filter(row => inDateRange(row, qStart, qEnd)),
      compareRows: (embedded.compareRaw || []).filter(row => inDateRange(row, cStart, cEnd)),
    };
  }

  function metricPack(rows, period) {
    const spend = rows.reduce((acc, row) => acc + Number(row.spend || 0), 0);
    const gmv = rows.reduce((acc, row) => acc + Number(row.gmv || 0), 0);
    const impressions = rows.reduce((acc, row) => acc + Number(row.impressions || 0), 0);
    const clicks = rows.reduce((acc, row) => acc + Number(row.clicks || 0), 0);
    const totalMaterialIds = new Set(
      rows
        .filter(row => Number(row.impressions || 0) > 0)
        .map(row => normalizeText(row.material_id))
        .filter(Boolean)
    );
    const activePeriod = period || { start: '', end: '' };
    const isTestNewRow = row => {
      const publishDate = normalizeText(row.publish_date);
      return publishDate && publishDate >= activePeriod.start && publishDate <= activePeriod.end;
    };
    const testNewRows = rows.filter(isTestNewRow);
    const testNewSpend = testNewRows.reduce((acc, row) => acc + Number(row.spend || 0), 0);
    const testNewExposedPostIds = new Set(
      testNewRows
        .filter(row => Number(row.impressions || 0) > 0)
        .map(row => normalizeText(row.material_id))
        .filter(Boolean)
    );
    return {
      spend,
      gmv,
      roi: spend ? gmv / spend : 0,
      impressions,
      cpm: impressions ? spend * 1000 / impressions : 0,
      total_materials: totalMaterialIds.size,
      ctr: impressions ? clicks / impressions : 0,
      test_new_spend: testNewSpend,
      test_new_posts: testNewExposedPostIds.size,
      test_new_spend_ratio: spend ? testNewSpend / spend : 0,
    };
  }

  function renderOverviewTable(sourceRows, compareSourceRows) {
    const wrap = document.getElementById('hechuangOverviewTable');
    if (!wrap) return;
    const queryPeriod = embedded.queryPeriod || { start: '', end: '' };
    const comparePeriod = embedded.comparePeriod || { start: '', end: '' };
    const rows = [
      { label: '\u5408\u521b\u8f6c\u5316', kind: 'conversion', ...metricPack(rowsByOverviewKind(sourceRows, 'conversion'), queryPeriod) },
      { label: '\u5408\u521b\u66dd\u5149', kind: 'exposure', ...metricPack(rowsByOverviewKind(sourceRows, 'exposure'), queryPeriod) },
    ];
    const compareRowsSource = compareSourceRows || [];
    const compareRows = [
      { label: '\u5408\u521b\u8f6c\u5316', kind: 'conversion', ...metricPack(rowsByOverviewKind(compareRowsSource, 'conversion'), comparePeriod) },
      { label: '\u5408\u521b\u66dd\u5149', kind: 'exposure', ...metricPack(rowsByOverviewKind(compareRowsSource, 'exposure'), comparePeriod) },
    ];
    const compareMap = new Map(compareRows.map(row => [row.kind, row]));
    const displayRows = rows.map(row => ({ ...row, __compare: compareMap.get(row.kind) || null }));
    wrap.innerHTML = `
      <table data-sortable="1">
        <thead><tr>
          <th data-key="ad_type" data-kind="text"><span class="th-label">\u5e7f\u544a\u7c7b\u578b</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="spend" data-kind="money"><span class="th-label">\u82b1\u8d39</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="roi" data-kind="score"><span class="th-label">ROI</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="cpm" data-kind="money"><span class="th-label">CPM</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="ctr" data-kind="pct"><span class="th-label">CTR</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="gmv" data-kind="money"><span class="th-label">GMV</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="impressions" data-kind="int"><span class="th-label">\u66dd\u5149</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="total_materials" data-kind="int"><span class="th-label">\u603b\u6295\u653e\u7d20\u6750\u6570\u91cf</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="test_new_posts" data-kind="int"><span class="th-label">\u6d4b\u65b0\u6570\u91cf</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="test_new_spend" data-kind="money"><span class="th-label">\u6d4b\u65b0\u82b1\u8d39</span><span class="col-resize-handle" data-resize="1"></span></th>
          <th data-key="test_new_spend_ratio" data-kind="pct"><span class="th-label">\u6d4b\u65b0\u82b1\u8d39\u5360\u6bd4</span><span class="col-resize-handle" data-resize="1"></span></th>
        </tr></thead>
        <tbody>
          ${displayRows.map(row => `
            <tr class="${activeOverviewKind === row.kind ? 'active-overview-row' : ''}">
              <td><button type="button" class="overview-chip" data-overview-kind="${row.kind}">${row.label}</button></td>
              <td class="num">${cellWithCompare(row, 'spend')}</td>
              <td class="num">${cellWithCompare(row, 'roi')}</td>
              <td class="num">${cellWithCompare(row, 'cpm')}</td>
              <td class="num">${cellWithCompare(row, 'ctr')}</td>
              <td class="num">${cellWithCompare(row, 'gmv')}</td>
              <td class="num">${cellWithCompare(row, 'impressions')}</td>
              <td class="num">${cellWithCompare(row, 'total_materials')}</td>
              <td class="num">${cellWithCompare(row, 'test_new_posts')}</td>
              <td class="num">${cellWithCompare(row, 'test_new_spend')}</td>
              <td class="num">${cellWithCompare(row, 'test_new_spend_ratio')}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
    wrap.querySelectorAll('[data-overview-kind]').forEach(btn => {
      btn.addEventListener('click', () => {
        const kind = btn.dataset.overviewKind || 'all';
        activeOverviewKind = activeOverviewKind === kind ? 'all' : kind;
        rerenderLinkedBoards();
      });
    });
  }

  function bindCompareTableLinkage(containerId, key, activeValue) {
    const wrap = document.getElementById(containerId);
    if (!wrap) return;
    const activeSet = Array.isArray(activeValue) ? new Set(activeValue.map(normalizeText)) : null;
    wrap.querySelectorAll('tbody tr').forEach(row => {
      const chip = row.querySelector('.compare-filter-chip');
      if (!chip) return;
      const value = normalizeText(chip.textContent);
      row.classList.add('linked-filter-row');
      const isActive = activeSet ? activeSet.has(value) : value === activeValue;
      row.classList.toggle('active-filter-row', isActive);
      chip.classList.toggle('active', isActive);
      chip.addEventListener('click', () => {
        if (key === 'system_type') {
          activeCompareSystemType = activeCompareSystemType === value ? 'all' : value;
          window.__activeCompareSystemType = activeCompareSystemType;
        } else if (key === 'platform') {
          activeComparePlatform = activeComparePlatform === value ? 'all' : value;
          window.__activeComparePlatform = activeComparePlatform;
        } else if (key === 'product') {
          const current = new Set(activeCompareProducts.map(normalizeText).filter(Boolean));
          if (current.has(value)) current.delete(value);
          else current.add(value);
          activeCompareProducts = Array.from(current);
          window.__activeCompareProducts = activeCompareProducts;
        }
        rerenderLinkedBoards();
      });
    });
  }

  function rerenderLinkedBoards() {
    const visibleRows = currentVisibleRows();
    if (typeof window.refreshAllBoards === 'function') {
      window.refreshAllBoards(visibleRows.queryRows, visibleRows.compareRows);
    }
  }

  function patchGlobalRefresh() {
    if (typeof window.refreshAllBoards !== 'function' || window.__originalRefreshAllBoards) return;
    window.__originalRefreshAllBoards = window.refreshAllBoards;
    window.refreshAllBoards = function(qRows, cRows) {
      const overviewQ = applyGlobalFilters(qRows, { includeOverview: false });
      const overviewC = applyGlobalFilters(cRows, { includeOverview: false });
      window.__systemCompareTableQueryRows = applyGlobalFilters(qRows, { includeSystemType: false });
      window.__systemCompareTableCompareRows = applyGlobalFilters(cRows, { includeSystemType: false });
      window.__platformCompareTableQueryRows = applyGlobalFilters(qRows, { includePlatform: false });
      window.__platformCompareTableCompareRows = applyGlobalFilters(cRows, { includePlatform: false });
      window.__productCompareTableQueryRows = rowsForProductCompare(applyGlobalFilters(qRows, { includeProduct: false }));
      window.__productCompareTableCompareRows = rowsForProductCompare(applyGlobalFilters(cRows, { includeProduct: false }));
      renderOverviewTable(overviewQ, overviewC);
      const scopedQ = applyGlobalFilters(qRows);
      const scopedC = applyGlobalFilters(cRows);
      const result = window.__originalRefreshAllBoards(scopedQ, scopedC);
      renderActiveFilterSummary();
      bindCompareTableLinkage('kolCompare', 'system_type', activeCompareSystemType);
      bindCompareTableLinkage('platformCompare', 'platform', activeComparePlatform);
      bindCompareTableLinkage('productCompare', 'product', activeCompareProducts);
      return result;
    };
  }

  patchGlobalRefresh();
  if (typeof window.refreshAllBoards === 'function') {
    window.refreshAllBoards(embedded.queryRaw || [], embedded.compareRaw || []);
  }
})();
</script>
"""
    return html_text.replace("</body>", inject_script + "\n</body>", 1)


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def extract_post_id_from_ad_name(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    match = re.search(r"甯栧瓙_(TH-[^_]+)", text, re.IGNORECASE)
    if match:
        return normalize_text(match.group(1))
    match = re.search(r"(TH-[A-Za-z0-9-]+)", text, re.IGNORECASE)
    if match:
        return normalize_text(match.group(1))
    return ""


def extract_kol_from_ad_name(value: object) -> str:
    text = normalize_text(value)
    if not text or "甯栧瓙_" not in text:
        return ""
    tail = text.split("甯栧瓙_", 1)[1]
    parts = [normalize_text(part) for part in tail.split("_")]
    if len(parts) < 2:
        return ""
    parts = parts[1:]
    if parts and re.fullmatch(r"(big|regular)\s*kol\s*绯荤粺", parts[-1], re.IGNORECASE):
        parts = parts[:-1]
    parts = [part for part in parts if part]
    return normalize_text("_".join(parts))


def normalize_match_key(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return re.sub(r"[_\s]+$", "", text)


def has_readable_material_code(value: object) -> bool:
    text = normalize_text(value)
    if not text or text == "-":
        return False
    if re.fullmatch(r"(big|regular)\s*kol\s*绯荤粺", text, re.IGNORECASE):
        return False
    if text == "鍏朵粬":
        return False
    return True


def first_non_empty(series: pd.Series, default: str = "") -> str:
    for value in series:
        text = normalize_text(value)
        if text:
            return text
    return default


def parse_gviz_json(text: str) -> dict:
    match = re.search(r"setResponse\((.*)\);", text, re.S)
    if not match:
        raise ValueError("Failed to parse DMS gviz response.")
    return json.loads(match.group(1))


def fetch_store_base_data() -> pd.DataFrame:
    workbook = pd.ExcelFile(WORKBOOK_XLSX_URL)
    target = next((name for name in workbook.sheet_names if normalize_text(name) == "搴楅摵鍩虹鏁版嵁"), None)
    if not target:
        raise ValueError("鏈壘鍒?sheet锛氬簵閾哄熀纭€鏁版嵁")
    df = workbook.parse(target)
    df.columns = [normalize_text(col) for col in df.columns]
    return df


def build_store_daily(store_df: pd.DataFrame) -> pd.DataFrame:
    required = ["鍟嗗搧鍚嶇О", "date", "Product Visitors (Visit)", "Sales (Confirmed Order) (THB)"]
    missing = [col for col in required if col not in store_df.columns]
    if missing:
        raise KeyError(f"Store base data is missing columns: {missing}")
    work = store_df[required].copy()
    work["鍟嗗搧鍚嶇О"] = work["鍟嗗搧鍚嶇О"].map(normalize_text)
    work["date"] = pd.to_datetime(work["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    work["store_visitors"] = pd.to_numeric(work["Product Visitors (Visit)"], errors="coerce").fillna(0)
    work["store_gmv_thb"] = (
        pd.to_numeric(work["Sales (Confirmed Order) (THB)"], errors="coerce").fillna(0) / 4.8
    )
    work = work[(work["鍟嗗搧鍚嶇О"] != "") & work["date"].notna()].copy()
    grouped = (
        work.groupby(["鍟嗗搧鍚嶇О", "date"], dropna=False)[["store_visitors", "store_gmv_thb"]]
        .sum()
        .reset_index()
        .rename(columns={"鍟嗗搧鍚嶇О": "product"})
    )
    return grouped


def inject_store_daily_embedded(html_text: str, store_daily: pd.DataFrame) -> str:
    match = re.search(r'(<script id="embeddedData" type="application/json">)(.*?)(</script>)', html_text, re.S)
    if not match:
        return html_text
    payload = json.loads(match.group(2))
    payload["storeDaily"] = [
        {
            "product": normalize_text(row["product"]),
            "date": normalize_text(row["date"]),
            "store_visitors": float(row["store_visitors"]),
            "store_gmv_thb": float(row["store_gmv_thb"]),
        }
        for _, row in store_daily.iterrows()
    ]
    replacement = match.group(1) + json.dumps(payload, ensure_ascii=False).replace("</", "<\\/") + match.group(3)
    return html_text[: match.start()] + replacement + html_text[match.end() :]


def enhance_html_with_store_metrics(html_text: str) -> str:
    inject_script = """
<script>
(function () {
  function normalizeStoreText(value) {
    return value === undefined || value === null ? '' : String(value).trim();
  }
  function storeMetricMap(start, end) {
    const map = new Map();
    (embedded.storeDaily || []).forEach(row => {
      const name = normalizeStoreText(row.product);
      const date = normalizeStoreText(row.date);
      if (!name || !date) return;
      if (start && date < start) return;
      if (end && date > end) return;
      const current = map.get(name) || { store_gmv_thb: 0, store_visitors: 0 };
      current.store_gmv_thb += Number(row.store_gmv_thb || 0);
      current.store_visitors += Number(row.store_visitors || 0);
      map.set(name, current);
    });
    return map;
  }

  const originalDeltaRows = deltaRows;
  deltaRows = function (qRows, cRows, key, options = {}) {
    const rows = originalDeltaRows(qRows, cRows, key);
    if (key !== 'product') return rows;
    const queryPeriod = options.queryPeriod || { start: '', end: '' };
    const comparePeriod = options.comparePeriod || { start: '', end: '' };
    const qMap = storeMetricMap(queryPeriod.start, queryPeriod.end);
    const cMap = storeMetricMap(comparePeriod.start, comparePeriod.end);
    return rows.map(row => {
      const name = normalizeStoreText(row.name);
      const qStore = qMap.get(name) || { store_gmv_thb: 0, store_visitors: 0 };
      const cStore = cMap.get(name) || { store_gmv_thb: 0, store_visitors: 0 };
      return {
        ...row,
        store_gmv_thb: qStore.store_gmv_thb,
        store_visitors: qStore.store_visitors,
        c: {
          ...(row.c || {}),
          store_gmv_thb: cStore.store_gmv_thb,
          store_visitors: cStore.store_visitors,
        },
      };
    });
  };

  const originalCompareTable = compareTable;
  compareTable = function (rows, label, limit, options = {}) {
    if (!['\\u4ea7\\u54c1', '浜у搧'].includes(label)) return originalCompareTable(rows, label, limit, options);
    if (!rows.length) return '<p class="small">\\u6682\\u65e0\\u73af\\u6bd4\\u6570\\u636e</p>';
    const activeValue = options.activeValue || 'all';
    const activeSet = Array.isArray(activeValue) ? new Set(activeValue.map(item => String(item || '').trim()).filter(Boolean)) : null;
    const forcedCompareCell = (row, key) => {
      const primary = cellValue(row, key);
      const compareRow = row.__compare || row.c || null;
      if (!compareRow) return primary;
      const meta = deltaMeta(row[key], compareRow[key]);
      if (!meta) return primary;
      return `<div class="cell-main">${primary}</div><div class="cell-sub ${meta.cls}"><strong>${meta.text}</strong></div>`;
    };
    const visibleRows = limit ? rows.slice(0, limit) : rows.slice();
    const normalized = visibleRows.map((row, index) => ({
      ...row,
      __compare: row.c || null,
      __order: index,
    }));
    const head = [
      ['name', label, 'text'],
      ['store_gmv_thb', '\\u5e97\\u94faGMV', 'money'],
      ['store_visitors', '\\u5e97\\u94fa\\u8bbf\\u5ba2', 'int'],
      ['spend', '\\u82b1\\u8d39', 'money'],
      ['spend_share', '\\u82b1\\u8d39\\u5360\\u6bd4', 'pct'],
      ['roi', 'ROI', 'roi'],
      ['cpm', 'CPM', 'money'],
      ['ctr', 'CTR', 'ctr_pct'],
      ['gmv', 'GMV', 'money'],
      ['impressions', '\\u66dd\\u5149', 'int'],
    ].map(([key, text, kind]) => `<th data-key="${key}" data-kind="${kind}"><span class="th-label">${text}</span><span class="col-resize-handle" data-resize="1"></span></th>`).join('');
    const body = normalized.map(row => {
      const name = row.name || '-';
      const isActive = activeSet ? activeSet.has(name) : (activeValue !== 'all' && name === activeValue);
      return `<tr class="linked-filter-row ${isActive ? 'active-filter-row' : ''}" data-order="${row.__order}">
        <td data-key="name" data-sort="${escapeHtml(name)}"><button type="button" class="compare-filter-chip ${isActive ? 'active' : ''}" data-filter-value="${escapeHtml(name)}">${escapeHtml(name)}</button></td>
        <td class="num" data-key="store_gmv_thb" data-sort="${escapeHtml(String(row.store_gmv_thb ?? ''))}">${forcedCompareCell(row, 'store_gmv_thb')}</td>
        <td class="num" data-key="store_visitors" data-sort="${escapeHtml(String(row.store_visitors ?? ''))}">${forcedCompareCell(row, 'store_visitors')}</td>
        <td class="num" data-key="spend" data-sort="${escapeHtml(String(row.spend ?? ''))}">${cellWithCompare(row, 'spend', 'money')}</td>
        <td class="num" data-key="spend_share" data-sort="${escapeHtml(String(row.spend_share ?? ''))}">${cellWithCompare(row, 'spend_share', 'pct')}</td>
        <td class="num" data-key="roi" data-sort="${escapeHtml(String(row.roi ?? ''))}">${cellWithCompare(row, 'roi', 'roi')}</td>
        <td class="num" data-key="cpm" data-sort="${escapeHtml(String(row.cpm ?? ''))}">${cellWithCompare(row, 'cpm', 'money')}</td>
        <td class="num" data-key="ctr" data-sort="${escapeHtml(String(row.ctr ?? ''))}">${cellWithCompare(row, 'ctr', 'ctr_pct')}</td>
        <td class="num" data-key="gmv" data-sort="${escapeHtml(String(row.gmv ?? ''))}">${cellWithCompare(row, 'gmv', 'money')}</td>
        <td class="num" data-key="impressions" data-sort="${escapeHtml(String(row.impressions ?? ''))}">${cellWithCompare(row, 'impressions', 'int')}</td>
      </tr>`;
    }).join('');
    return `<table class="compare-table" data-sortable="1"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
  };

  const originalRefreshDynamicCompare = refreshDynamicCompare;
  refreshDynamicCompare = function (qRows, cRows) {
    originalRefreshDynamicCompare(qRows, cRows);
    const qStart = document.getElementById('qStart')?.value || '';
    const qEnd = document.getElementById('qEnd')?.value || qStart;
    const cStart = document.getElementById('cStart')?.value || '';
    const cEnd = document.getElementById('cEnd')?.value || cStart;
    const productCompareQRows = window.__productCompareTableQueryRows || qRows;
    const productCompareCRows = window.__productCompareTableCompareRows || cRows;
    document.getElementById('productCompare').innerHTML = compareTable(
      deltaRows(productCompareQRows, productCompareCRows, 'product', {
        queryPeriod: { start: qStart, end: qEnd },
        comparePeriod: { start: cStart, end: cEnd },
      }),
      '\\u4ea7\\u54c1',
      null,
      { activeValue: window.__activeCompareProducts || [] }
    );
    applyDefaultTableSorts(document);
  };

  if (typeof currentRows === 'function' && typeof refreshAllBoards === 'function') {
    const rows = currentRows();
    refreshAllBoards(rows.qRows, rows.cRows);
  }
})();
</script>
"""
    return html_text.replace("</body>", inject_script + "\n</body>", 1)


def enhance_html_with_system_paid_materials_compare(html_text: str) -> str:
    return html_text


def enhance_html_with_batch_post_id_filter(html_text: str) -> str:
    replacements = [
        (
            r'''    .product-tools input[type="text"] {
      min-width: 220px;
      flex: 0 0 220px;
      height: 54px;
      border: 1px solid #d7dee8;
      border-radius: 12px;
      background: #fff;
      color: #445468;
      padding: 0 18px;
      font: inherit;
      font-size: 15px;
      font-weight: 500;
      box-sizing: border-box;
    }
    .product-tools input[type="text"]::placeholder { color: #8b98a9; }''',
            r'''    .product-tools textarea {
      min-width: 220px;
      flex: 0 0 220px;
      min-height: 54px;
      max-height: 140px;
      border: 1px solid #d7dee8;
      border-radius: 12px;
      background: #fff;
      color: #445468;
      padding: 10px 14px;
      font: inherit;
      font-size: 14px;
      font-weight: 500;
      line-height: 1.45;
      box-sizing: border-box;
      resize: vertical;
    }
    .product-tools textarea::placeholder { color: #8b98a9; }''',
        ),
        (
            "    let activeProductPostIdFilter = '';",
            "    let activeProductPostIdFilter = '';\n    let postIdFilterDebounceTimer = null;",
        ),
        (
            r'''    function filterProductBoardRows(materials, rawRows) {
      const postIdKeyword = String(activeProductPostIdFilter || '').trim().toLowerCase();
      const selectedProducts = new Set(activeProductSelections.map(item => String(item || '').trim()).filter(Boolean));
      const selectedCategories = new Set(activeProductCategorySelections.map(item => String(item || '').trim()).filter(Boolean));
      const scopedRawRows = rawRows.filter(row => {
        const product = row.product || row.material_product || '-';
        const category = productCategory(product);
        if (selectedProducts.size && !selectedProducts.has(String(product))) return false;
        if (selectedCategories.size && !selectedCategories.has(String(category))) return false;
        if (postIdKeyword) {
          const postId = String(row.post_id || '').trim().toLowerCase();
          if (!postId.includes(postIdKeyword)) return false;
        }
        return true;
      });
      const allowedMaterialIds = new Set(scopedRawRows.map(row => String(row.material_id || '').trim()).filter(Boolean));
      const scopedMaterials = materials.filter(row => {
        if (!allowedMaterialIds.has(String(row.material_id || '').trim())) return false;
        if (selectedProducts.size && !selectedProducts.has(String(row.product || ''))) return false;
        if (selectedCategories.size && !selectedCategories.has(String(productCategory(row.product)))) return false;
        if (postIdKeyword) {
          const joinedPostIds = Array.isArray(row.post_ids) ? row.post_ids.join(' ') : '';
          if (!joinedPostIds.toLowerCase().includes(postIdKeyword)) return false;
        }
        return true;
      });
      return { scopedMaterials, scopedRawRows };
    }''',
            r'''    function parsePostIdFilters(value) {
      const text = String(value || '').trim();
      if (!text) return [];
      const extractedIds = text.match(/TH-[A-Z0-9_-]+/gi);
      const candidates = extractedIds && extractedIds.length
        ? extractedIds
        : text.split(/[\s,，;；|]+/);
      return Array.from(new Set(candidates
        .map(item => String(item || '').trim().toLowerCase())
        .filter(item => item && !/^[-:]+$/.test(item))));
    }
    function matchesPostIdFilter(value, filters) {
      const postId = String(value || '').trim().toLowerCase();
      if (!filters.length) return true;
      if (!postId) return false;
      return filters.some(filter => filters.length === 1 ? postId.includes(filter) : postId === filter);
    }
    function filterProductBoardRows(materials, rawRows) {
      const postIdFilters = parsePostIdFilters(activeProductPostIdFilter);
      const selectedProducts = new Set(activeProductSelections.map(item => String(item || '').trim()).filter(Boolean));
      const selectedCategories = new Set(activeProductCategorySelections.map(item => String(item || '').trim()).filter(Boolean));
      const scopedRawRows = rawRows.filter(row => {
        const product = row.product || row.material_product || '-';
        const category = productCategory(product);
        if (selectedProducts.size && !selectedProducts.has(String(product))) return false;
        if (selectedCategories.size && !selectedCategories.has(String(category))) return false;
        if (!matchesPostIdFilter(row.post_id, postIdFilters)) return false;
        return true;
      });
      const allowedMaterialIds = new Set(scopedRawRows.map(row => String(row.material_id || '').trim()).filter(Boolean));
      const scopedMaterials = materials.filter(row => {
        if (!allowedMaterialIds.has(String(row.material_id || '').trim())) return false;
        if (selectedProducts.size && !selectedProducts.has(String(row.product || ''))) return false;
        if (selectedCategories.size && !selectedCategories.has(String(productCategory(row.product)))) return false;
        if (postIdFilters.length) {
          const postIds = Array.isArray(row.post_ids) ? row.post_ids : [];
          if (!postIds.some(postId => matchesPostIdFilter(postId, postIdFilters))) return false;
        }
        return true;
      });
      return { scopedMaterials, scopedRawRows };
    }''',
        ),
        (
            r'''<label for="productPostIdFilter">帖子编号</label><input id="productPostIdFilter" type="text" placeholder="输入帖子编号筛选" value="${escapeHtml(activeProductPostIdFilter)}">''',
            r'''<label for="productPostIdFilter">帖子编号（多值）</label><textarea id="productPostIdFilter" rows="2" spellcheck="false" placeholder="粘贴多个帖子编号">${escapeHtml(activeProductPostIdFilter)}</textarea>''',
        ),
        (
            r'''        postIdInput.addEventListener('input', () => {
          activeProductPostIdFilter = postIdInput.value || '';
          syncProductListFilterState();
          refreshAllBoards(currentRows().qRows, currentRows().cRows);
        });''',
            r'''        postIdInput.addEventListener('input', () => {
          activeProductPostIdFilter = postIdInput.value || '';
          syncProductListFilterState();
          clearTimeout(postIdFilterDebounceTimer);
          postIdFilterDebounceTimer = setTimeout(() => {
            refreshAllBoards(currentRows().qRows, currentRows().cRows);
            const restoredInput = document.getElementById('productPostIdFilter');
            if (restoredInput) {
              restoredInput.focus();
              restoredInput.setSelectionRange(restoredInput.value.length, restoredInput.value.length);
            }
          }, 180);
        });''',
        ),
    ]
    for old, new in replacements:
        if old not in html_text:
            raise ValueError(f"Batch post ID filter patch target not found: {old[:80]!r}")
        html_text = html_text.replace(old, new, 1)
    return html_text


def enhance_html_with_material_recommendations(html_text: str) -> str:
    replacements = [
        (
            '    details.product-detail .table-wrap { border-radius: 0; border-left: 0; border-right: 0; border-bottom: 0; }',
            '''    details.product-detail .table-wrap { border-radius: 0; border-left: 0; border-right: 0; border-bottom: 0; }
    .material-count-cell {
      padding: 11px 14px;
      background: #f2f5f8;
      color: #354152;
      font-weight: 700;
      text-align: left;
      border-top: 1px solid #dfe5ec;
    }''',
        ),
        (
            '''      addScores(list);
      return list.sort((a, b) => b.score - a.score);
    }
    function percentileScores(list, metric, high = true) {''',
            '''      addScores(list);
      return list.sort((a, b) => b.score - a.score);
    }
    function materialRecommendation(row, averageRoi, compareMap) {
      const spend = Number(row.spend || 0);
      const roi = Number(row.roi || 0);
      const compare = compareMap.get(String(row.material_id || ''));
      const compareRoi = compare ? Number(compare.roi || 0) : null;
      if (spend > 30 && roi > averageRoi * 1.5) return '潜力素材 P0';
      if (spend > 20 && roi > 4) return '进一步观察 P1';
      if (row.is_new && spend > 10 && roi < 2) return '新素材 立即停投 P2';
      if (spend < 10) return '低消耗素材';
      if (!row.is_new && spend > 50 && roi < averageRoi && compareRoi !== null && roi < compareRoi) return '老素材 建议停投 P3';
      return '常规观察';
    }
    function applyMaterialRecommendations(materials, compareMaterials, averageRoi = div(sum(materials, 'gmv'), sum(materials, 'spend'))) {
      const compareMap = new Map(compareMaterials.map(row => [String(row.material_id || ''), row]));
      return materials.map(row => ({
        ...row,
        action: materialRecommendation(row, averageRoi, compareMap)
      }));
    }
    function percentileScores(list, metric, high = true) {''',
        ),
        (
            '''    function productSelectorHtml(productRows, materials, rawRows, availableProducts, availableCategories) {
      const head = `''',
            '''    function productSelectorHtml(productRows, materials, rawRows, compareRawRows, averageRoi, availableProducts, availableCategories) {
      const head = `''',
        ),
        (
            '''      const details = productRows.map(summary => {
        const productRawRows = rawRows.filter(row => (row.product || row.material_product || '-') === summary.product);
        const rows = materialRows(productRawRows)
          .filter(row => {''',
            '''      const details = productRows.map(summary => {
        const productRawRows = rawRows.filter(row => (row.product || row.material_product || '-') === summary.product);
        const productCompareRawRows = compareRawRows.filter(row => (row.product || row.material_product || '-') === summary.product);
        const rows = applyMaterialRecommendations(materialRows(productRawRows), materialRows(productCompareRawRows), averageRoi)
          .filter(row => {''',
        ),
        (
            '          <div class="table-wrap">${tableHtml(rows, topColumns, null)}</div>',
            '          <div class="table-wrap">${tableHtml(rows, topColumns, null, { showMaterialCount: true })}</div>',
        ),
        (
            '''      const productBoardFiltered = filterProductBoardRows(materials, qRows);
      const productRows = productSummaryRows(productBoardFiltered.scopedMaterials, productBoardFiltered.scopedRawRows, materialPublishPeriod());
      document.getElementById('productSelectorBlock').innerHTML = productSelectorHtml(productRows, productBoardFiltered.scopedMaterials, productBoardFiltered.scopedRawRows, productOptions(qRows), categoryOptions(qRows));''',
            '''      const productBoardFiltered = filterProductBoardRows(materials, qRows);
      const productBoardCompareFiltered = filterProductBoardRows(compareMaterials, cRows);
      const averageRoi = div(sum(productBoardFiltered.scopedMaterials, 'gmv'), sum(productBoardFiltered.scopedMaterials, 'spend'));
      const recommendedMaterials = applyMaterialRecommendations(productBoardFiltered.scopedMaterials, productBoardCompareFiltered.scopedMaterials, averageRoi);
      const productRows = productSummaryRows(recommendedMaterials, productBoardFiltered.scopedRawRows, materialPublishPeriod());
      document.getElementById('productSelectorBlock').innerHTML = productSelectorHtml(productRows, recommendedMaterials, productBoardFiltered.scopedRawRows, productBoardCompareFiltered.scopedRawRows, averageRoi, productOptions(qRows), categoryOptions(qRows));''',
        ),
        (
            '''    function tableHtml(rows, columns, limit) {
      const part = limit ? rows.slice(0, limit) : rows;
      if (!part.length) return '<table><tbody><tr><td>暂无数据</td></tr></tbody></table>';
      const head = columns.map(([key, label, kind]) => `<th data-key="${key}" data-kind="${kind}"><span class="th-label">${label}</span><span class="col-resize-handle" data-resize="1"></span></th>`).join('');
      const body = part.map((row, index) => `<tr class="${row.is_new ? 'new-material' : ''}" data-order="${index}">` + columns.map(([key,, kind]) => {
        const raw = row[key];
        const sortValue = raw === undefined || raw === null ? '' : String(raw);
        return `<td class="${['money','int','score','roi','ctr_pct'].includes(kind) ? 'num' : ''}" data-key="${key}" data-sort="${escapeHtml(sortValue)}">${cellWithCompare(row, key, kind)}</td>`;
      }).join('') + '</tr>').join('');
      return `<table data-sortable="1"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
    }''',
            '''    function tableHtml(rows, columns, limit, options = {}) {
      const part = limit ? rows.slice(0, limit) : rows;
      const head = columns.map(([key, label, kind]) => `<th data-key="${key}" data-kind="${kind}"><span class="th-label">${label}</span><span class="col-resize-handle" data-resize="1"></span></th>`).join('');
      const body = part.length ? part.map((row, index) => `<tr class="${row.is_new ? 'new-material' : ''}" data-order="${index}">` + columns.map(([key,, kind]) => {
        const raw = row[key];
        const sortValue = raw === undefined || raw === null ? '' : String(raw);
        return `<td class="${['money','int','score','roi','ctr_pct'].includes(kind) ? 'num' : ''}" data-key="${key}" data-sort="${escapeHtml(sortValue)}">${cellWithCompare(row, key, kind)}</td>`;
      }).join('') + '</tr>').join('') : `<tr><td colspan="${columns.length}">暂无数据</td></tr>`;
      const foot = options.showMaterialCount
        ? `<tfoot><tr><td colspan="${columns.length}" class="material-count-cell">当前筛选共 ${fmtInt(rows.length)} 条素材</td></tr></tfoot>`
        : '';
      return `<table data-sortable="1"><thead><tr>${head}</tr></thead><tbody>${body}</tbody>${foot}</table>`;
    }''',
        ),
    ]
    for old, new in replacements:
        if old not in html_text:
            raise ValueError(f"Material recommendation patch target not found: {old[:80]!r}")
        html_text = html_text.replace(old, new, 1)
    return html_text


def enhance_html_with_ad_type_tabs(html_text: str) -> str:
    def replace_once(old: str, new: str, label: str) -> None:
        nonlocal html_text
        if old not in html_text:
            raise ValueError(f"Ad type tabs patch target not found: {label}")
        html_text = html_text.replace(old, new, 1)

    replace_once(
        """    .overview-chip {
      border: 1px solid var(--line);
      background: #f5fbfd;
      color: var(--ink);
      border-radius: 999px;
      padding: 8px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
    }
    .active-overview-row .overview-chip {
      background: var(--brand);
      color: #fff;
      border-color: var(--brand);
    }""",
        """    .ad-type-tabs {
      display: flex;
      gap: 4px;
      width: fit-content;
      margin: 20px 0 18px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #eef5f7;
    }
    .ad-type-tab {
      min-width: 132px;
      height: 40px;
      padding: 0 18px;
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-weight: 800;
      cursor: pointer;
      transition: color 160ms ease, background-color 160ms ease, box-shadow 160ms ease;
    }
    .ad-type-tab:hover { color: var(--ink); }
    .ad-type-tab.active {
      background: #fff;
      color: var(--brand);
      box-shadow: 0 1px 4px rgba(31, 65, 81, 0.14);
    }
    .ad-type-tab:focus-visible {
      outline: 3px solid rgba(24, 119, 147, 0.2);
      outline-offset: 2px;
    }
    .overview-label {
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      font-weight: 800;
      color: var(--brand);
      white-space: nowrap;
    }
    #hechuangOverviewTable table {
      min-width: 1120px;
    }""",
        "tab styles",
    )
    replace_once(
        """    <section id="hechuangOverviewSection">
      <h2>合创基础数据</h2>
      <p class="small">仅展示合创转化与合创曝光两行，点击可触发全局广告类型联动。</p>""",
        """    <nav class="ad-type-tabs" role="tablist" aria-label="合创数据类型">
      <button id="adTypeConversionTab" class="ad-type-tab active" type="button" role="tab" aria-selected="true" aria-controls="dashboardContent" data-ad-type-tab="conversion">合创转化</button>
      <button id="adTypeExposureTab" class="ad-type-tab" type="button" role="tab" aria-selected="false" aria-controls="dashboardContent" data-ad-type-tab="exposure">合创曝光</button>
    </nav>
    <section id="hechuangOverviewSection">
      <h2 id="activeAdTypeHeading">合创转化 · 基础数据</h2>
      <p class="small">当前看板仅统计所选广告类型；两个 Tab 的筛选条件会分别保留。</p>""",
        "tab markup",
    )
    replace_once("  <main>", "  <main id=\"dashboardContent\">", "dashboard main id")
    replace_once(
        """    function productSelectorHtml(productRows, materials, rawRows, compareRawRows, averageRoi, availableProducts, availableCategories) {
      const head = `""",
        """    function productSelectorHtml(productRows, materials, rawRows, compareRawRows, averageRoi, availableProducts, availableCategories) {
      const isConversionTab = window.__activeAdTypeTab !== 'exposure';
      const materialColumns = isConversionTab ? topColumns : topColumns.filter(([key]) => key !== 'action');
      const head = `""",
        "material columns",
    )
    replace_once(
        """        const productCompareRawRows = compareRawRows.filter(row => (row.product || row.material_product || '-') === summary.product);
        const rows = applyMaterialRecommendations(materialRows(productRawRows), materialRows(productCompareRawRows), averageRoi)
          .filter(row => {""",
        """        const productCompareRawRows = compareRawRows.filter(row => (row.product || row.material_product || '-') === summary.product);
        const productMaterials = materialRows(productRawRows);
        const rows = (isConversionTab
          ? applyMaterialRecommendations(productMaterials, materialRows(productCompareRawRows), averageRoi)
          : productMaterials)
          .filter(row => {""",
        "conditional recommendations",
    )
    replace_once(
        "${tableHtml(rows, topColumns, null, { showMaterialCount: true })}",
        "${tableHtml(rows, materialColumns, null, { showMaterialCount: true })}",
        "exposure detail columns",
    )
    replace_once(
        """      const recommendedMaterials = applyMaterialRecommendations(productBoardFiltered.scopedMaterials, productBoardCompareFiltered.scopedMaterials, averageRoi);
      const productRows = productSummaryRows(recommendedMaterials, productBoardFiltered.scopedRawRows, materialPublishPeriod());
      document.getElementById('productSelectorBlock').innerHTML = productSelectorHtml(productRows, recommendedMaterials, productBoardFiltered.scopedRawRows, productBoardCompareFiltered.scopedRawRows, averageRoi, productOptions(qRows), categoryOptions(qRows));""",
        """      const displayMaterials = window.__activeAdTypeTab === 'exposure'
        ? productBoardFiltered.scopedMaterials
        : applyMaterialRecommendations(productBoardFiltered.scopedMaterials, productBoardCompareFiltered.scopedMaterials, averageRoi);
      const productRows = productSummaryRows(displayMaterials, productBoardFiltered.scopedRawRows, materialPublishPeriod());
      document.getElementById('productSelectorBlock').innerHTML = productSelectorHtml(productRows, displayMaterials, productBoardFiltered.scopedRawRows, productBoardCompareFiltered.scopedRawRows, averageRoi, productOptions(qRows), categoryOptions(qRows));""",
        "conditional material analysis",
    )
    replace_once(
        """  let activeOverviewKind = 'all';
  let activeCompareSystemType = 'all';
  let activeComparePlatform = 'all';
  let activeCompareProducts = [];
  window.__activeCompareSystemType = activeCompareSystemType;""",
        """  let activeOverviewKind = 'conversion';
  let activeCompareSystemType = 'all';
  let activeComparePlatform = 'all';
  let activeCompareProducts = [];
  const emptyAdTypeTabState = () => ({
    systemType: 'all',
    platform: 'all',
    compareProducts: [],
    products: [],
    categories: [],
    postId: '',
    materialPublishStart: '',
    materialPublishEnd: '',
  });
  const adTypeTabStates = {
    conversion: emptyAdTypeTabState(),
    exposure: emptyAdTypeTabState(),
  };
  window.__activeAdTypeTab = activeOverviewKind;
  window.__activeCompareSystemType = activeCompareSystemType;""",
        "tab state",
    )
    replace_once(
        """  function rowsByOverviewKind(rows, kind) {
    if (kind === 'conversion') return rows.filter(row => normalizeText(row.ad_form2) === '收割广告');
    if (kind === 'exposure') return rows.filter(row => normalizeText(row.ad_form2) === '种草广告');
    return rows;
  }

  function rowsBySystemType""",
        """  function rowsByOverviewKind(rows, kind) {
    if (kind === 'conversion') return rows.filter(row => normalizeText(row.ad_form2) === '收割广告');
    if (kind === 'exposure') return rows.filter(row => normalizeText(row.ad_form2) === '种草广告');
    return rows;
  }

  function captureAdTypeTabState() {
    return {
      systemType: activeCompareSystemType,
      platform: activeComparePlatform,
      compareProducts: activeCompareProducts.slice(),
      products: (window.__activeProductSelections || []).slice(),
      categories: (window.__activeProductCategorySelections || []).slice(),
      postId: normalizeText(window.__activeProductPostIdFilter),
      materialPublishStart: normalizeText(window.__activeMaterialPublishStart),
      materialPublishEnd: normalizeText(window.__activeMaterialPublishEnd),
    };
  }

  function restoreAdTypeTabState(state) {
    const next = state || emptyAdTypeTabState();
    activeCompareSystemType = next.systemType || 'all';
    activeComparePlatform = next.platform || 'all';
    activeCompareProducts = Array.isArray(next.compareProducts) ? next.compareProducts.slice() : [];
    window.__activeCompareSystemType = activeCompareSystemType;
    window.__activeComparePlatform = activeComparePlatform;
    window.__activeCompareProducts = activeCompareProducts;
    if (typeof window.__setProductListFilterState === 'function') {
      window.__setProductListFilterState({
        products: Array.isArray(next.products) ? next.products : [],
        categories: Array.isArray(next.categories) ? next.categories : [],
        postId: next.postId || '',
        materialPublishStart: next.materialPublishStart || '',
        materialPublishEnd: next.materialPublishEnd || '',
      });
    }
  }

  function syncAdTypeTabs() {
    document.querySelectorAll('[data-ad-type-tab]').forEach(tab => {
      const isActive = tab.dataset.adTypeTab === activeOverviewKind;
      tab.classList.toggle('active', isActive);
      tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
      tab.tabIndex = isActive ? 0 : -1;
    });
    const heading = document.getElementById('activeAdTypeHeading');
    if (heading) heading.textContent = `${activeOverviewKind === 'exposure' ? '合创曝光' : '合创转化'} · 基础数据`;
  }

  function activateAdTypeTab(kind) {
    if (!['conversion', 'exposure'].includes(kind) || kind === activeOverviewKind) return;
    adTypeTabStates[activeOverviewKind] = captureAdTypeTabState();
    activeOverviewKind = kind;
    window.__activeAdTypeTab = kind;
    restoreAdTypeTabState(adTypeTabStates[kind]);
    syncAdTypeTabs();
    rerenderLinkedBoards();
  }

  function rowsBySystemType""",
        "tab state helpers",
    )
    replace_once("    if (activeOverviewKind === 'conversion') chips.push(filterChip('广告类型', '合创转化', 'overview'));\n", "", "conversion filter chip")
    replace_once("    if (activeOverviewKind === 'exposure') chips.push(filterChip('广告类型', '合创曝光', 'overview'));\n", "", "exposure filter chip")
    replace_once("        if (key === 'overview') activeOverviewKind = 'all';\n", "", "clear tab filter")
    replace_once("        activeOverviewKind = 'all';\n", "", "clear all tab filter")
    replace_once(
        """    if (!wrap) return;
    const queryPeriod = embedded.queryPeriod || { start: '', end: '' };
    const comparePeriod = embedded.comparePeriod || { start: '', end: '' };
    const rows = [
      { label: '合创转化', kind: 'conversion', ...metricPack(rowsByOverviewKind(sourceRows, 'conversion'), queryPeriod) },
      { label: '合创曝光', kind: 'exposure', ...metricPack(rowsByOverviewKind(sourceRows, 'exposure'), queryPeriod) },
    ];
    const compareRowsSource = compareSourceRows || [];
    const compareRows = [
      { label: '合创转化', kind: 'conversion', ...metricPack(rowsByOverviewKind(compareRowsSource, 'conversion'), comparePeriod) },
      { label: '合创曝光', kind: 'exposure', ...metricPack(rowsByOverviewKind(compareRowsSource, 'exposure'), comparePeriod) },
    ];""",
        """    if (!wrap) return;
    syncAdTypeTabs();
    const queryPeriod = embedded.queryPeriod || { start: '', end: '' };
    const comparePeriod = embedded.comparePeriod || { start: '', end: '' };
    const label = activeOverviewKind === 'exposure' ? '合创曝光' : '合创转化';
    const rows = [{ label, kind: activeOverviewKind, ...metricPack(rowsByOverviewKind(sourceRows, activeOverviewKind), queryPeriod) }];
    const compareRowsSource = compareSourceRows || [];
    const compareRows = [{ label, kind: activeOverviewKind, ...metricPack(rowsByOverviewKind(compareRowsSource, activeOverviewKind), comparePeriod) }];""",
        "active overview row",
    )
    replace_once(
        """            <tr class="${activeOverviewKind === row.kind ? 'active-overview-row' : ''}">
              <td><button type="button" class="overview-chip" data-overview-kind="${row.kind}">${row.label}</button></td>""",
        """            <tr>
              <td><span class="overview-label">${row.label}</span></td>""",
        "overview row label",
    )
    replace_once(
        """    wrap.querySelectorAll('[data-overview-kind]').forEach(btn => {
      btn.addEventListener('click', () => {
        const kind = btn.dataset.overviewKind || 'all';
        activeOverviewKind = activeOverviewKind === kind ? 'all' : kind;
        rerenderLinkedBoards();
      });
    });
""",
        "",
        "old overview click binding",
    )
    replace_once(
        """  patchGlobalRefresh();
  if (typeof window.refreshAllBoards === 'function') {""",
        """  patchGlobalRefresh();
  document.querySelectorAll('[data-ad-type-tab]').forEach(tab => {
    tab.addEventListener('click', () => activateAdTypeTab(tab.dataset.adTypeTab || 'conversion'));
    tab.addEventListener('keydown', event => {
      if (!['ArrowLeft', 'ArrowRight'].includes(event.key)) return;
      event.preventDefault();
      const nextKind = activeOverviewKind === 'conversion' ? 'exposure' : 'conversion';
      activateAdTypeTab(nextKind);
      document.querySelector(`[data-ad-type-tab="${nextKind}"]`)?.focus();
    });
  });
  syncAdTypeTabs();
  if (typeof window.refreshAllBoards === 'function') {""",
        "tab event binding",
    )
    return html_text


def enhance_html_with_trend_scaling(html_text: str) -> str:
    toolbar_css = """
    .trend-toolbar {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      margin-top: 14px;
    }
    .trend-toolbar-label {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .trend-mode-toggle {
      display: inline-grid;
      grid-template-columns: repeat(2, minmax(72px, 1fr));
      padding: 3px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #f4f7f9;
    }
    .trend-mode-button {
      min-height: 30px;
      padding: 0 12px;
      border: 0;
      border-radius: 5px;
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }
    .trend-mode-button:hover { color: var(--brand); }
    .trend-mode-button.active {
      background: #fff;
      color: var(--brand);
      box-shadow: 0 1px 4px rgba(31, 65, 81, 0.14);
    }
    .trend-mode-button:focus-visible {
      outline: 2px solid rgba(24, 119, 147, 0.28);
      outline-offset: 1px;
    }
"""
    toolbar_html = """<div class="trend-toolbar" aria-label="曲线显示方式">
        <span class="trend-toolbar-label">曲线显示</span>
        <div class="trend-mode-toggle" role="group" aria-label="曲线显示方式">
          <button class="trend-mode-button" type="button" data-trend-mode="actual" title="按指标真实数值绘制">实际值</button>
          <button class="trend-mode-button" type="button" data-trend-mode="index" title="各曲线首个有效日设为 100">趋势</button>
        </div>
      </div>
      """
    html_text = html_text.replace("</style>", toolbar_css + "\n</style>", 1)
    html_text = html_text.replace(
        '<div class="trend-canvas-wrap"><canvas id="compareTrendChart"></canvas></div>',
        toolbar_html + '<div class="trend-canvas-wrap"><canvas id="compareTrendChart"></canvas></div>',
        1,
    )

    replacements = [
        (
            "    let activeMetrics = ['impressions', 'ctr'];",
            """    let activeMetrics = ['impressions', 'ctr'];
    let trendScaleMode = 'auto';
    let lastTrendQueryRows = [];
    let lastTrendCompareRows = [];""",
        ),
        (
            """    function metricAxisGroup(metric) {
      return ['roi', 'cpm', 'ctr'].includes(metric) ? 'right' : 'left';
    }
    const hiddenTrendSeries = new Set();""",
            """    function metricAxisGroup(metric) {
      return ['roi', 'cpm', 'ctr'].includes(metric) ? 'right' : 'left';
    }
    function resolvedTrendScaleMode() {
      return trendScaleMode === 'auto'
        ? (activeMetricList().length > 1 ? 'index' : 'actual')
        : trendScaleMode;
    }
    function syncTrendScaleControls(mode = resolvedTrendScaleMode()) {
      document.querySelectorAll('.trend-mode-button[data-trend-mode]').forEach(button => {
        const active = button.dataset.trendMode === mode;
        button.classList.toggle('active', active);
        button.setAttribute('aria-pressed', active ? 'true' : 'false');
      });
    }
    function bindTrendScaleControls() {
      document.querySelectorAll('.trend-mode-button[data-trend-mode]').forEach(button => {
        if (button.dataset.trendModeBound === '1') return;
        button.dataset.trendModeBound = '1';
        button.addEventListener('click', () => {
          trendScaleMode = button.dataset.trendMode;
          const redraw = () => drawTrend(lastTrendQueryRows, lastTrendCompareRows);
          const wrap = document.querySelector('.trend-canvas-wrap');
          if (!wrap || window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
            redraw();
            return;
          }
          wrap.classList.remove('trend-entering');
          wrap.classList.add('trend-refreshing');
          requestAnimationFrame(() => requestAnimationFrame(() => {
            redraw();
            wrap.classList.remove('trend-refreshing');
            wrap.classList.add('trend-entering');
            window.setTimeout(() => wrap.classList.remove('trend-entering'), 420);
          }));
        });
      });
      syncTrendScaleControls();
    }
    const hiddenTrendSeries = new Set();""",
        ),
        (
            """    function drawTrend(qRows, cRows) {
      const canvas = document.getElementById('compareTrendChart');""",
            """    function drawTrend(qRows, cRows) {
      const canvas = document.getElementById('compareTrendChart');
      lastTrendQueryRows = qRows;
      lastTrendCompareRows = cRows;""",
        ),
        (
            """      const selectedMetrics = activeMetricList();
      const qDaily = byDay(qRows);
      const cDaily = byDay(cRows);
      const len = Math.max(qDaily.length, cDaily.length, 1);
      const leftMetrics = selectedMetrics.filter(metric => metricAxisGroup(metric) === 'left');
      const rightMetrics = selectedMetrics.filter(metric => metricAxisGroup(metric) === 'right');
      function axisRange(metricsForAxis) {
        const values = metricsForAxis.flatMap(metric =>
          qDaily.map(r => metricValue(r, metric)).concat(cDaily.map(r => metricValue(r, metric)))
        ).filter(v => Number.isFinite(v));
        if (!values.length) return { min: 0, max: 1 };
        const rawMin = Math.min(...values);
        const rawMax = Math.max(...values);
        const span = rawMax - rawMin;
        const onlySmall = metricsForAxis.every(metric => ['roi', 'cpm', 'ctr'].includes(metric));
        const padding = span === 0 ? Math.max(Math.abs(rawMax) * 0.2, onlySmall ? 0.001 : 1) : span * 0.18;
        let min = Math.max(0, rawMin - padding);
        let max = rawMax + padding;
        if (max === min) max = min + 1;
        return { min, max };
      }""",
            """      const selectedMetrics = activeMetricList();
      const scaleMode = resolvedTrendScaleMode();
      syncTrendScaleControls(scaleMode);
      const qDaily = byDay(qRows);
      const cDaily = byDay(cRows);
      const len = Math.max(qDaily.length, cDaily.length, 1);
      function firstValidBaseline(rows, metric) {
        const values = rows.map(row => metricValue(row, metric)).filter(value => Number.isFinite(value));
        return values.find(value => value !== 0) || 1;
      }
      const qBaselines = new Map(selectedMetrics.map(metric => [metric, firstValidBaseline(qDaily, metric)]));
      const cBaselines = new Map(selectedMetrics.map(metric => [metric, firstValidBaseline(cDaily, metric)]));
      function plotValue(row, metric, series) {
        const rawValue = metricValue(row, metric);
        if (scaleMode === 'actual') return rawValue;
        const baseline = series === 'query' ? qBaselines.get(metric) : cBaselines.get(metric);
        return baseline ? rawValue / baseline * 100 : 0;
      }
      function plotValueLabel(value, metric) {
        if (scaleMode === 'actual') return fmtAxis(value, metric);
        return Number(value || 0).toFixed(Math.abs(value) < 10 ? 1 : 0);
      }
      const useSingleAxis = scaleMode === 'index' || selectedMetrics.length === 1;
      const leftMetrics = useSingleAxis
        ? selectedMetrics
        : selectedMetrics.filter(metric => metricAxisGroup(metric) === 'left');
      const rightMetrics = useSingleAxis
        ? []
        : selectedMetrics.filter(metric => metricAxisGroup(metric) === 'right');
      function axisRange(metricsForAxis) {
        const values = metricsForAxis.flatMap(metric =>
          qDaily.map(row => plotValue(row, metric, 'query'))
            .concat(cDaily.map(row => plotValue(row, metric, 'compare')))
        ).filter(v => Number.isFinite(v));
        if (!values.length) return { min: 0, max: 1 };
        const rawMin = Math.min(...values);
        const rawMax = Math.max(...values);
        const span = rawMax - rawMin;
        const onlySmall = metricsForAxis.every(metric => ['roi', 'cpm', 'ctr'].includes(metric));
        const minimumPadding = scaleMode === 'index'
          ? 5
          : Math.max(Math.abs(rawMax) * 0.08, onlySmall ? 0.001 : 1);
        const padding = span === 0 ? minimumPadding : span * 0.1;
        let min = rawMin - padding;
        let max = rawMax + padding;
        if (scaleMode === 'actual' && rawMin >= 0 && min < 0 && rawMin <= padding) min = 0;
        if (max === min) max = min + 1;
        return { min, max };
      }""",
        ),
        (
            "      const yForMetric = metric => metricAxisGroup(metric) === 'right' ? yRight : yLeft;",
            "      const yForMetric = metric => rightMetrics.includes(metric) ? yRight : yLeft;",
        ),
        (
            "      ctx.fillText(selectedMetrics.map(metricDisplayName).join(' / '), pad.left, 14);",
            """      const chartTitle = scaleMode === 'index'
          ? '趋势指数（首个有效日 = 100）'
          : selectedMetrics.map(metricDisplayName).join(' / ');
        ctx.fillText(chartTitle, pad.left, 14);""",
        ),
        (
            "const leftLabel = fmtAxis(leftRange.min + (leftRange.max - leftRange.min) * t, leftMetrics[0]);",
            "const leftLabel = plotValueLabel(leftRange.min + (leftRange.max - leftRange.min) * t, leftMetrics[0]);",
        ),
        (
            "const rightLabel = fmtAxis(rightRange.min + (rightRange.max - rightRange.min) * t, rightMetrics[0]);",
            "const rightLabel = plotValueLabel(rightRange.min + (rightRange.max - rightRange.min) * t, rightMetrics[0]);",
        ),
        (
            "const py = yScale(metricValue(row, metric));",
            "const py = yScale(plotValue(row, metric, series));",
        ),
        (
            """            const px = x(idx);
            const py = yScale(metricValue(row, metric));
            ctx.beginPath();""",
            """            const px = x(idx);
            const indexedValue = plotValue(row, metric, series);
            const py = yScale(indexedValue);
            ctx.beginPath();""",
        ),
        (
            "hoverPoints.push({ idx, series, metric, date: row.date, value: metricValue(row, metric), px, py, color });",
            "hoverPoints.push({ idx, series, metric, date: row.date, value: metricValue(row, metric), indexedValue, px, py, color });",
        ),
        (
            """            ctx.font = '11px Microsoft YaHei, Arial';
            ctx.fillStyle = color;
            ctx.fillText(fmtAxis(metricValue(first, metric), metric), x(0) + 5, yScale(metricValue(first, metric)) - 6);
            const lastText = fmtAxis(metricValue(last, metric), metric);
            ctx.fillText(lastText, x(data.length - 1) - ctx.measureText(lastText).width - 6, yScale(metricValue(last, metric)) - 6);""",
            """            const firstValue = plotValue(first, metric, series);
            const lastValue = plotValue(last, metric, series);
            ctx.font = '11px Microsoft YaHei, Arial';
            ctx.fillStyle = color;
            ctx.fillText(plotValueLabel(firstValue, metric), x(0) + 5, yScale(firstValue) - 6);
            const lastText = plotValueLabel(lastValue, metric);
            ctx.fillText(lastText, x(data.length - 1) - ctx.measureText(lastText).width - 6, yScale(lastValue) - 6);""",
        ),
        (
            """            const label = `${metricDisplayName(point.metric)} ${periodLabel}`;
            return `<div class="trend-tooltip-row"><span class="trend-tooltip-dot" style="background:${escapeHtml(point.color)}"></span><span>${escapeHtml(label)} ${escapeHtml(metricFormatter(point.metric, point.value))}</span></div>`;""",
            """            const label = `${metricDisplayName(point.metric)} ${periodLabel}`;
            const indexText = scaleMode === 'index' ? ` · 指数 ${plotValueLabel(point.indexedValue, point.metric)}` : '';
            return `<div class="trend-tooltip-row"><span class="trend-tooltip-dot" style="background:${escapeHtml(point.color)}"></span><span>${escapeHtml(label)} ${escapeHtml(metricFormatter(point.metric, point.value))}${escapeHtml(indexText)}</span></div>`;""",
        ),
        (
            "    bindDatePickers();\n    bindMaterialPublishFilters();",
            "    bindDatePickers();\n    bindTrendScaleControls();\n    bindMaterialPublishFilters();",
        ),
    ]
    for old, new in replacements:
        if old not in html_text:
            raise ValueError(f"Trend scaling patch target not found: {old[:80]!r}")
        html_text = html_text.replace(old, new, 1)
    return html_text


def enhance_html_with_interaction_polish(html_text: str) -> str:
    polish_css = """
    .cmp-card {
      position: relative;
      overflow: hidden;
      transform: translateY(0) scale(1);
      transition:
        transform 180ms cubic-bezier(.2,.75,.25,1),
        border-color 220ms ease,
        box-shadow 220ms ease,
        background-color 220ms ease;
      will-change: transform;
    }
    .cmp-card:hover {
      transform: translateY(-2px);
      border-color: rgba(24, 119, 147, 0.42);
      box-shadow: 0 8px 22px rgba(31, 65, 81, 0.09);
    }
    .cmp-card:active { transform: translateY(0) scale(.985); }
    .cmp-card:focus-visible {
      outline: 3px solid rgba(24, 119, 147, 0.18);
      outline-offset: 2px;
    }
    .cmp-card.active {
      transform: translateY(-1px);
      box-shadow: inset 0 0 0 2px var(--brand), 0 8px 20px rgba(31, 65, 81, 0.08);
    }
    .cmp-card.active:active { transform: translateY(0) scale(.985); }
    .cmp-card.active b { color: var(--brand); }
    .trend-canvas-wrap {
      overflow: hidden;
      transform: translateY(0) scale(1);
      opacity: 1;
      transition: opacity 130ms ease, transform 130ms ease;
      will-change: opacity, transform;
    }
    .trend-canvas-wrap.trend-refreshing {
      opacity: .38;
      transform: translateY(5px) scale(.997);
    }
    .trend-canvas-wrap.trend-entering {
      animation: trend-chart-reveal 360ms cubic-bezier(.2,.75,.25,1) both;
    }
    @keyframes trend-chart-reveal {
      from { opacity: .38; transform: translateY(7px) scale(.997); }
      to { opacity: 1; transform: translateY(0) scale(1); }
    }
    @media (prefers-reduced-motion: reduce) {
      .cmp-card,
      .trend-canvas-wrap {
        transition: none !important;
        animation: none !important;
        transform: none !important;
      }
    }
"""
    polish_script = """
<script>
(function() {
  const cardsRoot = document.getElementById('compareCards');
  const trendWrap = document.querySelector('.trend-canvas-wrap');
  if (!cardsRoot || !trendWrap) return;

  let revealTimer = 0;
  const syncPressedState = () => {
    cardsRoot.querySelectorAll('.cmp-card[data-metric]').forEach(card => {
      card.setAttribute('aria-pressed', card.classList.contains('active') ? 'true' : 'false');
    });
  };

  syncPressedState();
  new MutationObserver(syncPressedState).observe(cardsRoot, { childList: true, subtree: true });

  cardsRoot.addEventListener('click', event => {
    const card = event.target.closest('.cmp-card[data-metric]');
    if (!card || !cardsRoot.contains(card)) return;
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      requestAnimationFrame(syncPressedState);
      return;
    }

    window.clearTimeout(revealTimer);
    trendWrap.classList.remove('trend-entering');
    trendWrap.classList.add('trend-refreshing');
    requestAnimationFrame(() => requestAnimationFrame(() => {
      trendWrap.classList.remove('trend-refreshing');
      trendWrap.classList.add('trend-entering');
      syncPressedState();
      revealTimer = window.setTimeout(() => trendWrap.classList.remove('trend-entering'), 420);
    }));
  }, true);
})();
</script>
"""
    html_text = html_text.replace("</style>", polish_css + "\n</style>", 1)
    return html_text.replace("</body>", polish_script + "\n</body>", 1)


def fetch_all_data() -> pd.DataFrame:
    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            response = requests.get(ALL_DATA_CSV_URL, timeout=120)
            response.raise_for_status()
            csv_text = response.content.decode("utf-8-sig")
            df = pd.read_csv(StringIO(csv_text), low_memory=False)
            df.columns = [normalize_text(col) for col in df.columns]
            return df
        except (requests.RequestException, OSError, ValueError) as exc:
            last_error = exc
            if attempt < 3:
                time.sleep(3 * attempt)
            else:
                raise RuntimeError("Failed to fetch ALL_data after 3 attempts") from last_error


def fetch_dms() -> pd.DataFrame:
    response = requests.get(DMS_GVIZ_URL, timeout=120)
    response.raise_for_status()
    payload = parse_gviz_json(response.text)
    cols = [normalize_text(col.get("label")) for col in payload["table"]["cols"]]
    rows = []
    for row in payload["table"]["rows"]:
        cells = row.get("c", [])
        values = []
        for i in range(len(cols)):
            cell = cells[i] if i < len(cells) else None
            values.append(None if cell is None else cell.get("v"))
        rows.append(values)
    return pd.DataFrame(rows, columns=cols)


def build_dms_lookup(dms: pd.DataFrame) -> pd.DataFrame:
    keep_cols = ["绱犳潗缂栧彿", "绮変笣鏁?, "杈句汉閲忕骇", "username", "url", "绯荤粺绫诲瀷", "鍐呭骞冲彴", "鍙戝竷鏃堕棿"]
    missing = [col for col in keep_cols if col not in dms.columns]
    if missing:
        raise KeyError(f"DMS 缂哄皯瀛楁锛歿missing}")
    work = dms[keep_cols].copy()
    work["绱犳潗缂栧彿"] = work["绱犳潗缂栧彿"].map(normalize_text)
    work = work[work["绱犳潗缂栧彿"] != ""].copy()
    work["match_key"] = work["绱犳潗缂栧彿"].map(normalize_match_key)
    work["绮変笣鏁?] = pd.to_numeric(work["绮変笣鏁?], errors="coerce").fillna(0)
    grouped = (
        work.groupby("match_key", dropna=False)
        .agg(
            绱犳潗缂栧彿=("绱犳潗缂栧彿", first_non_empty),
            绮変笣鏁?("绮変笣鏁?, "max"),
            杈句汉閲忕骇=("杈句汉閲忕骇", first_non_empty),
            username=("username", first_non_empty),
            url=("url", first_non_empty),
            绯荤粺绫诲瀷=("绯荤粺绫诲瀷", first_non_empty),
            鍐呭骞冲彴=("鍐呭骞冲彴", first_non_empty),
            鍙戝竷鏃堕棿=("鍙戝竷鏃堕棿", first_non_empty),
        )
        .reset_index()
    )
    return grouped


def default_periods(filtered: pd.DataFrame) -> tuple[base.Period, base.Period]:
    today = pd.Timestamp(datetime.now().date())
    query_end = today - pd.Timedelta(days=1)
    query_start = today.replace(day=1)
    if query_end < query_start:
        query_end = query_start
    prev_month_end = query_start - pd.Timedelta(days=1)
    compare_start = prev_month_end.replace(day=1)
    compare_end = min(compare_start + pd.Timedelta(days=query_end.day - 1), prev_month_end)
    return (
        base.Period(label="鏌ヨ", start=query_start, end=query_end),
        base.Period(label="鐜瘮", start=compare_start, end=compare_end),
    )


def build_source_frame(all_data: pd.DataFrame, dms_lookup: pd.DataFrame) -> pd.DataFrame:
    required_cols = [
        "鑺辫垂閲戦锛圲SD锛?,
        "灞曠ず娆℃暟",
        "鐐瑰嚮閲?,
        "璐墿杞寲浠峰€?,
        "璁㈠崟鏁?,
        "鏃ユ湡",
        "骞垮憡绫诲瀷",
        "浜у搧",
        "绱犳潗缂栧彿",
        "甯栧瓙ID",
        "绯荤粺绫诲瀷",
        "URL",
        "骞垮憡鍚嶇О",
        "鍓ф湰鍚嶇О",
        "绱犳潗鍒涘缓鏃堕棿",
        "鍒涘缓鏃堕棿",
        "骞冲彴",
        "KOL",
        "骞垮憡褰㈠紡2",
    ]
    missing = [col for col in required_cols if col not in all_data.columns]
    if missing:
        raise KeyError(f"ALL_data 缂哄皯瀛楁锛歿missing}")

    work = all_data.copy()
    work["骞垮憡绫诲瀷"] = work["骞垮憡绫诲瀷"].map(normalize_text)
    work = work[work["骞垮憡绫诲瀷"].str.contains("鍚堝垱骞垮憡", na=False)].copy()
    if work.empty:
        raise ValueError("ALL_data 涓病鏈夊箍鍛婄被鍨?鍚堝垱骞垮憡鐨勬暟鎹€?)

    for col in ["甯栧瓙ID", "绱犳潗缂栧彿", "浜у搧", "绯荤粺绫诲瀷", "URL", "骞垮憡鍚嶇О", "鍓ф湰鍚嶇О", "骞冲彴", "KOL"]:
        work[col] = work[col].map(normalize_text)
    work["骞垮憡褰㈠紡2"] = work["骞垮憡褰㈠紡2"].map(normalize_text)
    work["parsed_post_id"] = work["骞垮憡鍚嶇О"].map(extract_post_id_from_ad_name)
    work["parsed_kol_name"] = work["骞垮憡鍚嶇О"].map(extract_kol_from_ad_name)

    work["dms_match_id"] = work["甯栧瓙ID"]
    work.loc[work["dms_match_id"] == "", "dms_match_id"] = work["parsed_post_id"]
    work.loc[work["dms_match_id"] == "", "dms_match_id"] = work["绱犳潗缂栧彿"]
    work["dms_match_key"] = work["dms_match_id"].map(normalize_match_key)

    merged = work.merge(
        dms_lookup,
        left_on="dms_match_key",
        right_on="match_key",
        how="left",
        suffixes=("", "_dms"),
    )
    dms_url = merged["url"].map(normalize_text) if "url" in merged.columns else ""

    merged["date"] = pd.to_datetime(merged["鏃ユ湡"], errors="coerce").dt.normalize()
    merged = merged[merged["date"].notna()].copy()
    merged["post_id"] = merged["甯栧瓙ID"]
    merged.loc[merged["post_id"] == "", "post_id"] = merged["parsed_post_id"]

    merged["material_id"] = merged["甯栧瓙ID"]
    merged.loc[merged["material_id"] == "", "material_id"] = merged["parsed_post_id"]
    merged.loc[merged["material_id"] == "", "material_id"] = merged["绱犳潗缂栧彿"]
    merged.loc[merged["material_id"] == "", "material_id"] = merged["URL"]
    merged.loc[merged["material_id"] == "", "material_id"] = merged["骞垮憡鍚嶇О"]
    merged = merged[merged["material_id"] != ""].copy()

    merged["material_name"] = merged["鍓ф湰鍚嶇О"]
    merged.loc[merged["material_name"] == "", "material_name"] = merged["骞垮憡鍚嶇О"]
    merged.loc[merged["material_name"] == "", "material_name"] = merged["material_id"]
    merged["material_code"] = merged["绱犳潗缂栧彿_dms"].map(normalize_text) if "绱犳潗缂栧彿_dms" in merged.columns else ""
    merged.loc[merged["material_code"] == "", "material_code"] = merged["绱犳潗缂栧彿"]
    merged.loc[merged["material_code"] == "", "material_code"] = merged["material_id"]
    merged["has_readable_material_code"] = merged["material_code"].map(has_readable_material_code)
    merged.loc[~merged["has_readable_material_code"], "material_code"] = "-"

    merged["product"] = merged["浜у搧"]
    merged["material_product"] = merged["浜у搧"]
    merged["platform"] = merged["骞冲彴"]
    merged.loc[merged["platform"] == "", "platform"] = merged["鍐呭骞冲彴"].map(normalize_text)
    merged.loc[merged["platform"] == "", "platform"] = "-"

    merged["username"] = merged["username"].map(normalize_text)
    merged.loc[merged["username"] == "", "username"] = merged["KOL"]
    merged.loc[merged["username"] == "", "username"] = merged["parsed_kol_name"]
    merged.loc[merged["username"] == "", "username"] = "-"

    merged["system_type"] = merged["绯荤粺绫诲瀷"]
    merged.loc[merged["system_type"] == "", "system_type"] = merged["绯荤粺绫诲瀷_dms"].map(normalize_text)
    merged.loc[merged["system_type"] == "", "system_type"] = "鍏朵粬"

    merged["url"] = merged["URL"]
    merged.loc[merged["url"] == "", "url"] = dms_url
    merged.loc[merged["url"] == "", "url"] = "-"
    merged["source_url"] = merged["url"]

    merged["fans"] = pd.to_numeric(merged["绮変笣鏁?], errors="coerce").fillna(0)
    merged["scale_raw"] = merged["杈句汉閲忕骇"].map(normalize_text)

    for src, dst in [
        ("鑺辫垂閲戦锛圲SD锛?, "spend"),
        ("灞曠ず娆℃暟", "impressions"),
        ("鐐瑰嚮閲?, "clicks"),
        ("璁㈠崟鏁?, "orders"),
        ("璐墿杞寲浠峰€?, "gmv"),
    ]:
        merged[dst] = pd.to_numeric(merged[src], errors="coerce").fillna(0)
        merged.loc[merged[dst] < 0, dst] = 0

    merged["add_cart_gmv"] = 0.0
    merged["reach"] = merged["impressions"]
    merged["completion_rate"] = 0.0

    dms_publish_candidates = pd.to_datetime(merged["鍙戝竷鏃堕棿"], errors="coerce")
    publish_candidates = pd.to_datetime(merged["绱犳潗鍒涘缓鏃堕棿"], errors="coerce")
    create_candidates = pd.to_datetime(merged["鍒涘缓鏃堕棿"], errors="coerce")
    merged["earliest_publish_time"] = dms_publish_candidates.fillna(publish_candidates).fillna(create_candidates)
    merged["publish_time"] = merged["earliest_publish_time"]
    merged["publish_date"] = merged["publish_time"].dt.strftime("%Y-%m-%d").fillna("")

    keep = [
        "date",
        "post_id",
        "material_id",
        "material_code",
        "has_readable_material_code",
        "material_name",
        "material_product",
        "product",
        "earliest_publish_time",
        "publish_time",
        "publish_date",
        "spend",
        "impressions",
        "clicks",
        "orders",
        "gmv",
        "add_cart_gmv",
        "reach",
        "completion_rate",
        "source_url",
        "fans",
        "scale_raw",
        "username",
        "url",
        "platform",
        "system_type",
        "ad_form2",
        "骞垮憡绫诲瀷",
    ]
    merged["ad_form2"] = merged["骞垮憡褰㈠紡2"]
    return merged[keep].copy()


def filter_period(df: pd.DataFrame, period: base.Period) -> pd.DataFrame:
    return df[(df["date"] >= period.start) & (df["date"] <= period.end)].copy()


def main() -> None:
    print("姝ｅ湪璇诲彇 Google Sheet: ALL_data / DMS")
    all_data = fetch_all_data()
    dms = fetch_dms()
    dms_lookup = build_dms_lookup(dms)

    source = build_source_frame(all_data, dms_lookup)
    query_period, compare_period = default_periods(all_data[all_data["骞垮憡绫诲瀷"].map(normalize_text) == "鍚堝垱骞垮憡"])

    query_raw = filter_period(source, query_period)
    compare_raw = filter_period(source, compare_period)
    if query_raw.empty:
        raise ValueError("鏌ヨ鍛ㄦ湡鍐呮病鏈夋湁鏁堝悎鍒涘箍鍛婃暟鎹€?)
    if compare_raw.empty:
        raise ValueError("鐜瘮鍛ㄦ湡鍐呮病鏈夋湁鏁堝悎鍒涘箍鍛婃暟鎹€?)

    material = base.aggregate_material(query_raw, query_period)
    compare_material = base.aggregate_material(compare_raw, query_period)
    generated_at_text = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

    html_text = base.make_html(
        query_period=query_period,
        compare_period=compare_period,
        query_files=[Path("GoogleSheet/ALL_data")],
        compare_files=[Path("GoogleSheet/ALL_data"), Path("GoogleSheet/DMS")],
        query_raw=query_raw,
        compare_raw=compare_raw,
        material=material,
        compare_material=compare_material,
        generated_at_text=generated_at_text,
    )
    html_text = enhance_html_for_live_refresh(html_text)
    html_text = enhance_html_with_overview_filter(html_text)
    html_text = enhance_html_with_system_paid_materials_compare(html_text)
    html_text = enhance_html_with_batch_post_id_filter(html_text)
    html_text = enhance_html_with_material_recommendations(html_text)
    html_text = enhance_html_with_ad_type_tabs(html_text)
    html_text = enhance_html_with_trend_scaling(html_text)
    html_text = enhance_html_with_interaction_polish(html_text)
    OUTPUT_PATH.write_text(html_text, encoding="utf-8")
    print(f"宸茬敓鎴愶細{OUTPUT_PATH}")
    print(f"鏌ヨ鍛ㄦ湡锛歿query_period.text}")
    print(f"鐜瘮鍛ㄦ湡锛歿compare_period.text}")
    print(f"ALL_data 鍚堝垱骞垮憡琛屾暟锛歿len(source):,}")
    print(f"DMS 绮変笣鍖归厤鍚庨潪闆剁矇涓濈礌鏉愭暟锛歿int((material['fans'] > 0).sum()):,}")


if __name__ == "__main__":
    main()
