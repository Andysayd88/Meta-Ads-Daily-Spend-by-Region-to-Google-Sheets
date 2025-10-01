/******************************************************************************************
 * @name Meta Ads Daily Spend by Region to Google Sheets — Multi‑Account 
 *
 * @overview
 * Pulls daily spend by *region* for multiple Meta Ads accounts and adds to a Google Sheet.
 * Designed for agencies or in-house teams looking for automated spend reporting.
 * 
 * Regions are things like US States.
 *
 * Columns written:
 *   Account ID | Account name | Region | Day | Currency | Amount spent (USD) | Reporting starts | Reporting ends
 *
 * @instructions
 * 1) Create (or open) a Google Sheet and note the tab name in CONFIG.SHEET_NAME.
 * 2) In Google Apps Script: Extensions → Apps Script → New project (standalone or bound).
 * 3) Paste this file. Open the CONFIG section and review:
 *    - ACCESS_TOKEN: Meta Marketing API token with `ads_read` for all accounts.
 *    - API_VERSION: Graph API version (e.g., v23.0).
 *    - ALLOW_ACCOUNT_IDS: list of ad accounts to include (e.g., ['act_123...', ...]).
 *    - DAYS_BACK: how many days back (excluding today) to fetch/upsert.
 *    - SHEET_NAME / TIMEZONE and other options as needed.
 * 4) Entry point: runDailyRegionSpend()
 *    - Use Triggers → Add Trigger → Head: runDailyRegionSpend → Time‑driven (e.g., daily).
 * 5) Authorize and run once to create headers; verify the tab shows upserts (no dupes).
 *
 * @author Sam Lalonde — https://www.linkedin.com/in/samlalonde/ — sam@samlalonde.com
 *
 * @license
 * MIT — Free to use, modify, and distribute. See https://opensource.org/licenses/MIT
 *
 * @version
 * 1.0
 *
 * @changelog
 * - v1.0
 *   - Initial public release.
 *   - Multi‑account fetch with regional breakdowns; stable UPSERT per (account_id, region, day).
 *   - Backoff & retry for rate‑limits/transient errors; resilient pagination.
 *   - Header auto‑creation and safe append/replace behavior for target sheet.
 ******************************************************************************************/


/* ========================= SETTINGS ========================= */
const CONFIG = {
  ACCESS_TOKEN  : 'INSERT_TOKEN_HERE', // add your Marketing API token here
  API_VERSION   : 'v23.0', 
  SHEET_NAME    : 'Meta Daily Region Spend', //Sheet name to be updated
  TIMEZONE      : 'America/Toronto', // Timezone
  DAYS_BACK     : 3, // This look back at previous days except today.
  LEVEL         : 'account',
  FIELDS        : ['account_id','account_name','account_currency','spend','date_start','date_stop'],
  BREAKDOWNS    : ['region'],
  TIME_INCREMENT: 1,
  REQUEST_TIMEOUT_MS: 60000,

  // Graph endpoints need the act_ prefix. Add one or multiple account ID.
  ALLOW_ACCOUNT_IDS: [
    'act_123456789012345',
    'act_123456789012345',
    'act_123456789012345'
  ]
};

/* ========================= MAIN ========================= */

function runDailyRegionSpend() {
  const { since, until } = getLastNDays(CONFIG.TIMEZONE, CONFIG.DAYS_BACK);
  const rows = fetchInsightsForAccounts(
    CONFIG.ALLOW_ACCOUNT_IDS.map(id => ({ id })),
    { since, until }
  );

  // Normalize -> exact sheet columns
  const normalized = rows.map(r => {
    const dayISO = toDayISO(r.date_start);
    const region = (r.region || 'Unknown').trim();
    const rawId = (r.account_id || '').trim();
    const accountId = normalizeAccountId(rawId);
    const spend = Number(r.spend || 0);

    return {
      key: `${accountId}|${region}|${dayISO}`,
      spend,
      values: [
        accountId,
        r.account_name || rawId || '',
        region,
        dayISO,
        r.account_currency || '',
        spend,
        toDayISO(r.date_start) || '',
        toDayISO(r.date_stop)  || ''
      ]
    };
  });

  upsertToSheet(CONFIG.SHEET_NAME, normalized);
}

/* ====================== HELPERS ========================= */

function normalizeAccountId(id) {
  return String(id || '').trim().replace(/^act_/, '');
}

function toDayISO(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return s; // last resort
}

function getLastNDays(tz, n) {
  const now = new Date();
  const startOfToday = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 00:00:00'));
  const end = new Date(startOfToday.getTime() - 24*60*60*1000);
  const start = new Date(end.getTime() - (n-1)*24*60*60*1000);
  return {
    since: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    until: Utilities.formatDate(end,   tz, 'yyyy-MM-dd')
  };
}

function fetchInsightsForAccounts(accounts, { since, until }) {
  const headers = { Authorization: `Bearer ${CONFIG.ACCESS_TOKEN}` };
  const out = [];
  for (const acct of accounts) {
    const base = `https://graph.facebook.com/${CONFIG.API_VERSION}/${acct.id}/insights`;
    const q = {
      level: CONFIG.LEVEL,
      fields: CONFIG.FIELDS.join(','),
      breakdowns: CONFIG.BREAKDOWNS.join(','),
      time_increment: CONFIG.TIME_INCREMENT,
      time_range: JSON.stringify({ since, until })
    };
    let url = base + '?' + toQuery(q);
    while (url) {
      const res = backoffFetch(url, { headers });
      const body = JSON.parse(res.getContentText());
      if (body.data?.length) out.push(...body.data);
      url = body.paging?.next || null;
    }
  }
  return out;
}

function ensureSheetHeader(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const targetHeader = ['Account ID','Account name','Region','Day','Currency','Amount spent (USD)','Reporting starts','Reporting ends'];

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,targetHeader.length).setValues([targetHeader]);
    return sh;
  }

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);

  if (header[0] !== 'Account ID') {
    sh.insertColumnBefore(1);
    sh.getRange(1,1).setValue('Account ID');
  }
  sh.getRange(1,1,1,targetHeader.length).setValues([targetHeader]);
  return sh;
}

function dedupeIncoming(rows) {
  const seen = new Map();
  for (const r of rows) seen.set(r.key, r);
  return Array.from(seen.values());
}

function upsertToSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureSheetHeader(ss, sheetName);

  const HEADER = ['Account ID','Account name','Region','Day','Currency','Amount spent (USD)','Reporting starts','Reporting ends'];
  const COL = {
    ACCOUNT_ID: 1,
    ACCOUNT_NAME: 2,
    REGION: 3,
    DAY: 4,
    CURRENCY: 5,
    AMOUNT: 6,
    START: 7,
    STOP: 8
  };

  const batch = dedupeIncoming(rows);
  const lastRow = sh.getLastRow();
  const byId   = new Map();
  const byName = new Map();

  if (lastRow > 1) {
    const data = sh.getRange(2,1,lastRow-1,HEADER.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const rowNum = i + 2;
      const accountId = normalizeAccountId(r[COL.ACCOUNT_ID-1]);
      const accountName = String(r[COL.ACCOUNT_NAME-1] || '').trim();
      const region = String(r[COL.REGION-1] || '').trim();
      const dayISO = toDayISO(r[COL.DAY-1]); // <-- normalize Day read from sheet
      if (dayISO) {
        if (accountId) {
          byId.set(`${accountId}|${region}|${dayISO}`, rowNum);
        } else if (accountName && region) {
          byName.set(`${accountName}|${region}|${dayISO}`, rowNum);
        }
      }
    }
  }

  const toAppend = [];
  const toUpdate = [];

  for (const r of batch) {
    if (!isFinite(r.spend) || r.spend <= 0) continue;

    const accountId = normalizeAccountId(r.values[0]);
    const accountName = String(r.values[1] || '').trim();
    const region = String(r.values[2] || '').trim();
    const dayISO = toDayISO(r.values[3]); 

    const idKey   = `${accountId}|${region}|${dayISO}`;
    const nameKey = `${accountName}|${region}|${dayISO}`;

    let rowNum = byId.get(idKey);
    if (!rowNum) {
      rowNum = byName.get(nameKey) || null;
      if (rowNum) {
        const write = r.values.slice();
        if (!write[0]) write[0] = accountId; 
        toUpdate.push({ rowNum, values: write });
        byId.set(idKey, rowNum);
        byName.delete(nameKey);
        continue;
      }
    }

    if (rowNum) {
      toUpdate.push({ rowNum, values: r.values });
    } else {
      toAppend.push(r.values);
    }
  }

  for (const u of toUpdate) {
    sh.getRange(u.rowNum, 1, 1, HEADER.length).setValues([u.values]);
  }

  if (toAppend.length) {
    sh.getRange(sh.getLastRow()+1, 1, toAppend.length, HEADER.length).setValues(toAppend);
  }

  const newLastRow = sh.getLastRow();
  if (newLastRow > 1) {
    sh.getRange(2, COL.AMOUNT, newLastRow-1, 1).setNumberFormat('$#,##0.00');
  }
}

/* ===================== UTILITIES ======================== */

function toQuery(obj) {
  return Object.keys(obj).filter(k => obj[k] !== '' && obj[k] != null)
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(obj[k])}`).join('&');
}

function backoffFetch(url, opts) {
  const params = Object.assign({ muteHttpExceptions: true, timeout: CONFIG.REQUEST_TIMEOUT_MS }, opts || {});
  const waits = [0, 500, 1500, 3000, 6000];
  for (let i=0;i<waits.length;i++){
    const res = UrlFetchApp.fetch(url, params);
    if (res.getResponseCode() === 200) return res;
    const body = safeJson(res.getContentText());
    const code = res.getResponseCode();
    const err = body?.error?.code;
    if ([4,17,32].includes(err) || [500,502,503,504].includes(code)) Utilities.sleep(waits[i]);
    else throw new Error(`HTTP ${code} ${res.getContentText()}`);
  }
  throw new Error('Retries exhausted: ' + url);
}

function safeJson(t){ try { return JSON.parse(t); } catch { return null; } }
