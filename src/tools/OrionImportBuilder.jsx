// Model Audit Tool v2.1 — dynamic P6 row detection
import { useState, useCallback, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Master model list (from Orion) ────────────────────────────────────────────
// Populated at runtime from the uploaded master file
// ── Status config ─────────────────────────────────────────────────────────────
const STATUS = {
  OVER:           { label: "Overweight",      color: "#dc2626", bg: "#fef2f2", dot: "#dc2626" },
  UNDER:          { label: "Underweight",     color: "#d97706", bg: "#fffbeb", dot: "#d97706" },
  MATCH:          { label: "On Target",       color: "#16a34a", bg: "#f0fdf4", dot: "#16a34a" },
  MISSING_MASTER: { label: "Not in Master",   color: "#7c3aed", bg: "#f5f3ff", dot: "#7c3aed" },
  MISSING_TARGET: { label: "No Target Entry", color: "#64748b", bg: "#f8fafc", dot: "#64748b" },
};
const DEFAULT_THRESHOLD = 0.05;

// ── Ticker validation ─────────────────────────────────────────────────────────
// Real tickers: 1–5 chars, letters/digits/dot/hyphen only, no spaces
// Cash tickers (CASH-USD, CUSTODIAL_CASH etc.) are exempted from length limits
const TICKER_RE = /^[A-Z0-9][A-Z0-9.\-]{0,4}$/;

// Explicit segment/label words that pass the char-length check but are NOT tickers
const SEGMENT_BLOCKLIST = new Set([
  "EQUITY","EQUITIES","FIXED","INCOME","BOND","BONDS","CASH","ALTERNATIVE",
  "ALTERNATIVES","COMMODITY","COMMODITIES","CRYPTO","SECTOR","NUCLEAR",
  "TICKER","SYMBOL","DOMESTIC","INTERNATIONAL","ALLOCATION","TOTAL",
  "EMERGING","MARKETS","GROWTH","VALUE","BLEND","CORE","SHOTS","THEMES",
]);

// Normalize cash tickers — catches CASH-USD, CUSTODIAL_CASH, money market tickers etc.
const CASH_TICKER_VARIANTS = /^(CASH[-_]?.*|CUSTODIAL.?CASH|SGOV|MMKT|FDRXX|SPAXX|FDLXX|SWVXX|VMFXX)$/i;
function isCashTicker(s) {
  return CASH_TICKER_VARIANTS.test(s.trim());
}

function isValidTicker(s) {
  if (!s) return false;
  const v = s.trim().toUpperCase();
  if (v.includes(" ")) return false;              // "US LARGE CAP" etc.
  // Cash tickers bypass all other checks — they are always valid holdings
  if (isCashTicker(v)) return true;
  if (SEGMENT_BLOCKLIST.has(v)) return false;
  if (!TICKER_RE.test(v)) return false;           // max 5 chars, valid chars only
  return true;
}

// ── Asset-class section keywords ──────────────────────────────────────────────
const AC_KEYWORDS = ["equity","equities","fixed income","bond","bonds","alternative","alternatives","cash","real estate","commodity","commodities","crypto","cryptocurrency","international","domestic","us equities","us fixed income","sector","nuclear","emerging","themes","shots","inflation","defense","drones","robotics","big tech","cybersecurity"];

function looksLikeAssetClass(cellVal) {
  if (!cellVal) return false;
  const v = String(cellVal).trim().toLowerCase();
  // Cash tickers are never section headers
  if (isCashTicker(cellVal)) return false;
  // Multi-word strings with "&" or "/" are section headers (e.g. "Defense & Geopolitical Realignment")
  if (/[&\/]/.test(v) && v.length > 5) return true;
  return AC_KEYWORDS.some(k => v === k || v.startsWith(k) || v.endsWith(k));
}

// Normalize ticker for comparison — strips dots/hyphens (BRK.B == BRKB),
// and maps all cash variants to a single canonical key (CASH == CUSTODIAL_CASH == CASH-USD)
const CASH_CANONICAL = "CUSTODIAL_CASH";
function normalizeTicker(t) {
  const v = String(t || "").toUpperCase().replace(/[.\-]/g, "");
  return isCashTicker(v) ? CASH_CANONICAL : v;
}

// Deduplicate positions by normalized ticker within a model — sum targets
function dedupePositions(positions) {
  const map = {};
  for (const p of positions) {
    const key = normalizeTicker(p.ticker);
    if (map[key]) {
      map[key].target += p.target;
    } else {
      map[key] = { ...p, _normKey: key };
    }
  }
  return Object.values(map);
}

function toTitleCase(s) {
  return String(s).replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
}

// ── Fuzzy string match (dice coefficient) ─────────────────────────────────────
function bigrams(s) {
  const b = new Set();
  for (let i = 0; i < s.length - 1; i++) b.add(s.slice(i, i+2));
  return b;
}
function diceScore(a, b) {
  if (!a || !b) return 0;
  const A = bigrams(a.toLowerCase()), B = bigrams(b.toLowerCase());
  let inter = 0;
  for (const bg of A) if (B.has(bg)) inter++;
  return (2 * inter) / (A.size + B.size);
}
// Normalize a label for matching: strip sheet prefix, expand known abbreviations
function normalizeForMatch(label) {
  if (label === "__skip__") return "__skip__";
  // P6 labels are already master-ready (emitted directly by parseP6Format)
  if (label.toLowerCase().startsWith("advisor model - p6")) return label.toLowerCase();

  const s = label.toLowerCase();
  const dashIdx = s.indexOf(" — ");
  const colLabel   = dashIdx >= 0 ? s.slice(dashIdx + 3) : s;
  const sheetPart  = dashIdx >= 0 ? s.slice(0, dashIdx) : "";
  const isTaxAware = s.includes("tax aw");
  const taxSuffix  = isTaxAware && !colLabel.includes("tax") ? " - tax aware" : "";

  // STP: "Savvy Total Portfolios (Tax Awa — Aggressive" → "stp - aggressive (tax aware)"
  if (sheetPart.includes("savvy total portfolios")) {
    const suffix = isTaxAware && !colLabel.includes("tax") ? " (tax aware)" : "";
    return `stp - ${colLabel}${suffix}`;
  }
  // Tilt variants: "Savvy Value Tilt (Tactical) — 50/50" → "savvy tactical model 50/50 - value tilt"
  if (sheetPart.includes("value tilt"))    return `savvy tactical model ${colLabel} - value tilt${taxSuffix}`;
  if (sheetPart.includes("growth tilt"))   return `savvy tactical model ${colLabel} - growth tilt${taxSuffix}`;
  if (sheetPart.includes("dividend tilt")) return `savvy tactical model ${colLabel} - dividend tilt${taxSuffix}`;
  // Tactical: "Savvy Core (Tactical) — 50/50" → "savvy tactical model 50/50"
  if (sheetPart.includes("tactical"))  return `savvy tactical model ${colLabel}${taxSuffix}`;
  // Strategic: "Savvy Core (Strategic) — 50/50" → "savvy strategic model 50/50"
  if (sheetPart.includes("strategic")) return `savvy strategic model ${colLabel}${taxSuffix}`;

  // STP midpoint labels are already master-ready
  if (label.match(/^STP - \d+\/\d+/)) return label.toLowerCase();
  // Savvy Strategic/Tactical normalized labels — pass through directly
  if (label.toLowerCase().startsWith("savvy strategic model")) return label.toLowerCase();
  if (label.toLowerCase().startsWith("savvy tactical model")) return label.toLowerCase();

  return colLabel;
}

function fuzzyMatch(query, candidates) {
  const normalizedQuery = normalizeForMatch(query);
  if (normalizedQuery === "__skip__") return null;
  const isTaxAwareQuery  = normalizedQuery.includes("(tax aware)") || normalizedQuery.includes("- tax aware");
  const queryIsTactical  = normalizedQuery.includes("tactical");
  const queryIsStrategic = normalizedQuery.includes("strategic");

  const normalizedNoTax = normalizedQuery
    .replace(/\s*\(tax aware\)\s*$/i, "")
    .replace(/\s*-\s*tax aware\s*$/i, "")
    .trim();

  // Extract ratio (e.g. "70/30") for hard filtering to prevent 0/100 ↔ 100/0 confusion
  const ratioMatch  = normalizedQuery.match(/\b(\d+\/\d+)\b/);
  const queryRatio  = ratioMatch ? ratioMatch[1] : null;

  // Tax-aware master exists if any active tax-aware candidate scores strongly
  const taxAwareMasterExists = isTaxAwareQuery && candidates.some(c =>
    c.name.toLowerCase().includes("tax aware") &&
    c.active !== false &&
    diceScore(normalizedQuery, c.name.toLowerCase()) > 0.85
  );

  const fallback      = isTaxAwareQuery && !taxAwareMasterExists;
  const queryToUse    = fallback ? normalizedNoTax : normalizedQuery;
  const rawQueryToUse = fallback ? normalizedNoTax : query.toLowerCase();

  // Restrict to candidates containing the exact ratio (if present)
  const ratioCandidates = queryRatio
    ? candidates.filter(c => c.name.toLowerCase().includes(queryRatio))
    : candidates;
  const pool = ratioCandidates.length > 0 ? ratioCandidates : candidates;

  // Precompute best raw score across pool — O(n)
  let bestRawScore = 0;
  for (const c of pool) {
    const s = Math.max(diceScore(queryToUse, c.name.toLowerCase()), diceScore(rawQueryToUse, c.name.toLowerCase()));
    if (s > bestRawScore) bestRawScore = s;
  }

  // Precompute whether a same-type active model scores within 3% of best — O(n)
  let hasSameTypeActiveNearBest = false;
  for (const c of pool) {
    if (c.active === false) continue;
    const cName = c.name.toLowerCase();
    if (queryIsTactical  && cName.includes("strategic")) continue;
    if (queryIsStrategic && cName.includes("tactical"))  continue;
    const s = Math.max(diceScore(queryToUse, cName), diceScore(rawQueryToUse, cName));
    if (s >= bestRawScore * 0.97) { hasSameTypeActiveNearBest = true; break; }
  }

  // Score each candidate — O(n), no inner loops
  let best = null, bestScore = 0;
  for (const c of pool) {
    const cName        = c.name.toLowerCase();
    const cIsTaxAware  = cName.includes("tax aware");
    const cIsStrategic = cName.includes("strategic");
    const cIsTactical  = cName.includes("tactical");
    const isActive     = c.active !== false;

    let s = Math.max(diceScore(queryToUse, cName), diceScore(rawQueryToUse, cName));

    // Tax parity penalty (0.749 breaks ties in favour of correct parity)
    if (isTaxAwareQuery  && !cIsTaxAware)  s *= 0.749;
    if (!isTaxAwareQuery && cIsTaxAware)   s *= 0.749;

    // Tactical/strategic type mismatch penalty
    if (queryIsTactical  && cIsStrategic)  s *= 0.5;
    if (queryIsStrategic && cIsTactical)   s *= 0.5;

    // Inactive penalty — only when a same-type active model is close enough
    if (!isActive && hasSameTypeActiveNearBest) s *= 0.4;

    if (s > bestScore) { bestScore = s; best = c; }
  }
  return bestScore > 0.2 ? { match: best, score: bestScore } : null;
}

// ── File reading ──────────────────────────────────────────────────────────────
function readXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: "array" });
  return wb;
}

function readFileBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onerror = rej;
    r.onload = e => res(e.target.result);
    r.readAsArrayBuffer(file);
  });
}

// ── Parse master file ─────────────────────────────────────────────────────────
function parseMasterFile(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  if (!rows.length) return { models: [], holdings: [] };

  const keys = Object.keys(rows[0]);
  const findCol = (...cs) => keys.find(k => cs.some(c => k.toLowerCase().includes(c.toLowerCase())));
  const tickerCol    = findCol("ticker","symbol");
  const nameCol      = findCol("product name","name","security");
  const targetCol    = findCol("target percent","target");
  const actualCol    = findCol("actual percent","actual");
  const modelNameCol = findCol("model aggregation name","model name");
  const modelIdCol   = findCol("model aggregation id","model id","eclipse");

  const modelsMap = {};
  const modelActuals = {};
  const holdings = [];
  for (const r of rows) {
    const ticker = String(r[tickerCol] || "").trim().toUpperCase();
    if (!ticker) continue;
    const modelId   = parseInt(r[modelIdCol]) || null;
    const modelName = String(r[modelNameCol] || "").trim();
    if (modelId && modelName && !modelsMap[modelId]) modelsMap[modelId] = modelName;
    if (modelId) modelActuals[modelId] = (modelActuals[modelId] || 0) + (parseFloat(r[actualCol]) || 0);
    // Use short keys to keep storage footprint small
    holdings.push({
      t:   ticker,
      n:   String(r[nameCol] || "").trim().slice(0, 40),
      p:   parseFloat(r[targetCol]) || 0,
      mid: modelId,
    });
  }

  const models = Object.entries(modelsMap).map(([id, name]) => ({
    id: parseInt(id),
    name,
    active: (modelActuals[parseInt(id)] || 0) > 0,
  }));
  return { models, holdings };
}

// Expand slim holdings back to full format for audit use
function expandHoldings(holdings, modelsMap) {
  return holdings.map(h => ({
    ticker:    h.t,
    name:      h.n,
    target:    h.p,
    modelId:   h.mid,
    modelName: modelsMap[h.mid] || "",
  }));
}

// ── Scale detection: are values already % (sum≈100) or decimal (sum≈1)? ──────
function detectScale(nums) {
  const sum = nums.reduce((s, n) => s + n, 0);
  if (sum < 5) return "decimal";   // sum near 1 → decimals, multiply by 100
  return "percent";                 // sum near 100 → already percent, use as-is
}
function applyScale(num, scale) {
  return scale === "decimal" ? num * 100 : num;
}

// ── Paragon "security set" format parser ─────────────────────────────────────
// Models sheet: row 1 = header (col A = model name, cols B+ = set names)
//               rows 2+ = model rows with set allocation weights (decimals)
// Remaining sheets: each is a security set with ticker + weight
//
// Effective ticker target in a model = sum over sets of (set_weight × ticker_weight × 100)

function isParagonFormat(wb) {
  // Heuristic: has a sheet called "Models" whose first row has non-numeric string headers in cols B+
  if (!wb.SheetNames.includes("Models")) return false;
  const ws = wb.Sheets["Models"];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!raw.length) return false;
  const header = raw[0];
  // Col A should be a label column, cols B+ should be security set names (strings)
  return header.length > 1 && typeof header[1] === "string" && header[1].length > 3;
}

function parseParagonFormat(wb) {
  const results = [];

  // ── 1. Parse Models sheet ──────────────────────────────────────────────────
  const modelsWs = wb.Sheets["Models"];
  const modelsRaw = XLSX.utils.sheet_to_json(modelsWs, { header: 1, defval: null });
  const header = modelsRaw[0]; // ["Paragon Asset Allocation Models", "Set A", "Set B", ...]
  // Map col index → set name (col 0 is model name label)
  const setColMap = {}; // colIdx → setName
  for (let ci = 1; ci < header.length; ci++) {
    if (header[ci]) setColMap[ci] = String(header[ci]).trim();
  }

  // Each model row: col 0 = model name, cols 1+ = set weight (decimal or null)
  const modelRows = []; // [{ name, sets: { setName: weight } }]
  for (let ri = 1; ri < modelsRaw.length; ri++) {
    const row = modelsRaw[ri];
    if (!row || !row[0]) continue;
    const modelName = String(row[0]).trim();
    const sets = {};
    for (const [ci, setName] of Object.entries(setColMap)) {
      const w = parseFloat(row[ci]);
      if (!isNaN(w) && w > 0) sets[setName] = w;
    }
    if (Object.keys(sets).length > 0) modelRows.push({ name: modelName, sets });
  }

  // ── 2. Parse each security set sheet ──────────────────────────────────────
  const setHoldings = {}; // setName → [{ ticker, name, weight (0-100) }]

  for (const sheetName of wb.SheetNames) {
    if (sheetName === "Models") continue;
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw.length) continue;

    const holdings = [];
    for (const row of raw) {
      if (!row) continue;
      const c0 = String(row[0] || "").trim();
      const c1 = String(row[1] || "").trim();
      const c2 = row[2] != null ? parseFloat(row[2]) : NaN;
      const c1num = parseFloat(c1);
      let ticker = null, name = null, rawWeight = null;

      if (!isNaN(c2) && c2 > 0 && isNaN(c1num) && c1.length >= 2 && c1.length <= 5) {
        // Format B: name | ticker | weight
        name = c0; ticker = c1.toUpperCase(); rawWeight = c2;
      } else if (!isNaN(c1num) && c1num > 0 && c0.length >= 1 && c0.length <= 7) {
        // Format A: ticker | weight
        ticker = c0.toUpperCase(); rawWeight = c1num;
      } else { continue; }

      if (!ticker) continue;
      ticker = ticker.replace(/\//g, ".");
      holdings.push({ ticker, name: name || "", rawWeight });
    }

    // Detect scale from all raw weights in this sheet
    const allWeights = holdings.map(h => h.rawWeight).filter(w => w > 0);
    const scale = detectScale(allWeights);
    const parsed = holdings.map(h => ({
      ticker: h.ticker,
      name: h.name,
      target: applyScale(h.rawWeight, scale),
    }));

    // Normalise so weights sum to 100
    const total = parsed.reduce((s, h) => s + h.target, 0);
    const factor = total > 0 ? 100 / total : 1;
    for (const h of parsed) h.target *= factor;
    setHoldings[sheetName] = parsed;
  }

  // ── 3. Blend sets into effective model positions ───────────────────────────
  for (const { name: modelName, sets } of modelRows) {
    const blended = {}; // normTicker → { ticker, name, target }

    for (const [setName, setWeight] of Object.entries(sets)) {
      const holdings = setHoldings[setName];
      if (!holdings) continue;
      for (const h of holdings) {
        const key = normalizeTicker(h.ticker);
        const contrib = setWeight * h.target; // set_weight(decimal) × ticker%(0-100) → effective %
        if (blended[key]) {
          blended[key].target += contrib;
        } else {
          blended[key] = { ticker: h.ticker, name: h.name, target: contrib, assetClass: "Equity" };
        }
      }
    }

    const positions = dedupePositions(Object.values(blended));
    if (!positions.length) continue;

    results.push({
      modelKey:   `paragon__${modelName}`,
      modelLabel: modelName,
      sheetName:  "Models",
      colLabel:   null,
      positions,
    });
  }

  return results;
}

// ── P6 "sleeve model" format parser ──────────────────────────────────────────
// Each sheet = a model family. Row 10 has "Models: X". Row 13 has Asset Class/Security/Symbol.
// Ticker in col D, model weights (decimals) in cols G/H/I and right-side cols.
// Labels are emitted ready for fuzzy matching against master — no normalizeForMatch needed.

function isP6Format(wb) {
  // Structural signature: scan for "Models:" row and "Asset Class"+"Symbol" row
  for (const sn of wb.SheetNames) {
    const ws  = wb.Sheets[sn];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (raw.length < 10) continue;
    const hasModelsRow  = raw.slice(0, 15).some(r => (r||[]).some(c => String(c||"").includes("Models:")));
    const hasAssetClass = raw.slice(0, 20).some(r => (r||[]).some(c => String(c||"") === "Asset Class") && (r||[]).some(c => String(c||"") === "Symbol"));
    if (hasModelsRow && hasAssetClass) return true;
  }
  return false;
}

// Build the master-model label for a P6 sheet+column combination.
// familyName comes from "Models:  Core+" → "Core+"
// sheetName is the tab name (e.g. "TaxEfficient", "ESG")
// colHeader is the column header (e.g. "Moderate", "#1", "Aggressive")
function p6MasterLabel(sheetName, colHeader) {
  const sheet = sheetName.toLowerCase().replace(/\s+/g,"");
  const col   = (colHeader || "").trim();
  const cl    = col.toLowerCase();

  // Skip columns that don't correspond to master models
  if (["tactical","strategic model","strategic","tactical model"].includes(cl)) return "__skip__";

  // Risk level mapping (#1/#2/#3 → Moderate/Mod Agg/Aggressive)
  const riskMap = { "#1":"Moderate", "#2":"Mod Agg", "#3":"Aggressive" };
  const risk = riskMap[col] || col;

  // Sheet → master name segment
  if (sheet === "taxefficient" || sheet.includes("taxeff")) {
    return `Advisor Model - P6 - TAX EFF ${risk}`;
  }
  if (sheet === "esg") {
    return `Advisor Model - P6 - ESG ${risk}`;
  }
  if (sheet === "top20" || sheet === "top 20" || sheetName.toLowerCase().includes("top 20")) {
    return `Advisor Model - P6 ${risk} - LC Stock Sub`;
  }
  // Core+ and any other sheet → base P6 model
  return `Advisor Model - P6 ${risk}`;
}

function parseP6Format(wb) {
  const results = [];

  for (const sheetName of wb.SheetNames) {
    const ws  = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw.length) continue;

    // ── Locate key rows dynamically (row offsets vary per sheet) ─────────────
    // Find "Models:" row
    let modelsRowIdx = -1;
    for (let ri = 0; ri < Math.min(raw.length, 15); ri++) {
      if ((raw[ri]||[]).some(c => String(c||"").includes("Models:"))) { modelsRowIdx = ri; break; }
    }
    if (modelsRowIdx < 0) continue;

    // Find "Asset Class"/"Symbol" header row (the main table header)
    let headerRowIdx = -1;
    for (let ri = modelsRowIdx; ri < Math.min(raw.length, modelsRowIdx + 8); ri++) {
      const row = raw[ri] || [];
      if (row.some(c => String(c||"") === "Asset Class") && row.some(c => String(c||"") === "Symbol")) {
        headerRowIdx = ri; break;
      }
    }
    if (headerRowIdx < 0) continue;

    // Sleeve weights are 3 rows before the "Models:" row (rows 7-9 = modelsRowIdx - 3 to -1)
    // Actually they're at fixed offsets relative to modelsRowIdx
    // Scan for rows with "Bond (Adjusted)"/"Stock (Adjusted)"/"Over/Under" in col C (index 2)
    const sleeveWeights = {};
    for (let ri = 0; ri < modelsRowIdx + 5; ri++) {
      const row = raw[ri] || [];
      const label = String(row[2] || "").trim();
      if (["Bond (Adjusted)", "Stock (Adjusted)", "Over/Under"].includes(label)) {
        sleeveWeights[label] = [
          parseFloat(row[6]) || 0,
          parseFloat(row[7]) || 0,
          parseFloat(row[8]) || 0,
        ];
      }
    }

    // Model column headers: scan first 4 rows for the staggered model name labels
    // They appear in cols 5-11 in the top rows of the sheet
    const modelColDefs = [];
    const leftCIs = [];
    const labelFound = {}; // ci → label (take first non-null found scanning top rows)
    for (let ri = 0; ri < Math.min(raw.length, 5); ri++) {
      const row = raw[ri] || [];
      for (let ci = 5; ci < 12; ci++) {
        const v = row[ci];
        if (v && typeof v === "string" && !v.includes("Model Date") && !labelFound[ci]) {
          labelFound[ci] = v;
        }
      }
    }
    for (let ci = 5; ci < 12; ci++) {
      const label = labelFound[ci];
      if (label) {
        const master = p6MasterLabel(sheetName, label);
        if (master !== "__skip__") {
          modelColDefs.push({ ci, weightIdx: leftCIs.length, master });
          leftCIs.push(ci);
        }
      }
    }
    if (!modelColDefs.length) continue;

    const dataStartRow = headerRowIdx + 1; // row after the header

    // ── 3. Parse security sets ────────────────────────────────────────────────
    // Find import section (after data rows, look for "IMPORT FILE" or "Model Name")
    let importStart = -1;
    for (let ri = dataStartRow + 10; ri < raw.length; ri++) {
      const row = raw[ri];
      if (row && (String(row[1]||"").includes("IMPORT FILE") || String(row[1]||"") === "Model Name")) {
        importStart = ri; break;
      }
    }
    while (importStart >= 0 && importStart < raw.length) {
      const r = raw[importStart];
      if (r && (String(r[1]||"") === "Model Name" || String(r[1]||"").includes("IMPORT FILE"))) { importStart++; continue; }
      break;
    }

    const sets = {}; // setName → [{ticker, weight}]
    let currentSet = null;

    if (importStart >= 0) {
      // Named import section
      for (let ri = importStart; ri < raw.length; ri++) {
        const row = raw[ri];
        if (!row) continue;
        const colB = String(row[1] || "").trim();
        const colD = String(row[3] || "").trim().toUpperCase();
        const colE = parseFloat(row[4]);

        const isMetaRow = !colB || colB.length <= 1 || colB.startsWith("To Do") ||
          /^[123][).]/.test(colB) || /^\d+$/.test(colB) ||
          ["IMPORT FILE","Model Name","VLOOK Column"].includes(colB);

        if (!isMetaRow && colB) {
          currentSet = colB;
          if (!sets[currentSet]) sets[currentSet] = [];
        }
        if (currentSet && colD && !isNaN(colE) && colE > 0) {
          sets[currentSet].push({ ticker: colD, weight: colE });
        }
      }

      // ESG fallback: no set names in col B — tickers listed directly
      if (Object.keys(sets).length === 0) {
        const sleeveNames = ["Fixed Income", "Equity", "O/U"];
        let group = [], sleeveIdx = 0;
        for (let ri = importStart; ri < raw.length; ri++) {
          const row = raw[ri];
          if (!row) continue;
          const colD = String(row[3] || "").trim().toUpperCase();
          const colE = parseFloat(row[4]);
          if (colD && !isNaN(colE) && colE > 0) {
            group.push({ ticker: colD, weight: colE });
          } else if (group.length) {
            const total = group.reduce((s, h) => s + h.weight, 0);
            if (total >= 0.9) {
              sets[`${sheetName} ${sleeveNames[sleeveIdx] || "Set"}`] = group;
              group = []; sleeveIdx++;
            }
          }
        }
        if (group.length) sets[`${sheetName} ${sleeveNames[sleeveIdx] || "Set"}`] = group;
      }
    }

    // No import section — build sets from the main table rows 14-38
    // Col B = sleeve/set name, col D = ticker, col E = weight
    // Top 20 Sleeve is expanded using the sub-table (rows 41+) col E values
    if (Object.keys(sets).length === 0) {
      results.push({ modelKey: "debug__" + sheetName, modelLabel: "DEBUG " + sheetName + ": reached fallback, importStart=" + importStart, sheetName, colLabel: null, positions: [{ticker:"X",name:"debug",target:1,assetClass:"debug"}] });
      // Find and parse the Top 20 sub-table first
      let t20SubHoldings = []; // [{ticker, weight}] using col E (% of sleeve)
      let t20Start = -1, t20End = -1;
      for (let ri = 0; ri < raw.length; ri++) {
        const row = raw[ri];
        if (!row) continue;
        if (String(row[3]||"").trim() === "% of Top 20") t20Start = ri + 1;
        if (t20Start > 0 && ri > t20Start && (row[3] === 1 || String(row[3]||"").trim() === "1")) {
          t20End = ri; break;
        }
      }
      if (t20Start > 0) {
        const end = t20End > 0 ? t20End : raw.length;
        for (let ri = t20Start; ri < end; ri++) {
          const row = raw[ri];
          if (!row) continue;
          const ticker = String(row[2]||"").trim().toUpperCase();
          const weight = parseFloat(row[4]); // col E = % of sleeve
          if (ticker && !isNaN(weight) && weight > 0) t20SubHoldings.push({ ticker, weight });
        }
      }

      // Parse main table rows from dataStartRow to (t20Start-2) as security sets
      const SLEEVE_NAMES = ["bond","stock","tactical","over/under"];
      const endRow = t20Start > 0 ? t20Start - 2 : raw.length;
      let mainSet = null;
      for (let ri = dataStartRow; ri < endRow; ri++) {
        const row = raw[ri];
        if (!row) continue;
        const colB = String(row[1] || "").trim();
        const colD = String(row[3] || "").trim().toUpperCase();
        const colE = parseFloat(row[4]);

        if (colB && !colD && SLEEVE_NAMES.some(kw => colB.toLowerCase() === kw)) {
          mainSet = colB;
          if (!sets[mainSet]) sets[mainSet] = [];
          continue;
        }
        if (!colD) continue;

        if (colD === "TOP 20 SLEEVE") {
          // Inline-expand with sub-table holdings
          if (mainSet) {
            for (const h of t20SubHoldings) {
              sets[mainSet].push({ ticker: h.ticker, weight: h.weight });
            }
          }
          continue;
        }

        if (mainSet && !isNaN(colE) && colE > 0) {
          sets[mainSet].push({ ticker: colD, weight: colE });
        }
      }
    }

    // ── 4. Map sleeve labels to set names ──────────────────────────────────────
    function findSet(sleeveLabel) {
      const sl = sleeveLabel.toLowerCase();
      const isBond  = sl.includes("bond");
      const isStock = sl.includes("stock");
      const isOU    = sl.includes("over") || sl.includes("under");
      for (const name of Object.keys(sets)) {
        const nl = name.toLowerCase();
        if (isBond  && (nl.includes("fixed") || nl.includes("bond") || nl.includes("income"))) return name;
        if (isStock && (nl.includes("equity") || nl.includes("stock"))) return name;
        if (isOU    && (nl.includes("o/u") || nl.includes("over") || nl.includes("under"))) return name;
      }
      return null;
    }

    // ── 5. Blend sleeves × sets into model positions ──────────────────────────
    for (const { ci, weightIdx, master } of modelColDefs) {
      const blended = {};

      for (const [sleeveLabel, weights] of Object.entries(sleeveWeights)) {
        const w = weights[weightIdx] || 0;
        if (w === 0) continue;

        const isStockSleeve = sleeveLabel.toLowerCase().includes("stock");
        const setName = findSet(sleeveLabel);
        let holdings = setName ? sets[setName] : null;
        if (!holdings || !holdings.length) continue;

        // Normalise set to sum=1
        const total = holdings.reduce((s, h) => s + h.weight, 0);
        const factor = total > 0 ? 1 / total : 1;

        const sl = sleeveLabel.toLowerCase();
        const ac = sl.includes("bond") ? "Fixed Income" : "Equity";

        for (const h of holdings) {
          const key = normalizeTicker(h.ticker);
          const contrib = w * h.weight * factor * 100;
          if (blended[key]) blended[key].target += contrib;
          else blended[key] = { ticker: h.ticker, name: "", target: contrib, assetClass: ac };
        }
      }

      const positions = dedupePositions(Object.values(blended));
      const tot = positions.reduce((s,p)=>s+p.target,0).toFixed(1);
      if (!positions.length) continue;
      results.push({ modelKey: "p6__" + master, modelLabel: master, sheetName, colLabel: null, positions });
    }
  }
  return results;
}

// ── Parse target file: returns array of { modelKey, modelLabel, positions[] } ─
// Each position: { ticker, name, target (%), assetClass }
function parseTargetFile(wb, masterModels = []) {
  // Build set of master model names for existence filtering
  const masterNameSet = new Set(masterModels.map(m => m.name.toLowerCase()));
  // Fast-path: Paragon security-set format
  if (isParagonFormat(wb)) return parseParagonFormat(wb);
  // Fast-path: P6 sleeve model format
  if (isP6Format(wb)) return parseP6Format(wb);
  const results = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw.length) continue;

    // Find header row: look for "Ticker" or "Symbol" in col 0 or 1
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(raw.length, 8); i++) {
      const row = raw[i];
      const flat = row.map(c => String(c || "").trim().toLowerCase());
      if (flat.some(v => v === "ticker" || v === "symbol")) { headerRowIdx = i; break; }
    }

    // Detect ticker column
    let tickerColIdx = 0, nameColIdx = 1;
    if (headerRowIdx >= 0) {
      raw[headerRowIdx].forEach((h, idx) => {
        const v = String(h || "").trim().toLowerCase();
        if (v === "ticker" || v === "symbol") tickerColIdx = idx;
        if (v.includes("name") || v.includes("fund")) nameColIdx = idx;
      });
    }

    // Detect model columns: all numeric-header columns after the name col
    const modelCols = []; // { colIdx, label }
    if (headerRowIdx >= 0) {
      const headerRow = raw[headerRowIdx];
      for (let ci = nameColIdx + 1; ci < headerRow.length; ci++) {
        const h = headerRow[ci];
        if (h !== null && h !== undefined && String(h).trim() !== "") {
          modelCols.push({ colIdx: ci, label: String(h).trim() });
        }
      }
    }

    const dataStartRow = headerRowIdx >= 0 ? headerRowIdx + 1 : 0;

    // Primary rule: col B populated = holding row, col B empty = section header row
    // This is reliable across all sheet formats in this file.
    function isHoldingRow(row) {
      const nameVal = nameColIdx >= 0 ? row[nameColIdx] : null;
      return nameVal !== null && nameVal !== undefined && String(nameVal).trim() !== "";
    }

    // Helper: collect only holding rows for scale detection
    function collectHoldingRows(colIdx) {
      const nums = [];
      for (let ri = dataStartRow; ri < raw.length; ri++) {
        const row = raw[ri];
        if (!row) continue;
        if (!isHoldingRow(row)) continue;
        const num = parseFloat(row[colIdx]);
        if (!isNaN(num) && num > 0) nums.push(num);
      }
      return nums;
    }

    if (modelCols.length > 1) {
      // Multi-model sheet — detect scale per column from holding rows only
      const scaleByCol = {};
      for (const mc of modelCols) {
        scaleByCol[mc.colIdx] = detectScale(collectHoldingRows(mc.colIdx));
      }

      const posMap = {};
      for (const mc of modelCols) posMap[mc.colIdx] = [];

      let currentAC = "Equity";
      let currentSubSection = "";
      // Sub-section keywords for STP variant derivation
      const SUB_SECTION_HEADERS = ["us large cap","us small cap","intl developed","emerging markets",
        "u.s. investment grade","international bonds","high yield corporate","high yield municipal",
        "u.s. investment grade municipal","alternative","commodities","cash equivalents","fixed income"];

      for (let ri = dataStartRow; ri < raw.length; ri++) {
        const row = raw[ri];
        if (!row || row.every(c => c === null || c === undefined || String(c).trim() === "")) continue;
        const cellA = row[tickerColIdx];
        const cellB = nameColIdx >= 0 ? row[nameColIdx] : null;
        const colAStr = String(cellA || "").trim();

        if (!isHoldingRow(row)) {
          // Section header — update asset class and sub-section context
          if (colAStr) {
            currentAC = toTitleCase(colAStr);
            const lower = colAStr.toLowerCase();
            if (SUB_SECTION_HEADERS.some(k => lower.includes(k) || k.includes(lower))) {
              currentSubSection = lower;
            }
          }
          continue;
        }

        const ticker = colAStr.toUpperCase();
        if (!ticker) continue;
        const secName = String(cellB || "").trim();

        for (const mc of modelCols) {
          const rawVal = row[mc.colIdx];
          if (rawVal === null || rawVal === undefined) continue;
          const num = parseFloat(rawVal);
          if (isNaN(num) || num === 0) continue;
          const pct = applyScale(num, scaleByCol[mc.colIdx]);
          const assetClass = isCashTicker(ticker) ? "Cash" : currentAC;
          posMap[mc.colIdx].push({ ticker, name: secName, target: pct, assetClass, subSection: currentSubSection });
        }
      }

      for (const mc of modelCols) {
        const positions = dedupePositions(posMap[mc.colIdx]);
        if (!positions.length) continue;
        results.push({
          modelKey: `${sheetName}__${mc.label}`,
          modelLabel: `${sheetName} — ${mc.label}`,
          sheetName,
          colLabel: mc.label,
          positions,
        });
      }
    } else {
      // Single-model sheet
      let allocColIdx = 2;
      if (headerRowIdx >= 0) {
        raw[headerRowIdx].forEach((h, idx) => {
          if (idx > nameColIdx) {
            const v = String(h || "").toLowerCase().trim();
            if (v.includes("alloc") || v.includes("target") || v.includes("weight") || v.includes("%") || v === "round") allocColIdx = idx;
          }
        });
      }

      // Detect scale from holding rows only
      const scale = detectScale(collectHoldingRows(allocColIdx));

      const positions = [];
      let currentAC = "Equity";
      for (let ri = dataStartRow; ri < raw.length; ri++) {
        const row = raw[ri];
        if (!row) continue;
        const cellA = row[tickerColIdx];
        const cellB = nameColIdx >= 0 ? row[nameColIdx] : null;
        const colAStr = String(cellA || "").trim();
        if (!colAStr) continue;

        if (!isHoldingRow(row)) {
          if (colAStr) currentAC = toTitleCase(colAStr);
          continue;
        }
        const ticker = colAStr.toUpperCase();
        if (!ticker) continue;
        const rawVal = row[allocColIdx];
        if (rawVal === null || rawVal === undefined) continue;
        const num = parseFloat(rawVal);
        if (isNaN(num) || num === 0) continue;
        const pct = applyScale(num, scale);
        const assetClass = isCashTicker(ticker) ? "Cash" : currentAC;
        positions.push({ ticker, name: String(cellB || "").trim(), target: pct, assetClass });
      }

      if (positions.length) {
        results.push({
          modelKey: sheetName,
          modelLabel: sheetName,
          sheetName,
          colLabel: null,
          positions: dedupePositions(positions),
        });
      }
    }
  }

  // ── Derive all STP variants from base models ───────────────────────────────
  const hasSTP = results.some(r => r.modelLabel.includes("Savvy Total Portfolios"));
  if (hasSTP) {

    // Parse stock model tabs → holdings map
    const stockModelTabs = {
      "Core Stock Model":     "Savvy Core S&P 500 Stock Model",
      "Growth Stock Model":   "Savvy LC Growth Stock Model",
      "Dividend Stock Model": "Savvy Dividend Stock Model",
      "Value Stock Model":    "Savvy Value Stock Model",
    };
    const stockModelHoldings = {}; // variantName → [{ticker, name, weight (normalised to sum=1)}]
    for (const [variantName, tabName] of Object.entries(stockModelTabs)) {
      const ws2 = wb.Sheets[tabName];
      if (!ws2) continue;
      const raw2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null });
      const holdings = [];
      for (let ri = 1; ri < raw2.length; ri++) {
        const row = raw2[ri];
        if (!row) continue;
        // Format A: ticker(0) name(1) alloc(2)  OR  ticker in col 0, weight in col 2
        const ticker = String(row[0] || "").trim().toUpperCase();
        if (!ticker || ticker.length > 6) continue;
        const weight = parseFloat(row[2]);
        if (isNaN(weight) || weight <= 0) continue;
        holdings.push({ ticker, name: String(row[1] || "").trim(), weight });
      }
      // Normalise to sum=1
      const total = holdings.reduce((s, h) => s + h.weight, 0);
      if (total > 0) for (const h of holdings) h.weight /= total;
      stockModelHoldings[variantName] = holdings;
    }

    // Helper: get positions from a named result
    function getBase(modelLabel) {
      return results.find(r => r.modelLabel === modelLabel);
    }

    // Helper: build a derived model from a base + transformation
    function makeVariant(label, baseLabel, transform) {
      const base = getBase(baseLabel);
      if (!base) return null;
      const positions = transform(base.positions);
      if (!positions || !positions.length) return null;
      return { modelKey: `derived__${label}`, modelLabel: label, sheetName: "Derived", colLabel: null, positions };
    }

    // Transformation helpers
    function reweight(positions) {
      const total = positions.reduce((s, p) => s + p.target, 0);
      if (total <= 0) return positions;
      return positions.map(p => ({ ...p, target: (p.target / total) * 100 }));
    }

    function excludeSubSection(positions, ...sections) {
      return reweight(positions.filter(p => !sections.some(s => (p.subSection || "").includes(s))));
    }

    function keepSubSections(positions, ...sections) {
      return reweight(positions.filter(p => sections.some(s => (p.subSection || "").includes(s))));
    }

    function substituteUSLC(positions, stockVariant) {
      const stockHoldings = stockModelHoldings[stockVariant];
      if (!stockHoldings) return positions;
      // Sum weight of US Large Cap positions
      const ulcWeight = positions
        .filter(p => (p.subSection || "").includes("us large cap"))
        .reduce((s, p) => s + p.target, 0);
      // Remove US Large Cap ETFs, add stock model holdings scaled to ulcWeight
      const without = positions.filter(p => !(p.subSection || "").includes("us large cap"));
      const stocks = stockHoldings.map(h => ({
        ticker: h.ticker, name: h.name,
        target: h.weight * ulcWeight,
        assetClass: "Equity", subSection: "us large cap"
      }));
      return dedupePositions([...without, ...stocks]);
    }

    // Base risk level labels (regular sheet)
    const BASE_LABELS = [
      "All Fixed", "Conservative", "Moderately Conservative",
      "Moderate", "Moderately Aggressive", "Aggressive", "All Equity"
    ];
    const STP_SHEET    = "Savvy Total Portfolios";
    const STP_TAX_SHEET = "Savvy Total Portfolios (Tax Awa";

    const derived = [];

    for (const risk of BASE_LABELS) {
      const regKey  = `${STP_SHEET} — ${risk}`;
      const taxKey  = `${STP_TAX_SHEET} — ${risk}`;
      const regName = `STP - ${risk}`;  // base — already in results
      const taxName = `STP - ${risk} (Tax Aware)`;

      // ex-USLC regular
      if (!["All Fixed"].includes(risk)) {
        derived.push(makeVariant(`STP - ${risk} (ex-USLC)`, regKey,
          p => excludeSubSection(p, "us large cap")));
        // ex-USLC Tax Aware
        derived.push(makeVariant(`STP - ${risk} (ex-USLC) - Tax Aware`, taxKey,
          p => excludeSubSection(p, "us large cap")));
      }

      // US Equity Only — small cap stays as-is, international/emerging excluded,
      // their weight is redistributed pro-rata to US Large Cap holdings only
      if (!["All Fixed"].includes(risk)) {
        const usEquityOnly = p => {
          const isUSLargeCap = pos => (pos.subSection || "").includes("us large cap");
          const isUSSmallCap = pos => (pos.subSection || "").includes("us small cap");
          const isIntlOrEM   = pos => {
            const ss = pos.subSection || "";
            return ss.includes("intl developed") || ss.includes("emerging");
          };

          // Weight being removed (intl + emerging)
          const removedWeight = p.filter(isIntlOrEM).reduce((s, pos) => s + pos.target, 0);
          // Current US Large Cap total weight
          const ulcWeight = p.filter(isUSLargeCap).reduce((s, pos) => s + pos.target, 0);

          const result = p
            .filter(pos => !isIntlOrEM(pos))
            .map(pos => {
              if (isUSLargeCap(pos) && ulcWeight > 0) {
                // Scale each US Large Cap holding up by (ulcWeight + removedWeight) / ulcWeight
                return { ...pos, target: pos.target * (ulcWeight + removedWeight) / ulcWeight };
              }
              return pos; // small cap, FI, alts, cash — unchanged
            });
          return dedupePositions(result);
        };
        derived.push(makeVariant(`STP - ${risk} - US Equity Only`, regKey, usEquityOnly));
        derived.push(makeVariant(`STP - ${risk} - US Equity Only (Tax Aware)`, taxKey, usEquityOnly));
      }

      // Stock model substitutions (replace US Large Cap with individual stocks)
      for (const [variantName] of Object.entries(stockModelTabs)) {
        const suffix = variantName; // e.g. "Core Stock Model"
        derived.push(makeVariant(`STP - ${risk} - ${suffix}`, regKey,
          p => substituteUSLC(p, variantName)));
        derived.push(makeVariant(`STP - ${risk} (Tax Aware) - ${suffix}`, taxKey,
          p => substituteUSLC(p, variantName)));
      }
    }

    // Filter out nulls, avoid duplicates, and ONLY keep models that exist in master
    // This prevents generating variants that have no master counterpart
    const existingKeys = new Set(results.map(r => r.modelLabel));
    for (const d of derived) {
      if (d && !existingKeys.has(d.modelLabel) &&
          (masterNameSet.size === 0 || masterNameSet.has(d.modelLabel.toLowerCase()))) {
        results.push(d);
        existingKeys.add(d.modelLabel);
      }
    }

    // ── Midpoint models ────────────────────────────────────────────────────────
    function avgPositions(leftPos, rightPos) {
      const rightIdx = {};
      for (const p of rightPos) rightIdx[normalizeTicker(p.ticker)] = p;
      const blended = {};
      for (const p of leftPos) {
        const key = normalizeTicker(p.ticker);
        const rt = rightIdx[key] ? rightIdx[key].target : 0;
        blended[key] = { ...p, target: (p.target + rt) / 2 };
      }
      for (const p of rightPos) {
        const key = normalizeTicker(p.ticker);
        if (!blended[key]) blended[key] = { ...p, target: p.target / 2 };
      }
      return dedupePositions(Object.values(blended).filter(p => p.target > 0));
    }

    function makeMidpoint(label, leftKey, rightKey) {
      const leftModel  = results.find(r => r.modelLabel === leftKey);
      const rightModel = results.find(r => r.modelLabel === rightKey);
      if (!leftModel || !rightModel) return null;
      const positions = avgPositions(leftModel.positions, rightModel.positions);
      return positions.length ? { modelKey: `midpoint__${label}`, modelLabel: label, sheetName: "Midpoint", colLabel: null, positions } : null;
    }

    const STP_MIDPOINTS = [
      { label: "STP - 30/70",              left: `${STP_SHEET} — Conservative`,            right: `${STP_SHEET} — Moderately Conservative` },
      { label: "STP - 50/50",              left: `${STP_SHEET} — Moderately Conservative`, right: `${STP_SHEET} — Moderate` },
      { label: "STP - 70/30",              left: `${STP_SHEET} — Moderate`,                right: `${STP_SHEET} — Moderately Aggressive` },
      { label: "STP - 10/90 (Tax Aware)",  left: `${STP_TAX_SHEET} — All Fixed`,               right: `${STP_TAX_SHEET} — Conservative` },
      { label: "STP - 50/50 (Tax Aware)",  left: `${STP_TAX_SHEET} — Moderately Conservative`, right: `${STP_TAX_SHEET} — Moderate` },
      { label: "STP - 70/30 (Tax Aware)",  left: `${STP_TAX_SHEET} — Moderate`,                right: `${STP_TAX_SHEET} — Moderately Aggressive` },
    ];

    for (const mp of STP_MIDPOINTS) {
      const generated = makeMidpoint(mp.label, mp.left, mp.right);
      if (generated && !existingKeys.has(generated.modelLabel) &&
          (masterNameSet.size === 0 || masterNameSet.has(generated.modelLabel.toLowerCase()))) {
        results.push(generated);
        existingKeys.add(generated.modelLabel);
      }
    }
  }

  return results;
}
function buildInitialMappings(targetModels, masterModels) {
  const filtered = targetModels.filter(tm => normalizeForMatch(tm.modelLabel) !== "__skip__");
  if (!filtered.length) return [];

  // Build exact-match index: normalized master name → master model
  const exactIndex = {};
  for (const m of masterModels) {
    exactIndex[m.name.toLowerCase()] = m;
  }

  // If all target labels are already master-ready P6 names, restrict candidates to P6 models only
  const allP6 = filtered.every(tm => tm.modelLabel.toLowerCase().includes("p6") ||
    tm.modelKey.startsWith("p6__"));
  const candidates = allP6
    ? masterModels.filter(m => m.name.includes("P6"))
    : masterModels;

  return filtered.map(tm => {
    const normalizedLabel = normalizeForMatch(tm.modelLabel);

    // 1. Try exact match first (case-insensitive)
    const exactMatch = exactIndex[normalizedLabel] || exactIndex[tm.modelLabel.toLowerCase()];
    if (exactMatch) {
      // If exact match is inactive, check if an active .e variant exists
      const dotE = exactIndex[(exactMatch.name + ".e").toLowerCase()];
      const preferred = (exactMatch.active === false && dotE && dotE.active !== false) ? dotE : exactMatch;
      return {
        targetKey: tm.modelKey,
        targetLabel: tm.modelLabel,
        masterModelId: preferred.id,
        masterModelName: preferred.name,
        confidence: 1.0,
      };
    }

    // 2. Fall back to fuzzy match
    const match = fuzzyMatch(tm.modelLabel, candidates);
    return {
      targetKey: tm.modelKey,
      targetLabel: tm.modelLabel,
      masterModelId: match ? match.match.id : null,
      masterModelName: match ? match.match.name : null,
      confidence: match ? match.score : 0,
    };
  });
}

// ── Audit computation ─────────────────────────────────────────────────────────
function runAudit(mapping, targetModel, masterHoldings, excludedClasses, threshold) {
  const masterModelId = mapping.masterModelId;
  const masterForModel = masterHoldings.filter(h => h.modelId === masterModelId);

  // Index master by normalized ticker
  const mIdx = {};
  for (const h of masterForModel) mIdx[normalizeTicker(h.ticker)] = h;

  const included = targetModel.positions.filter(p => !excludedClasses.includes(p.assetClass));
  const excluded = targetModel.positions.filter(p =>  excludedClasses.includes(p.assetClass));

  const sumIncluded = included.reduce((s, p) => s + (p.target > 0 ? p.target : 0), 0);

  // Index target by normalized ticker
  const adjMap = {};
  for (const p of included) {
    const adj = sumIncluded > 0 && p.target > 0 ? (p.target / sumIncluded) * 100 : 0;
    adjMap[normalizeTicker(p.ticker)] = { ...p, adjTarget: adj };
  }

  // Union of all normalized tickers from both sides
  const allNormTickers = new Set([
    ...masterForModel.map(h => normalizeTicker(h.ticker)),
    ...included.map(p => normalizeTicker(p.ticker)),
  ]);

  const rows = [];
  for (const normTicker of allNormTickers) {
    const mRow = mIdx[normTicker];
    const tRow = adjMap[normTicker];
    const masterTarget = mRow ? mRow.target : null;
    const adjTarget    = tRow ? tRow.adjTarget : null;
    const rawTarget    = tRow ? tRow.target    : null;
    // Prefer master ticker for display (canonical), fall back to target ticker
    const displayTicker = mRow?.ticker || tRow?.ticker || normTicker;
    const name          = mRow?.name || tRow?.name || displayTicker;
    const assetClass    = tRow?.assetClass || "—";
    let status, diff = null;
    if (masterTarget === null)    status = "MISSING_MASTER";
    else if (adjTarget === null)  status = "MISSING_TARGET";
    else {
      diff = masterTarget - adjTarget;
      status = Math.abs(diff) <= threshold ? "MATCH" : diff > 0 ? "OVER" : "UNDER";
    }
    rows.push({ ticker: displayTicker, name, assetClass, masterTarget, adjTarget, rawTarget, diff, status });
  }
  rows.sort((a, b) => {
    const order = ["OVER","UNDER","MISSING_MASTER","MISSING_TARGET","MATCH"];
    const d = order.indexOf(a.status) - order.indexOf(b.status);
    return d !== 0 ? d : Math.abs(b.diff||0) - Math.abs(a.diff||0);
  });
  const summary = {
    total: rows.length,
    over:    rows.filter(r => r.status==="OVER").length,
    under:   rows.filter(r => r.status==="UNDER").length,
    match:   rows.filter(r => r.status==="MATCH").length,
    missing: rows.filter(r => r.status==="MISSING_MASTER"||r.status==="MISSING_TARGET").length,
  };
  return { rows, summary, excluded, sumIncluded };
}

// ── Drag-drop hook ────────────────────────────────────────────────────────────
function useDrop(onFiles) {
  const [dragging, setDragging] = useState(false);
  const counter = useRef(0);
  return {
    dragging,
    onDragEnter: e => { e.preventDefault(); counter.current++; setDragging(true); },
    onDragLeave: e => { e.preventDefault(); if (--counter.current===0) setDragging(false); },
    onDragOver:  e => e.preventDefault(),
    onDrop: e => {
      e.preventDefault(); counter.current=0; setDragging(false);
      const files = Array.from(e.dataTransfer.files).filter(f => /\.(xlsx|xls)$/i.test(f.name));
      if (files.length) onFiles(files);
    },
  };
}

const STORAGE_KEY = "master_file_data";

// ── Main component ────────────────────────────────────────────────────────────
export default function ModelAuditTool() {
  const [masterData, setMasterData]     = useState(null); // { models, holdings, fileName, savedAt }
  const [masterFromCache, setMasterFromCache] = useState(false);
  const [targetModels, setTargetModels] = useState([]); // parsed target models
  const [debugInfo, setDebugInfo] = useState(null);
  const [targetFileName, setTargetFileName] = useState(null);
  const [mappings, setMappings]         = useState([]); // { targetKey, masterModelId, ... }
  const [loadingM, setLoadingM]         = useState(false);
  const [loadingT, setLoadingT]         = useState(false);
  const [storageLoading, setStorageLoading] = useState(true);
  const [errorM, setErrorM]             = useState(null);
  const [errorT, setErrorT]             = useState(null);
  const [threshold, setThreshold]       = useState(DEFAULT_THRESHOLD);
  const [excludedClasses, setExcludedClasses] = useState([]);
  const [expanded, setExpanded]         = useState({});
  const [filters, setFilters]           = useState({});
  const [mappingOpen, setMappingOpen]   = useState(false);

  // ── Load master from storage on mount ──
  useEffect(() => {
    (async () => {
      try {
        const result = await window.storage.get(STORAGE_KEY);
        if (result?.value) {
          const data = JSON.parse(result.value);
          // Expand slim holdings if stored in new format
          const modelsMap = Object.fromEntries((data.models || []).map(m => [m.id, m.name]));
          const holdings = data.holdings?.[0]?.t !== undefined
            ? expandHoldings(data.holdings, modelsMap)
            : data.holdings;
          setMasterData({ ...data, holdings });
          setMasterFromCache(true);
        }
      } catch { /* no cached data */ }
      setStorageLoading(false);
    })();
  }, []);

  // ── Save master to storage ──
  const saveMasterToStorage = useCallback(async (data) => {
    try {
      await window.storage.set(STORAGE_KEY, JSON.stringify(data));
    } catch(e) { console.warn("Storage save failed:", e); }
  }, []);

  const clearMasterStorage = useCallback(async () => {
    try { await window.storage.delete(STORAGE_KEY); } catch {}
  }, []);

  // ── Load master from file ──
  const handleMasterFiles = useCallback(async (files) => {
    setLoadingM(true); setErrorM(null);
    try {
      const buf = await readFileBuffer(files[0]);
      const wb = readXLSX(buf);
      const { models, holdings: slimHoldings } = parseMasterFile(wb);
      // Expand slim holdings for in-memory use
      const modelsMap = Object.fromEntries(models.map(m => [m.id, m.name]));
      const holdings = expandHoldings(slimHoldings, modelsMap);
      const data = { models, holdings, fileName: files[0].name, savedAt: new Date().toLocaleDateString() };
      setMasterData(data);
      setMasterFromCache(false);
      // Save slim format to storage (half the size)
      await saveMasterToStorage({ models, holdings: slimHoldings, fileName: files[0].name, savedAt: data.savedAt });
      if (targetModels.length) setMappings(buildInitialMappings(targetModels, models));
    } catch(e) { setErrorM("Could not parse Master File: " + e.message); }
    setLoadingM(false);
  }, [targetModels, saveMasterToStorage]);

  // ── Load target ──
  const handleTargetFiles = useCallback(async (files) => {
    setLoadingT(true); setErrorT(null);
    // Clear previous target state before loading new file
    setTargetModels([]); setTargetFileName(null); setMappings([]);
    setExpanded({}); setFilters({}); setExcludedClasses([]);
    try {
      const buf = await readFileBuffer(files[0]);
      const wb = readXLSX(buf);
      const parsed = parseTargetFile(wb, masterData?.models || []);
      setTargetModels(parsed);
      setTargetFileName(files[0].name);
      const mappings = masterData ? buildInitialMappings(parsed, masterData.models) : [];
      if (masterData) setMappings(mappings);
      setDebugInfo({
        format: isParagonFormat(wb) ? "Paragon" : isP6Format(wb) ? "P6" : "Standard",
        parsedModels: parsed.map(m => ({ label: m.modelLabel, positions: m.positions.length, total: m.positions.reduce((s,p)=>s+p.target,0).toFixed(1) })),
        mappings: mappings.map(m => ({ from: m.targetLabel, to: m.masterModelName || "UNMATCHED" })),
      });
    } catch(e) { setErrorT("Could not parse Target File: " + e.message); }
    setLoadingT(false);
  }, [masterData]);

  const updateMapping = (targetKey, masterModelId) => {
    const model = masterData?.models.find(m => m.id === parseInt(masterModelId));
    setMappings(prev => prev.map(m => m.targetKey === targetKey
      ? { ...m, masterModelId: model?.id || null, masterModelName: model?.name || null }
      : m));
  };

  const allAssetClasses = useMemo(() => {
    const s = new Set(targetModels.flatMap(tm => tm.positions.map(p => p.assetClass)).filter(c => c && c !== "—"));
    return [...s].sort();
  }, [targetModels]);

  const toggleExclude = cls => setExcludedClasses(prev => prev.includes(cls) ? prev.filter(c=>c!==cls) : [...prev, cls]);
  const toggleExpand = key => setExpanded(prev => ({ ...prev, [key]: !prev[key] }));
  const allExpanded = mappings.length > 0 && mappings.every(m => expanded[m.targetKey] === true);
  const collapseAll = () => { const e = {}; mappings.forEach(m => { e[m.targetKey] = false; }); setExpanded(e); };
  const expandAll   = () => { const e = {}; mappings.forEach(m => { e[m.targetKey] = true;  }); setExpanded(e); };
  const [sortBy, setSortBy] = useState("default"); // "default" | "issues" | "clean" | "missing"
  const getFilter = key => filters[key] || "ALL";
  const setFilter = (key, val) => setFilters(prev => ({ ...prev, [key]: val }));

  const masterDrop = useDrop(handleMasterFiles);
  const targetDrop = useDrop(handleTargetFiles);

  const ready = masterData && targetModels.length && mappings.length;

  // Sorted mappings — computed outside JSX to avoid IIFE complexity
  const sortedMappings = useMemo(() => {
    if (sortBy === "default" || !masterData || !mappings.length) return mappings;
    const scoreCache = new Map();
    const getScore = (mapping) => {
      if (scoreCache.has(mapping.targetKey)) return scoreCache.get(mapping.targetKey);
      const s = [0, 0, 0];
      try {
        if (mapping.masterModelId) {
          const tm = targetModels.find(t => t.modelKey === mapping.targetKey);
          if (tm && tm.positions && tm.positions.length) {
            const { rows } = runAudit(mapping, tm, masterData.holdings, excludedClasses, threshold);
            s[0] = rows.filter(r => r.status==="OVER"||r.status==="UNDER").length;
            s[1] = rows.filter(r => r.status==="MISSING_MASTER"||r.status==="MISSING_TARGET").length;
            s[2] = rows.filter(r => r.status==="MATCH").length;
          }
        }
      } catch { /* ignore scoring errors */ }
      scoreCache.set(mapping.targetKey, s);
      return s;
    };
    return [...mappings].sort((a, b) => {
      const [ai, am, ac] = getScore(a);
      const [bi, bm, bc] = getScore(b);
      if (sortBy === "issues")  return (bi + bm) - (ai + am);
      if (sortBy === "missing") return bm - am;
      if (sortBy === "clean")   return bc - ac;
      return 0;
    });
  }, [mappings, sortBy, masterData, targetModels, excludedClasses, threshold]);

  return (
    <div style={{ fontFamily: "'Inter', system-ui, sans-serif", background: "#f1f5f9", minHeight: "100vh" }}>

      {/* Header */}
      <div style={{ background: "#0f172a", padding: "16px 28px", display: "flex", alignItems: "center", gap: 14 }}>
        <div style={{ width: 34, height: 34, background: "#3b82f6", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2"><path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 1-2-2V9m0 0h18"/></svg>
        </div>
        <div>
          <div style={{ color: "white", fontWeight: 700, fontSize: 16 }}>Model Audit Tool</div>
          <div style={{ color: "#94a3b8", fontSize: 11 }}>Orion Custom Indexing · Savvy</div>
        </div>
        {ready && (
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
            <div style={{ display: "flex", gap: 4, background: "#1e293b", borderRadius: 8, padding: 3 }}>
              {[
                { key: "default", label: "Default" },
                { key: "issues",  label: "Issues First" },
                { key: "clean",   label: "On Target First" },
                { key: "missing", label: "Missing First" },
              ].map(s => (
                <button key={s.key} onClick={() => setSortBy(s.key)} style={{ padding: "4px 10px", borderRadius: 6, border: "none", fontSize: 11, fontWeight: 600, cursor: "pointer", background: sortBy === s.key ? "#3b82f6" : "transparent", color: sortBy === s.key ? "white" : "#94a3b8", transition: "all 0.15s" }}>
                  {s.label}
                </button>
              ))}
            </div>
            <button onClick={allExpanded ? collapseAll : expandAll} style={{ padding: "6px 14px", borderRadius: 8, background: "#334155", border: "none", color: "white", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              {allExpanded ? "⊟ Collapse All" : "⊞ Expand All"}
            </button>
            <button onClick={() => setMappingOpen(true)} style={{ padding: "6px 14px", borderRadius: 8, background: "#1e40af", border: "none", color: "white", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              ⚙ Model Mappings {mappings.filter(m=>!m.masterModelId).length > 0 && `(${mappings.filter(m=>!m.masterModelId).length} unmatched)`}
            </button>
          </div>
        )}
      </div>

      <div style={{ padding: "20px 28px", maxWidth: 1400, margin: "0 auto" }}>

        {/* Upload + Settings row */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 260px", gap: 14, marginBottom: 16 }}>

          <DropZone label="Master File" sublabel="Orion export — source of truth for model names & holdings"
            accent="#3b82f6" loading={loadingM || storageLoading} error={errorM} dragging={masterDrop.dragging}
            dropProps={masterDrop} onFiles={handleMasterFiles} hint="Drop or click · .xlsx">
            {masterData && <>
              <FileChip
                name={masterData.fileName}
                sub={`${masterData.models.length} models · ${masterData.holdings.length} rows · saved ${masterData.savedAt}`}
                color="#3b82f6"
                badge={masterFromCache ? { label: "Cached", color: "#16a34a", bg: "#f0fdf4" } : { label: "Updated", color: "#2563eb", bg: "#eff6ff" }}
                onRemove={async () => { setMasterData(null); setMappings([]); setMasterFromCache(false); await clearMasterStorage(); }}
              />
              {masterFromCache && mappings.length > 0 && mappings.filter(m => !m.masterModelId).length > 0 && (
                <div style={{ marginTop: 6, padding: "6px 10px", background: "#fff7ed", borderRadius: 7, fontSize: 11, color: "#92400e" }}>
                  ⚠ {mappings.filter(m => !m.masterModelId).length} model(s) unmatched — master file may be stale. Re-upload to refresh.
                </div>
              )}
            </>}
          </DropZone>

          <DropZone label="Target File" sublabel="Your model library — drop a new file to replace"
            accent="#8b5cf6" loading={loadingT} error={errorT} dragging={targetDrop.dragging}
            dropProps={targetDrop} onFiles={handleTargetFiles} hint="Drop or click · .xlsx">
            {targetFileName && <FileChip name={targetFileName} sub={`${targetModels.length} models across ${[...new Set(targetModels.map(m=>m.sheetName))].length} tabs`} color="#8b5cf6" onRemove={() => { setTargetModels([]); setTargetFileName(null); setMappings([]); setFilters({}); setExcludedClasses([]); }} />}
          </DropZone>

          {/* Settings */}
          <div style={{ background: "white", borderRadius: 12, padding: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.06)", borderTop: "3px solid #0ea5e9" }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#0f172a", marginBottom: 12 }}>Audit Settings</div>
            <div style={{ marginBottom: 14 }}>
              <label style={lbl}>Tolerance (±%)</label>
              <input type="number" step="0.01" min="0" value={threshold}
                onChange={e => setThreshold(parseFloat(e.target.value)||0)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1.5px solid #e2e8f0", fontSize: 13, color: "#0f172a", boxSizing: "border-box" }} />
            </div>
            {allAssetClasses.length > 0 && (
              <div>
                <label style={lbl}>Exclude Asset Classes</label>
                <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                  {allAssetClasses.map(cls => (
                    <label key={cls} style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, color: excludedClasses.includes(cls) ? "#dc2626" : "#334155" }}>
                      <input type="checkbox" checked={excludedClasses.includes(cls)} onChange={() => toggleExclude(cls)} style={{ width: 14, height: 14, accentColor: "#dc2626" }} />
                      {cls}
                    </label>
                  ))}
                </div>
                {excludedClasses.length > 0 && <div style={{ marginTop: 8, fontSize: 11, color: "#94a3b8" }}>Non-zero remaining targets re-weighted to 100%.</div>}
              </div>
            )}
          </div>
        </div>

        {/* Mapping modal */}
        {mappingOpen && masterData && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}
            onClick={e => { if (e.target === e.currentTarget) setMappingOpen(false); }}>
            <div style={{ background: "white", borderRadius: 16, padding: 28, maxWidth: 720, width: "90%", maxHeight: "80vh", overflowY: "auto" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 16, color: "#0f172a" }}>Model Mappings</div>
                  <div style={{ fontSize: 12, color: "#94a3b8" }}>Match each target model column to its Orion master model</div>
                </div>
                <button onClick={() => setMappingOpen(false)} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#64748b" }}>×</button>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {mappings.map(m => (
                  <div key={m.targetKey} style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr", gap: 10, alignItems: "center", padding: "10px 14px", background: m.masterModelId ? "#f8fafc" : "#fff7ed", borderRadius: 10, border: `1px solid ${m.masterModelId ? "#e2e8f0" : "#fed7aa"}` }}>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#0f172a" }}>{m.targetLabel}</div>
                      {m.confidence > 0 && <div style={{ fontSize: 10, color: "#94a3b8" }}>Confidence: {(m.confidence*100).toFixed(0)}%</div>}
                    </div>
                    <div style={{ fontSize: 16, color: "#94a3b8" }}>→</div>
                    <select value={m.masterModelId || ""} onChange={e => updateMapping(m.targetKey, e.target.value)}
                      style={{ padding: "6px 8px", borderRadius: 7, border: `1.5px solid ${m.masterModelId ? "#e2e8f0" : "#f97316"}`, fontSize: 12, color: "#0f172a", background: "white" }}>
                      <option value="">— Not mapped —</option>
                      <optgroup label="Active models">
                        {masterData.models.filter(mod => mod.active).map(mod => (
                          <option key={mod.id} value={mod.id}>{mod.name}</option>
                        ))}
                      </optgroup>
                      <optgroup label="Inactive (no accounts)">
                        {masterData.models.filter(mod => !mod.active).map(mod => (
                          <option key={mod.id} value={mod.id}>{mod.name}</option>
                        ))}
                      </optgroup>
                    </select>
                  </div>
                ))}
              </div>
              <div style={{ marginTop: 20, display: "flex", justifyContent: "flex-end" }}>
                <button onClick={() => setMappingOpen(false)} style={{ padding: "8px 20px", borderRadius: 8, background: "#3b82f6", border: "none", color: "white", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
                  Save & Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Debug panel */}
        {debugInfo && (
          <details style={{ marginBottom: 12, background: "#0f172a", borderRadius: 10, padding: "10px 16px", fontSize: 11, color: "#94a3b8", fontFamily: "monospace" }}>
            <summary style={{ cursor: "pointer", color: "#60a5fa", fontWeight: 600, marginBottom: 6 }}>
              🔍 Parse Debug — {debugInfo.format} format · {debugInfo.parsedModels.length} models detected
            </summary>
            <div style={{ marginTop: 8 }}>
              <div style={{ color: "#e2e8f0", marginBottom: 6, fontWeight: 600 }}>Parsed models:</div>
              {debugInfo.parsedModels.map((m, i) => (
                <div key={i} style={{ color: m.positions === 0 ? "#f87171" : "#86efac" }}>
                  {m.positions === 0 ? "✗" : "✓"} {m.label} → {m.positions} positions ({m.total}%)
                </div>
              ))}
              <div style={{ color: "#e2e8f0", marginTop: 8, marginBottom: 6, fontWeight: 600 }}>Mappings:</div>
              {debugInfo.mappings.map((m, i) => (
                <div key={i} style={{ color: m.to === "UNMATCHED" ? "#f87171" : "#86efac" }}>
                  {m.to === "UNMATCHED" ? "✗" : "✓"} {m.from} → {m.to}
                </div>
              ))}
            </div>
          </details>
        )}

        {/* Empty state */}
        {!masterData && !targetModels.length && !storageLoading && <EmptyState icon text="Upload a Master File and Target File to begin the audit." />}
        {!masterData && !targetModels.length && storageLoading && <EmptyState icon text="Loading saved master file…" />}
        {masterData && !targetModels.length && <EmptyState text={`Master file loaded (${masterData.fileName}). Now upload your Target File to run the audit.`} />}
        {!masterData && targetModels.length > 0 && <EmptyState text="Now upload the Master File from Orion to map models and run the audit." />}

        {/* Results */}
        {ready && sortedMappings.map(mapping => {
          if (!mapping.masterModelId) return (
            <div key={mapping.targetKey} style={{ background: "#fff7ed", borderRadius: 12, padding: "14px 20px", marginBottom: 10, display: "flex", alignItems: "center", gap: 12, border: "1px solid #fed7aa" }}>
              <span style={{ fontSize: 13, color: "#92400e" }}>⚠ <b>{mapping.targetLabel}</b> — not mapped to a master model.</span>
              <button onClick={() => setMappingOpen(true)} style={{ marginLeft: "auto", padding: "4px 12px", borderRadius: 6, background: "#f97316", border: "none", color: "white", fontSize: 12, cursor: "pointer" }}>Fix mapping</button>
            </div>
          );

          const targetModel = targetModels.find(tm => tm.modelKey === mapping.targetKey);
          if (!targetModel) return null;
          const { rows, summary, excluded, sumIncluded } = runAudit(mapping, targetModel, masterData.holdings, excludedClasses, threshold);

          const isOpen = expanded[mapping.targetKey] === true;
          const filt = getFilter(mapping.targetKey);
          const filtRows = filt === "ALL" ? rows : filt === "DISCREPANCY" ? rows.filter(r => r.status !== "MATCH") : rows.filter(r => r.status === filt);
          const hasIssues = summary.over + summary.under + summary.missing > 0;

          return (
            <div key={mapping.targetKey} style={{ background: "white", borderRadius: 12, boxShadow: "0 1px 3px rgba(0,0,0,0.06)", marginBottom: 10, overflow: "hidden" }}>
              {/* Accordion header */}
              <button onClick={() => toggleExpand(mapping.targetKey)} style={{ width: "100%", background: "none", border: "none", cursor: "pointer", padding: "13px 20px", display: "flex", alignItems: "center", gap: 12, textAlign: "left" }}>
                <span style={{ fontSize: 13, transform: isOpen ? "rotate(90deg)" : "none", transition: "transform 0.2s", color: "#94a3b8", display: "inline-block" }}>▶</span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>{mapping.masterModelName}</div>
                  <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 1 }}>
                    Target: <span style={{ color: "#64748b" }}>{mapping.targetLabel}</span>
                    {excluded.length > 0 && ` · ${excluded.length} excluded`}
                    {excludedClasses.length > 0 && sumIncluded > 0 && ` · re-weighted from ${sumIncluded.toFixed(1)}%`}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                  {summary.over > 0    && <Pill label={`${summary.over} over`}    color="#dc2626" bg="#fef2f2" />}
                  {summary.under > 0   && <Pill label={`${summary.under} under`}  color="#d97706" bg="#fffbeb" />}
                  {summary.missing > 0 && <Pill label={`${summary.missing} missing`} color="#7c3aed" bg="#f5f3ff" />}
                  {!hasIssues          && <Pill label="All on target" color="#16a34a" bg="#f0fdf4" />}
                  <Pill label={`${summary.total} positions`} color="#64748b" bg="#f8fafc" />
                </div>
              </button>

              {isOpen && (
                <>
                  {/* Filter bar */}
                  <div style={{ padding: "0 20px 10px", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", borderBottom: "1px solid #f1f5f9" }}>
                    {["ALL","DISCREPANCY","OVER","UNDER","MISSING_MASTER","MISSING_TARGET","MATCH"].map(f => {
                      const active = filt === f;
                      const count = f==="ALL" ? summary.total : f==="DISCREPANCY" ? (summary.total-summary.match) : f==="OVER" ? summary.over : f==="UNDER" ? summary.under : f==="MATCH" ? summary.match : rows.filter(r=>r.status===f).length;
                      return (
                        <button key={f} onClick={() => setFilter(mapping.targetKey, f)} style={{ padding: "4px 10px", borderRadius: 20, border: `1.5px solid ${active?"#3b82f6":"#e2e8f0"}`, background: active?"#eff6ff":"white", color: active?"#1d4ed8":"#64748b", fontWeight: active?600:400, fontSize: 11, cursor: "pointer" }}>
                          {f==="DISCREPANCY"?"Issues":f==="ALL"?"All":STATUS[f]?.label||f}
                          <span style={{ marginLeft: 4, background: active?"#3b82f6":"#e2e8f0", color: active?"white":"#64748b", borderRadius: 8, padding: "0 5px", fontSize: 10 }}>{count}</span>
                        </button>
                      );
                    })}
                    <span style={{ marginLeft: "auto", fontSize: 11, color: "#94a3b8" }}>±{threshold}% tolerance</span>
                  </div>

                  {/* Excluded notice */}
                  {excluded.length > 0 && (
                    <div style={{ margin: "6px 20px 0", padding: "7px 12px", background: "#fff7ed", borderRadius: 8, fontSize: 12, color: "#92400e" }}>
                      ⊘ Excluded ({[...new Set(excluded.map(r=>r.assetClass))].join(", ")}): {excluded.map(r=>r.ticker).join(", ")}
                    </div>
                  )}

                  {/* Table */}
                  {filtRows.length > 0 ? (
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: "#f8fafc" }}>
                            {["Status","Ticker","Security","Asset Class","Master Target %","Adj. Target %","Diff"].map(h => (
                              <th key={h} style={{ padding: "9px 14px", textAlign: ["Master Target %","Adj. Target %","Diff"].includes(h)?"right":"left", fontSize: 10, fontWeight: 600, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.04em", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {filtRows.map((row, i) => {
                            const s = STATUS[row.status];
                            return (
                              <tr key={row.ticker+i} style={{ background: i%2===0?"white":"#fafbfc", borderBottom: "1px solid #f1f5f9" }}>
                                <td style={{ padding: "8px 14px", whiteSpace: "nowrap" }}>
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: s.bg, color: s.color, padding: "2px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>
                                    <span style={{ width: 5, height: 5, borderRadius: "50%", background: s.dot, flexShrink: 0 }} />{s.label}
                                  </span>
                                </td>
                                <td style={{ padding: "8px 14px", fontWeight: 600, color: "#0f172a", fontFamily: "monospace", fontSize: 12 }}>{row.ticker}</td>
                                <td style={{ padding: "8px 14px", color: "#334155", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.name}</td>
                                <td style={{ padding: "8px 14px", color: "#64748b", fontSize: 12, whiteSpace: "nowrap" }}>{row.assetClass}</td>
                                <td style={{ padding: "8px 14px", textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#0f172a" }}>
                                  {row.masterTarget !== null ? `${row.masterTarget.toFixed(2)}%` : <span style={{ color: "#cbd5e1" }}>—</span>}
                                </td>
                                <td style={{ padding: "8px 14px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>
                                  {row.adjTarget !== null
                                    ? <span title={row.rawTarget !== null ? `Raw: ${row.rawTarget.toFixed(2)}%` : ""} style={{ color: "#0f172a", cursor: row.rawTarget !== row.adjTarget ? "help" : "default", borderBottom: row.rawTarget !== row.adjTarget ? "1px dashed #94a3b8" : "none" }}>
                                        {row.adjTarget.toFixed(2)}%
                                      </span>
                                    : <span style={{ color: "#cbd5e1" }}>—</span>}
                                </td>
                                <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 600, fontVariantNumeric: "tabular-nums", color: row.diff===null?"#cbd5e1":row.diff>0?"#dc2626":row.diff<0?"#d97706":"#16a34a" }}>
                                  {row.diff===null ? "—" : `${row.diff>0?"+":""}${row.diff.toFixed(2)}%`}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div style={{ padding: 24, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>No holdings match this filter.</div>
                  )}
                </>
              )}
            </div>
          );
        })}
        {/* Legend */}
        {ready && (
          <div style={{ marginTop: 8, display: "flex", gap: 14, flexWrap: "wrap", fontSize: 11, color: "#94a3b8" }}>
            <span>● <b style={{ color:"#dc2626" }}>Overweight</b>: Master &gt; Adj. target beyond tolerance</span>
            <span>● <b style={{ color:"#d97706" }}>Underweight</b>: Master &lt; Adj. target beyond tolerance</span>
            <span>● <b style={{ color:"#7c3aed" }}>Not in Master</b>: In your targets, absent from Master</span>
            <span>● <b style={{ color:"#64748b" }}>No Target Entry</b>: In Master, not in target file</span>
            <span style={{ marginLeft: "auto" }}>Hover Adj. Target % to see raw pre-reweight value</span>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Sub-components ────────────────────────────────────────────────────────────
const lbl = { fontSize: 11, fontWeight: 600, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.05em", display: "block", marginBottom: 6 };

function DropZone({ label, sublabel, accent, loading, error, dragging, dropProps, onFiles, hint, children }) {
  const hasChildren = Array.isArray(children) ? children.filter(Boolean).length > 0 : !!children;
  return (
    <div style={{ background: "white", borderRadius: 12, padding: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.06)", borderTop: `3px solid ${accent}` }}>
      <div style={{ marginBottom: 10 }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: "#0f172a" }}>{label}</div>
        <div style={{ fontSize: 11, color: "#94a3b8" }}>{sublabel}</div>
      </div>
      {hasChildren && <div style={{ marginBottom: 8 }}>{children}</div>}
      <label {...dropProps} style={{ display: "block", border: `1.5px dashed ${dragging?accent:"#e2e8f0"}`, borderRadius: 8, padding: 12, textAlign: "center", cursor: "pointer", background: dragging ? accent + "10" : "transparent", transition: "all 0.15s" }}>
        <input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => onFiles(Array.from(e.target.files))} />
        {loading ? <span style={{ color: "#94a3b8", fontSize: 12 }}>Reading…</span>
          : <span style={{ fontSize: 12, color: dragging?accent:"#64748b", fontWeight: dragging?600:400 }}>{dragging?"Release to upload":hasChildren?"+ Replace file":hint}</span>}
      </label>
      {error && <div style={{ marginTop: 6, fontSize: 11, color: "#dc2626" }}>{error}</div>}
    </div>
  );
}

function FileChip({ name, sub, color, badge, onRemove }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#f8fafc", borderRadius: 7, padding: "7px 10px", gap: 8 }}>
      <div style={{ minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 1 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#0f172a", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{name}</div>
          {badge && <span style={{ fontSize: 10, fontWeight: 600, color: badge.color, background: badge.bg, padding: "1px 6px", borderRadius: 8, whiteSpace: "nowrap", flexShrink: 0 }}>{badge.label}</span>}
        </div>
        <div style={{ fontSize: 11, color: "#94a3b8", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sub}</div>
      </div>
      <button onClick={onRemove} style={{ flexShrink: 0, background: "none", border: "none", cursor: "pointer", color: "#94a3b8", fontSize: 16, lineHeight: 1, padding: "2px 4px" }}>×</button>
    </div>
  );
}

function Pill({ label, color, bg }) {
  return <span style={{ fontSize: 11, fontWeight: 600, color, background: bg, padding: "3px 8px", borderRadius: 10, whiteSpace: "nowrap" }}>{label}</span>;
}

function EmptyState({ text, icon }) {
  return (
    <div style={{ background: "white", borderRadius: 12, padding: 48, textAlign: "center" }}>
      {icon && <div style={{ width: 44, height: 44, background: "#f1f5f9", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5"><path d="M9 17v-2m3 2v-4m3 4v-6M13 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V9l-6-6z"/></svg></div>}
      <div style={{ fontSize: 13, color: "#64748b", maxWidth: 380, margin: "0 auto" }}>{text}</div>
    </div>
  );
}
