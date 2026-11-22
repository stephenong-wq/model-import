// Model Audit Tool v3.0 — STP model review, sleeve strategy review, and the original general audit
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

  // Fast-path: Flat format — same structure as master file
  // Detected by header containing "Target Percent" + "Model Aggregation Name"
  const firstWs  = wb.Sheets[wb.SheetNames[0]];
  const firstRaw = XLSX.utils.sheet_to_json(firstWs, { header: 1, defval: null });
  const h0 = (firstRaw[0] || []).map(c => String(c || "").toLowerCase());
  if (h0.some(v => v.includes("target percent")) && h0.some(v => v.includes("model aggregation name"))) {
    const rows = XLSX.utils.sheet_to_json(firstWs, { defval: null });
    if (rows.length) {
      const keys        = Object.keys(rows[0]);
      const findCol     = (...cs) => keys.find(k => cs.some(c => k.toLowerCase().includes(c.toLowerCase())));
      const tickerCol   = findCol("ticker", "symbol");
      const nameCol     = findCol("product name", "name", "security");
      const targetCol   = findCol("target percent", "target");
      const modelCol    = findCol("model aggregation name", "model name");

      const byModel = {};
      for (const r of rows) {
        const ticker    = String(r[tickerCol] || "").trim().toUpperCase();
        const modelName = String(r[modelCol]  || "").trim();
        const target    = parseFloat(r[targetCol]) || 0;
        if (!ticker || !modelName || target <= 0) continue;
        if (!byModel[modelName]) byModel[modelName] = [];
        byModel[modelName].push({ ticker, name: String(r[nameCol] || "").trim(), target, assetClass: "Equity", subSection: "" });
      }

      return Object.entries(byModel).map(([modelName, positions]) => ({
        modelKey: `flat__${modelName}`, modelLabel: modelName,
        sheetName: wb.SheetNames[0], colLabel: null, positions: dedupePositions(positions),
      }));
    }
  }

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

  // ── Derive all STP variants dynamically from master model list ────────────
  const hasSTP = results.some(r => r.modelLabel.includes("Savvy Total Portfolios"));
  if (hasSTP && masterModels.length) {

    // Parse stock model tabs
    const stockModelTabs = {
      "Core Stock Model":     "Savvy Core S&P 500 Stock Model",
      "Growth Stock Model":   "Savvy LC Growth Stock Model",
      "Dividend Stock Model": "Savvy Dividend Stock Model",
      "Value Stock Model":    "Savvy Value Stock Model",
    };
    const stockModelHoldings = {};
    for (const [variantName, tabName] of Object.entries(stockModelTabs)) {
      const ws2 = wb.Sheets[tabName];
      if (!ws2) continue;
      const raw2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null });
      const holdings = [];
      for (let ri = 1; ri < raw2.length; ri++) {
        const row = raw2[ri];
        if (!row) continue;
        const ticker = String(row[0] || "").trim().toUpperCase();
        if (!ticker || ticker.length > 6) continue;
        const weight = parseFloat(row[2]);
        if (isNaN(weight) || weight <= 0) continue;
        holdings.push({ ticker, name: String(row[1] || "").trim(), weight });
      }
      const total = holdings.reduce((s, h) => s + h.weight, 0);
      if (total > 0) for (const h of holdings) h.weight /= total;
      stockModelHoldings[variantName] = holdings;
    }

    function substituteUSLC(positions, stockVariant) {
      const stockHoldings = stockModelHoldings[stockVariant];
      if (!stockHoldings) return positions;
      const ulcWeight = positions.filter(p => (p.subSection||"").includes("us large cap")).reduce((s,p)=>s+p.target,0);
      const without   = positions.filter(p => !(p.subSection||"").includes("us large cap"));
      const stocks    = stockHoldings.map(h => ({ ticker: h.ticker, name: h.name, target: h.weight * ulcWeight, assetClass: "Equity", subSection: "us large cap" }));
      return dedupePositions([...without, ...stocks]);
    }

    function reweightAll(positions) {
      const total = positions.reduce((s,p)=>s+p.target,0);
      return total > 0 ? positions.map(p=>({...p, target: p.target/total*100})) : positions;
    }

    function excludeSS(positions, ...sections) {
      return reweightAll(positions.filter(p => !sections.some(s => (p.subSection||"").toLowerCase().includes(s.toLowerCase()))));
    }

    const STP_SHEET     = "Savvy Total Portfolios";
    const STP_TAX_SHEET = "Savvy Total Portfolios (Tax Awa";

    // Base risk → source model label in results
    const BASE_RISK_MAP = {};
    const TAX_RISK_MAP  = {};
    const BASE_ORDER = ["All Fixed","All Equity","Moderately Conservative","Moderately Aggressive","Conservative","Moderate","Aggressive"];
    for (const risk of BASE_ORDER) {
      BASE_RISK_MAP[risk] = `${STP_SHEET} — ${risk}`;
      TAX_RISK_MAP[risk]  = `${STP_TAX_SHEET} — ${risk}`;
    }

    // Build midpoints and add to maps
    function avgPositions(lp, rp) {
      const ri = {};
      for (const p of rp) ri[normalizeTicker(p.ticker)] = p;
      const b = {};
      for (const p of lp) {
        const k = normalizeTicker(p.ticker);
        b[k] = { ...p, target: (p.target + (ri[k]?ri[k].target:0)) / 2 };
      }
      for (const p of rp) {
        const k = normalizeTicker(p.ticker);
        if (!b[k]) b[k] = { ...p, target: p.target / 2 };
      }
      return dedupePositions(Object.values(b).filter(p=>p.target>0));
    }

    const MIDPOINT_NEIGHBORS = {
      "10/90": ["All Fixed","Conservative"],
      "30/70": ["Conservative","Moderately Conservative"],
      "50/50": ["Moderately Conservative","Moderate"],
      "70/30": ["Moderate","Moderately Aggressive"],
    };
    const midpointCache = {}; // ratio → {reg: positions, tax: positions}
    for (const [ratio, [left, right]] of Object.entries(MIDPOINT_NEIGHBORS)) {
      const lReg = results.find(r=>r.modelLabel===BASE_RISK_MAP[left]);
      const rReg = results.find(r=>r.modelLabel===BASE_RISK_MAP[right]);
      const lTax = results.find(r=>r.modelLabel===TAX_RISK_MAP[left]);
      const rTax = results.find(r=>r.modelLabel===TAX_RISK_MAP[right]);
      midpointCache[ratio] = {
        reg: lReg && rReg ? avgPositions(lReg.positions, rReg.positions) : null,
        tax: lTax && rTax ? avgPositions(lTax.positions, rTax.positions) : null,
      };
      if (midpointCache[ratio].reg) BASE_RISK_MAP[ratio] = `__mp_reg_${ratio}`;
      if (midpointCache[ratio].tax) TAX_RISK_MAP[ratio]  = `__mp_tax_${ratio}`;
    }

    function getBasePositions(base, taxAware) {
      const map = taxAware ? TAX_RISK_MAP : BASE_RISK_MAP;
      const key = map[base];
      if (!key) return null;
      if (key.startsWith("__mp_reg_")) return midpointCache[base]?.reg || null;
      if (key.startsWith("__mp_tax_")) return midpointCache[base]?.tax || null;
      const model = results.find(r=>r.modelLabel===key);
      return model ? model.positions : null;
    }

    // Decode STP model name
    function decodeSTPName(name) {
      let base = null;
      const allBases = [...BASE_ORDER, ...Object.keys(MIDPOINT_NEIGHBORS)];
      // Sort longest first to avoid partial matches (Moderately Conservative before Conservative)
      const sorted = allBases.sort((a,b)=>b.length-a.length);
      for (const b of sorted) {
        if (name.includes(`STP - ${b}`)) { base = b; break; }
      }
      return {
        base,
        taxAware:     /\(Tax Aware\)/i.test(name) || /- Tax Aware/i.test(name),
        exUslc:       /\(ex-USLC\)|\(Ex-USLC\)/i.test(name) && !/Ex-Int Dev/i.test(name),
        exAll:        /Ex-USLC, Ex-Int Dev, Ex-EM/i.test(name),
        usEquityOnly: /US Equity Only/i.test(name),
        stockVariant: Object.keys(stockModelTabs).find(sv=>name.includes(sv)) || null,
        isDotE:       name.endsWith(".e"),
      };
    }

    // Apply all transformations to a set of positions
    function applyTransforms(pos, d) {
      // ex-USLC: remove US LC, reweight all remaining pro-rata
      if (d.exUslc || d.exAll) pos = excludeSS(pos, "us large cap");
      // ex-All: also remove intl + emerging
      if (d.exAll) pos = excludeSS(pos, "intl developed", "emerging");
      // US Equity Only: remove intl/emerging, freed weight goes to US LC (or pro-rata if no US LC)
      if (d.usEquityOnly) {
        const removedWt = pos.filter(p=>(p.subSection||"").match(/intl|emerging/i)).reduce((s,p)=>s+p.target,0);
        const ulcWt     = pos.filter(p=>(p.subSection||"").includes("us large cap")).reduce((s,p)=>s+p.target,0);
        if (ulcWt > 0) {
          const sf = (ulcWt + removedWt) / ulcWt;
          pos = dedupePositions(pos
            .filter(p=>!(p.subSection||"").match(/intl|emerging/i))
            .map(p=>(p.subSection||"").includes("us large cap") ? {...p,target:p.target*sf} : p));
        } else {
          pos = reweightAll(pos.filter(p=>!(p.subSection||"").match(/intl|emerging/i)));
        }
      }
      // Stock substitution: replace US LC with individual stocks
      if (d.stockVariant) pos = substituteUSLC(pos, d.stockVariant);
      return dedupePositions(pos);
    }

    const existingKeys = new Set(results.map(r=>r.modelLabel));

    // Iterate every STP model in master and derive positions
    for (const masterModel of masterModels) {
      const name = masterModel.name;
      if (!name.startsWith("STP - ")) continue;
      if (existingKeys.has(name)) continue;
      const d = decodeSTPName(name);
      if (!d.base || d.isDotE) continue;

      const basePos = getBasePositions(d.base, d.taxAware);
      if (!basePos || !basePos.length) continue;

      const positions = applyTransforms([...basePos], d);
      if (!positions.length) continue;

      results.push({ modelKey: `derived__${name}`, modelLabel: name, sheetName: "Derived", colLabel: null, positions });
      existingKeys.add(name);
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

// ═══════════════════════════════════════════════════════════════════════════
// Shared with OrionImportBuilder.jsx — the model-name rules and the sleeve
// strategy logic. Keep in sync with that file. (Two small differences, marked
// below: matchModelVariant takes a family table, and computeSleeveUpdates also
// reports the detail the look-through review needs.)
// ═══════════════════════════════════════════════════════════════════════════
const MODEL_FAMILIES = [
  { id:"paragon", test:/^\s*paragon\/stp\b/i, source:"savvy", sheetPrefix:"Savvy Total Portfolios", libraryLabel:"Savvy" },
  { id:"tacticalEtf", test:/^\s*stp\s+tactical\s+etf\b/i,
    skip:"STP Tactical ETF models are a single Security Set with no categories or classes — nothing to update here" },
  { id:"direct", test:/^\s*stp\s+direct\s+hh\b/i, source:"direct", sheetPrefix:"STP Direct Med", libraryLabel:"STP Direct" },
  { id:"stp", test:/^\s*stp\b/i, source:"savvy", sheetPrefix:"Savvy Total Portfolios", libraryLabel:"Savvy" },
];

function normalizeKey(k) { return (k||"").toString().trim().toLowerCase(); }

function normAlnum(s) {
  return (s||"").toString().toUpperCase().replace(/[^A-Z0-9]/g,"");
}

function libraryShortLabel(label) {
  return (label||"").replace(/\([^)]*\)/g," ").split(" - ")[0].replace(/\s{2,}/g," ").trim();
}

function matchLibraryLabel(sourceStr, libraryRows) {
  const ns = normAlnum(sourceStr);
  if (!ns) return null;
  const pick = labelOf => {
    let best = null, bestEnd = -1, bestLen = 0;
    libraryRows.forEach(row => {
      const nl = normAlnum(labelOf(row));
      if (!nl) return;
      const idx = ns.lastIndexOf(nl);
      if (idx === -1) return;
      const end = idx + nl.length;
      if (end > bestEnd || (end === bestEnd && nl.length > bestLen)) {
        best = row; bestEnd = end; bestLen = nl.length;
      }
    });
    return best;
  };
  return pick(row => row.label) || pick(row => libraryShortLabel(row.label));
}

function analyzeModelName(modelName) {
  const core = (modelName||"").replace(/\([^)]*\)/g,"").replace(/\s{2,}/g," ").trim();
  const flags = {
    isTaxAware: /tax/i.test(modelName||""),
    isExUslc: /ex[-\s]?uslc/i.test(modelName||""),
    isExFI: /\bex[-\s]?fi\b/i.test(modelName||""),
    isExIntDev: /\bex[-\s]?int(?:l|ernational)?\.?\s*dev/i.test(modelName||""),
    isExEM: /\bex[-\s]?em\b/i.test(modelName||""),
    isUsEquityOnly: /us\s+equity\s+only/i.test(modelName||""),
    isParagon: /^\s*paragon\/stp\b/i.test(modelName||""),
    holdingsAssetClass: null, // "equity" | "fixedIncome" | "other" | null
  };
  let segments;
  if (flags.isParagon) {
    // "Paragon/STP All Equity (ex-USLC)" and "Paragon/STP All Equity + Core Stock
    // Model" have no " - " separator: the prefix is followed directly by the
    // variant, optionally joined to a stock-model tag with " + ".
    let rest = core.replace(/^paragon\/stp\s*/i,"").trim();
    const plus = rest.match(/^(.*?)\s*\+\s*(?:core|growth|value|dividend)\s+stock\s+model$/i);
    if (plus) { rest = plus[1].trim(); flags.holdingsAssetClass = "equity"; }
    segments = ["Paragon/STP", rest].filter(Boolean);
  } else {
    segments = core.split(" - ").map(s=>s.trim()).filter(Boolean);
  }
  let changed = true;
  while (changed && segments.length > 2) {
    changed = false;
    const last = segments[segments.length-1];
    if (/^tax\s*aware$/i.test(last)) { flags.isTaxAware = true; segments.pop(); changed = true; }
    else if (/^us\s+equity\s+only$/i.test(last)) { flags.isUsEquityOnly = true; segments.pop(); changed = true; }
    else if (/^(core|growth|value|dividend)\s+stock\s+model$/i.test(last) || /^enhanced\s+dividend\s+stock$/i.test(last)) {
      flags.holdingsAssetClass = "equity"; segments.pop(); changed = true;
    }
    else if (/^bond\s+ladder$/i.test(last)) { flags.holdingsAssetClass = "fixedIncome"; segments.pop(); changed = true; }
    else if (/^custom\s+model$/i.test(last)) { flags.holdingsAssetClass = flags.holdingsAssetClass || "other"; segments.pop(); changed = true; }
  }
  return { segments, flags };
}

function matchModelVariant(modelName, librarySheets, families = MODEL_FAMILIES) {
  const { segments, flags } = analyzeModelName(modelName);
  const family = families.find(f => f.test.test(modelName||""));
  if (!family) {
    return { match:null, flags, reason:"doesn't start with a recognized model family (STP, STP Direct HH, Paragon/STP)" };
  }
  if (family.skip) return { match:null, flags, reason:family.skip };

  // An exclusion tag the tool doesn't know would silently produce wrong numbers
  // (the model would be computed as if it weren't excluded) — skip it instead.
  const unknownEx = (modelName||"")
    .replace(/ex[-\s]?uslc/gi,"").replace(/\bex[-\s]?fi\b/gi,"")
    .replace(/\bex[-\s]?int(?:l|ernational)?\.?\s*dev/gi,"").replace(/\bex[-\s]?em\b/gi,"");
  if (/\bex[-\s]+[a-z]/i.test(unknownEx)) {
    return { match:null, flags, reason:"has an exclusion tag the tool doesn't recognize (known: Ex-USLC, Ex-FI, Ex-Int Dev, Ex-EM)" };
  }
  if (segments.length !== 2) {
    return { match:null, flags, reason:`${segments.length} segment(s) left after removing recognized tags (${segments.join(" / ")||"none"}) — doesn't reduce to a single "{prefix} - {variant}" form` };
  }
  const candidate = segments[1];

  const fromLibrary = librarySheets.filter(s => !s.source || s.source === family.source);
  if (fromLibrary.length === 0) {
    return { match:null, flags, reason:`the ${family.libraryLabel} library hasn't been uploaded` };
  }
  const sheetPrefix = typeof family.sheetPrefix === "function" ? family.sheetPrefix(modelName) : family.sheetPrefix;
  const prefixN = normAlnum(sheetPrefix);
  const searchIn = fromLibrary.filter(s => normAlnum(s.name).startsWith(prefixN) && s.isTaxAware === flags.isTaxAware);
  if (searchIn.length === 0) {
    return { match:null, flags, reason:`no ${flags.isTaxAware?"Tax Aware ":""}"${sheetPrefix}" tab found in the ${family.libraryLabel} library` };
  }

  for (const sheet of searchIn) {
    const variant = sheet.variants.find(v => v.toLowerCase() === candidate.toLowerCase());
    if (variant) return { match:{ kind:"single", sheet, variant }, flags, reason:null };
  }

  // 30/70-style blend: verified against real data as the average of the two
  // neighboring risk-ladder variants — pair index = (firstNumber-10)/20 along
  // the library's own column order (10/90=AllFixed+Conservative, 30/70=
  // Conservative+ModConservative, 50/50=ModConservative+Moderate, 70/30=
  // Moderate+ModAggressive). Only reached when the library has no column of
  // that exact name — a 50/50 or 70/30 column in the library is used directly.
  const blendMatch = candidate.match(/^(\d{1,3})\s*\/\s*(\d{1,3})$/);
  if (blendMatch) {
    const n1 = parseInt(blendMatch[1],10), n2 = parseInt(blendMatch[2],10);
    if (n1+n2 !== 100) {
      return { match:null, flags, reason:`blend ratio "${candidate}" doesn't sum to 100 — can't place it on the risk ladder` };
    }
    const pairIndex = (n1-10)/20;
    if (!Number.isInteger(pairIndex)) {
      return { match:null, flags, reason:`blend ratio "${candidate}" doesn't align to the library's 20-point variant spacing` };
    }
    for (const sheet of searchIn) {
      if (pairIndex>=0 && pairIndex+1 < sheet.variants.length) {
        return { match:{ kind:"blend", sheet, lowerVariant: sheet.variants[pairIndex], upperVariant: sheet.variants[pairIndex+1] }, flags, reason:null };
      }
    }
    return { match:null, flags, reason:`blend ratio "${candidate}" falls outside the library's variant range` };
  }

  return { match:null, flags, reason:`"${candidate}" doesn't match any library column name or a recognized blend ratio` };
}

function getMatchFrac(row, match) {
  if (!row) return null;
  if (match.kind === "single") return row.values[match.variant];
  const a = row.values[match.lowerVariant], b = row.values[match.upperVariant];
  if ((a===null||a===undefined) && (b===null||b===undefined)) return null;
  return ((a??0) + (b??0)) / 2;
}

const SLEEVE_AGG_KEY = "sleeve-agg-map";

async function loadSleeveAggMap() {
  try {
    const res = await window.storage.get(SLEEVE_AGG_KEY);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch { return null; }
}

async function saveSleeveAggMap(obj) {
  try { await window.storage.set(SLEEVE_AGG_KEY, JSON.stringify(obj)); } catch {}
}

function parseSleeveExport(buffer) {
  const wb = XLSX.read(buffer, { type:"array" });
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:null });
  let headerIdx = -1;
  for (let i=0; i<Math.min(raw.length,10); i++) {
    if ((raw[i]||[]).map(normalizeKey).includes("model agg id")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) throw new Error("Couldn't find a 'Model Agg ID' column — this should be the Sleeve Strategy Export file.");
  const head = raw[headerIdx].map(normalizeKey);
  const idCol = head.indexOf("model agg id");
  let nameCol = head.indexOf("model agg. name");
  if (nameCol === -1) nameCol = head.indexOf("model agg name");
  if (nameCol === -1) nameCol = head.indexOf("model");
  if (nameCol === -1) throw new Error("Couldn't find the 'Model Agg. Name' column in the Sleeve Strategy Export.");
  const map = {};
  for (let i=headerIdx+1; i<raw.length; i++) {
    const row = raw[i]; if (!row) continue;
    const id = row[idCol], nm = row[nameCol];
    if (id === null || id === undefined || id === "" || !nm) continue;
    map[String(id).trim()] = String(nm).trim();
  }
  if (Object.keys(map).length === 0) throw new Error("No model aggregates found in the Sleeve Strategy Export.");
  return map;
}

function parseSleeveImport(buffer) {
  const wb = XLSX.read(buffer, { type:"array" });
  const sheetName = wb.SheetNames[0];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:null });
  let headerIdx = -1;
  for (let i=0; i<Math.min(raw.length,10); i++) {
    if ((raw[i]||[]).map(normalizeKey).includes("sleeve strategy id")) { headerIdx = i; break; }
  }
  if (headerIdx === -1) throw new Error("Couldn't find a 'Sleeve Strategy ID' column — this should be the Sleeve Strategy Import file.");
  const headers = [];
  const names = raw[headerIdx].map(c => (c===null||c===undefined) ? null : c.toString().trim());
  names.forEach(h => { if (h && !headers.includes(h)) headers.push(h); });
  for (const need of ["Sleeve Strategy ID","Name","SleeveType","Modelagg Id","Target Allocation"]) {
    if (!headers.includes(need)) throw new Error(`The Sleeve Strategy Import is missing the '${need}' column.`);
  }
  const rows = [];
  for (let i=headerIdx+1; i<raw.length; i++) {
    const row = raw[i];
    if (!row || row.every(c=>c===null||c==="")) continue;
    const obj = {};
    names.forEach((h,idx)=>{ if (h) obj[h] = row[idx]!==undefined ? row[idx] : null; });
    if (obj["Sleeve Strategy ID"]===null || obj["Sleeve Strategy ID"]==="") continue;
    rows.push(obj);
  }
  if (rows.length === 0) throw new Error("No sleeve strategies found in the import file.");
  return { headers, rows, sheetName };
}

function sleeveKind(n) {
  if (/thematic/i.test(n)) return "THEMATIC";
  if (/Large Cap 500/i.test(n)) return "DI_SPX";
  if (/Total Portfolios Model US Large Cap Equity/i.test(n)) return "DI_USLC";
  if (/Total Portfolios Global Equity/i.test(n)) return "DI_GLOBAL";
  if (/Corporate Bond Strategy/i.test(n)) return "CORP_BOND";
  if (/Municipal Bond Strategy/i.test(n)) return "MUNI_BOND";
  if (/Paragon Core ETF All Fixed/i.test(n)) return "PARAGON_FI";
  if (/^Paragon\/STP/i.test(n)) return "PARAGON_EQ";
  if (/^STP Tactical ETF/i.test(n)) return "TACT_ETF";
  if (/^STP - /i.test(n)) return /ex-uslc|ex-fi/i.test(n) ? "STP_EX" : "STP";
  if (/Bond Ladder/i.test(n)) return "LADDER";
  if (/Custom Indexing|Orion Custom/i.test(n)) return "DI_OTHER";
  if (/Stock Model|Single Stock/i.test(n)) return "STOCK";
  return "OTHER";
}

const SLEEVE_EQUITY_DI = new Set(["DI_SPX","DI_USLC","DI_GLOBAL","DI_OTHER"]);

const SLEEVE_FI = new Set(["CORP_BOND","MUNI_BOND"]);

const sleeveRound = x => +(Math.round(x*100 + 1e-9)/100).toFixed(2);

const isNormalSleeve = r => String(r["SleeveType"] || "").trim().toLowerCase() === "normal"; // only these rows carry the targets; Custodial / Contribution / Distribution rows are ignored

const SLEEVE_TOLERANCE = 1e-9; // exact: any real difference is brought to the library value (float noise only)

function parseDirectStrategyName(name) {
  const m = (name||"").match(/^STP Direct (Large|Medium)\s*\((NQ|Qualified)\)\s*-\s*(.+?)\s*$/i);
  if (!m) return null;
  let profile = m[3].trim(), isTA = false;
  const t = profile.match(/^(.*?)\s*\((?:TA|Tax[-\s]?Aware)\)\s*$/i);
  if (t) { profile = t[1].trim(); isTA = true; }
  const alias = { "mod agg":"Moderately Aggressive", "mod cons":"Moderately Conservative" };
  profile = alias[profile.toLowerCase()] || profile;
  return { size: /large/i.test(m[1]) ? "Large" : "Medium", acct: m[2], profile, isTA };
}

function sleeveExcludedFraction(aggName, librarySheets) {
  const { match, flags, reason } = matchModelVariant(aggName, librarySheets);
  if (!match) return { reason };
  const classRows = match.sheet.rows.filter(x=>x.level==="class");
  const catRows = match.sheet.rows.filter(x=>x.level==="category");
  const uslcRow = matchLibraryLabel("US Large Cap", classRows);
  const intlRow = matchLibraryLabel("Intl Developed ex-US Market", classRows);
  const emRow = matchLibraryLabel("Emerging Markets", classRows);
  const raw = row => (row ? (getMatchFrac(row, match) || 0) : 0);
  // Paragon/STP: international is halved and the other half sits in US Large Cap.
  const frac = row => {
    const v = raw(row);
    if (!flags.isParagon) return v;
    if (row === intlRow || row === emRow) return v/2;
    if (row === uslcRow) return v + raw(intlRow)/2 + raw(emRow)/2;
    return v;
  };
  let ex = 0;
  if (flags.isExUslc) { if (!uslcRow) return { reason:"the library has no US Large Cap row" }; ex += frac(uslcRow); }
  if (flags.isExIntDev) { if (!intlRow) return { reason:"the library has no Intl Developed row" }; ex += frac(intlRow); }
  if (flags.isExEM) { if (!emRow) return { reason:"the library has no Emerging Markets row" }; ex += frac(emRow); }
  if (flags.isExFI) {
    const fiRow = matchLibraryLabel("FIXED INCOME", catRows);
    if (!fiRow) return { reason:"the library has no Fixed Income row" };
    ex += raw(fiRow);
  }
  return { frac: ex };
}

function computeSleeveUpdates(rows, aggMap, librarySheets) {
  const updated = rows.map(r => ({...r}));
  const strategies = new Map(); // id -> { id, name, idx: [row indexes of its Normal sleeves] }
  rows.forEach((r, i) => {
    const id = r["Sleeve Strategy ID"];
    if (id === null || id === undefined || id === "") return;
    if (!strategies.has(id)) strategies.set(id, { id, name: r["Name"], idx: [] });
    if (isNormalSleeve(r)) strategies.get(id).idx.push(i);
  });
  const changes = [], statuses = [];

  for (const s of strategies.values()) {
    const comps = s.idx.map(i => {
      const aggId = rows[i]["Modelagg Id"];
      const aggName = aggMap[String(aggId)];
      return { i, aggId, aggName, kind: aggName ? sleeveKind(aggName) : null, old: parseFloat(rows[i]["Target Allocation"]) || 0 };
    });
    const skip = (group, reason) => statuses.push({ id:s.id, name:s.name, status:"skipped", group, reason,
      comps: comps.map(c => ({ aggId:c.aggId, aggName:c.aggName, kind:c.kind, current:c.old, expected:c.old })) });
    // newVals: Map(comp -> new target). Applies only what actually moved.
    const apply = (family, newVals, ctx) => {
      let changed = 0;
      for (const [c, v] of newVals) {
        if (Math.abs(v - c.old) > SLEEVE_TOLERANCE) {
          updated[c.i]["Target Allocation"] = v;
          changes.push({ strategyId:s.id, strategyName:s.name, family, label:c.aggName, oldVal:c.old, newVal:v });
          changed++;
        }
      }
      statuses.push({ id:s.id, name:s.name, status: changed ? "updated" : "unchanged", family, ctx,
        comps: comps.map(c => ({ aggId:c.aggId, aggName:c.aggName, kind:c.kind, current:c.old, expected: newVals.has(c) ? newVals.get(c) : c.old })) });
    };

    if (!/STP|Paragon/i.test(s.name||"")) { skip("Not touched", "No STP or Paragon in the name"); continue; }
    const unknown = comps.filter(c => !c.aggName).map(c => c.aggId);
    if (unknown.length) {
      skip("Needs a fresh Sleeve Strategy Export", `Model Agg ID ${[...new Set(unknown)].join(", ")} isn't in the stored export`);
      continue;
    }

    // ── 1. STP Direct ───────────────────────────────────────────────────────
    if (/^STP Direct\b/i.test(s.name)) {
      const info = parseDirectStrategyName(s.name);
      if (!info) { skip("Couldn't read", "Name isn't in the form 'STP Direct Large|Medium (NQ|Qualified) - Profile'"); continue; }
      const allowed = new Set(["DI_GLOBAL","STOCK","CORP_BOND","MUNI_BOND","TACT_ETF"]);
      const odd = comps.find(c => !allowed.has(c.kind) || (info.size==="Medium" && SLEEVE_FI.has(c.kind)));
      if (odd) { skip("Couldn't read", `Unexpected sleeve: ${odd.aggName}`); continue; }
      const globals = comps.filter(c => c.kind==="DI_GLOBAL" || c.kind==="STOCK");
      const bonds = comps.filter(c => SLEEVE_FI.has(c.kind));
      const tacts = comps.filter(c => c.kind==="TACT_ETF");
      if (tacts.length !== 1 || globals.length > 1 || bonds.length > 1) { skip("Couldn't read", "Expected one Tactical ETF sleeve and at most one Global Equity and one bond sleeve"); continue; }
      const prefix = normAlnum(info.size==="Large" ? "STP Direct Large" : "STP Direct Med");
      const sheet = librarySheets.find(x => (!x.source || x.source==="direct") && normAlnum(x.name).startsWith(prefix) && x.isTaxAware===info.isTA);
      if (!sheet) { skip("Library not available", `no ${info.isTA?"Tax Aware ":""}'STP Direct ${info.size==="Large"?"Large":"Med"}' tab (is the STP Direct library uploaded?)`); continue; }
      const variant = sheet.variants.find(v => v.toLowerCase()===info.profile.toLowerCase());
      if (!variant) { skip("Library not available", `'${info.profile}' isn't a column in ${sheet.name}`); continue; }
      const classRows = sheet.rows.filter(x=>x.level==="class");
      const gRow = matchLibraryLabel("Global Equity", classRows), igRow = matchLibraryLabel("US Investment Grade", classRows);
      const pct = row => row ? sleeveRound((row.values[variant] ?? 0) * 100) : 0;
      const globalPct = pct(gRow), igPct = info.size==="Large" ? pct(igRow) : 0;
      if (globalPct > 0 && !globals.length) { skip("Couldn't read", "The library has Global Equity for this profile but the strategy has no Global Equity sleeve"); continue; }
      if (igPct > 0 && !bonds.length) { skip("Couldn't read", "The library has US Investment Grade for this profile but the strategy has no bond sleeve"); continue; }
      const m = new Map();
      if (globals.length) m.set(globals[0], globalPct);
      if (bonds.length) m.set(bonds[0], igPct);
      m.set(tacts[0], sleeveRound(100 - (globals.length?globalPct:0) - (bonds.length?igPct:0)));
      apply("STP Direct", m, { type:"direct", info, tact: tacts[0] });
      continue;
    }

    // ── 2. ex- model + the sleeve that replaces what it excludes ─────────────
    const exComps = comps.filter(c => c.kind==="STP_EX" || (c.kind==="PARAGON_EQ" && /ex-uslc/i.test(c.aggName)));
    if (exComps.length === 0) { skip("Fixed splits", "Percentages are fixed (not derived from the model library)"); continue; }
    if (exComps.length > 1) { skip("Couldn't read", "More than one ex- model in the strategy"); continue; }
    const ex = exComps[0];
    const { flags } = analyzeModelName(ex.aggName);
    const equityEx = flags.isExUslc || flags.isExIntDev || flags.isExEM, fiEx = flags.isExFI;
    if (equityEx && fiEx) { skip("Couldn't read", `${ex.aggName} excludes both equity and Fixed Income — that needs two replacement sleeves`); continue; }
    if (!equityEx && !fiEx) { skip("Couldn't read", `No recognizable exclusion in ${ex.aggName}`); continue; }
    const wanted = equityEx ? (k => SLEEVE_EQUITY_DI.has(k)) : (k => SLEEVE_FI.has(k));
    const reps = comps.filter(c => c !== ex && wanted(c.kind));
    if (reps.length === 0) { skip("Fixed splits", `Nothing library-driven: no ${equityEx?"DI":"bond"} sleeve takes the weight ${ex.aggName} excludes`); continue; }
    if (reps.length > 1) { skip("Couldn't read", `More than one ${equityEx?"DI":"bond"} sleeve to take the excluded weight`); continue; }
    const rep = reps[0];
    const fixedTotal = comps.filter(c => c !== ex && c !== rep).reduce((t, c) => t + c.old, 0);
    const remainder = 100 - fixedTotal;
    if (remainder < 0) { skip("Couldn't read", "The fixed sleeves already add to more than 100"); continue; }
    const exFrac = sleeveExcludedFraction(ex.aggName, librarySheets);
    if (exFrac.reason) { skip("Library not available", `${ex.aggName}: ${exFrac.reason}`); continue; }
    const repPct = sleeveRound(remainder * exFrac.frac);
    const m = new Map([[rep, repPct], [ex, sleeveRound(remainder - repPct)]]);
    apply(flags.isParagon ? "Paragon" : "STP ex- model + sleeve", m, { type:"split", ex, rep, fixedTotal, remainder });
  }

  return { updatedRows: updated, changes, statuses, strategyCount: strategies.size };
}

// ═══════════════════════════════════════════════════════════════════════════
// STP review engine
// Expected ticker-level holdings for every STP-family model, derived from the
// model libraries with the same name rules as the Import Builder, compared to
// the underlying allocations Orion reports (the Model Holding Weighting file).
// ═══════════════════════════════════════════════════════════════════════════

// The audit also routes Tactical ETF models (the Import Builder skips them: they
// have no categories/classes to update, but their tickers can be reviewed).
const AUDIT_FAMILIES = [
  { id:"paragon", test:/^\s*paragon\/stp\b/i, source:"savvy", sheetPrefix:"Savvy Total Portfolios", libraryLabel:"Savvy" },
  { id:"tacticalEtf", test:/^\s*stp\s+tactical\s+etf\b/i, source:"direct", libraryLabel:"STP Direct",
    sheetPrefix: n => /\(\s*large\s*\)/i.test(n) ? "STP Direct Large" : "STP Direct Med" },
  { id:"direct", test:/^\s*stp\s+direct\s+hh\b/i, source:"direct", sheetPrefix:"STP Direct Med", libraryLabel:"STP Direct" },
  { id:"stp", test:/^\s*stp\s+-\s/i, source:"savvy", sheetPrefix:"Savvy Total Portfolios", libraryLabel:"Savvy" },
];

// The three filters you use in Orion to find the models that need reviewing.
const DEFAULT_SCOPE_FILTERS = [
  { id:"notCustom",  mode:"not", term:"custom model", on:true },
  { id:"notOptions", mode:"not", term:"options",      on:true },
  { id:"hasStp",     mode:"has", term:"STP",          on:true },
];
function inReviewScope(modelName, filters) {
  const n = (modelName||"").toLowerCase();
  return filters.every(f => {
    const term = (f.term||"").trim().toLowerCase();
    if (!f.on || !term) return true;
    return f.mode === "has" ? n.includes(term) : !n.includes(term);
  });
}

// ── Library tabs, with ticker rows (the Import Builder's parser ignores them) ──
// Same shape the shared sleeve/matching code expects (name, variants, rows,
// isTaxAware, source) plus `tickers`: [{ ticker, name, category, cls, values }].
function parseLibraryTabs(wb, source) {
  const tabs = [];
  for (const sheetName of wb.SheetNames) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:null });
    let headerIdx = -1;
    for (let i=0; i<Math.min(raw.length,10); i++) {
      if ((raw[i]||[]).map(normalizeKey).includes("ticker")) { headerIdx = i; break; }
    }
    if (headerIdx === -1) continue;
    const headerRow = raw[headerIdx];
    const variantCols = [];
    for (let c=2; c<headerRow.length; c++) {
      const v = headerRow[c];
      if (typeof v === "string" && v.trim()) variantCols.push({ idx:c, name:v.trim() });
      else if (variantCols.length>0) break;
    }
    const rows = [], tickers = [];
    let category = null, cls = null;
    for (let i=headerIdx+1; i<raw.length; i++) {
      const row = raw[i]; if (!row) continue;
      const label = row[0];
      if (label === null || label === undefined || label === "") continue;
      const labelStr = label.toString().trim(); if (!labelStr) continue;
      const fund = row[1];
      const isTicker = fund !== null && fund !== undefined && fund !== "";
      const values = {};
      variantCols.forEach(({idx,name}) => {
        const n = parseFloat(row[idx]);
        values[name] = Number.isFinite(n) ? n : null;
      });
      if (isTicker) { tickers.push({ ticker: labelStr.toUpperCase(), name: String(fund).trim(), category, cls, values }); continue; }
      const isCategory = labelStr === labelStr.toUpperCase() && labelStr !== labelStr.toLowerCase();
      if (isCategory) { category = labelStr; cls = null; rows.push({ level:"category", label:labelStr, values }); }
      else { if (!/expense ratio/i.test(labelStr)) cls = labelStr; rows.push({ level:"class", label:labelStr, values, category }); }
    }
    if (variantCols.length < 2 || !rows.some(r => r.level === "category")) continue;
    tabs.push({ name:sheetName, variants:variantCols.map(v => v.name), rows, tickers, isTaxAware:/tax|\(\s*ta\s*\)/i.test(sheetName), source });
  }
  return tabs;
}

// Stock-model tabs in the Savvy library (holdings that replace US Large Cap).
const STOCK_MODEL_TABS = {
  "Core Stock Model":     "Savvy Core S&P 500 Stock Model",
  "Growth Stock Model":   "Savvy LC Growth Stock Model",
  "Dividend Stock Model": "Savvy Dividend Stock Model",
  "Value Stock Model":    "Savvy Value Stock Model",
};
const STOCK_VARIANT_RES = [
  ["Core Stock Model", /core\s+stock\s+model/i], ["Growth Stock Model", /growth\s+stock\s+model/i],
  ["Dividend Stock Model", /dividend\s+stock\s+model/i], ["Value Stock Model", /value\s+stock\s+model/i],
];
function parseStockModels(wb) {
  const out = {};
  for (const [variant, tab] of Object.entries(STOCK_MODEL_TABS)) {
    const ws = wb.Sheets[tab]; if (!ws) continue;
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    const holdings = [];
    for (let ri=1; ri<raw.length; ri++) {
      const row = raw[ri]; if (!row) continue;
      const ticker = String(row[0] || "").trim().toUpperCase();
      if (!ticker || ticker.length > 6) continue;
      const weight = parseFloat(row[2]);
      if (!Number.isFinite(weight) || weight <= 0) continue;
      holdings.push({ ticker, name:String(row[1] || "").trim(), weight });
    }
    const total = holdings.reduce((s,h) => s + h.weight, 0);
    if (total > 0) { holdings.forEach(h => { h.weight /= total; }); out[variant] = holdings; }
  }
  return out;
}

// ── Expected holdings for one model ──────────────────────────────────────────
// Returns { positions, notes } — whole-model weights in % (not re-scaled, so what
// an ex- tag removes is visible) — or { reason } when the model can't be derived.
const GLOBAL_EQUITY_TICKER = "SAVVY GLOBAL EQUITY"; // the library's Global Equity row; ACWI in the HH models
const isUSLC = p => /us large cap/i.test(p.cls);
const isIntlDev = p => /intl developed/i.test(p.cls);
const isEmMkt = p => /emerging/i.test(p.cls);
const isFixedIncome = p => /fixed income/i.test(p.assetClass);

function expectedHoldings(modelName, lib) {
  if (/^\s*<\s*copy of/i.test(modelName)) return { reason:"A copy of another model — review the original" };
  const clean = modelName.replace(/\.e\s*$/i, "").trim();
  const m = matchModelVariant(clean, lib.tabs, AUDIT_FAMILIES);
  if (!m.match) return { reason:m.reason };
  const { match, flags } = m;
  const family = AUDIT_FAMILIES.find(f => f.test.test(clean));
  if (flags.holdingsAssetClass === "fixedIncome") return { reason:"Bond Ladder — the ladder's bonds aren't in the library, so it can't be reviewed ticker by ticker" };
  if (flags.holdingsAssetClass === "other") return { reason:"Custom model — not defined by the library" };

  const value = t => match.kind === "single"
    ? (t.values[match.variant] || 0)
    : ((t.values[match.lowerVariant] || 0) + (t.values[match.upperVariant] || 0)) / 2;
  let pos = match.sheet.tickers
    .map(t => ({ ticker:t.ticker, name:t.name, assetClass:toTitleCase(t.category || ""), cls:t.cls || "", target:value(t) * 100 }))
    .filter(p => p.target > 0);
  const notes = [`${match.sheet.name} · ${match.kind === "single" ? match.variant : `${match.lowerVariant} + ${match.upperVariant} (average)`}`];
  const sum = f => pos.filter(f).reduce((s,p) => s + p.target, 0);

  if (family.id === "direct") {
    pos = pos.map(p => p.ticker === GLOBAL_EQUITY_TICKER ? { ...p, ticker:"ACWI", name:"iShares MSCI ACWI" } : p);
  }
  if (family.id === "tacticalEtf") {
    const large = /\(\s*large\s*\)/i.test(clean);
    pos = pos.filter(p => p.ticker !== GLOBAL_EQUITY_TICKER && !(large && /us investment grade/i.test(p.cls)));
    notes.push(large ? "excludes Global Equity and US Investment Grade" : "excludes Global Equity");
  }
  if (flags.isParagon) {
    // Paragon/STP: international (Intl Developed + EM) is half the library's; the other half goes to US Large Cap.
    const us = sum(isUSLC);
    if (us <= 0) return { reason:"The library column has no US Large Cap to add Paragon's international half to" };
    const add = (sum(isIntlDev) + sum(isEmMkt)) / 2;
    pos = pos.map(p => (isIntlDev(p) || isEmMkt(p)) ? { ...p, target:p.target / 2 } : isUSLC(p) ? { ...p, target:p.target * (1 + add/us) } : p);
    notes.push("international halved, the other half added to US Large Cap");
  }
  if (flags.isUsEquityOnly) {
    const removed = sum(p => isIntlDev(p) || isEmMkt(p)), us = sum(isUSLC);
    pos = pos.filter(p => !(isIntlDev(p) || isEmMkt(p)));
    if (us > 0) pos = pos.map(p => isUSLC(p) ? { ...p, target:p.target * (us + removed) / us } : p);
    notes.push("US equity only");
  }
  if (flags.isExUslc)   { pos = pos.filter(p => !isUSLC(p));      notes.push("ex-US Large Cap"); }
  if (flags.isExIntDev) { pos = pos.filter(p => !isIntlDev(p));   notes.push("ex-Intl Developed"); }
  if (flags.isExEM)     { pos = pos.filter(p => !isEmMkt(p));     notes.push("ex-Emerging Markets"); }
  if (flags.isExFI)     { pos = pos.filter(p => !isFixedIncome(p)); notes.push("ex-Fixed Income"); }

  const sv = STOCK_VARIANT_RES.find(([, re]) => re.test(clean));
  if (sv || flags.holdingsAssetClass === "equity") {
    const holdings = sv ? lib.stock[sv[0]] : null;
    if (!holdings) return { reason: sv ? `The library has no '${STOCK_MODEL_TABS[sv[0]]}' tab to read the ${sv[0]} holdings from` : "A stock-model holdings substitution the library doesn't define" };
    // STP stock-model variants replace the whole US Large Cap class with the stocks. The Paragon
    // variant swaps ONLY the VOO (S&P 500) ETF — SCHG, RSP and XLK stay as ETFs.
    const swap = flags.isParagon ? (p => isUSLC(p) && p.ticker === "VOO") : isUSLC;
    const swapped = sum(swap);
    if (swapped > 0) {
      pos = [...pos.filter(p => !swap(p)), ...holdings.map(h => ({ ticker:h.ticker, name:h.name, assetClass:"Equity", cls:"US Large Cap", target:h.weight * swapped }))];
      notes.push(flags.isParagon ? `VOO replaced by the ${sv[0]}` : `US Large Cap replaced by the ${sv[0]}`);
    }
  }
  return { positions: dedupePositions(pos), notes, family: family.id };
}

// ── Compare expected vs actual weights ───────────────────────────────────────
// actual/expected: [{ ticker, name, assetClass?, target }]. `rescale` re-bases the
// expected weights to 100 (Orion's model targets always total 100).
function compareWeights(expected, actual, threshold, { rescale } = {}) {
  const idx = list => { const m = {}; for (const p of list) { if (!(p.target > 0)) continue; const k = normalizeTicker(p.ticker); m[k] = m[k] ? { ...m[k], target:m[k].target + p.target } : { ...p }; } return m; };
  const E = idx(expected), A = idx(actual);
  const eSum = Object.values(E).reduce((s,p) => s + p.target, 0);
  const rows = [];
  for (const k of new Set([...Object.keys(E), ...Object.keys(A)])) {
    const e = E[k], a = A[k];
    const exp = e ? (rescale && eSum > 0 ? e.target / eSum * 100 : e.target) : null;
    const act = a ? a.target : null;
    let status, diff = null;
    if (act === null) status = "MISSING_MASTER";
    else if (exp === null) status = "MISSING_TARGET";
    else { diff = act - exp; status = Math.abs(diff) <= threshold + 1e-9 ? "MATCH" : diff > 0 ? "OVER" : "UNDER"; } // +1e-9: a difference of exactly the tolerance is a match
    rows.push({ ticker:a?.ticker || e?.ticker || k, name:a?.name || e?.name || k, assetClass:e?.assetClass || "—", masterTarget:act, adjTarget:exp, rawTarget:e ? e.target : null, diff, status });
  }
  const order = ["OVER","UNDER","MISSING_MASTER","MISSING_TARGET","MATCH"];
  rows.sort((x, y) => { const d = order.indexOf(x.status) - order.indexOf(y.status); return d !== 0 ? d : Math.abs(y.diff||0) - Math.abs(x.diff||0); });
  const count = s => rows.filter(r => r.status === s).length;
  return { rows, summary:{ total:rows.length, over:count("OVER"), under:count("UNDER"), match:count("MATCH"), missing:count("MISSING_MASTER") + count("MISSING_TARGET") } };
}

// ── Model review ─────────────────────────────────────────────────────────────
function indexHoldingsByModel(holdings) {
  const by = {};
  for (const h of holdings) { if (h.modelId == null) continue; (by[h.modelId] = by[h.modelId] || []).push(h); }
  return by;
}
// Orion keeps a stale, empty aggregation under a model's plain name and names the live
// one "<name>.e". When "<name>.e" has accounts and the plain one doesn't, the .e is the
// model to review.
function supersededBy(model, modelsByName) {
  if (/\.e\s*$/i.test(model.name)) return null;
  const live = modelsByName[model.name + ".e"];
  return live && live.active && !model.active ? live : null;
}
function reviewModel(model, byModelId, lib, threshold, modelsByName) {
  const live = modelsByName ? supersededBy(model, modelsByName) : null;
  if (live) return { model, reason:`Superseded by '${live.name}' — the live aggregation (this one has no accounts); it's reviewed there` };
  const exp = expectedHoldings(model.name, lib);
  if (exp.reason) return { model, reason:exp.reason };
  const actual = (byModelId[model.id] || []).map(h => ({ ticker:h.ticker, name:h.name, target:h.target }));
  return { model, notes:exp.notes, ...compareWeights(exp.positions, actual, threshold, { rescale:true }) };
}

// ── Sleeve strategy review ───────────────────────────────────────────────────
// 1) Targets: each library-driven strategy's sleeve targets vs what the libraries say.
// 2) Look-through: the strategy's blended underlying holdings (each sleeve's target
//    x that model's holdings in the holding file) vs the library's whole-portfolio
//    ticker weights — an end-to-end check that sleeve targets and model holdings
//    together reproduce the library. Only the ETF-based sleeve can be checked ticker
//    by ticker (DI / bond / thematic sleeves are represented by proxy holdings).
function sleeveLookThrough(st, byAggId, lib, threshold) {
  const ctx = st.ctx;
  if (!ctx) return { skipped:"Not a library-driven strategy" };
  let sleeve, sleeveName, expected, label;
  if (ctx.type === "direct") {
    sleeve = st.comps.find(c => c.kind === "TACT_ETF"); sleeveName = sleeve && sleeve.aggName;
    const info = ctx.info;
    const tab = lib.tabs.find(t => t.source === "direct" && normAlnum(t.name).startsWith(normAlnum(info.size === "Large" ? "STP Direct Large" : "STP Direct Med")) && t.isTaxAware === info.isTA);
    const variant = tab && tab.variants.find(v => v.toLowerCase() === info.profile.toLowerCase());
    if (!tab || !variant) return { skipped:"The library column for this profile isn't available" };
    expected = tab.tickers
      .filter(t => t.ticker !== GLOBAL_EQUITY_TICKER && !(info.size === "Large" && /us investment grade/i.test(t.cls || "")))
      .map(t => ({ ticker:t.ticker, name:t.name, assetClass:toTitleCase(t.category || ""), target:(t.values[variant] || 0) * 100 }));
    label = "Tactical ETF sleeve vs the library's ETF weights";
  } else {
    sleeve = st.comps.find(c => c.aggName === ctx.ex.aggName); sleeveName = ctx.ex.aggName;
    const exp = expectedHoldings(sleeveName, lib);
    if (exp.reason) return { skipped:exp.reason };
    // The ex- model carries (1 - what it excludes) of the profile; the profile fills `remainder`% of the strategy.
    expected = exp.positions.map(p => ({ ...p, target:p.target * ctx.remainder / 100 }));
    label = "ex- model sleeve vs the library's weights";
  }
  if (!sleeve) return { skipped:"The strategy has no such sleeve" };
  const h = byAggId[sleeve.aggId];
  if (!h || !h.length) return { skipped:`'${sleeveName}' isn't in the holding file` };
  const actual = h.filter(x => x.target > 0).map(x => ({ ticker:x.ticker, name:x.name, target: x.target * sleeve.current / 100 }));
  return { label, sleeveName, ...compareWeights(expected, actual, threshold) };
}

function reviewSleeves(importRows, aggMap, lib, holdings, threshold) {
  const upd = computeSleeveUpdates(importRows, aggMap, lib.tabs);
  const byAggId = {};
  for (const h of (holdings || [])) { if (h.modelId != null) (byAggId[h.modelId] = byAggId[h.modelId] || []).push(h); }
  return upd.statuses.map(st => {
    if (st.status === "skipped") return { ...st, targetsOk:null, look:null };
    const diffs = st.comps.filter(c => Math.abs(c.expected - c.current) > 1e-9);
    return { ...st, targetsOk: diffs.length === 0, diffs, look: holdings ? sleeveLookThrough(st, byAggId, lib, threshold) : { skipped:"Upload the holdings file to run the look-through" } };
  });
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
function GeneralAudit() {
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

// ═══════════════════════════════════════════════════════════════════════════
// Review modes: STP models and sleeve strategies
// ═══════════════════════════════════════════════════════════════════════════
const MODES = [
  { key:"models",  label:"STP model review" },
  { key:"sleeves", label:"Sleeve strategy review" },
  { key:"general", label:"General audit" },
];
const SCOPE_LABEL = { not:"Model Name does not contain", has:"Model Name contains" };
const FAMILY_ORDER = ["STP", "STP Direct HH", "STP Tactical ETF", "Paragon/STP"];
function modelFamily(name) {
  if (/^\s*paragon\/stp/i.test(name)) return "Paragon/STP";
  if (/^\s*stp\s+direct\s+hh/i.test(name)) return "STP Direct HH";
  if (/^\s*stp\s+tactical\s+etf/i.test(name)) return "STP Tactical ETF";
  if (/^\s*stp\s+-\s/i.test(name)) return "STP";
  return "Other";
}

function downloadWorkbook(sheets, filename) {
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of sheets) {
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Note:"Nothing to report" }]);
    ws["!cols"] = Object.keys(rows[0] || { Note:"" }).map(k => ({ wch: Math.max(k.length + 2, 14) }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  }
  XLSX.writeFile(wb, filename);
}
const r3 = x => (x === null || x === undefined) ? "" : Math.round(x * 1000) / 1000;

// Everything the two review modes share: the holdings file, both libraries, the
// sleeve files and the tolerance. Uploaded once, used by both.
function useAuditWorkspace() {
  const [masterData, setMasterData] = useState(null); // { models, holdings, fileName, savedAt }
  const [masterFromCache, setMasterFromCache] = useState(false);
  const [storageLoading, setStorageLoading] = useState(true);
  const [savvy, setSavvy] = useState(null); // { fileName, tabs, stock }
  const [direct, setDirect] = useState(null); // { fileName, tabs }
  const [sleeveImport, setSleeveImport] = useState(null); // { fileName, headers, rows }
  const [aggMeta, setAggMeta] = useState(null); // { map, fileName, savedAt } — stored; shared with the Import Builder
  const [loading, setLoading] = useState({});
  const [errors, setErrors] = useState({});
  const [threshold, setThreshold] = useState(DEFAULT_THRESHOLD);
  const setBusy = (k, v) => setLoading(p => ({ ...p, [k]: v }));
  const setErr = (k, v) => setErrors(p => ({ ...p, [k]: v }));

  useEffect(() => {
    (async () => {
      try {
        const result = await window.storage.get(STORAGE_KEY);
        if (result?.value) {
          const data = JSON.parse(result.value);
          const modelsMap = Object.fromEntries((data.models || []).map(m => [m.id, m.name]));
          const holdings = data.holdings?.[0]?.t !== undefined ? expandHoldings(data.holdings, modelsMap) : data.holdings;
          setMasterData({ ...data, holdings }); setMasterFromCache(true);
        }
      } catch { /* nothing cached */ }
      try {
        const saved = await loadSleeveAggMap();
        if (saved && saved.map) setAggMeta({ map:saved.map, fileName:saved.fileName, savedAt:saved.savedAt });
      } catch { /* nothing stored */ }
      setStorageLoading(false);
    })();
  }, []);

  const run = (key, fn) => async files => {
    setBusy(key, true); setErr(key, null);
    try { await fn(files[0]); } catch (e) { setErr(key, e.message); }
    setBusy(key, false);
  };
  const handleMaster = useCallback(run("master", async file => {
    const wb = readXLSX(await readFileBuffer(file));
    const { models, holdings: slim } = parseMasterFile(wb);
    if (!models.length) throw new Error("No models found — is this the Model Holding Weighting export?");
    const modelsMap = Object.fromEntries(models.map(m => [m.id, m.name]));
    const data = { models, holdings: expandHoldings(slim, modelsMap), fileName:file.name, savedAt:new Date().toLocaleDateString() };
    setMasterData(data); setMasterFromCache(false);
    try { await window.storage.set(STORAGE_KEY, JSON.stringify({ models, holdings:slim, fileName:file.name, savedAt:data.savedAt })); } catch {}
  }), []);
  const handleSavvy = useCallback(run("savvy", async file => {
    const wb = readXLSX(await readFileBuffer(file));
    const tabs = parseLibraryTabs(wb, "savvy");
    if (!tabs.length) throw new Error("No model matrix tabs found (looked for a 'Ticker' column).");
    setSavvy({ fileName:file.name, tabs, stock:parseStockModels(wb) });
  }), []);
  const handleDirect = useCallback(run("direct", async file => {
    const wb = readXLSX(await readFileBuffer(file));
    const tabs = parseLibraryTabs(wb, "direct");
    if (!tabs.length) throw new Error("No model matrix tabs found in the STP Direct library (looked for a 'Ticker' column).");
    setDirect({ fileName:file.name, tabs });
  }), []);
  const handleSleeveImport = useCallback(run("sleeveImport", async file => {
    const parsed = parseSleeveImport(new Uint8Array(await readFileBuffer(file)));
    setSleeveImport({ fileName:file.name, headers:parsed.headers, rows:parsed.rows });
  }), []);
  const handleSleeveExport = useCallback(run("sleeveExport", async file => {
    const map = parseSleeveExport(new Uint8Array(await readFileBuffer(file)));
    const meta = { map, fileName:file.name, savedAt:new Date().toISOString() };
    setAggMeta(meta);
    saveSleeveAggMap({ map, savedAt:meta.savedAt, fileName:file.name });
  }), []);
  const clearMaster = useCallback(async () => {
    setMasterData(null); setMasterFromCache(false);
    try { await window.storage.delete(STORAGE_KEY); } catch {}
  }, []);

  const lib = useMemo(() => (savvy || direct)
    ? { tabs:[...(savvy ? savvy.tabs : []), ...(direct ? direct.tabs : [])], stock:savvy ? savvy.stock : {} }
    : null, [savvy, direct]);

  return { masterData, masterFromCache, storageLoading, savvy, direct, sleeveImport, aggMeta, loading, errors, threshold, setThreshold,
           lib, handleMaster, handleSavvy, handleDirect, handleSleeveImport, handleSleeveExport, clearMaster, setSavvy, setDirect, setSleeveImport };
}

// ── The top-level tool: a mode bar, then one of the three views ──────────────
export default function ModelAuditTool() {
  const [mode, setMode] = useState("models");
  const ws = useAuditWorkspace();
  return (
    <div style={{ fontFamily:"'Inter', system-ui, sans-serif", background:"#f1f5f9", minHeight:"100vh" }}>
      <div style={{ background:"white", borderBottom:"1px solid #e2e8f0", padding:"8px 28px", display:"flex", gap:6, alignItems:"center" }}>
        {MODES.map(m => (
          <button key={m.key} onClick={() => setMode(m.key)}
            style={{ padding:"6px 14px", borderRadius:8, border:"none", fontSize:12, fontWeight:600, cursor:"pointer",
              background:mode === m.key ? "#0f172a" : "transparent", color:mode === m.key ? "white" : "#64748b" }}>
            {m.label}
          </button>
        ))}
      </div>
      {mode === "general" ? <GeneralAudit /> : <ReviewShell mode={mode} ws={ws} />}
    </div>
  );
}

function ReviewShell({ mode, ws }) {
  const dMaster = useDrop(ws.handleMaster), dSavvy = useDrop(ws.handleSavvy), dDirect = useDrop(ws.handleDirect);
  const dImport = useDrop(ws.handleSleeveImport), dExport = useDrop(ws.handleSleeveExport);
  const sleeves = mode === "sleeves";
  const { masterData, savvy, direct, sleeveImport, aggMeta, loading, errors } = ws;
  return (
    <div>
      <div style={{ background:"#0f172a", padding:"16px 28px", display:"flex", alignItems:"center", gap:14 }}>
        <div style={{ width:34, height:34, background:"#3b82f6", borderRadius:8, display:"flex", alignItems:"center", justifyContent:"center" }}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2"><path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 1-2-2V9m0 0h18"/></svg>
        </div>
        <div>
          <div style={{ color:"white", fontWeight:700, fontSize:16 }}>{sleeves ? "Sleeve Strategy Review" : "STP Model Review"}</div>
          <div style={{ color:"#94a3b8", fontSize:11 }}>{sleeves ? "Sleeve targets and underlying holdings vs the model libraries" : "Underlying holdings in Orion vs the model libraries"}</div>
        </div>
      </div>

      <div style={{ padding:"20px 28px", maxWidth:1400, margin:"0 auto" }}>
        <div style={{ display:"grid", gridTemplateColumns:sleeves ? "repeat(3, 1fr)" : "repeat(4, 1fr)", gap:14, marginBottom:16 }}>
          <DropZone label="Model Holding Weighting" sublabel="Orion export — underlying ticker allocations for every model"
            accent="#3b82f6" loading={loading.master || ws.storageLoading} error={errors.master} dragging={dMaster.dragging}
            dropProps={dMaster} onFiles={ws.handleMaster} hint="Drop or click · .xlsx">
            {masterData && <FileChip name={masterData.fileName}
              sub={`${masterData.models.length} models · ${masterData.holdings.length} rows · saved ${masterData.savedAt}`} color="#3b82f6"
              badge={ws.masterFromCache ? { label:"Cached — re-upload after pushing updates", color:"#92400e", bg:"#fffbeb" } : { label:"Updated", color:"#2563eb", bg:"#eff6ff" }}
              onRemove={ws.clearMaster} />}
          </DropZone>
          <DropZone label="Savvy model library" sublabel="STP and Paragon/STP targets"
            accent="#0aa89c" loading={loading.savvy} error={errors.savvy} dragging={dSavvy.dragging}
            dropProps={dSavvy} onFiles={ws.handleSavvy} hint="Drop or click · .xlsx">
            {savvy && <FileChip name={savvy.fileName} sub={`${savvy.tabs.length} tabs · ${Object.keys(savvy.stock).length} stock models`} color="#0aa89c" onRemove={() => ws.setSavvy(null)} />}
          </DropZone>
          <DropZone label="STP Direct library" sublabel="STP Direct HH and Tactical ETF targets"
            accent="#7c3aed" loading={loading.direct} error={errors.direct} dragging={dDirect.dragging}
            dropProps={dDirect} onFiles={ws.handleDirect} hint="Drop or click · .xlsx">
            {direct && <FileChip name={direct.fileName} sub={`${direct.tabs.length} tabs`} color="#7c3aed" onRemove={() => ws.setDirect(null)} />}
          </DropZone>
          {sleeves && (
            <DropZone label="Sleeve Strategy Import" sublabel="The sleeve strategies to review"
              accent="#f59e0b" loading={loading.sleeveImport} error={errors.sleeveImport} dragging={dImport.dragging}
              dropProps={dImport} onFiles={ws.handleSleeveImport} hint="Drop or click · .xlsx">
              {sleeveImport && <FileChip name={sleeveImport.fileName}
                sub={`${new Set(sleeveImport.rows.map(r => r["Sleeve Strategy ID"])).size} strategies · ${sleeveImport.rows.length} rows`} color="#f59e0b" onRemove={() => ws.setSleeveImport(null)} />}
            </DropZone>
          )}
          {sleeves && (
            <DropZone label="Sleeve Strategy Export" sublabel="Model Agg ID → name — stored between sessions"
              accent="#f59e0b" loading={loading.sleeveExport || ws.storageLoading} error={errors.sleeveExport} dragging={dExport.dragging}
              dropProps={dExport} onFiles={ws.handleSleeveExport} hint="Drop or click · .xlsx">
              {aggMeta && <FileChip name={aggMeta.fileName || "stored export"}
                sub={`${Object.keys(aggMeta.map).length} model aggregates · saved ${new Date(aggMeta.savedAt).toLocaleDateString()}`} color="#f59e0b"
                badge={{ label:"Stored", color:"#16a34a", bg:"#f0fdf4" }} />}
            </DropZone>
          )}
          <div style={{ background:"white", borderRadius:12, padding:16, boxShadow:"0 1px 3px rgba(0,0,0,0.06)", borderTop:"3px solid #0ea5e9" }}>
            <div style={{ fontWeight:700, fontSize:13, color:"#0f172a", marginBottom:12 }}>Audit Settings</div>
            <label style={lbl}>Tolerance (±%)</label>
            <input type="number" step="0.01" min="0" value={ws.threshold}
              onChange={e => ws.setThreshold(parseFloat(e.target.value) || 0)}
              style={{ width:"100%", padding:"7px 10px", borderRadius:7, border:"1.5px solid #e2e8f0", fontSize:13, color:"#0f172a", boxSizing:"border-box" }} />
            <div style={{ marginTop:8, fontSize:11, color:"#94a3b8", lineHeight:1.5 }}>
              A difference up to this many points counts as on target. The libraries carry weights to 0.1%, so differences under 0.05 are rounding.
            </div>
          </div>
        </div>
        {sleeves ? <SleeveReviewView ws={ws} /> : <ModelReviewView ws={ws} />}
      </div>
    </div>
  );
}

// A ticker comparison table (used by both reviews)
function TickerTable({ rows, actualLabel, expectedLabel }) {
  if (!rows.length) return <div style={{ padding:24, textAlign:"center", color:"#94a3b8", fontSize:13 }}>No holdings match this filter.</div>;
  return (
    <div style={{ overflowX:"auto" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
        <thead>
          <tr style={{ background:"#f8fafc" }}>
            {["Status","Ticker","Security","Asset Class",actualLabel,expectedLabel,"Diff"].map(h => (
              <th key={h} style={{ padding:"9px 14px", textAlign:[actualLabel,expectedLabel,"Diff"].includes(h) ? "right" : "left", fontSize:10, fontWeight:600, color:"#64748b", textTransform:"uppercase", letterSpacing:"0.05em", whiteSpace:"nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => {
            const s = STATUS[row.status];
            return (
              <tr key={row.ticker + i} style={{ background:i % 2 === 0 ? "white" : "#fafbfc", borderBottom:"1px solid #f1f5f9" }}>
                <td style={{ padding:"8px 14px", whiteSpace:"nowrap" }}>
                  <span style={{ display:"inline-flex", alignItems:"center", gap:5, background:s.bg, color:s.color, padding:"2px 8px", borderRadius:10, fontSize:11, fontWeight:600 }}>
                    <span style={{ width:5, height:5, borderRadius:"50%", background:s.dot, flexShrink:0 }} />{s.label}
                  </span>
                </td>
                <td style={{ padding:"8px 14px", fontWeight:600, color:"#0f172a", fontFamily:"monospace", fontSize:12 }}>{row.ticker}</td>
                <td style={{ padding:"8px 14px", color:"#334155", maxWidth:220, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{row.name}</td>
                <td style={{ padding:"8px 14px", color:"#64748b", fontSize:12, whiteSpace:"nowrap" }}>{row.assetClass}</td>
                <td style={{ padding:"8px 14px", textAlign:"right", fontVariantNumeric:"tabular-nums", color:"#0f172a" }}>
                  {row.masterTarget !== null ? `${row.masterTarget.toFixed(2)}%` : <span style={{ color:"#cbd5e1" }}>—</span>}
                </td>
                <td style={{ padding:"8px 14px", textAlign:"right", fontVariantNumeric:"tabular-nums", color:"#0f172a" }}>
                  {row.adjTarget !== null ? `${row.adjTarget.toFixed(2)}%` : <span style={{ color:"#cbd5e1" }}>—</span>}
                </td>
                <td style={{ padding:"8px 14px", textAlign:"right", fontWeight:600, fontVariantNumeric:"tabular-nums", color:row.diff === null ? "#cbd5e1" : row.diff > 0 ? "#dc2626" : row.diff < 0 ? "#d97706" : "#16a34a" }}>
                  {row.diff === null ? "—" : `${row.diff > 0 ? "+" : ""}${row.diff.toFixed(2)}%`}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function StatCard({ label, value, color }) {
  return (
    <div style={{ background:"white", borderRadius:10, padding:"12px 16px", boxShadow:"0 1px 3px rgba(0,0,0,0.06)" }}>
      <div style={{ fontSize:22, fontWeight:700, color }}>{value}</div>
      <div style={{ fontSize:11, color:"#64748b" }}>{label}</div>
    </div>
  );
}

function NotReviewed({ items }) {
  // items: [{ name, reason }] grouped by reason
  if (!items.length) return null;
  const grouped = items.reduce((g, x) => { (g[x.reason] = g[x.reason] || []).push(x.name); return g; }, {});
  return (
    <details style={{ marginTop:14, background:"white", borderRadius:12, boxShadow:"0 1px 3px rgba(0,0,0,0.06)", padding:"12px 20px" }}>
      <summary style={{ cursor:"pointer", fontSize:13, fontWeight:600, color:"#64748b" }}>Not reviewed — {items.length}</summary>
      <ul style={{ margin:"10px 0 0", paddingLeft:18, fontSize:12, color:"#64748b", lineHeight:1.7 }}>
        {Object.entries(grouped).map(([why, names], i) => (
          <li key={i} style={{ marginBottom:4 }}><strong style={{ color:"#334155" }}>{why}</strong> ({names.length}): {names.join("; ")}</li>
        ))}
      </ul>
    </details>
  );
}

function Prompt({ text }) { return <EmptyState text={text} />; }

// ── STP model review ─────────────────────────────────────────────────────────
function ModelReviewView({ ws }) {
  const { masterData, lib, threshold } = ws;
  const [filters, setFilters] = useState(DEFAULT_SCOPE_FILTERS);
  const [issuesOnly, setIssuesOnly] = useState(true);
  const [family, setFamily] = useState("ALL");
  const [expanded, setExpanded] = useState({});
  const [rowFilter, setRowFilter] = useState({});

  const review = useMemo(() => {
    if (!masterData || !lib) return null;
    const by = indexHoldingsByModel(masterData.holdings);
    const byName = Object.fromEntries(masterData.models.map(m => [m.name, m]));
    const scoped = masterData.models.filter(m => inReviewScope(m.name, filters));
    const results = scoped.map(m => reviewModel(m, by, lib, threshold, byName));
    const withIssues = r => !r.reason && (r.summary.over + r.summary.under + r.summary.missing) > 0;
    return { scoped:scoped.length, results, withIssues };
  }, [masterData, lib, threshold, filters]);

  const setFilter = (id, patch) => setFilters(fs => fs.map(f => f.id === id ? { ...f, ...patch } : f));

  const scopePanel = (
    <div style={{ background:"white", borderRadius:12, padding:"14px 18px", boxShadow:"0 1px 3px rgba(0,0,0,0.06)", marginBottom:14 }}>
      <div style={{ fontWeight:700, fontSize:13, color:"#0f172a", marginBottom:2 }}>Review scope</div>
      <div style={{ fontSize:11, color:"#94a3b8", marginBottom:10 }}>The same filters you use in Orion to pick the models that need reviewing — the holdings file lists every model.</div>
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {filters.map(f => (
          <label key={f.id} style={{ display:"flex", alignItems:"center", gap:10, fontSize:13, color:f.on ? "#0f172a" : "#94a3b8" }}>
            <input type="checkbox" checked={f.on} onChange={() => setFilter(f.id, { on:!f.on })} style={{ width:15, height:15, accentColor:"#3b82f6" }} />
            <span style={{ minWidth:200 }}>{SCOPE_LABEL[f.mode]}</span>
            <input type="text" value={f.term} onChange={e => setFilter(f.id, { term:e.target.value })}
              style={{ padding:"4px 8px", borderRadius:6, border:"1.5px solid #e2e8f0", fontSize:12, width:170, color:"#0f172a" }} />
          </label>
        ))}
      </div>
    </div>
  );

  if (ws.storageLoading) return <>{scopePanel}<Prompt text="Loading saved holdings file…" /></>;
  if (!masterData) return <>{scopePanel}<Prompt text="Upload the Model Holding Weighting file to begin." /></>;
  if (!lib) return <>{scopePanel}<Prompt text="Upload the Savvy model library and the STP Direct library to review the models." /></>;

  const { results, withIssues } = review;
  const reviewed = results.filter(r => !r.reason);
  const issues = reviewed.filter(withIssues);
  const notReviewed = results.filter(r => r.reason).map(r => ({ name:r.model.name, reason:r.reason }));
  const inFamily = r => family === "ALL" || modelFamily(r.model.name) === family;
  const shown = reviewed.filter(inFamily).filter(r => !issuesOnly || withIssues(r))
    .sort((a, b) => (withIssues(b) - withIssues(a)) || (b.summary.over + b.summary.under + b.summary.missing) - (a.summary.over + a.summary.under + a.summary.missing));
  const familiesPresent = [...new Set(reviewed.map(r => modelFamily(r.model.name)))].sort((a, b) => FAMILY_ORDER.indexOf(a) - FAMILY_ORDER.indexOf(b));
  const allOpen = shown.length > 0 && shown.every(r => expanded[r.model.id]);

  const exportReport = () => {
    const disc = [], summary = [];
    for (const r of results) {
      if (r.reason) { summary.push({ Model:r.model.name, Result:"Not reviewed", Detail:r.reason }); continue; }
      const bad = r.rows.filter(x => x.status !== "MATCH");
      summary.push({ Model:r.model.name, Result:bad.length ? "Discrepancies" : "On target", Detail:`${r.notes.join(" · ")}`, "Positions":r.summary.total, Over:r.summary.over, Under:r.summary.under, Missing:r.summary.missing });
      for (const x of bad) disc.push({ Model:r.model.name, Ticker:x.ticker, Security:x.name, "Asset class":x.assetClass, "In Orion %":r3(x.masterTarget), "Expected %":r3(x.adjTarget), "Diff (pp)":r3(x.diff), Status:STATUS[x.status].label });
    }
    downloadWorkbook([["Discrepancies", disc], ["Models", summary]], `STP_Model_Review_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <>
      {scopePanel}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(5, 1fr)", gap:12, marginBottom:14 }}>
        <StatCard label="Models in scope" value={review.scoped} color="#0f172a" />
        <StatCard label="Reviewed" value={reviewed.length} color="#3b82f6" />
        <StatCard label="On target" value={reviewed.length - issues.length} color="#16a34a" />
        <StatCard label="With discrepancies" value={issues.length} color={issues.length ? "#dc2626" : "#94a3b8"} />
        <StatCard label="Not reviewed" value={notReviewed.length} color="#94a3b8" />
      </div>

      <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap", marginBottom:12 }}>
        {["ALL", ...familiesPresent].map(f => (
          <button key={f} onClick={() => setFamily(f)}
            style={{ padding:"5px 12px", borderRadius:20, border:`1.5px solid ${family === f ? "#3b82f6" : "#e2e8f0"}`, background:family === f ? "#eff6ff" : "white", color:family === f ? "#1d4ed8" : "#64748b", fontSize:12, fontWeight:600, cursor:"pointer" }}>
            {f === "ALL" ? "All families" : f}
          </button>
        ))}
        <label style={{ marginLeft:12, display:"flex", alignItems:"center", gap:6, fontSize:12, color:"#334155", cursor:"pointer" }}>
          <input type="checkbox" checked={issuesOnly} onChange={() => setIssuesOnly(v => !v)} style={{ accentColor:"#3b82f6" }} /> Discrepancies only
        </label>
        <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
          <button onClick={() => { const e = {}; shown.forEach(r => { e[r.model.id] = !allOpen; }); setExpanded(e); }}
            style={{ padding:"6px 14px", borderRadius:8, background:"#334155", border:"none", color:"white", fontSize:12, fontWeight:600, cursor:"pointer" }}>{allOpen ? "⊟ Collapse All" : "⊞ Expand All"}</button>
          <button onClick={exportReport}
            style={{ padding:"6px 14px", borderRadius:8, background:"#1d4ed8", border:"none", color:"white", fontSize:12, fontWeight:600, cursor:"pointer" }}>Export report ↓</button>
        </div>
      </div>

      {shown.length === 0 && <EmptyState text={issues.length === 0 ? "Every reviewed model is on target." : "No models match this view."} />}
      {shown.map(r => {
        const isOpen = expanded[r.model.id] === true, filt = rowFilter[r.model.id] || "DISCREPANCY";
        const rows = filt === "ALL" ? r.rows : filt === "DISCREPANCY" ? r.rows.filter(x => x.status !== "MATCH") : r.rows.filter(x => x.status === filt);
        const bad = r.summary.over + r.summary.under + r.summary.missing;
        return (
          <div key={r.model.id} style={{ background:"white", borderRadius:12, boxShadow:"0 1px 3px rgba(0,0,0,0.06)", marginBottom:10, overflow:"hidden" }}>
            <button onClick={() => setExpanded(p => ({ ...p, [r.model.id]: !isOpen }))}
              style={{ width:"100%", background:"none", border:"none", cursor:"pointer", padding:"13px 20px", display:"flex", alignItems:"center", gap:12, textAlign:"left" }}>
              <span style={{ fontSize:13, transform:isOpen ? "rotate(90deg)" : "none", transition:"transform 0.2s", color:"#94a3b8", display:"inline-block" }}>▶</span>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontWeight:700, fontSize:14, color:"#0f172a" }}>{r.model.name}</div>
                <div style={{ fontSize:11, color:"#94a3b8", marginTop:1 }}>Library: <span style={{ color:"#64748b" }}>{r.notes.join(" · ")}</span></div>
              </div>
              <div style={{ display:"flex", gap:6, flexShrink:0 }}>
                {r.summary.over > 0 && <Pill label={`${r.summary.over} over`} color="#dc2626" bg="#fef2f2" />}
                {r.summary.under > 0 && <Pill label={`${r.summary.under} under`} color="#d97706" bg="#fffbeb" />}
                {r.summary.missing > 0 && <Pill label={`${r.summary.missing} missing`} color="#7c3aed" bg="#f5f3ff" />}
                {bad === 0 && <Pill label="All on target" color="#16a34a" bg="#f0fdf4" />}
                <Pill label={`${r.summary.total} positions`} color="#64748b" bg="#f8fafc" />
              </div>
            </button>
            {isOpen && (
              <>
                <div style={{ padding:"0 20px 10px", display:"flex", alignItems:"center", gap:8, flexWrap:"wrap", borderBottom:"1px solid #f1f5f9" }}>
                  {[["DISCREPANCY", "Discrepancies"], ["ALL", "All"]].map(([f, label]) => (
                    <button key={f} onClick={() => setRowFilter(p => ({ ...p, [r.model.id]: f }))}
                      style={{ padding:"4px 10px", borderRadius:20, border:`1.5px solid ${filt === f ? "#3b82f6" : "#e2e8f0"}`, background:filt === f ? "#eff6ff" : "white", color:filt === f ? "#1d4ed8" : "#64748b", fontSize:11, fontWeight:600, cursor:"pointer" }}>{label}</button>
                  ))}
                  <span style={{ marginLeft:"auto", fontSize:11, color:"#94a3b8" }}>±{threshold}% tolerance</span>
                </div>
                <TickerTable rows={rows} actualLabel="In Orion %" expectedLabel="Expected %" />
              </>
            )}
          </div>
        );
      })}
      <NotReviewed items={notReviewed} />
    </>
  );
}

// ── Sleeve strategy review ───────────────────────────────────────────────────
function SleeveReviewView({ ws }) {
  const { masterData, lib, threshold, sleeveImport, aggMeta } = ws;
  const [issuesOnly, setIssuesOnly] = useState(true);
  const [expanded, setExpanded] = useState({});

  const results = useMemo(() => (sleeveImport && aggMeta && lib)
    ? reviewSleeves(sleeveImport.rows, aggMeta.map, lib, masterData ? masterData.holdings : null, threshold)
    : null, [sleeveImport, aggMeta, lib, masterData, threshold]);

  if (!sleeveImport) return <Prompt text="Upload the Sleeve Strategy Import to review it." />;
  if (!aggMeta) return <Prompt text="Upload the Sleeve Strategy Export once — it supplies the model names for the sleeves, and is remembered." />;
  if (!lib) return <Prompt text="Upload the Savvy model library and the STP Direct library to review the sleeve strategies." />;

  const reviewed = results.filter(s => s.status !== "skipped");
  const lookBad = s => s.look && s.look.summary && (s.look.summary.over + s.look.summary.under + s.look.summary.missing) > 0;
  const hasIssue = s => !s.targetsOk || lookBad(s);
  const issues = reviewed.filter(hasIssue);
  const notReviewed = results.filter(s => s.status === "skipped").map(s => ({ name:s.name, reason:`${s.group} — ${s.reason}` }));
  const shown = reviewed.filter(s => !issuesOnly || hasIssue(s));
  shown.sort((a, b) => (hasIssue(b) - hasIssue(a)) || String(a.name).localeCompare(String(b.name)));
  const allOpen = shown.length > 0 && shown.every(s => expanded[s.id]);

  const exportReport = () => {
    const targets = [], look = [], summary = [];
    for (const s of results) {
      if (s.status === "skipped") { summary.push({ Strategy:s.name, "Strategy ID":s.id, Result:"Not reviewed", Detail:`${s.group} — ${s.reason}` }); continue; }
      summary.push({ Strategy:s.name, "Strategy ID":s.id, Family:s.family,
        "Sleeve targets":s.targetsOk ? "Match the libraries" : "Differ from the libraries",
        "Look-through":s.look.skipped ? `Not reviewed — ${s.look.skipped}` : lookBad(s) ? "Discrepancies" : "On target" });
      for (const c of s.comps) targets.push({ Strategy:s.name, "Strategy ID":s.id, Sleeve:c.aggName, "Current %":r3(c.current), "Libraries say %":r3(c.expected), "Diff (pp)":r3(c.current - c.expected), Status:Math.abs(c.expected - c.current) > 1e-9 ? "Differs" : "Match" });
      if (s.look.rows) for (const x of s.look.rows.filter(r => r.status !== "MATCH")) look.push({ Strategy:s.name, "Strategy ID":s.id, Sleeve:s.look.sleeveName, Ticker:x.ticker, Security:x.name, "Blended in Orion %":r3(x.masterTarget), "Library %":r3(x.adjTarget), "Diff (pp)":r3(x.diff), Status:STATUS[x.status].label });
    }
    downloadWorkbook([["Strategies", summary], ["Sleeve targets", targets], ["Look-through discrepancies", look]], `Sleeve_Strategy_Review_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(5, 1fr)", gap:12, marginBottom:14 }}>
        <StatCard label="Strategies in file" value={results.length} color="#0f172a" />
        <StatCard label="Reviewed" value={reviewed.length} color="#3b82f6" />
        <StatCard label="On target" value={reviewed.length - issues.length} color="#16a34a" />
        <StatCard label="Need attention" value={issues.length} color={issues.length ? "#dc2626" : "#94a3b8"} />
        <StatCard label="Not reviewed" value={notReviewed.length} color="#94a3b8" />
      </div>
      {!masterData && (
        <div style={{ background:"#fffbeb", border:"1px solid #fde68a", borderRadius:10, padding:"10px 14px", fontSize:12, color:"#92400e", marginBottom:12 }}>
          The holdings file isn't loaded, so only the sleeve targets are checked. Upload it to also check the underlying holdings.
        </div>
      )}
      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
        <label style={{ display:"flex", alignItems:"center", gap:6, fontSize:12, color:"#334155", cursor:"pointer" }}>
          <input type="checkbox" checked={issuesOnly} onChange={() => setIssuesOnly(v => !v)} style={{ accentColor:"#3b82f6" }} /> Strategies needing attention only
        </label>
        <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
          <button onClick={() => { const e = {}; shown.forEach(s => { e[s.id] = !allOpen; }); setExpanded(e); }}
            style={{ padding:"6px 14px", borderRadius:8, background:"#334155", border:"none", color:"white", fontSize:12, fontWeight:600, cursor:"pointer" }}>{allOpen ? "⊟ Collapse All" : "⊞ Expand All"}</button>
          <button onClick={exportReport}
            style={{ padding:"6px 14px", borderRadius:8, background:"#1d4ed8", border:"none", color:"white", fontSize:12, fontWeight:600, cursor:"pointer" }}>Export report ↓</button>
        </div>
      </div>

      {shown.length === 0 && <EmptyState text={issues.length === 0 ? "Every reviewed strategy is on target." : "No strategies match this view."} />}
      {shown.map(s => {
        const isOpen = expanded[s.id] === true;
        const lookSkipped = s.look && s.look.skipped, nBad = s.look && s.look.summary ? s.look.summary.over + s.look.summary.under + s.look.summary.missing : 0;
        return (
          <div key={s.id} style={{ background:"white", borderRadius:12, boxShadow:"0 1px 3px rgba(0,0,0,0.06)", marginBottom:10, overflow:"hidden" }}>
            <button onClick={() => setExpanded(p => ({ ...p, [s.id]: !isOpen }))}
              style={{ width:"100%", background:"none", border:"none", cursor:"pointer", padding:"13px 20px", display:"flex", alignItems:"center", gap:12, textAlign:"left" }}>
              <span style={{ fontSize:13, transform:isOpen ? "rotate(90deg)" : "none", transition:"transform 0.2s", color:"#94a3b8", display:"inline-block" }}>▶</span>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontWeight:700, fontSize:14, color:"#0f172a" }}>{s.name} <span style={{ fontWeight:400, color:"#94a3b8", fontSize:12 }}>(ID {s.id})</span></div>
                <div style={{ fontSize:11, color:"#94a3b8", marginTop:1 }}>{s.family}</div>
              </div>
              <div style={{ display:"flex", gap:6, flexShrink:0 }}>
                {s.targetsOk ? <Pill label="Targets match" color="#16a34a" bg="#f0fdf4" /> : <Pill label={`${s.diffs.length} target${s.diffs.length === 1 ? "" : "s"} differ`} color="#dc2626" bg="#fef2f2" />}
                {lookSkipped ? <Pill label="Holdings not checked" color="#64748b" bg="#f8fafc" />
                  : nBad === 0 ? <Pill label="Holdings on target" color="#16a34a" bg="#f0fdf4" />
                  : <Pill label={`${nBad} holding${nBad === 1 ? "" : "s"} off`} color="#d97706" bg="#fffbeb" />}
              </div>
            </button>
            {isOpen && (
              <div style={{ padding:"0 20px 16px" }}>
                <div style={{ fontSize:11, fontWeight:600, color:"#64748b", textTransform:"uppercase", letterSpacing:"0.05em", margin:"4px 0 6px" }}>Sleeve targets</div>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13, marginBottom:14 }}>
                  <thead><tr style={{ background:"#f8fafc" }}>
                    {["Sleeve","Current %","Libraries say %","Diff"].map(h => <th key={h} style={{ padding:"8px 12px", textAlign:h === "Sleeve" ? "left" : "right", fontSize:10, fontWeight:600, color:"#64748b", textTransform:"uppercase", letterSpacing:"0.05em" }}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {s.comps.map((c, i) => {
                      const d = c.current - c.expected, off = Math.abs(d) > 1e-9;
                      return (
                        <tr key={i} style={{ borderBottom:"1px solid #f1f5f9", background:off ? "#fffbeb" : "transparent" }}>
                          <td style={{ padding:"7px 12px", color:"#334155" }}>{c.aggName}</td>
                          <td style={{ padding:"7px 12px", textAlign:"right", fontVariantNumeric:"tabular-nums" }}>{c.current}%</td>
                          <td style={{ padding:"7px 12px", textAlign:"right", fontVariantNumeric:"tabular-nums" }}>{c.expected}%</td>
                          <td style={{ padding:"7px 12px", textAlign:"right", fontWeight:600, fontVariantNumeric:"tabular-nums", color:off ? (d > 0 ? "#dc2626" : "#d97706") : "#16a34a" }}>{off ? `${d > 0 ? "+" : ""}${d.toFixed(2)}` : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <div style={{ fontSize:11, fontWeight:600, color:"#64748b", textTransform:"uppercase", letterSpacing:"0.05em", margin:"4px 0 2px" }}>Underlying holdings</div>
                {lookSkipped
                  ? <div style={{ padding:"10px 0", fontSize:12, color:"#94a3b8" }}>Not checked — {s.look.skipped}.</div>
                  : <>
                      <div style={{ fontSize:11, color:"#94a3b8", marginBottom:6 }}>{s.look.label} · {s.look.sleeveName} · ±{threshold}% tolerance</div>
                      <div style={{ border:"1px solid #f1f5f9", borderRadius:8, overflow:"hidden" }}>
                        <TickerTable rows={s.look.rows.filter(x => x.status !== "MATCH" || nBad === 0)} actualLabel="Blended in Orion %" expectedLabel="Library %" />
                      </div>
                    </>}
              </div>
            )}
          </div>
        );
      })}
      <NotReviewed items={notReviewed} />
    </>
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
