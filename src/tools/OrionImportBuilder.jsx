import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

// ── Constants ────────────────────────────────────────────────────────────────

const TEMPLATE_COLS = [
  "Model ID","* Model Name","Model Description","* Security Set ID",
  "* Security Set SubModel Name","* Security Set Target %","Security Set Band/Range",
  "Security Set Upper %","Security Set Lower %","* Dynamic","Management Style",
  "Sleeve Suffix","Team ID","* Name Space","Tags",
  "Category SubModel Name","Category Asset Class Type","Category Team ID","Category Namespace",
  "Category Target %","Category Band/Range","Category Upper %","Category Lower %",
  "Class SubModel Name","Class Asset Class Type","Class Team ID","Class Namespace",
  "Class Target %","Class Band/Range","Class Upper %","Class Lower %",
  "Subclass SubModel Name","Subclass Asset Class Type","Subclass Team ID","Subclass Namespace",
  "Subclass Target %","Subclass Band/Range","Subclass Upper %","Subclass Lower %"
];

const COL_ALIASES = {
  "model id":"Model ID","* model name":"* Model Name","model name":"* Model Name",
  "model description":"Model Description","* security set id":"* Security Set ID",
  "security set id":"* Security Set ID","* security set submodel name":"* Security Set SubModel Name",
  "security set submodel name":"* Security Set SubModel Name",
  "* security set target %":"* Security Set Target %","security set target %":"* Security Set Target %",
  "security set band/range":"Security Set Band/Range","security set upper %":"Security Set Upper %",
  "security set lower %":"Security Set Lower %","* dynamic":"* Dynamic","dynamic":"* Dynamic",
  "management style":"Management Style","sleeve suffix":"Sleeve Suffix","team id":"Team ID",
  "* name space":"* Name Space","name space":"* Name Space","namespace":"* Name Space","tags":"Tags",
  "category submodel name":"Category SubModel Name","category asset class type":"Category Asset Class Type",
  "category team id":"Category Team ID","category namespace":"Category Namespace",
  "category target %":"Category Target %","category band/range":"Category Band/Range",
  "category upper %":"Category Upper %","category lower %":"Category Lower %",
  "class submodel name":"Class SubModel Name","class asset class type":"Class Asset Class Type",
  "class team id":"Class Team ID","class namespace":"Class Namespace",
  "class target %":"Class Target %","class band/range":"Class Band/Range",
  "class upper %":"Class Upper %","class lower %":"Class Lower %",
  "subclass submodel name":"Subclass SubModel Name","subclass asset class type":"Subclass Asset Class Type",
  "subclass team id":"Subclass Team ID","subclass namespace":"Subclass Namespace",
  "subclass target %":"Subclass Target %","subclass band/range":"Subclass Band/Range",
  "subclass upper %":"Subclass Upper %","subclass lower %":"Subclass Lower %",
};

// Node type colors matching Orion's teal/blue/green/yellow palette
const NODE_COLORS = {
  root:     { fill:"#0dd3c5", stroke:"#0aa89c", text:"#003d3a" },
  category: { fill:"#1a6fb5", stroke:"#145490", text:"#e0f0ff" },
  class:    { fill:"#c8b400", stroke:"#a09000", text:"#2d2600" },
  ss:       { fill:"#1a56db", stroke:"#1240a8", text:"#e8f0fe" },
};

// ── Parsing ──────────────────────────────────────────────────────────────────

function normalizeKey(k) { return (k||"").toString().trim().toLowerCase(); }

function parseWorkbookRows(buffer) {
  const wb = XLSX.read(buffer, { type:"array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  let headerIdx = -1;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const norm = (raw[i]||[]).map(c => normalizeKey(c));
    if (norm.includes("model id")||norm.includes("* model name")||norm.includes("model name")) { headerIdx=i; break; }
  }
  if (headerIdx===-1) throw new Error("Could not find a header row with 'Model ID' or 'Model Name'.");
  const headers = raw[headerIdx].map(c => COL_ALIASES[normalizeKey(c)] || (c?c.toString().trim():null));
  const rows = [];
  for (let i=headerIdx+1; i<raw.length; i++) {
    const row = raw[i];
    if (!row||row.every(c=>c===null||c==="")) continue;
    const obj = {};
    headers.forEach((h,idx)=>{ if(h) obj[h]=row[idx]!==undefined?row[idx]:null; });
    if (!obj["* Model Name"]&&!obj["* Security Set SubModel Name"]) continue;
    rows.push(obj);
  }
  // Preserve the file's own header order (deduped, nulls dropped) so a
  // round-tripped export doesn't drop columns the fixed template doesn't know about.
  const orderedHeaders = [];
  headers.forEach(h => { if (h && !orderedHeaders.includes(h)) orderedHeaders.push(h); });
  return { headers: orderedHeaders, rows };
}

function parseExcel(buffer) {
  return parseWorkbookRows(buffer).rows;
}

// ── Tree Builder ─────────────────────────────────────────────────────────────
// Each node stores TWO sets of band values:
//   own:   band/upper/lower  → written to that node's own column (Category Band/Range etc.)
//   child: childBand/childUpper/childLower → written to the CHILD tier's band columns
//
// Editing "Model" → you set each Category's own target+band (category.band → Category Band/Range)
// Editing "Category" → you set each Class's target+band (class.band → Class Band/Range)
// Editing "Class" → you set SS tolerance (class.childBand → Security Set Band/Range)

function buildTree(modelName, rows) {
  const root = { id:"root", type:"root", label:modelName, target:100, children:[] };
  const catMap = {};
  const classMap = {};

  rows.forEach((r, ri) => {
    const catName   = r["Category SubModel Name"];
    const className = r["Class SubModel Name"];
    const ssName    = r["* Security Set SubModel Name"];

    // Category — owns its Category Band/Range/Upper/Lower columns
    let catNode = catMap[catName];
    if (catName && !catNode) {
      catNode = {
        id: `cat_${catName}`, type:"category", label:catName,
        target: parseFloat(r["Category Target %"])   || 0,
        band:   parseFloat(r["Category Band/Range"]) || 0,
        upper:  parseFloat(r["Category Upper %"])    || 0,
        lower:  parseFloat(r["Category Lower %"])    || 0,
        assetClassType: r["Category Asset Class Type"] || "",
        children:[]
      };
      catMap[catName] = catNode;
      root.children.push(catNode);
    }

    // Class — owns its Class Band/Range/Upper/Lower columns
    // Also stores childBand/childUpper/childLower → written to Security Set Band/Range
    const classKey = `${catName||""}__${className||""}`;
    let classNode = classMap[classKey];
    if (className && !classNode) {
      classNode = {
        id: `cls_${classKey}`, type:"class", label:className,
        target:    parseFloat(r["Class Target %"])        || 0,
        band:      parseFloat(r["Class Band/Range"])      || 0,
        upper:     parseFloat(r["Class Upper %"])         || 0,
        lower:     parseFloat(r["Class Lower %"])         || 0,
        // child bands = what gets written to SS level
        childBand:  parseFloat(r["Security Set Band/Range"]) || 0,
        childUpper: parseFloat(r["Security Set Upper %"])    || 0,
        childLower: parseFloat(r["Security Set Lower %"])    || 0,
        children:[]
      };
      classMap[classKey] = classNode;
      const parent = catNode || root;
      parent.children.push(classNode);
    }

    // Security Set — display only, read from SS columns
    if (ssName) {
      const ssNode = {
        id: `ss_${ri}_${ssName}`, type:"ss", label:ssName,
        target: parseFloat(r["* Security Set Target %"])  || 100,
        band:   parseFloat(r["Security Set Band/Range"])  || 0,
        upper:  parseFloat(r["Security Set Upper %"])     || 0,
        lower:  parseFloat(r["Security Set Lower %"])     || 0,
        rowIndex: ri, children:[]
      };
      const parent = classNode || catNode || root;
      parent.children.push(ssNode);
    }
  });

  return root;
}

// ── Apply tree → rows ─────────────────────────────────────────────────────────
// category.target/band/upper/lower  → Category Target/Band/Upper/Lower %
// class.target/band/upper/lower     → Class Target/Band/Upper/Lower %
// class.childBand/childUpper/childLower → Security Set Band/Range/Upper/Lower %
// ss → read-only, nothing written

function applyTreeToRows(rows, tree) {
  const updated = rows.map(r => ({...r}));

  function walk(node) {
    if (node.type === "root") {
      node.children.forEach(walk);

    } else if (node.type === "category") {
      rows.forEach((r,i) => {
        if (r["Category SubModel Name"] === node.label) {
          updated[i]["Category Target %"]   = node.target;
          updated[i]["Category Band/Range"] = node.band;
          updated[i]["Category Upper %"]    = node.upper;
          updated[i]["Category Lower %"]    = node.lower;
        }
      });
      node.children.forEach(walk);

    } else if (node.type === "class") {
      rows.forEach((r,i) => {
        if (r["Class SubModel Name"] === node.label) {
          updated[i]["Class Target %"]          = node.target;
          updated[i]["Class Band/Range"]        = node.band;
          updated[i]["Class Upper %"]           = node.upper;
          updated[i]["Class Lower %"]           = node.lower;
          // class's childBand controls the SS tier
          updated[i]["Security Set Band/Range"] = node.childBand  ?? node.band;
          updated[i]["Security Set Upper %"]    = node.childUpper ?? node.upper;
          updated[i]["Security Set Lower %"]    = node.childLower ?? node.lower;
        }
      });
      // ss children are read-only, no walk needed
    }
  }

  walk(tree);
  return updated;
}

// ── Layout Engine ─────────────────────────────────────────────────────────────
const NODE_R = 38;
const V_GAP  = 90;

function layoutTree(root) {
  // Assign x positions bottom-up (leaf spreading), y by depth
  const positions = {};
  let leafX = 0;

  function measureWidth(node) {
    if (node.children.length===0) {
      node._x = leafX * (NODE_R*2+18);
      leafX++;
      return;
    }
    node.children.forEach(measureWidth);
    const xs = node.children.map(c=>c._x);
    node._x = (Math.min(...xs)+Math.max(...xs))/2;
  }

  function assignY(node, depth) {
    node._y = depth * (NODE_R*2 + V_GAP);
    node.children.forEach(c=>assignY(c, depth+1));
  }

  measureWidth(root);
  assignY(root, 0);

  // Collect all nodes
  const all = [];
  function collect(node) { all.push(node); node.children.forEach(collect); }
  collect(root);

  const minX = Math.min(...all.map(n=>n._x));
  const maxX = Math.max(...all.map(n=>n._x));
  const maxY = Math.max(...all.map(n=>n._y));
  const pad = 60;

  return { all, minX, maxX, maxY, pad };
}

// ── Shared band/target input row helpers ──────────────────────────────────────

function useBandState(initTarget, initBand, initUpper, initLower) {
  const [target, setTargetRaw] = useState(initTarget ?? 0);
  const [band,   setBandRaw]   = useState(initBand   ?? 0);
  const [upper,  setUpper]     = useState(initUpper  ?? 0);
  const [lower,  setLower]     = useState(initLower  ?? 0);

  const setTarget = useCallback((raw) => {
    const t = parseFloat(raw) || 0;
    setTargetRaw(t);
    setBandRaw(b => { if (b > 0) { const a = +(t * b / 100).toFixed(2); setUpper(a); setLower(a); } return b; });
  }, []);

  const setBand = useCallback((raw) => {
    const b = parseFloat(raw) || 0;
    setBandRaw(b);
    setTargetRaw(t => { const a = +(t * b / 100).toFixed(2); setUpper(a); setLower(a); return t; });
  }, []);

  const setUpperDirect = useCallback((raw) => {
    const u = parseFloat(raw) || 0;
    setUpper(u);
    setTargetRaw(t => { if (t > 0) setBandRaw(+(u / t * 100).toFixed(2)); return t; });
  }, []);

  const setLowerDirect = useCallback((raw) => {
    setLower(parseFloat(raw) || 0);
  }, []);

  return { target, band, upper, lower, setTarget, setBand, setUpperDirect, setLowerDirect };
}

function BandFields({ state, showTarget=true, siblingTotal=null, hideTargetSum=false }) {
  const { target, band, upper, lower, setTarget, setBand, setUpperDirect, setLowerDirect } = state;
  const sumOk = siblingTotal === null || Math.abs(siblingTotal - 100) < 0.01;
  return (
    <div>
      {showTarget && (
        <div style={{marginBottom:10}}>
          <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Target %</label>
          <input type="number" step="0.01" min="0" max="100" value={target}
            onChange={e=>setTarget(e.target.value)}
            style={{width:"100%",padding:"6px 8px",border:`0.5px solid ${sumOk?"#d1d5db":"#fca5a5"}`,borderRadius:6,fontSize:13,boxSizing:"border-box"}}
          />
          {!hideTargetSum && siblingTotal !== null && (
            <div style={{fontSize:10,marginTop:3,color:sumOk?"#16a34a":"#dc2626"}}>
              Siblings total: {siblingTotal.toFixed(2)}% {sumOk?"✓":"← must equal 100%"}
            </div>
          )}
        </div>
      )}
      <div style={{marginBottom:6}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
          <label style={{fontSize:11,fontWeight:600,color:"#374151"}}>Band</label>
          <span style={{fontSize:10,color:"#9ca3af"}}>% of target</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:5}}>
          <input type="number" step="0.5" min="0" max="100" value={band}
            onChange={e=>setBand(e.target.value)}
            style={{flex:1,padding:"6px 8px",border:"0.5px solid #d1d5db",borderRadius:6,fontSize:13,boxSizing:"border-box"}}
          />
          <span style={{fontSize:11,color:"#6b7280"}}>%</span>
        </div>
      </div>
      <div style={{background:"#f9fafb",borderRadius:6,padding:"8px",marginBottom:4}}>
        <div style={{fontSize:10,fontWeight:600,color:"#6b7280",marginBottom:6,textTransform:"uppercase",letterSpacing:"0.05em"}}>Tolerance (absolute %)</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
          {[["Upper",upper,setUpperDirect],["Lower",lower,setLowerDirect]].map(([lbl,val,fn])=>(
            <div key={lbl}>
              <label style={{fontSize:10,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>{lbl} %</label>
              <input type="number" step="0.01" min="0" value={val}
                onChange={e=>fn(e.target.value)}
                style={{width:"100%",padding:"4px 6px",border:"0.5px solid #d1d5db",borderRadius:4,fontSize:12,boxSizing:"border-box"}}
              />
              <div style={{fontSize:10,color:"#9ca3af",marginTop:1}}>
                {lbl==="Upper"?`=${+(target+val).toFixed(2)}%`:`=${+(target-val).toFixed(2)}%`}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Children table row — target + band per child, fully independent ──

function ChildRow({ child, onChange }) {
  const [target, setTargetRaw] = useState(child.target ?? 0);
  const [band,   setBandRaw]   = useState(child.band   ?? 0);
  const [upper,  setUpper]     = useState(child.upper  ?? 0);
  const [lower,  setLower]     = useState(child.lower  ?? 0);

  function setTarget(raw) {
    const t = parseFloat(raw) || 0;
    setTargetRaw(t);
    if (band > 0) { const a = +(t * band / 100).toFixed(2); setUpper(a); setLower(a); }
  }
  function setBand(raw) {
    const b = parseFloat(raw) || 0;
    setBandRaw(b);
    const a = +(target * b / 100).toFixed(2);
    setUpper(a); setLower(a);
  }
  function setUpperDirect(raw) {
    const u = parseFloat(raw) || 0;
    setUpper(u);
    if (target > 0) setBandRaw(+(u / target * 100).toFixed(2));
  }
  function setLowerDirect(raw) { setLower(parseFloat(raw) || 0); }

  useEffect(() => { onChange(child.id, { target, band, upper, lower }); }, [target, band, upper, lower]);

  const col = NODE_COLORS[child.type] || NODE_COLORS.ss;
  return (
    <div style={{borderBottom:"0.5px solid #f0f0f0",paddingBottom:12,marginBottom:12}}>
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
        <div style={{width:8,height:8,borderRadius:"50%",background:col.fill,border:`1.5px solid ${col.stroke}`,flexShrink:0}}/>
        <span style={{fontSize:12,fontWeight:600,color:"#374151",lineHeight:1.3}}>{child.label}</span>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
        <div>
          <label style={{fontSize:10,fontWeight:600,color:"#6b7280",display:"block",marginBottom:2}}>Target %</label>
          <input type="number" step="0.01" min="0" max="100" value={target}
            onChange={e=>setTarget(e.target.value)}
            style={{width:"100%",padding:"5px 7px",border:"0.5px solid #d1d5db",borderRadius:5,fontSize:12,boxSizing:"border-box"}}
          />
        </div>
        <div>
          <label style={{fontSize:10,fontWeight:600,color:"#6b7280",display:"block",marginBottom:2}}>Band %</label>
          <input type="number" step="0.5" min="0" max="100" value={band}
            onChange={e=>setBand(e.target.value)}
            style={{width:"100%",padding:"5px 7px",border:"0.5px solid #d1d5db",borderRadius:5,fontSize:12,boxSizing:"border-box"}}
          />
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
        {[["Upper",upper,setUpperDirect],["Lower",lower,setLowerDirect]].map(([lbl,val,fn])=>(
          <div key={lbl}>
            <label style={{fontSize:10,fontWeight:600,color:"#6b7280",display:"block",marginBottom:2}}>{lbl} (abs %)</label>
            <input type="number" step="0.01" min="0" value={val}
              onChange={e=>fn(e.target.value)}
              style={{width:"100%",padding:"5px 7px",border:"0.5px solid #d1d5db",borderRadius:5,fontSize:12,boxSizing:"border-box"}}
            />
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Edit Panel ────────────────────────────────────────────────────────────────
// Band semantics — each node's band controls the tier BELOW it:
//   root band/upper/lower      → Category tolerance (written to Category Band/Range, Upper, Lower)
//   category band/upper/lower  → Class tolerance    (written to Class Band/Range, Upper, Lower)
//   class band/upper/lower     → SS tolerance       (written to SS Band/Range, Upper, Lower)
//   ss                         → read-only, not editable

const BAND_CONTROLS_LABEL = {
  root:     "Category tolerance",
  category: "Class tolerance",
  class:    "Security Set tolerance",
};

function EditPanel({ node, onSave, onClose, siblingSum }) {
  const isReadOnly = node.type === "ss";
  const isParent   = node.type === "root" || node.type === "category";

  // Own band state — this node's own Band/Range/Upper/Lower column
  const ownState = useBandState(node.target, node.band, node.upper, node.lower);

  // Child band state — only for class nodes; independently controls Security Set Band/Range
  const childBandState = useBandState(node.target, node.childBand ?? 0, node.childUpper ?? 0, node.childLower ?? 0);

  // Only non-SS children are editable in the children table
  const editableChildren = (node.children||[]).filter(c => c.type !== "ss");

  const [childEdits, setChildEdits] = useState(() => {
    const m = {};
    editableChildren.forEach(c => { m[c.id] = { target: c.target, band: c.band, upper: c.upper, lower: c.lower }; });
    return m;
  });

  function handleChildChange(childId, vals) {
    setChildEdits(prev => ({ ...prev, [childId]: vals }));
  }

  const childTotal = editableChildren.reduce((s,c) => s + (childEdits[c.id]?.target ?? c.target), 0);
  const childSumOk = editableChildren.length === 0 || Math.abs(childTotal - 100) < 0.01;

  const otherSum  = siblingSum - node.target;
  const ownTotal  = +(otherSum + ownState.target).toFixed(4);
  const ownSumOk  = Math.abs(ownTotal - 100) < 0.01;

  const typeLabel      = { root:"Model", category:"Category", class:"Class", ss:"Security Set" }[node.type];
  const childTypeName  = node.type === "root" ? "Categories" : "Classes";
  const bandLabel      = BAND_CONTROLS_LABEL[node.type];

  function handleSave() {
    onSave({
      ownVals: {
        target: ownState.target, band: ownState.band, upper: ownState.upper, lower: ownState.lower,
        childBand: childBandState.band, childUpper: childBandState.upper, childLower: childBandState.lower,
      },
      childEdits
    });
  }

  return (
    <div style={{
      position:"absolute", top:0, right:0, width:310, background:"#fff",
      border:"0.5px solid #e5e7eb", borderRadius:10, boxShadow:"0 4px 24px rgba(0,0,0,0.13)",
      padding:"16px", zIndex:100, maxHeight:"80vh", overflowY:"auto"
    }}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
        <div style={{fontSize:11,fontWeight:700,color:"#111827",textTransform:"uppercase",letterSpacing:"0.06em"}}>{typeLabel}</div>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:16,lineHeight:1,padding:0}}>✕</button>
      </div>
      <div style={{fontSize:11,color:"#6b7280",marginBottom:12,lineHeight:1.4,wordBreak:"break-word",borderBottom:"0.5px solid #f3f4f6",paddingBottom:10}}>
        {node.label}
      </div>

      {/* ── Security Set: read-only reference ── */}
      {isReadOnly && (
        <div>
          <div style={{fontSize:10,fontWeight:700,color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:8}}>
            Reference only — built separately in Orion
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6,marginBottom:14}}>
            {[["SS Target %", node.target+"%"],["Band", node.band+"%"],["Upper %", node.upper+"%"],["Lower %", node.lower+"%"]].map(([lbl,val])=>(
              <div key={lbl} style={{background:"#f9fafb",border:"0.5px solid #e5e7eb",borderRadius:5,padding:"6px 8px"}}>
                <div style={{fontSize:10,color:"#9ca3af",marginBottom:2}}>{lbl}</div>
                <div style={{fontSize:13,fontWeight:600,color:"#6b7280"}}>{val}</div>
              </div>
            ))}
          </div>
          <button onClick={onClose} style={{width:"100%",padding:"7px",border:"0.5px solid #d1d5db",borderRadius:6,background:"none",fontSize:12,cursor:"pointer",color:"#374151"}}>Close</button>
        </div>
      )}

      {/* ── Category: read-only model-level allocation tiles ── */}
      {node.type === "category" && (
        <div style={{marginBottom:14,paddingBottom:14,borderBottom:"0.5px solid #e5e7eb"}}>
          <div style={{fontSize:10,fontWeight:700,color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:8}}>Model-level allocation</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:5}}>
            {[["Target",node.target+"%"],["Band",node.band+"%"],["Upper",node.upper+"%"],["Lower",node.lower+"%"]].map(([lbl,val])=>(
              <div key={lbl} style={{background:"#f9fafb",border:"0.5px solid #e5e7eb",borderRadius:5,padding:"5px 7px",textAlign:"center"}}>
                <div style={{fontSize:9,color:"#9ca3af",marginBottom:2}}>{lbl}</div>
                <div style={{fontSize:12,fontWeight:600,color:"#374151"}}>{val}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── Root header for children table ── */}
      {node.type === "root" && (
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{fontSize:11,fontWeight:700,color:"#374151"}}>{childTypeName}</div>
          <div style={{fontSize:11,color:childSumOk?"#16a34a":"#dc2626",fontWeight:500}}>
            {childTotal.toFixed(2)}% {childSumOk?"✓":"← needs 100%"}
          </div>
        </div>
      )}

      {/* ── Editable children table (target only — bands live on this node) ── */}
      {!isReadOnly && editableChildren.length > 0 && (
        <div style={{marginBottom:14}}>
          {node.type === "category" && (
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div style={{fontSize:11,fontWeight:700,color:"#374151"}}>{childTypeName}</div>
              <div style={{fontSize:11,color:childSumOk?"#16a34a":"#dc2626",fontWeight:500}}>
                {childTotal.toFixed(2)}% {childSumOk?"✓":"← needs 100%"}
              </div>
            </div>
          )}
          {editableChildren.map(child => (
            <ChildRow key={child.id} child={child} onChange={handleChildChange} />
          ))}
        </div>
      )}

      {/* ── Class: own target (its sibling allocation within the category) ── */}
      {node.type === "class" && (
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Class target %</label>
          <input type="number" step="0.01" min="0" max="100" value={ownState.target}
            onChange={e=>ownState.setTarget(e.target.value)}
            style={{width:"100%",padding:"6px 8px",border:`0.5px solid ${ownSumOk?"#d1d5db":"#fca5a5"}`,borderRadius:6,fontSize:13,boxSizing:"border-box"}}
          />
          <div style={{fontSize:10,marginTop:3,color:ownSumOk?"#16a34a":"#dc2626"}}>
            Siblings total: {ownTotal.toFixed(2)}% {ownSumOk?"✓":"← must equal 100%"}
          </div>
        </div>
      )}

      {/* ── Class: band/tolerance for Security Set tier — separate from class's own band ── */}
      {node.type === "class" && (
        <div style={{background:"#f8fafc",border:"0.5px solid #e2e8f0",borderRadius:7,padding:"10px",marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#0369a1",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:8}}>
            Security Set tolerance
          </div>
          <div style={{marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <label style={{fontSize:11,fontWeight:600,color:"#374151"}}>Band</label>
              <span style={{fontSize:10,color:"#9ca3af"}}>% of target → sets Upper &amp; Lower</span>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              <input type="number" step="0.5" min="0" max="100" value={childBandState.band}
                onChange={e=>childBandState.setBand(e.target.value)}
                style={{flex:1,padding:"6px 8px",border:"0.5px solid #d1d5db",borderRadius:6,fontSize:13,boxSizing:"border-box"}}
              />
              <span style={{fontSize:11,color:"#6b7280"}}>%</span>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
            {[["Upper",childBandState.upper,childBandState.setUpperDirect],["Lower",childBandState.lower,childBandState.setLowerDirect]].map(([lbl,val,fn])=>(
              <div key={lbl}>
                <label style={{fontSize:10,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>{lbl} (absolute %)</label>
                <input type="number" step="0.01" min="0" value={val}
                  onChange={e=>fn(e.target.value)}
                  style={{width:"100%",padding:"4px 6px",border:"0.5px solid #d1d5db",borderRadius:4,fontSize:12,boxSizing:"border-box"}}
                />
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Save/Cancel */}
      {!isReadOnly && (
        <div style={{display:"flex",gap:8,position:"sticky",bottom:0,background:"#fff",paddingTop:8,borderTop:"0.5px solid #f3f4f6"}}>
          <button onClick={onClose} style={{flex:1,padding:"7px",border:"0.5px solid #d1d5db",borderRadius:6,background:"none",fontSize:12,cursor:"pointer",color:"#374151"}}>Cancel</button>
          <button onClick={handleSave} style={{flex:1,padding:"7px",border:"none",borderRadius:6,background:"#1a56db",fontSize:12,fontWeight:600,cursor:"pointer",color:"#fff"}}>Save</button>
        </div>
      )}
    </div>
  );
}

// ── Model Tree View ───────────────────────────────────────────────────────────

function ModelTree({ modelName, rows, onRowsChange }) {
  const [tree, setTree] = useState(()=>buildTree(modelName, rows));
  const [selectedId, setSelectedId] = useState(null);
  const [scrollLeft, setScrollLeft] = useState(0);
  const containerRef = useRef();

  const { all, minX, maxX, maxY, pad } = useMemo(()=>layoutTree(tree),[tree]);
  const svgW = maxX - minX + pad*2 + NODE_R*2;
  const svgH = maxY + pad*2 + NODE_R*2;
  const ox = -minX + pad;

  function getSiblingSum(node) {
    let parent = null;
    function findParent(n) {
      if (n.children.some(c=>c.id===node.id)) { parent=n; return; }
      n.children.forEach(findParent);
    }
    findParent(tree);
    if (!parent) return 0;
    return parent.children.reduce((s,c)=>s+c.target,0);
  }

  function handleSave(nodeId, payload) {
    const { ownVals, childEdits } = payload;
    function updateNode(n) {
      if (n.id === nodeId) {
        const updated = {
          ...n,
          target: ownVals.target, band: ownVals.band, upper: ownVals.upper, lower: ownVals.lower,
          childBand: ownVals.childBand ?? n.childBand,
          childUpper: ownVals.childUpper ?? n.childUpper,
          childLower: ownVals.childLower ?? n.childLower,
        };
        if (childEdits && Object.keys(childEdits).length > 0) {
          updated.children = n.children.map(c => {
            const edit = childEdits[c.id];
            return edit ? { ...c, target: edit.target, band: edit.band, upper: edit.upper, lower: edit.lower } : c;
          });
        }
        return updated;
      }
      return { ...n, children: n.children.map(updateNode) };
    }
    const newTree = updateNode(tree);
    setTree(newTree);
    const newRows = applyTreeToRows(rows, newTree);
    onRowsChange(newRows);
    setSelectedId(null);
  }

  const selectedNode = all.find(n=>n.id===selectedId);
  const siblingSum = selectedNode ? getSiblingSum(selectedNode) : 0;

  // Validation: check if any sibling group doesn't sum to 100
  function getValidation(node) {
    if (node.type==="root"||node.type==="ss") return true;
    if (node.children.length===0) return true;
    const sum = node.children.reduce((s,c)=>s+c.target,0);
    return Math.abs(sum-100)<0.5;
  }

  return (
    <div style={{position:"relative"}}>
      <div style={{fontSize:13,fontWeight:700,color:"#111827",marginBottom:8,display:"flex",alignItems:"center",gap:8}}>
        {modelName}
        <span style={{fontSize:11,fontWeight:400,color:"#6b7280"}}>Click any node to edit</span>
      </div>
      <div ref={containerRef} style={{overflowX:"auto",overflowY:"visible",background:"#0f172a",borderRadius:10,padding:"8px 0"}}>
        <svg width={Math.max(svgW,400)} height={svgH} style={{display:"block"}}>
          {/* Connector lines */}
          {all.map(node =>
            node.children.map(child => (
              <line key={`${node.id}-${child.id}`}
                x1={node._x+ox+NODE_R} y1={node._y+pad+NODE_R}
                x2={child._x+ox+NODE_R} y2={child._y+pad}
                stroke="#334155" strokeWidth="1.5"
              />
            ))
          )}
          {/* Nodes */}
          {all.map(node => {
            const cx = node._x+ox+NODE_R;
            const cy = node._y+pad+NODE_R;
            const col = NODE_COLORS[node.type]||NODE_COLORS.ss;
            const isSelected = node.id===selectedId;
            const valid = getValidation(node);
            const shortLabel = node.label.length>14 ? node.label.slice(0,13)+"…" : node.label;
            return (
              <g key={node.id} style={{cursor:"pointer"}}
                onClick={()=>{ setSelectedId(node.id===selectedId?null:node.id); }}>
                {/* Outer ring for selection / error */}
                <circle cx={cx} cy={cy} r={NODE_R+4}
                  fill="none"
                  stroke={isSelected?"#60a5fa":!valid?"#f87171":"transparent"}
                  strokeWidth={isSelected?2:1.5}
                />
                <circle cx={cx} cy={cy} r={NODE_R} fill={col.fill} stroke={col.stroke} strokeWidth="2"/>
                <text x={cx} y={cy-7} textAnchor="middle" fontSize="9.5" fontWeight="600" fill={col.text} fontFamily="system-ui">
                  {shortLabel}
                </text>
                <text x={cx} y={cy+7} textAnchor="middle" fontSize="10" fontWeight="700" fill={col.text} fontFamily="system-ui">
                  {node.target}%
                </text>
                {node.type!=="root" && node.type!=="ss" && (
                  <text x={cx} y={cy+18} textAnchor="middle" fontSize="8.5" fill={col.text} fontFamily="system-ui" opacity="0.8">
                    ±{node.band||node.upper}
                  </text>
                )}
                {!valid && (
                  <text x={cx+NODE_R-6} y={cy-NODE_R+6} textAnchor="middle" fontSize="12" fill="#f87171">!</text>
                )}
              </g>
            );
          })}
        </svg>
      </div>

      {/* Edit panel */}
      {selectedNode && (
        <EditPanel
          key={selectedNode.id}
          node={selectedNode}
          siblingSum={siblingSum}
          onSave={(vals)=>handleSave(selectedNode.id, vals)}
          onClose={()=>setSelectedId(null)}
        />
      )}

      {/* Legend */}
      <div style={{display:"flex",gap:12,marginTop:8,flexWrap:"wrap"}}>
        {[["root","Model"],["category","Category"],["class","Class"],["ss","Security Set"]].map(([type,lbl])=>(
          <div key={type} style={{display:"flex",alignItems:"center",gap:5}}>
            <div style={{width:12,height:12,borderRadius:"50%",background:NODE_COLORS[type].fill,border:`1.5px solid ${NODE_COLORS[type].stroke}`}}/>
            <span style={{fontSize:11,color:"#6b7280"}}>{lbl}</span>
          </div>
        ))}
        <div style={{display:"flex",alignItems:"center",gap:5}}>
          <div style={{width:12,height:12,borderRadius:"50%",background:"transparent",border:"1.5px solid #f87171"}}/>
          <span style={{fontSize:11,color:"#6b7280"}}>Targets don't sum to 100%</span>
        </div>
      </div>

      {/* Band defaults reminder */}
      <div style={{marginTop:10,display:"inline-flex",alignItems:"center",gap:0,background:"#f8fafc",border:"0.5px solid #e2e8f0",borderRadius:7,overflow:"hidden",fontSize:11}}>
        <div style={{padding:"6px 10px",background:"#1a6fb5",color:"#e0f0ff",fontWeight:600,whiteSpace:"nowrap"}}>
          Band defaults
        </div>
        {[
          {label:"Category",detail:"5 absolute Upper/Lower",color:"#1a6fb5",bg:"#eef4fb"},
          {label:"Class",detail:"25% band",color:"#8a7a00",bg:"#fdfbe8"},
          {label:"Security Set",detail:"50% band",color:"#1240a8",bg:"#eef1fd"},
        ].map(({label,detail,color,bg},i,arr)=>(
          <div key={label} style={{display:"flex",alignItems:"center"}}>
            <div style={{padding:"6px 12px",background:bg,color:"#374151",whiteSpace:"nowrap"}}>
              <span style={{fontWeight:600,color}}>{label}</span>
              <span style={{color:"#6b7280",marginLeft:4}}>{detail}</span>
            </div>
            {i < arr.length-1 && (
              <div style={{color:"#94a3b8",fontSize:13,padding:"0 2px",background:"#f1f5f9",alignSelf:"stretch",display:"flex",alignItems:"center"}}>›</div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Excel output ──────────────────────────────────────────────────────────────

function buildOutputRows(inputRows) {
  return inputRows.map(r => {
    const out = {};
    TEMPLATE_COLS.forEach(col => {
      const val = r[col];
      out[col] = (val===undefined||val===null||val==="") ? null : val;
    });
    return out;
  });
}

function downloadXlsx(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows, { header:TEMPLATE_COLS });
  ws["!cols"] = TEMPLATE_COLS.map(c=>({ wch:Math.max(c.length+2,14) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
  XLSX.writeFile(wb, filename);
}

function groupByModel(rows) {
  const models = {};
  rows.forEach(r => {
    const name = r["* Model Name"]||"(unnamed)";
    if (!models[name]) models[name]=[];
    models[name].push(r);
  });
  return models;
}

// ── Model Library (targets) parsing ────────────────────────────────────────
// The library workbook has one sheet per "portfolio family" (e.g. base vs.
// Tax Aware). Each sheet is a matrix: rows are Category (ALL CAPS, no fund
// name), Class (mixed case, no fund name), or Ticker (has a fund name —
// security-level, ignored: this tool only updates Category/Class targets).
// Columns after Ticker/Fund Name are model variants (e.g. "Conservative"),
// each holding that variant's allocation as a fraction of the whole model.

function normAlnum(s) {
  return (s||"").toString().toUpperCase().replace(/[^A-Z0-9]/g,"");
}

function parseLibraryWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type:"array" });
  const sheets = [];
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    let headerIdx = -1;
    for (let i=0; i<Math.min(raw.length,10); i++) {
      const norm = (raw[i]||[]).map(c=>normalizeKey(c));
      if (norm.includes("ticker")) { headerIdx=i; break; }
    }
    if (headerIdx===-1) return; // not a recognizable matrix sheet — skip it
    const headerRow = raw[headerIdx];
    const variantCols = [];
    for (let c=2; c<headerRow.length; c++) {
      const v = headerRow[c];
      if (typeof v === "string" && v.trim()) variantCols.push({ idx:c, name:v.trim() });
      else if (variantCols.length>0) break;
    }
    const rows = [];
    let currentCategory = null;
    for (let i=headerIdx+1; i<raw.length; i++) {
      const row = raw[i];
      if (!row) continue;
      const label = row[0];
      if (label===null || label===undefined || label==="") continue;
      const labelStr = label.toString().trim();
      if (!labelStr) continue;
      const fundName = row[1];
      const isTicker = fundName !== null && fundName !== undefined && fundName !== "";
      if (isTicker) continue; // security/ticker level — not managed by this tool
      const isCategory = labelStr === labelStr.toUpperCase() && labelStr !== labelStr.toLowerCase();
      const values = {};
      variantCols.forEach(({idx,name}) => {
        const v = row[idx];
        values[name] = (v===null||v===undefined||v==="") ? null : parseFloat(v);
      });
      if (isCategory) {
        currentCategory = labelStr;
        rows.push({ level:"category", label:labelStr, values });
      } else {
        // category tracks which category (by label) this class belongs to,
        // based on the sheet's own row order — needed so exclusion rules
        // (e.g. ex-USLC) know which classes are siblings of the excluded one.
        rows.push({ level:"class", label:labelStr, values, category:currentCategory });
      }
    }
    // Guard against unrelated sheets that also happen to have a "Ticker"
    // column — e.g. a different fund family's own allocation matrix, or a
    // flat ticker/allocation holdings list for a stock model. A genuine STP
    // variant matrix always has at least one real ALL-CAPS category row
    // (EQUITY, FIXED INCOME, etc.) and more than one variant column; sheets
    // without that structural signature are skipped regardless of name, so
    // this doesn't depend on guessing which sheet names to allow.
    if (variantCols.length < 2 || !rows.some(r => r.level === "category")) return;

    sheets.push({ name, variants: variantCols.map(v=>v.name), rows, isTaxAware: /tax/i.test(name) });
  });
  return sheets;
}

// Finds the library row whose label is contained within a Category/Class
// SubModel Name — e.g. library "Commodities" inside "Savvy ... - Commodities",
// or library "U.S. Investment Grade" inside "... U.S. Investment Grade FI".
// Real-world naming isn't always an exact match, so containment (rather than
// equality) is the rule. Ambiguous cases exist — e.g. "STP - All Equity -
// Cash" contains both "EQUITY" and "CASH" — so among candidates the one
// ending closest to the end of the string wins (the actual category/class
// segment is always the trailing part of the name); ties go to the longer label.
function matchLibraryLabel(sourceStr, libraryRows) {
  const ns = normAlnum(sourceStr);
  if (!ns) return null;
  let best = null, bestEnd = -1, bestLen = 0;
  libraryRows.forEach(row => {
    const nl = normAlnum(row.label);
    if (!nl) return;
    const idx = ns.lastIndexOf(nl);
    if (idx === -1) return;
    const end = idx + nl.length;
    if (end > bestEnd || (end === bestEnd && nl.length > bestLen)) {
      best = row; bestEnd = end; bestLen = nl.length;
    }
  });
  return best;
}

// Strips recognized modifier tags off a Model Name's trailing segments (in
// any order/combination) and reports which ones were found. Tags can appear
// either parenthesized ("(Tax Aware)") or as a bare trailing " - X" segment
// ("... - Tax Aware", "... - Core Stock Model") — real files mix both styles.
//   isTaxAware         → use the Tax Aware library sheet instead of the base one
//   isExUslc           → excludes US Large Cap; see the flatten-and-rescale rule below
//   isUsEquityOnly     → excludes non-US equity; US Small Cap keeps its normal
//                        value, US Large Cap absorbs whatever's left
//   holdingsAssetClass → a holdings substitution: "equity" (Core/Growth/
//                        Value/Dividend Stock Model, Enhanced Dividend Stock)
//                        or "fixedIncome" (Bond Ladder). Category/Class
//                        targets are identical to the base variant — only the
//                        holdings differ — but the substitute product name
//                        sometimes replaces the category label too (e.g.
//                        "Savvy Bond Ladder" instead of "...Fixed Income"),
//                        so matching falls back to the known asset-class
//                        token when the renamed label doesn't contain it.
//                        Bond Ladder specifically collapses its whole category
//                        into one row at 100% (verified: Category Target %
//                        equals the base model's Fixed Income value, Class
//                        Target % is always 100 — there's nothing else in it).
//                        "Custom Model" is tagged the same way but with no
//                        specific asset class, since it's only ever seen
//                        alongside an otherwise-unrecognized name already
//                        left alone.
// Whatever's left after peeling should reduce to exactly "{prefix} - {variant}";
// anything else (extra segments, unrecognized suffixes) is a genuine
// derivative the library doesn't cover.
function analyzeModelName(modelName) {
  const core = (modelName||"").replace(/\([^)]*\)/g,"").replace(/\s{2,}/g," ").trim();
  const segments = core.split(" - ").map(s=>s.trim()).filter(Boolean);
  const flags = {
    isTaxAware: /tax/i.test(modelName||""),
    isExUslc: /ex[-\s]?uslc/i.test(modelName||""),
    isUsEquityOnly: false,
    holdingsAssetClass: null, // "equity" | "fixedIncome" | "other" | null
  };
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

// Matches a Model Name to a specific {sheet, variant} (or blend of two
// neighboring variants — see the risk-ladder note further down) in the
// library, after peeling off any recognized modifier tags via analyzeModelName.
// Returns { match, flags, reason }: match is null with a human-readable
// reason when the library genuinely doesn't cover this model.
function matchModelVariant(modelName, librarySheets) {
  const { segments, flags } = analyzeModelName(modelName);
  if (segments.length !== 2) {
    return { match:null, flags, reason:`${segments.length} segment(s) left after removing recognized tags (${segments.join(" / ")||"none"}) — doesn't reduce to a single "{prefix} - {variant}" form` };
  }
  const candidate = segments[1];
  const searchIn = librarySheets.filter(s => s.isTaxAware === flags.isTaxAware);
  if (searchIn.length === 0) {
    return { match:null, flags, reason:`no ${flags.isTaxAware?"Tax Aware":"base"} sheet found in the library file` };
  }

  for (const sheet of searchIn) {
    const variant = sheet.variants.find(v => v.toLowerCase() === candidate.toLowerCase());
    if (variant) return { match:{ kind:"single", sheet, variant }, flags, reason:null };
  }

  // 30/70-style blend: verified against real data as the average of the two
  // neighboring risk-ladder variants — pair index = (firstNumber-10)/20 along
  // the library's own column order (10/90=AllFixed+Conservative, 30/70=
  // Conservative+ModConservative, 50/50=ModConservative+Moderate, 70/30=
  // Moderate+ModAggressive).
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

// Reads a category/class row's fraction for whatever matched — a single
// variant column, or the average of two neighboring variants for a blend
// model. Blend averaging treats a missing/blank side as 0% (verified: 10/90's
// Equity = avg(0, Conservative's 22) = 11, since All Fixed has no Equity row).
function getMatchFrac(row, match) {
  if (!row) return null;
  if (match.kind === "single") return row.values[match.variant];
  const a = row.values[match.lowerVariant], b = row.values[match.upperVariant];
  if ((a===null||a===undefined) && (b===null||b===undefined)) return null;
  return ((a??0) + (b??0)) / 2;
}

// Compares an existing-model export against the library and computes new
// Category Target % / Class Target % values. Bands are never touched.
//
// Standard models:
//   new Category Target % = category's fraction of the whole model × 100
//   new Class Target %    = class's fraction of the whole model, re-based as
//                            a share of its own category (so siblings still
//                            sum to 100%) — matches how the existing file
//                            already stores class-level targets.
//
// ex-USLC (confirmed rule; note this changes the file's current ex-USLC
// numbers, which were built on an earlier whole-portfolio version):
//   US Large Cap is excluded and its share is redistributed pro-rata among
//   the *other classes in its own category (Equity) only* — Fixed Income,
//   Alternatives, and Cash are untouched. Category Target % still reflects
//   each class's true underlying category value (e.g. 62/32/5/1), even
//   though the file's Category SubModel Name text is one flattened label
//   shared across every row of the model.
//
// US Equity Only (verified against STP - Moderate - US Equity Only):
//   Category Target %s are untouched. US Small Cap keeps its normal
//   classFrac/catFrac value; US Large Cap = 100 - that Small Cap value
//   (they're the only two classes left in the category).
//
// Differences within `tolerance` percentage points are treated as rounding
// noise, not a real change — the original value is left exactly as-is.
// Defaults to 0.05pp but is user-adjustable in the review screen.
const DEFAULT_CHANGE_TOLERANCE = 0.05;

function computeLibraryUpdates(rows, librarySheets, tolerance = DEFAULT_CHANGE_TOLERANCE) {
  const updated = rows.map(r => ({...r}));
  const changes = [];
  const changeKeys = new Set();
  const skippedModels = new Map(); // modelName -> reason
  const matchedModels = new Set();
  const unmatchedCategories = [];
  const unmatchedClasses = [];
  const modelMatchCache = {};

  function recordChange(i, field, modelName, label, oldVal, newVal, note) {
    if (Math.abs(newVal - (parseFloat(oldVal)||0)) > tolerance) {
      // newVal is part of the key: ex-USLC rows can share the exact same
      // Category label text across genuinely different underlying categories
      // (the file flattens the label, not the number), so two rows with the
      // same label but different correct values must NOT be deduped away.
      const key = `${modelName}|${field}|${label}|${newVal}`;
      if (!changeKeys.has(key)) {
        changeKeys.add(key);
        changes.push({ modelName, level: field==="Category Target %"?"Category":"Class", label, oldVal, newVal, note });
      }
      updated[i][field] = newVal;
    }
  }

  rows.forEach((r, i) => {
    const modelName = r["* Model Name"];
    if (!modelName) return;
    if (!(modelName in modelMatchCache)) modelMatchCache[modelName] = matchModelVariant(modelName, librarySheets);
    const { match, flags, reason } = modelMatchCache[modelName];
    if (!match) { skippedModels.set(modelName, reason); return; }
    matchedModels.add(modelName);

    const { sheet } = match;
    const catRows = sheet.rows.filter(x=>x.level==="category");
    const classRows = sheet.rows.filter(x=>x.level==="class");
    const catSubName = r["Category SubModel Name"];
    const classSubName = r["Class SubModel Name"];

    if (flags.isExUslc) {
      // Verified against real data: US Large Cap is excluded and its
      // whole-model fraction is redistributed pro-rata across EVERY other
      // class in the model (not just Equity) — new Class % = classFrac ×
      // 1/(1-largeCapFrac) × 100. Category Target % is intentionally left
      // untouched — the file's Category SubModel Name for these rows is a
      // flattened placeholder, not real structured data.
      const usLargeCapRow = matchLibraryLabel("US Large Cap", classRows);
      const largeCapFrac = usLargeCapRow ? (getMatchFrac(usLargeCapRow, match) || 0) : 0;
      const scale = largeCapFrac < 1 ? 1/(1-largeCapFrac) : 1;

      if (classSubName) {
        const classMatch = matchLibraryLabel(classSubName, classRows);
        if (!classMatch) {
          unmatchedClasses.push({ modelName, label: classSubName });
        } else {
          const classFrac = getMatchFrac(classMatch, match);
          if (classFrac !== null && classFrac !== undefined) {
            recordChange(i, "Class Target %", modelName, classSubName, r["Class Target %"], +(classFrac*scale*100).toFixed(2));
          }
        }
      }
      return;
    }

    let catMatch = catSubName ? matchLibraryLabel(catSubName, catRows) : null;
    if (!catMatch && catSubName && flags.holdingsAssetClass === "equity") {
      catMatch = matchLibraryLabel("EQUITY", catRows);
    } else if (!catMatch && catSubName && flags.holdingsAssetClass === "fixedIncome") {
      catMatch = matchLibraryLabel("FIXED INCOME", catRows);
    }

    if (catSubName) {
      if (!catMatch) {
        unmatchedCategories.push({ modelName, label: catSubName });
      } else {
        const frac = getMatchFrac(catMatch, match);
        if (frac !== null && frac !== undefined) {
          recordChange(i, "Category Target %", modelName, catSubName, r["Category Target %"], +(frac*100).toFixed(2));
        }
      }
    }

    if (classSubName) {
      let classMatch = matchLibraryLabel(classSubName, classRows);
      let forcedFullAllocation = false;
      if (!classMatch && flags.holdingsAssetClass === "equity") {
        // Substituted holdings row (stock model) stands in for US Large Cap.
        classMatch = matchLibraryLabel("US Large Cap", classRows);
      } else if (!classMatch && flags.holdingsAssetClass === "fixedIncome") {
        // Bond Ladder collapses its whole category into this one row (verified: always 100%).
        forcedFullAllocation = true;
      }
      if (!classMatch && !forcedFullAllocation) {
        unmatchedClasses.push({ modelName, label: classSubName });
      } else {
        const catFrac = catMatch ? getMatchFrac(catMatch, match) : null;
        let newVal = null;
        if (forcedFullAllocation) {
          newVal = 100;
        } else if (flags.isUsEquityOnly && normAlnum(classMatch.label)==="USLARGECAP") {
          const smallCapRow = matchLibraryLabel("US Small Cap", classRows);
          const smallCapFrac = smallCapRow ? getMatchFrac(smallCapRow, match) : null;
          if (smallCapFrac !== null && smallCapFrac !== undefined && catFrac) {
            newVal = +(100 - +(smallCapFrac/catFrac*100).toFixed(2)).toFixed(2);
          }
        } else {
          const classFrac = getMatchFrac(classMatch, match);
          if (classFrac !== null && classFrac !== undefined && catFrac) newVal = +(classFrac/catFrac*100).toFixed(2);
        }
        if (newVal !== null) recordChange(i, "Class Target %", modelName, classSubName, r["Class Target %"], newVal);
      }
    }
  });

  return {
    updatedRows: updated, changes,
    skippedModels: [...skippedModels.entries()].map(([modelName, reason]) => ({ modelName, reason })),
    matchedModelCount: matchedModels.size,
    unmatchedCategories, unmatchedClasses,
  };
}

function downloadXlsxWithHeaders(rows, headers, filename) {
  const ws = XLSX.utils.json_to_sheet(rows, { header:headers });
  ws["!cols"] = headers.map(c=>({ wch:Math.max((c||"").length+2,14) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
  XLSX.writeFile(wb, filename);
}

// ── Update Existing Models Flow ─────────────────────────────────────────────
// Two-file workflow: an existing Orion export (current state) + a model
// library file (new targets). Matches models/categories/classes between them
// and lets the user review every value that would change before exporting.

function FilePickBox({ label, hint, file, onFile, accentColor="#1a56db" }) {
  const ref = useRef();
  const [dragging, setDragging] = useState(false);
  return (
    <div
      onDragOver={e=>{e.preventDefault();setDragging(true);}}
      onDragLeave={()=>setDragging(false)}
      onDrop={e=>{e.preventDefault();setDragging(false);onFile(e.dataTransfer.files[0]);}}
      onClick={()=>ref.current.click()}
      style={{
        border:`2px dashed ${dragging?accentColor:file?"#86efac":"#d1d5db"}`,
        borderRadius:10, background:dragging?"#eff6ff":file?"#f0fdf4":"#f9fafb",
        padding:"22px 16px", textAlign:"center", cursor:"pointer", transition:"all 0.15s", flex:1,
      }}
    >
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}}
        onChange={e=>{onFile(e.target.files[0]);e.target.value="";}} />
      <div style={{fontSize:13,fontWeight:600,color:"#374151",marginBottom:3}}>{label}</div>
      <div style={{fontSize:11,color:"#9ca3af",marginBottom:10}}>{hint}</div>
      {file ? (
        <div style={{fontSize:12,color:"#16a34a",fontWeight:600}}>✓ {file.name}</div>
      ) : (
        <div style={{display:"inline-block",background:accentColor,color:"#fff",padding:"6px 16px",borderRadius:6,fontSize:12,fontWeight:500}}>
          Choose file
        </div>
      )}
    </div>
  );
}

function UpdateModelsFlow({ onBack }) {
  const [stage, setStage] = useState("upload"); // upload | review | done
  const [existingFile, setExistingFile] = useState(null);
  const [existing, setExisting] = useState(null); // {headers, rows}
  const [libraryFile, setLibraryFile] = useState(null);
  const [library, setLibrary] = useState(null); // sheets[]
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [lastDownload, setLastDownload] = useState("full"); // "full" | "changed"
  const [tolerance, setTolerance] = useState(DEFAULT_CHANGE_TOLERANCE);

  function handleExistingFile(file) {
    if (!file) return;
    setError(null); setExistingFile(file); setExisting(null);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const parsed = parseWorkbookRows(new Uint8Array(e.target.result));
        if (parsed.rows.length===0) throw new Error("No data rows found in the current-model file.");
        setExisting(parsed);
      } catch(err) { setError(err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleLibraryFile(file) {
    if (!file) return;
    setError(null); setLibraryFile(file); setLibrary(null);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const sheets = parseLibraryWorkbook(new Uint8Array(e.target.result));
        if (sheets.length===0) throw new Error("Couldn't find a model matrix sheet (looked for a row with a 'Ticker' column).");
        setLibrary(sheets);
      } catch(err) { setError(err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function runCompare() {
    const res = computeLibraryUpdates(existing.rows, library, tolerance);
    setResult(res);
    setStage("review");
  }

  function handleToleranceChange(newTolerance) {
    setTolerance(newTolerance);
    if (existing && library) {
      setResult(computeLibraryUpdates(existing.rows, library, newTolerance));
    }
  }

  function handleDownload(changedOnly) {
    const baseName = existingFile.name.replace(/\.[^.]+$/,"");
    const changedModelNames = new Set(result.changes.map(c=>c.modelName));
    const rowsToExport = changedOnly
      ? result.updatedRows.filter(r => changedModelNames.has(r["* Model Name"]))
      : result.updatedRows;
    const suffix = changedOnly ? "_ChangedModelsOnly" : "_Updated";
    downloadXlsxWithHeaders(rowsToExport, existing.headers, `${baseName}${suffix}.xlsx`);
    setLastDownload(changedOnly ? "changed" : "full");
    setStage("done");
  }

  function reset() {
    setStage("upload"); setExistingFile(null); setExisting(null);
    setLibraryFile(null); setLibrary(null); setError(null); setResult(null);
  }

  const bothReady = existing && library;

  if (stage === "upload") {
    return (
      <div>
        <div style={{display:"flex",gap:12,marginBottom:14}}>
          <FilePickBox label="Current model export" hint="How the models are set up today (.xlsx)"
            file={existingFile} onFile={handleExistingFile} />
          <FilePickBox label="Model library / targets" hint="New targets to apply (.xlsx)"
            file={libraryFile} onFile={handleLibraryFile} accentColor="#0aa89c" />
        </div>
        {error && <div style={{marginBottom:14,background:"#fee2e2",border:"0.5px solid #fca5a5",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#991b1b"}}><strong>Error:</strong> {error}</div>}
        <div style={{background:"#f0f9ff",border:"0.5px solid #bae6fd",borderRadius:8,padding:"12px 16px",fontSize:12,color:"#0c4a6e",lineHeight:1.6,marginBottom:16}}>
          <strong style={{color:"#0369a1"}}>What this does:</strong> matches each model to its column in the library, recalculates Category and Class targets from it, and leaves bands, Security Sets, and anything the library doesn't cover untouched. You'll see every change before exporting.
        </div>
        <div style={{display:"flex",justifyContent:"space-between"}}>
          <button onClick={onBack} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 16px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back</button>
          <button onClick={runCompare} disabled={!bothReady}
            style={{background:bothReady?"#1a56db":"#93c5fd",border:"none",borderRadius:6,padding:"8px 20px",fontSize:13,fontWeight:600,color:"#fff",cursor:bothReady?"pointer":"default"}}>
            Compare & review →
          </button>
        </div>
      </div>
    );
  }

  if (stage === "review") {
    const { changes, skippedModels, matchedModelCount, unmatchedCategories, unmatchedClasses } = result;
    const grouped = {};
    changes.forEach(c => { (grouped[c.modelName] = grouped[c.modelName]||[]).push(c); });
    const modelNames = Object.keys(grouped);

    return (
      <div>
        <div style={{display:"flex",gap:16,marginBottom:16,alignItems:"stretch"}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,flex:1}}>
            {[
              ["Models matched", matchedModelCount, "#1a56db"],
              ["Models skipped", skippedModels.length, "#9ca3af"],
              ["Unmatched labels", unmatchedCategories.length+unmatchedClasses.length, unmatchedCategories.length+unmatchedClasses.length?"#dc2626":"#9ca3af"],
            ].map(([lbl,val,color])=>(
              <div key={lbl} style={{background:"#f9fafb",border:"0.5px solid #e5e7eb",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:19,fontWeight:700,color}}>{val}</div>
                <div style={{fontSize:11,color:"#6b7280"}}>{lbl}</div>
              </div>
            ))}
          </div>
          <div style={{background:"#f9fafb",border:"0.5px solid #e5e7eb",borderRadius:8,padding:"10px 12px",minWidth:150}}>
            <label style={{fontSize:11,color:"#6b7280",display:"block",marginBottom:4}}>Change tolerance (pp)</label>
            <input type="number" step="0.01" min="0" value={tolerance}
              onChange={e=>{
                const v = e.target.value === "" ? 0 : parseFloat(e.target.value);
                if (!isNaN(v)) handleToleranceChange(v);
              }}
              style={{width:"100%",border:"0.5px solid #d1d5db",borderRadius:6,padding:"4px 8px",fontSize:14,fontWeight:700,color:"#111827"}} />
          </div>
        </div>


        {modelNames.length===0 && (
          <div style={{background:"#f0fdf4",border:"0.5px solid #bbf7d0",borderRadius:8,padding:"14px",fontSize:13,color:"#166534",marginBottom:16}}>
            No target values differ from what's already in the current model file — nothing to change.
          </div>
        )}

        {modelNames.map(name => (
          <div key={name} style={{border:"0.5px solid #e5e7eb",borderRadius:8,marginBottom:10,overflow:"hidden"}}>
            <div style={{background:"#f9fafb",padding:"8px 12px",fontSize:12,fontWeight:700,color:"#111827",borderBottom:"0.5px solid #e5e7eb"}}>{name}</div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <tbody>
                {grouped[name].map((c,i)=>{
                  const delta = Math.abs(c.newVal - (parseFloat(c.oldVal)||0));
                  const notable = delta > 0.5;
                  return (
                  <tr key={i} style={{borderTop:i>0?"0.5px solid #f3f4f6":"none", background:notable?"#fffbeb":"transparent"}}>
                    <td style={{padding:"6px 12px",color:"#6b7280",width:70}}>{c.level}</td>
                    <td style={{padding:"6px 12px",color:"#374151"}}>
                      {c.label}
                      {c.note && <span style={{marginLeft:6,fontSize:10,color:"#6b7280",fontStyle:"italic"}}>({c.note})</span>}
                      {notable && <span style={{marginLeft:6,fontSize:9,fontWeight:700,color:"#b45309",background:"#fef3c7",borderRadius:4,padding:"1px 5px"}}>Δ {delta.toFixed(2)}pt</span>}
                    </td>
                    <td style={{padding:"6px 12px",color:"#dc2626",textAlign:"right",width:70}}>{c.oldVal ?? "—"}%</td>
                    <td style={{padding:"6px 4px",color:"#9ca3af",width:20}}>→</td>
                    <td style={{padding:"6px 12px",color:"#16a34a",fontWeight:700,textAlign:"right",width:70}}>{c.newVal}%</td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        ))}

        {(skippedModels.length>0 || unmatchedCategories.length>0 || unmatchedClasses.length>0) && (
          <details style={{marginTop:8,marginBottom:20}}>
            <summary style={{cursor:"pointer",fontSize:12,fontWeight:600,color:"#6b7280"}}>
              Not updated — {skippedModels.length} model(s), {unmatchedCategories.length+unmatchedClasses.length} unmatched label(s)
            </summary>
            <div style={{marginTop:8,padding:"10px 12px",background:"#f9fafb",border:"0.5px solid #e5e7eb",borderRadius:8,fontSize:11,color:"#6b7280",lineHeight:1.7}}>
              {skippedModels.length>0 && (
                <div style={{marginBottom:8}}>
                  <strong style={{color:"#374151"}}>Not matched to the library</strong> (kept as-is):
                  <ul style={{margin:"4px 0 0",paddingLeft:18}}>
                    {skippedModels.map((s,i)=>(
                      <li key={i} style={{marginBottom:2}}><strong style={{color:"#374151"}}>{s.modelName}</strong> — {s.reason}</li>
                    ))}
                  </ul>
                </div>
              )}
              {unmatchedCategories.length>0 && (
                <div style={{marginBottom:8}}>
                  <strong style={{color:"#374151"}}>Category not found in library:</strong>{" "}
                  {unmatchedCategories.map((u,i)=>`${u.modelName} → ${u.label}`).join("; ")}
                </div>
              )}
              {unmatchedClasses.length>0 && (
                <div>
                  <strong style={{color:"#374151"}}>Class not found in library:</strong>{" "}
                  {unmatchedClasses.map((u,i)=>`${u.modelName} → ${u.label}`).join("; ")}
                </div>
              )}
            </div>
          </details>
        )}

        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderTop:"0.5px solid #e5e7eb",paddingTop:16}}>
          <button onClick={()=>setStage("upload")} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 16px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back</button>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>handleDownload(true)} disabled={modelNames.length===0}
              style={{background:"none",border:`0.5px solid ${modelNames.length?"#1a56db":"#d1d5db"}`,borderRadius:6,padding:"8px 18px",fontSize:13,fontWeight:600,color:modelNames.length?"#1a56db":"#9ca3af",cursor:modelNames.length?"pointer":"default"}}>
              Export changed models only ↓
            </button>
            <button onClick={()=>handleDownload(false)} style={{background:"#1a56db",border:"none",borderRadius:6,padding:"8px 20px",fontSize:13,fontWeight:600,color:"#fff",cursor:"pointer"}}>
              Export full file ↓
            </button>
          </div>
        </div>
      </div>
    );
  }

  // stage === "done"
  return (
    <div style={{textAlign:"center",padding:"48px 24px"}}>
      <div style={{width:56,height:56,borderRadius:"50%",background:"#dcfce7",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}>
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
      </div>
      <div style={{fontSize:18,fontWeight:700,color:"#111827",marginBottom:6}}>File downloaded</div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:6}}>
        {lastDownload==="changed"
          ? `Only the ${new Set(result.changes.map(c=>c.modelName)).size} model(s) with updated values were included.`
          : `All ${result.matchedModelCount + result.skippedModels.length} models included, ${result.changes.length} value(s) updated.`}
      </div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:28}}>Ready to re-import into Orion Eclipse</div>
      <div style={{display:"flex",gap:10,justifyContent:"center",flexWrap:"wrap"}}>
        <button onClick={()=>setStage("review")} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back to review</button>
        <button onClick={()=>handleDownload(true)} style={{background:"none",border:"0.5px solid #1a56db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#1a56db",cursor:"pointer"}}>Download changed only</button>
        <button onClick={()=>handleDownload(false)} style={{background:"none",border:"0.5px solid #1a56db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#1a56db",cursor:"pointer"}}>Download full file</button>
        <button onClick={reset} style={{background:"#1a56db",border:"none",borderRadius:6,padding:"8px 18px",fontSize:13,fontWeight:600,color:"#fff",cursor:"pointer"}}>Process another pair</button>
      </div>
    </div>
  );
}

// ── Advisor Custom Model (ACM) builder ──────────────────────────────────────
// Turns an advisor's raw (non-proportional) fund/model matrix into: (1) a
// readable digest of current allocations + computed proportional targets,
// editable and re-importable, then (2) a final Model + Security Set import
// pair. See engine notes inline — every rule here was verified against a
// real advisor file (Fortify) before being generalized.

const ACM_OUTLIER_Z = 2.4;
const ACM_ROUNDING_FALLBACK_THRESHOLD = 0.024; // pp; above this, use 1% grid instead of 5%

// Parses the standardized ACM import template: a header row with "Ticker",
// optional "Category"/"Class" columns (filled in directly, since that's now
// the source of truth — no more guessing from a label or a remembered
// lookup), then N model columns (optionally in adjacent Reg/"X NQ" pairs —
// tax-status variants can use different tickers entirely, e.g. munis only
// in NQ). No labels, subtotals, or blank-ticker rows to guess around; every
// row (including Cash) has a real ticker.
function parseAdvisorTemplate(buffer, sheetName) {
  const wb = XLSX.read(buffer, { type:"array" });
  const name = sheetName || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });

  let headerIdx = -1, tickerCol = -1;
  for (let i=0; i<Math.min(raw.length,10); i++) {
    const row = raw[i] || [];
    const idx = row.findIndex(c => /^ticker$/i.test((c||"").toString().trim()));
    if (idx !== -1) { headerIdx = i; tickerCol = idx; break; }
  }
  if (headerIdx === -1) throw new Error('Could not find a header row with "Ticker".');
  const headerRow = raw[headerIdx];
  const dataStart = headerIdx + 1;

  const categoryCol = headerRow.findIndex(h => /^category$/i.test((h||"").toString().trim()));
  const classCol = headerRow.findIndex(h => /^class$/i.test((h||"").toString().trim()));

  const modelCols = [];
  for (let i=0; i<headerRow.length; i++) {
    if (i===tickerCol || i===categoryCol || i===classCol) continue;
    const h = headerRow[i];
    if (h===null || h===undefined || h==="") continue;
    modelCols.push({ i, h: h.toString().trim() });
  }

  // Pair adjacent "X" / "X NQ" columns into families.
  const families = {};
  const consumed = new Set();
  for (let k=0; k<modelCols.length; k++) {
    const { i, h } = modelCols[k];
    if (consumed.has(i) || /nq$/i.test(h)) continue;
    families.Reg = families.Reg || [];
    families.Reg.push({ model:h, col:i });
    consumed.add(i);
    const next = modelCols[k+1];
    if (next && /nq$/i.test(next.h) && !consumed.has(next.i)) {
      families.NQ = families.NQ || [];
      families.NQ.push({ model:h, col:next.i });
      consumed.add(next.i);
    }
  }
  for (const { i, h } of modelCols) {
    if (consumed.has(i)) continue;
    families.NQ = families.NQ || [];
    families.NQ.push({ model:h.replace(/\s*nq$/i,"").trim(), col:i });
    consumed.add(i);
  }
  if (!families.Reg && families.NQ) { families.Reg = families.NQ; delete families.NQ; }
  if (Object.keys(families).length===1 && families.Reg) { families.Standard = families.Reg; delete families.Reg; }

  const securities = [];
  for (let r=dataStart; r<raw.length; r++) {
    const row = raw[r]; if (!row) continue;
    const tickerRaw = row[tickerCol];
    const ticker = tickerRaw!==null && tickerRaw!==undefined ? tickerRaw.toString().trim() : "";
    if (!ticker) continue;
    const hasAnyValue = Object.values(families).some(cols => cols.some(({col}) => typeof row[col]==="number"));
    if (!hasAnyValue) continue;

    const rawByFamily = {};
    for (const [fam, cols] of Object.entries(families)) {
      rawByFamily[fam] = {};
      cols.forEach(({model,col}) => { const v = row[col]; rawByFamily[fam][model] = typeof v==="number" ? v : 0; });
    }
    const categoryRaw = categoryCol>=0 && row[categoryCol] ? row[categoryCol].toString().trim() : "";
    const classRaw = classCol>=0 && row[classCol] ? row[classCol].toString().trim() : "";
    securities.push({
      ticker, raw:rawByFamily,
      category: categoryRaw ? categoryRaw.toUpperCase() : null,
      class: classRaw || null,
    });
  }

  const familyModelOrder = {};
  for (const [fam, cols] of Object.entries(families)) familyModelOrder[fam] = cols.map(c=>c.model);
  return { securities, familyModelOrder };
}

// Persistent ticker → {category, class} lookup, so a ticker only needs to be
// categorized once, ever — future imports (of any advisor) auto-fill from
// what's already been confirmed. Uses the same window.storage API the Model
// Audit Tool already relies on (real in the Claude artifact sandbox; backed
// by localStorage once deployed via storageShim.js).
const ACM_TICKER_LOOKUP_KEY = "acm-ticker-lookup";
async function loadTickerLookup() {
  try {
    const res = await window.storage.get(ACM_TICKER_LOOKUP_KEY);
    return res && res.value ? JSON.parse(res.value) : {};
  } catch { return {}; }
}
async function saveTickerLookup(lookup) {
  try { await window.storage.set(ACM_TICKER_LOOKUP_KEY, JSON.stringify(lookup)); } catch {}
}

// Proportional targets: zero-skip, then exclude statistical outliers (z>2.5,
// "judgment not rigor" per spec), average what's left, round to nearest 5% —
// falling back to nearest 1% for that whole sibling group if 5% would distort
// the true computed value by more than ~2.4pp. Verified against real data.
function acmRawGroupMeans(siblingRawArrays) {
  const numModels = siblingRawArrays[0].length;
  const parentTotals = Array.from({length:numModels}, (_,m) => siblingRawArrays.reduce((s,arr)=>s+(arr[m]||0),0));
  return siblingRawArrays.map(arr => {
    const ratios = [];
    for (let m=0;m<numModels;m++) if (parentTotals[m]>1e-9) ratios.push(arr[m]/parentTotals[m]);
    if (ratios.length===0) return 0;
    const mean = ratios.reduce((a,b)=>a+b,0)/ratios.length;
    const stdev = Math.sqrt(ratios.reduce((a,b)=>a+(b-mean)**2,0)/ratios.length);
    let kept = ratios;
    if (stdev>0 && ratios.length>=4) {
      kept = ratios.filter(r => Math.abs((r-mean)/stdev) <= ACM_OUTLIER_Z);
      if (kept.length===0) kept = ratios;
    }
    return kept.reduce((a,b)=>a+b,0)/kept.length;
  });
}
function acmRoundAndReconcile(fractions, inc) {
  const rounded = fractions.map(f => Math.round(f/inc)*inc);
  let diffSteps = Math.round((1 - rounded.reduce((a,b)=>a+b,0))/inc);
  if (diffSteps===0) return rounded.map(v=>+v.toFixed(4));
  const dist = fractions.map(f => Math.abs(((f/inc)%1) - 0.5));
  const order = fractions.map((_,i)=>i).sort((a,b)=>dist[a]-dist[b]);
  const result = [...rounded];
  const step = diffSteps>0 ? inc : -inc;
  let remaining = Math.abs(diffSteps), i=0;
  while (remaining>0) { result[order[i%order.length]] += step; remaining--; i++; }
  return result.map(v=>+v.toFixed(4));
}
function computeProportionalTargets(siblingRawArrays) {
  if (siblingRawArrays.length===1) return { targets:[1], avgs:[1] };
  const rawMeans = acmRawGroupMeans(siblingRawArrays);
  const sum = rawMeans.reduce((a,b)=>a+b,0);
  const normalized = sum>0 ? rawMeans.map(t=>t/sum) : rawMeans.map(()=>1/siblingRawArrays.length);
  const rounded5 = normalized.map(f => Math.round(f/0.05)*0.05);
  const maxDeviation = Math.max(...normalized.map((f,i)=>Math.abs(f-rounded5[i])));
  const increment = maxDeviation > ACM_ROUNDING_FALLBACK_THRESHOLD ? 0.01 : 0.05;
  return { targets: acmRoundAndReconcile(normalized, increment), avgs: normalized };
}

// Builds the per-family Category > Class > Ticker tree with current (raw)
// totals per model and computed proportional targets. Categories are never
// smoothed (they're the model's true risk-profile definition); Class and
// Ticker levels get smoothed only when they have 2+ siblings under the same
// parent — a single child just inherits 100% of its parent, no smoothing
// possible or needed. avgPct is the raw computed average *before* rounding —
// kept alongside targetPct so it's visible for troubleshooting.
//
// Builds trees for ALL families at once (rather than one family in
// isolation) so that Security Sets sharing the exact same ticker
// composition across families (e.g. US Equity/International/Alternatives
// usually hold identical funds regardless of tax status — only Fixed Income
// commonly differs, e.g. munis-only in NQ) can pool their model data before
// running the z-score/average computation. Small per-family sample sizes
// (Reg's 9 models vs NQ's 8, say) were causing the exact same underlying
// ratio to land on opposite sides of the outlier threshold depending only on
// which family it was judged against — pooling gives one shared, more
// statistically grounded answer instead, applied identically to every
// family that has that same Security Set. Groups whose ticker composition
// genuinely differs between families (different bond tickers, etc.) are
// still computed independently per family, same as before.
function buildAcmFamilyTrees(familyData) {
  const famNames = Object.keys(familyData);

  const famGroups = {}; // famName -> { catName -> { clsKey -> [securities] } }
  famNames.forEach(fam => {
    const byCategory = {};
    familyData[fam].securities.forEach(s => {
      const cat = s.category || "UNCATEGORIZED";
      const cls = s.class || "__direct__";
      byCategory[cat] = byCategory[cat] || {};
      byCategory[cat][cls] = byCategory[cat][cls] || [];
      byCategory[cat][cls].push(s);
    });
    famGroups[fam] = byCategory;
  });

  // A ticker/class present with 0% across every model in a family isn't
  // really "in" that family — e.g. the template lists DFNM in both a Reg and
  // NQ column pair, but if Reg's advisor never uses munis, DFNM is 0% there
  // the whole way across. Treating it as a real (if small) sibling would
  // both give it a nonsensical nonzero weight and falsely make the ticker
  // set "look the same" across families for merge-eligibility purposes, so
  // it's excluded from both — it just gets 0% directly, no averaging.
  const secTotal = (fam, sec) => familyData[fam].models.reduce((s,m)=>s+(sec.raw[m]||0),0);
  const isActiveSec = (fam, sec) => secTotal(fam, sec) > 1e-9;
  const classTotal = (fam, catName, cn) => famGroups[fam][catName][cn].reduce((s,sec)=>s+secTotal(fam,sec),0);
  const isActiveClass = (fam, catName, cn) => classTotal(fam, catName, cn) > 1e-9;
  const tickerSig = (fam, secs) => secs.filter(s=>isActiveSec(fam,s)).map(s=>s.ticker).slice().sort().join("|");

  const allCatNames = new Set();
  famNames.forEach(fam => Object.keys(famGroups[fam]).forEach(c=>allCatNames.add(c)));

  const trees = {};
  famNames.forEach(fam => { trees[fam] = { models: familyData[fam].models, categories: [] }; });

  allCatNames.forEach(catName => {
    const famsWithCat = famNames.filter(f => famGroups[f][catName]);

    // Class-level merge check: do all families with this category define
    // the exact same set of ACTIVE (nonzero) class names?
    const activeClassNameSets = famsWithCat.map(f =>
      Object.keys(famGroups[f][catName]).filter(cn=>isActiveClass(f,catName,cn)).slice().sort().join("|"));
    const classSetsMatch = famsWithCat.length>1 && activeClassNameSets.every(s=>s===activeClassNameSets[0]) && activeClassNameSets[0]!=="";
    let classTargetsShared=null, classAvgsShared=null, classNamesOrdered=null;
    if (classSetsMatch) {
      classNamesOrdered = Object.keys(famGroups[famsWithCat[0]][catName]).filter(cn=>isActiveClass(famsWithCat[0],catName,cn));
      if (classNamesOrdered.length>1 && classNamesOrdered[0]!=="__direct__") {
        const combinedArrays = classNamesOrdered.map(cn => {
          let combined = [];
          famsWithCat.forEach(f => {
            const models = familyData[f].models;
            const totals = models.map(m => famGroups[f][catName][cn].reduce((s,sec)=>s+(sec.raw[m]||0),0));
            combined = combined.concat(totals);
          });
          return combined;
        });
        ({ targets: classTargetsShared, avgs: classAvgsShared } = computeProportionalTargets(combinedArrays));
      }
    }

    famsWithCat.forEach(fam => {
      const models = familyData[fam].models;
      const classes = famGroups[fam][catName];
      const allClassNames = Object.keys(classes);
      const activeClassNames = allClassNames.filter(cn=>isActiveClass(fam,catName,cn));
      const catTotals = models.map(m => Object.values(classes).flat().reduce((s,sec)=>s+(sec.raw[m]||0),0));

      let classTargets, classAvgs;
      if (classSetsMatch && classTargetsShared) {
        classTargets = classTargetsShared; classAvgs = classAvgsShared;
      } else if (activeClassNames.length>1 && activeClassNames[0]!=="__direct__") {
        const classArrays = activeClassNames.map(cn => models.map(m => classes[cn].reduce((s,sec)=>s+(sec.raw[m]||0),0)));
        ({ targets: classTargets, avgs: classAvgs } = computeProportionalTargets(classArrays));
      } else {
        classTargets = activeClassNames.map(()=>1); classAvgs = activeClassNames.map(()=>1);
      }
      const activeClassIdx = cn => (classSetsMatch && classTargetsShared) ? classNamesOrdered.indexOf(cn) : activeClassNames.indexOf(cn);

      const classNodes = allClassNames.map((cn) => {
        const secs = classes[cn];
        const classTotals = models.map(m => secs.reduce((s,sec)=>s+(sec.raw[m]||0),0));
        const active = isActiveClass(fam, catName, cn);
        const cIdx = activeClassIdx(cn);
        const clsTargetPct = active && cIdx>=0 ? classTargets[cIdx] : 0;
        const clsAvgPct = active && cIdx>=0 ? classAvgs[cIdx] : 0;

        // Ticker-level merge check, using only ACTIVE tickers per family.
        const activeSecs = secs.filter(s=>isActiveSec(fam,s));
        const famsWithThisClass = famNames.filter(f => famGroups[f][catName] && famGroups[f][catName][cn]);
        const activeTickerSigs = famsWithThisClass.map(f => tickerSig(f, famGroups[f][catName][cn]));
        const tickersMatch = famsWithThisClass.length>1 && activeTickerSigs.every(s=>s===activeTickerSigs[0]) && activeTickerSigs[0]!=="";

        let tickerTargets=[], tickerAvgs=[], tickerNamesOrdered=null;
        if (activeSecs.length>1) {
          if (tickersMatch) {
            tickerNamesOrdered = famGroups[famsWithThisClass[0]][catName][cn].filter(s=>isActiveSec(famsWithThisClass[0],s)).map(s=>s.ticker);
            const combined = tickerNamesOrdered.map(()=>[]);
            famsWithThisClass.forEach(f => {
              const fModels = familyData[f].models;
              const fSecs = famGroups[f][catName][cn];
              tickerNamesOrdered.forEach((tname, ti) => {
                const sec = fSecs.find(s=>s.ticker===tname);
                combined[ti] = combined[ti].concat(fModels.map(m => sec ? (sec.raw[m]||0) : 0));
              });
            });
            ({ targets: tickerTargets, avgs: tickerAvgs } = computeProportionalTargets(combined));
          } else {
            const arrays = activeSecs.map(s => models.map(m => s.raw[m]||0));
            ({ targets: tickerTargets, avgs: tickerAvgs } = computeProportionalTargets(arrays));
          }
        } else if (activeSecs.length===1) {
          tickerTargets = [1]; tickerAvgs = [1];
        }

        return {
          name: cn==="__direct__" ? catName : cn,
          isDirect: cn==="__direct__",
          targetPct: clsTargetPct,
          avgPct: clsAvgPct,
          totals: classTotals,
          tickers: secs.map((s) => {
            if (!isActiveSec(fam, s)) return { ticker: s.ticker, targetPct: 0, avgPct: 0, currentByModel: models.map(m=>s.raw[m]||0) };
            const tIdx = (tickersMatch && tickerNamesOrdered) ? tickerNamesOrdered.indexOf(s.ticker) : activeSecs.indexOf(s);
            return {
              ticker: s.ticker,
              targetPct: tickerTargets[tIdx],
              avgPct: tickerAvgs[tIdx],
              currentByModel: models.map(m => s.raw[m]||0),
            };
          }),
        };
      });

      trees[fam].categories.push({ name: catName, totals: catTotals, classes: classNodes });
    });
  });

  return trees;
}

// ── Digest export (Step 2): current + suggested + difference, one sheet per
// family — styled to match the advisor's own reference workbook (black
// header bars, brown category shading, tan class shading, thin borders,
// percent formatting). "Target %" is the single editable source of truth
// per ticker/class; Current/Suggested/Diff per model are reference columns
// computed FROM that share × the category's actual raw weight for that
// model, so re-importing just needs the Category/Class/Ticker/Target %
// columns. Families are named after the advisor, not "Reg"/"NQ" internally
// (e.g. "Fortify Wealth" / "Fortify Wealth (NQ)").
function acmFamilyDisplayName(fam, advisorName) {
  const base = advisorName || "Advisor";
  if (/^nq$/i.test(fam)) return `${base} (NQ)`;
  return base;
}

const ACM_STYLE = {
  black: "FF000000", white: "FFFFFFFF", categoryFill: "FF8F7E57", classFill: "FFC7BDA1",
};

// Writes one sheet's worth of rows (Category/Class/Ticker/Label + a value per
// model) with the reference file's visual scheme: black header bar, brown
// category rows, tan class rows, thin ticker borders. Category/Class cells
// are left blank on ticker rows (relying on the colored bars above for
// grouping, not repeated text) — merged across A:B on category/class summary
// rows so it reads cleanly instead of looking staggered.
function acmColLetter(n) { // 1-indexed column number -> Excel column letter
  let s = "";
  while (n > 0) { const m = (n-1)%26; s = String.fromCharCode(65+m)+s; n = Math.floor((n-1)/26); }
  return s;
}

// Writes one sheet's worth of rows (Category/Class/Ticker + a value per
// model) with the reference file's visual scheme: black header bar, brown
// category rows, tan class rows, thin ticker borders. Category/Class cells
// are left blank on ticker rows (relying on the colored bars above for
// grouping, not repeated text) — merged across A:B on category/class summary
// rows so it reads cleanly instead of looking staggered.
//
// getValue may return either a plain fraction, or a {formula, result} object
// for a live Excel formula (result is the cached fallback value shown before
// any recalculation). The context passed to getValue includes selfCatRow/
// selfClsRow — the row numbers already written for this category/class
// *within this same sheet* — so a caller building a formula-based sheet can
// self-reference rows just written above without needing an external map.
// The returned rowMap (keyed "cat", "cat|cls", or "cat|cls|ticker") lets a
// *later* sheet build cross-sheet formulas pointing back at this one.
// Excel sheet names are capped at 31 characters, so a long advisor/family
// name plus a " (Suggested)"/" Difference"/" (NQ)" suffix can truncate down
// to the exact same string for what were meant to be different sheets —
// exceljs throws "Worksheet name already exists" in that case. This
// guarantees uniqueness by shortening further and appending a disambiguator
// (~2, ~3, ...) rather than just hoping names stay short enough.
function acmUniqueSheetName(rawName, usedNames) {
  const base = rawName.replace(/[\\/*?:[\]]/g,"");
  const first = base.slice(0,31);
  if (!usedNames.has(first)) { usedNames.add(first); return first; }
  for (let i=2; i<1000; i++) {
    const suffix = `~${i}`;
    const candidate = base.slice(0, 31-suffix.length) + suffix;
    if (!usedNames.has(candidate)) { usedNames.add(candidate); return candidate; }
  }
  throw new Error(`Could not generate a unique sheet name for "${rawName}" — too many collisions.`);
}

function acmWriteSheet(wb, sheetName, title, models, categories, getValue, extraCols, usedNames) {
  const ws = wb.addWorksheet(acmUniqueSheetName(sheetName, usedNames || new Set()));
  extraCols = extraCols || []; // [{header, getValue({level,cat,cls,t,mi})}] — placed after the model columns
  const totalCols = 1 + models.length + extraCols.length;
  const modelColStart = 2;
  const extraColStart = modelColStart + models.length;

  // Title bar: filled across the full width but NOT merged (unmerged cells
  // sort/filter more predictably in Excel) — text sits in A1 only.
  for (let c=1; c<=totalCols; c++) {
    const cell = ws.getCell(1,c);
    cell.fill = { type:"pattern", pattern:"solid", fgColor:{argb:ACM_STYLE.black} };
  }
  const titleCell = ws.getCell(1,1);
  titleCell.value = title;
  titleCell.font = { name:"Arial", size:14, bold:true, color:{argb:ACM_STYLE.white} };
  titleCell.alignment = { horizontal:"left", vertical:"middle" };
  ws.getRow(1).height = 22;

  const headerRowIdx = 2;
  ["Ticker", ...models, ...extraCols.map(c=>c.header)].forEach((h, ci) => {
    const cell = ws.getCell(headerRowIdx, ci+1);
    cell.value = h;
    cell.font = { name:"Arial", size:10, bold:true, color:{argb:ACM_STYLE.white} };
    cell.fill = { type:"pattern", pattern:"solid", fgColor:{argb:ACM_STYLE.black} };
    cell.alignment = { horizontal:"center", vertical:"middle" };
  });
  ws.getRow(headerRowIdx).height = 20;

  const pctCell = (r,c,val) => {
    const cell = ws.getCell(r,c);
    if (val && typeof val==="object" && "formula" in val) cell.value = { formula: val.formula, result: val.result };
    else cell.value = val;
    cell.numFmt = "0.0%";
    cell.alignment = { horizontal:"center" };
  };
  const fillRow = (r, argb) => { for (let c=1;c<=totalCols;c++) ws.getCell(r,c).fill = { type:"pattern", pattern:"solid", fgColor:{argb} }; };
  const writeExtras = (r, ctx) => extraCols.forEach((ec,i) => pctCell(r, extraColStart+i, ec.getValue(ctx)));
  const rowMap = new Map();

  let r = headerRowIdx + 1;
  categories.forEach(cat => {
    fillRow(r, ACM_STYLE.categoryFill);
    const catCell = ws.getCell(r,1);
    catCell.value = cat.name;
    catCell.font = { name:"Arial", size:12, bold:true, color:{argb:ACM_STYLE.white} };
    const catRow = r;
    rowMap.set(cat.name, catRow);
    models.forEach((m,mi) => pctCell(r, modelColStart+mi, getValue({ level:"category", cat, mi, selfCatRow:catRow })));
    r++;

    cat.classes.forEach(cls => {
      let clsRow = catRow; // isDirect: the "class row" is the category row itself
      if (!cls.isDirect) {
        fillRow(r, ACM_STYLE.classFill);
        const clsCell = ws.getCell(r,1);
        clsCell.value = cls.name;
        clsCell.font = { name:"Arial", size:10, bold:true, color:{argb:"FF000000"} };
        clsRow = r;
        rowMap.set(`${cat.name}|${cls.name}`, clsRow);
        models.forEach((m,mi) => pctCell(r, modelColStart+mi, getValue({ level:"class", cat, cls, mi, selfCatRow:catRow })));
        writeExtras(r, { level:"class", cat, cls });
        r++;
      }
      cls.tickers.forEach(t => {
        const tickerCell = ws.getCell(r,1);
        tickerCell.value = t.ticker;
        tickerCell.alignment = { horizontal:"center" };
        tickerCell.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
        rowMap.set(`${cat.name}|${cls.name}|${t.ticker}`, r);
        models.forEach((m,mi) => pctCell(r, modelColStart+mi, getValue({ level:"ticker", cat, cls, t, mi, selfCatRow:catRow, selfClsRow:clsRow })));
        writeExtras(r, { level:"ticker", cat, cls, t });
        r++;
      });
    });
  });

  ws.getColumn(1).width = 10;
  for (let mi=0; mi<models.length; mi++) ws.getColumn(modelColStart+mi).width = 13;
  for (let i=0; i<extraCols.length; i++) ws.getColumn(extraColStart+i).width = 11;
  ws.views = [{ state:"frozen", xSplit:1, ySplit:headerRowIdx }];
  return { ws, sheetName: ws.name, rowMap, modelColStart, extraColStart, headerRowIdx };
}


// Generates the standard blank ACM import template (Step 1's expected shape)
// — a starter file to hand to whoever reformats an advisor's raw data, with
// one filled-in example model column, several "ENTER MODEL" placeholders to
// rename, and two example ticker rows demonstrating the fill pattern.
async function downloadAcmTemplate() {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Template");
  const numPlaceholders = 16;
  const totalCols = 4 + numPlaceholders; // Ticker, Category, Class, one example model + placeholders

  ws.mergeCells(1,1,1,totalCols);
  const title = ws.getCell(1,1);
  title.value = "Advisor Custom Model Import Builder Template";
  title.font = { name:"Arial", size:14, bold:true, color:{argb:ACM_STYLE.white} };
  title.fill = { type:"pattern", pattern:"solid", fgColor:{argb:ACM_STYLE.black} };
  title.alignment = { horizontal:"left", vertical:"middle" };
  ws.getRow(1).height = 22;

  ["TICKER","Category","Class"].forEach((h, i) => {
    const cell = ws.getCell(2, i+1);
    cell.value = h;
    cell.font = { name:"Arial", size:10, bold:true, color:{argb:ACM_STYLE.white} };
    cell.fill = { type:"pattern", pattern:"solid", fgColor:{argb:ACM_STYLE.categoryFill} };
    cell.alignment = { horizontal:"center", vertical:"middle" };
  });

  const modelHeaders = ["STP - Moderate", ...Array(numPlaceholders).fill("ENTER MODEL")];
  modelHeaders.forEach((h, i) => {
    const cell = ws.getCell(2, i+4);
    cell.value = h;
    cell.font = { name:"Arial", size:10, bold:true, color:{argb:"FF000000"} };
    cell.fill = { type:"pattern", pattern:"solid", fgColor:{argb:ACM_STYLE.classFill} };
    cell.alignment = { horizontal:"center", vertical:"middle" };
  });
  ws.getRow(2).height = 20;

  [["APPL","Equity","US Equity",0.5],["IBIT","Equity","US Equity",0.5]].forEach(([ticker,cat,cls,val], ri) => {
    const r = 3+ri;
    ws.getCell(r,1).value = ticker;
    ws.getCell(r,1).alignment = { horizontal:"center" };
    ws.getCell(r,1).border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
    ws.getCell(r,2).value = cat;
    ws.getCell(r,3).value = cls;
    ws.getCell(r,4).value = val;
  });

  ws.getColumn(1).width = 10; ws.getColumn(2).width = 14; ws.getColumn(3).width = 14;
  for (let i=0; i<modelHeaders.length; i++) ws.getColumn(i+4).width = 15;
  ws.views = [{ state:"frozen", xSplit:3, ySplit:2 }];

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type:"application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "ModelImportTemplate.xlsx";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function downloadAcmDigest(familyTrees, advisorName) {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  const usedNames = new Set(); // shared across every sheet in this workbook, so truncated names never collide

  Object.entries(familyTrees).forEach(([fam, tree]) => {
    const displayName = acmFamilyDisplayName(fam, advisorName);
    const { models, categories } = tree;
    const ref = (sheetName, col, row) => `'${sheetName}'!${col}${row}`;

    // Current tab — every level (category/class/ticker) expressed as a %
    // of the WHOLE model, consistently, so a class's number visibly equals
    // the sum of its tickers' numbers, and a category's equals the sum of
    // its classes'. Plus Target % (editable — this is what re-import reads
    // back, and what the Suggested/Difference formulas below key off of) and
    // Avg % (the raw computed average before rounding, for troubleshooting).
    const current = acmWriteSheet(wb, displayName, displayName, models, categories, ({level, cat, cls, t, mi}) => {
      if (level==="category") return cat.totals[mi];
      if (level==="class") return cls.totals[mi];
      return t.currentByModel[mi];
    }, [
      { header:"Target %", getValue: ({level,cls,t}) => level==="ticker" ? t.targetPct : (cls ? cls.targetPct : null) },
      { header:"Avg %", getValue: ({level,cls,t}) => level==="ticker" ? t.avgPct : (cls ? cls.avgPct : null) },
    ], usedNames);
    const targetColLetter = acmColLetter(current.extraColStart);
    const modelColLetter = mi => acmColLetter(current.modelColStart+mi);

    // Suggested tab — LIVE FORMULAS cascading from the category's actual
    // (unsmoothed) raw weight in the Current tab through each Target % cell
    // there, so editing a Target % on the Current tab (e.g. to adjust
    // rounding) recalculates this tab automatically, no re-export needed.
    // Class rows reference Current's category row × Current's own Target %
    // for that class; ticker rows reference either Current's category row
    // (if the class is "direct") or *this sheet's own* just-written class
    // row (self-reference) × Current's Target % for that ticker.
    const suggested = acmWriteSheet(wb, `${displayName} (Suggested)`, `${displayName} — Suggested`, models, categories,
      ({level, cat, cls, t, mi, selfClsRow}) => {
        const catRow = current.rowMap.get(cat.name);
        const catRef = ref(current.sheetName, modelColLetter(mi), catRow);
        const catResult = cat.totals[mi];
        if (level==="category") return { formula: catRef, result: catResult };

        if (level==="class") {
          if (cls.isDirect) return { formula: catRef, result: catResult };
          const targetRow = current.rowMap.get(`${cat.name}|${cls.name}`);
          const targetRef = ref(current.sheetName, targetColLetter, targetRow);
          return { formula: `${catRef}*${targetRef}`, result: catResult*cls.targetPct };
        }

        // ticker
        const clsResult = cls.isDirect ? catResult : catResult*cls.targetPct;
        const parentRef = cls.isDirect ? catRef : `${modelColLetter(mi)}${selfClsRow}`; // self-sheet ref when not direct
        const tickerTargetRow = current.rowMap.get(`${cat.name}|${cls.name}|${t.ticker}`);
        const tickerTargetRef = ref(current.sheetName, targetColLetter, tickerTargetRow);
        return { formula: `${parentRef}*${tickerTargetRef}`, result: clsResult*t.targetPct };
      }, null, usedNames);

    // Difference tab — LIVE FORMULA, Suggested minus Current, so it also
    // updates automatically when Target % changes.
    acmWriteSheet(wb, `${displayName} Difference`, `${displayName} — Difference (Suggested − Current)`, models, categories,
      ({level, cat, cls, t, mi}) => {
        if (level==="category") return { formula: "0", result: 0 };

        const key = cls.isDirect ? cat.name : `${cat.name}|${cls.name}`;
        const catResult = cat.totals[mi];
        const clsSuggested = cls.isDirect ? catResult : catResult*cls.targetPct;

        if (level==="class") {
          const sugRow = suggested.rowMap.get(key), curRow = current.rowMap.get(key);
          const sugRef = ref(suggested.sheetName, modelColLetter(mi), sugRow);
          const curRef = ref(current.sheetName, modelColLetter(mi), curRow);
          return { formula: `${sugRef}-${curRef}`, result: clsSuggested - cls.totals[mi] };
        }

        // ticker
        const tKey = `${cat.name}|${cls.name}|${t.ticker}`;
        const sugRow = suggested.rowMap.get(tKey), curRow = current.rowMap.get(tKey);
        const sugRef = ref(suggested.sheetName, modelColLetter(mi), sugRow);
        const curRef = ref(current.sheetName, modelColLetter(mi), curRow);
        return { formula: `${sugRef}-${curRef}`, result: clsSuggested*t.targetPct - t.currentByModel[mi] };
      }, null, usedNames);
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type:"application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `${advisorName || "Advisor"}_ACM_Digest.xlsx`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Re-imports a (possibly advisor-edited) digest file. Reads the "Current"
// tabs specifically (skips any "(Suggested)"/"Difference" tabs, which are
// derived reference-only). Row type (category/class/ticker) is read from
// each row's fill color — the same brown/tan/none used when writing the
// sheet — since there's no longer a Label column to infer it from text case.
function parseAcmDigestForReimport(buffer) {
  const wb = XLSX.read(buffer, { type:"array", cellStyles:true });
  const families = {};
  wb.SheetNames.forEach(sheetName => {
    if (/\(Suggested\)$/.test(sheetName) || / Difference$/.test(sheetName)) return; // reference-only tabs, skip
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    if (raw.length < 2) return;
    // Row 0 is the title bar; the actual column headers are row 1.
    const header = raw[1].map(h=>(h||"").toString());
    const idx = { ticker: header.indexOf("Ticker"), target: header.indexOf("Target %") };
    if (idx.ticker===-1 || idx.target===-1) return; // not a recognizable digest sheet
    const models = header.slice(idx.ticker+1, idx.target).filter(Boolean);
    const modelColFor = (modelIdx) => idx.ticker + 1 + modelIdx;
    // Values are stored as fractions (0.6 with a "60.0%" display format) — ×100 to get plain percentages.
    // Digest cells already store native fractions (0.5 = 50%) — read as-is.
    const asPct = v => typeof v==="number" ? v : 0;
    const fillOf = (r) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c: idx.ticker })];
      return cell && cell.s && cell.s.fgColor ? cell.s.fgColor.rgb : null;
    };

    let currentCategory = null, currentClass = null;
    const categoryTotals = {}; // catName -> [val per model]
    const classTargets = {};   // catName -> { clsName -> targetPct }
    const tickers = [];        // { category, class, ticker, targetPct }

    for (let r=2; r<raw.length; r++) {
      const row = raw[r]; if (!row) continue;
      const text = row[idx.ticker] ? row[idx.ticker].toString().trim() : "";
      const target = asPct(row[idx.target]);
      if (!text) continue;

      const fill = fillOf(r);
      if (fill === ACM_STYLE.categoryFill.slice(2)) {
        // Category header row: pull each model's value as the true (unsmoothed) category total.
        currentCategory = text; currentClass = null;
        categoryTotals[text] = models.map((_,mi) => asPct(row[modelColFor(mi)]));
        continue;
      }
      if (fill === ACM_STYLE.classFill.slice(2)) {
        // Class header row: Target % is that class's fixed share of its category.
        currentClass = text;
        classTargets[currentCategory] = classTargets[currentCategory] || {};
        classTargets[currentCategory][text] = target;
        continue;
      }
      tickers.push({ category: currentCategory, class: currentClass, ticker: text, targetPct: target });
    }
    families[sheetName] = { models, categoryTotals, classTargets, tickers };
  });
  return families;
}

// ── Final export (Step 3): Model import (reusing TEMPLATE_COLS) + Security
// Set import (Orion's dedicated Security Set template). Security Sets = the
// Class-level groupings; one Security Set per Class, shared across every
// model in a family since ticker composition doesn't vary by risk level —
// only each model's Category/Class target percentages do.
const ACM_SS_TEMPLATE_COLS = [
  "Security Set ID","Name","Team ID","Security Set Do Not TLH","Security Do Not TLH",
  "Equivalent Of(symbol)","Dynamic","Description","Symbol","Rank","Allocation %","Range",
  "Fix Band %","Lower Tol %","Upper Tol %","Min Trade Amt","Min Initial Buy $",
  "Alternate Custodian 1","T - Alternate 1","T - Min Trade Amt 1","T - Min Initial Buy Amt 1",
  "TD - Alternate 1","TD - Min Trade Amt 1","TD - Min Initial Buy Amt 1",
  "TE - Alternate 1","TE - Min Trade Amt 1","TE - Min Initial Buy Amt 1",
  "Alternate Custodian 2","T - Alternate 2","T - Min Trade Amt 2","T - Min Initial Buy Amt 2",
  "TD - Alternate 2","TD - Min Trade Amt 2","TD - Min Initial Buy Amt 2",
  "TE - Alternate 2","TE - Min Trade Amt 2","TE - Min Initial Buy Amt 2",
  "Alternate Custodian 3","T - Alternate 3","T - Min Trade Amt 3","T - Min Initial Buy Amt 3",
  "TD - Alternate 3","TD - Min Trade Amt 3","TD - Min Initial Buy Amt 3",
  "TE - Alternate 3","TE - Min Trade Amt 3","TE - Min Initial Buy Amt 3",
  "Buy Priority","Sell Priority",
  "Security Set TLH Symbol 1","Security Set TLH Custodian 1","Security Set TLH Priority 1",
  "Security Set TLH Symbol 2","Security Set TLH Custodian 2","Security Set TLH Priority 2",
  "Security Set TLH Symbol 3","Security Set TLH Custodian 3","Security Set TLH Priority 3",
  "TLH Symbol 1","TLH Custodian 1","TLH Priority 1","TLH Symbol 2","TLH Custodian 2","TLH Priority 2",
  "TLH Symbol 3","TLH Custodian 3","TLH Priority 3",
  "Group Equivalence Type 1","Group Equivalent 1","Group Buy Priority 1","Group Sell Priority 1",
  "Group Equivalence Type 2","Group Equivalent 2","Group Buy Priority 2","Group Sell Priority 2",
  "Group Equivalence Type 3","Group Equivalent 3","Group Buy Priority 3","Group Sell Priority 3",
];

// Strips a family's tax-status suffix (e.g. "(NQ)") to get its base name —
// used when two families' Security Sets turn out identical, so the shared
// name doesn't awkwardly carry one family's suffix.
// Categories are stored internally as ALL CAPS (EQUITY, FIXED INCOME) for
// matching/grouping — this is purely for display in the final export.
function acmTitleCase(s) {
  return (s||"").toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function acmBaseFamilyName(familyName) {
  return familyName.replace(/\s*\(NQ\)\s*$/i, "").trim();
}

// Reference fields confirmed against a real Orion export: Category-level
// bands follow a fixed business rule by category type — Equity/Fixed Income
// (the main categories) always get a flat ±5 percentage-point band
// regardless of target size, while Alternatives/Cash (smaller satellite
// categories) get a "100% of target" relative band instead (so a 3% target
// gets a ±3pt band, a 1% target gets ±1pt). Unrecognized/custom categories
// default to the flat-band treatment. Class-level bands are always a flat
// 25%-of-target relative band, and Security Set-level bands are always
// Fix Band % = 50 (from the separate Security Set import).
// Categories/classes that map to an already-existing Orion Security Set
// (shared across every advisor, not created per-advisor) rather than a new
// one this tool would normally generate — referenced by ID + its real bare
// name, and never written to the Security Set import since it already exists.
const ACM_EXISTING_SECURITY_SETS = {
  "CASH": { id: 57, name: "Cash" },
};

const ACM_CATEGORY_META = {
  "EQUITY": { display: "Equity", bandType: "fixed5" },
  "FIXED INCOME": { display: "Fixed Income", bandType: "fixed5" },
  "ALTERNATIVES": { display: "Alternatives", bandType: "relative100" },
  "CASH": { display: "Cash & Cash Equivalents", bandType: "relative100" },
};
function acmCategoryMeta(catKey) {
  return ACM_CATEGORY_META[catKey] || { display: acmTitleCase(catKey), bandType: "fixed5" };
}

function buildAcmFinalExport(reimportedFamilies) {
  const modelRows = [];

  // ── Pass 1: determine, per (Category, Class), whether its ticker+
  // allocation composition is IDENTICAL across 2+ families. If so, every
  // family shares ONE Security Set (named after the base/stripped family
  // name) instead of each getting its own — this is exactly the point of a
  // shared Security Set: edit it once, every model that references it
  // updates together. Classes that genuinely differ (e.g. taxable vs. muni
  // bond tickers) keep separate, family-specific Security Sets. The Class
  // (and, when every class in it is shared, the Category) SubModel Name
  // follows the same shared/not-shared call, so the naming never implies a
  // tax-status distinction that the underlying Security Set doesn't actually
  // have.
  const classSignature = (catName, clsKey, tickers) => {
    const tks = tickers.filter(t => t.category===catName && (t.class||"__direct__")===clsKey);
    return tks.map(t => `${t.ticker}:${(+t.targetPct.toFixed(6))}`).sort().join(",");
  };

  const sigToFamilies = new Map(); // "catName|clsKey|signature" -> Set of families sharing that exact signature

  Object.entries(reimportedFamilies).forEach(([familyName, { tickers }]) => {
    const seen = new Set();
    tickers.forEach(t => {
      const clsKey = t.class || "__direct__";
      const key = `${t.category}|${clsKey}`;
      if (seen.has(key)) return;
      seen.add(key);
      const sig = classSignature(t.category, clsKey, tickers);
      const fullKey = `${key}|${sig}`;
      if (!sigToFamilies.has(fullKey)) sigToFamilies.set(fullKey, new Set());
      sigToFamilies.get(fullKey).add(familyName);
    });
  });

  // Returns the shared base family name if 2+ families have this exact
  // (Category, Class) composition, otherwise this family's own full name.
  function sharedOrOwnFamilyName(familyName, catName, clsKey, tickers) {
    const sig = classSignature(catName, clsKey, tickers);
    const fullKey = `${catName}|${clsKey}|${sig}`;
    const sharingFamilies = sigToFamilies.get(fullKey);
    if (sharingFamilies && sharingFamilies.size > 1) {
      return acmBaseFamilyName([...sharingFamilies].sort()[0]);
    }
    return familyName;
  }

  // ── Pass 2: build Model rows (Category/Class SubModel Names are
  // family-specific but NOT model-specific — the number varies per model
  // row already, the name doesn't need to) and Security Set rows (deduped
  // by their final consolidated name).
  const ssRowsByName = new Map();

  Object.entries(reimportedFamilies).forEach(([familyName, { models, categoryTotals, classTargets, tickers }]) => {
    const byCategory = {};
    tickers.forEach(t => {
      byCategory[t.category] = byCategory[t.category] || {};
      const clsKey = t.class || "__direct__";
      byCategory[t.category][clsKey] = byCategory[t.category][clsKey] || [];
      byCategory[t.category][clsKey].push(t);
    });

    models.forEach((modelName, mi) => {
      const isNQFamily = /\(NQ\)/i.test(familyName);
      const fullModelName = `Strategic Advisor Model - ${acmBaseFamilyName(familyName)} ${modelName}${isNQFamily ? " (NQ)" : ""}`;
      Object.entries(byCategory).forEach(([catName, classes]) => {
        const catTotal = (categoryTotals[catName] && categoryTotals[catName][mi]) || 0;
        if (catTotal <= 1e-9) return; // 0% category for this model — nothing to allocate, skip entirely
        const catDisplay = acmTitleCase(catName);
        const classKeys = Object.keys(classes);
        // Category is only shared (no family suffix) when EVERY class in it
        // is independently shared too — a category with any genuinely
        // family-specific class (like Fixed Income's differing bonds) keeps
        // its own family-specific name.
        const allClassesShared = classKeys.length>0 && classKeys.every(ck => {
          const sig = classSignature(catName, ck, tickers);
          const fam = sigToFamilies.get(`${catName}|${ck}|${sig}`);
          return fam && fam.size > 1;
        });
        const categoryPrefix = allClassesShared ? acmBaseFamilyName(familyName) : familyName;

        Object.entries(classes).forEach(([clsKey, tks]) => {
          const isDirect = clsKey === "__direct__";
          const clsName = isDirect ? catName : clsKey;
          const clsDisplay = isDirect ? catDisplay : clsKey;
          const clsTargetFrac = isDirect ? 1 : ((classTargets[catName] && classTargets[catName][clsName]) ?? 0);
          const classPrefix = sharedOrOwnFamilyName(familyName, catName, clsKey, tickers);

          // Cash always uses Orion's existing shared "Cash" Security Set
          // (ID 57) rather than creating a new advisor-specific one — the
          // Category/Class SubModel Names stay advisor-specific as usual,
          // only the Security Set reference points at the pre-built one.
          const existing = ACM_EXISTING_SECURITY_SETS[catName];
          const ssName = existing ? existing.name : `${classPrefix} - ${clsDisplay}`;
          const ssId = existing ? existing.id : null;

          const catPct = +(catTotal*100).toFixed(2);
          const clsPct = +(clsTargetFrac*100).toFixed(2);
          const meta = acmCategoryMeta(catName);
          const catBand = meta.bandType==="relative100" ? { band:100, upper:catPct, lower:catPct } : { band:null, upper:5, lower:5 };
          const clsBand = { band:25, upper:+(clsPct*0.25).toFixed(2), lower:+(clsPct*0.25).toFixed(2) };

          modelRows.push({
            "* Model Name": fullModelName,
            "* Security Set ID": ssId,
            "Category SubModel Name": `${categoryPrefix} - ${catDisplay}`,
            "Category Asset Class Type": meta.display,
            "Category Namespace": "Default Team",
            "Category Target %": catPct,
            "Category Band/Range": catBand.band,
            "Category Upper %": catBand.upper,
            "Category Lower %": catBand.lower,
            "Class SubModel Name": `${classPrefix} - ${clsDisplay}`,
            "Class Namespace": "Default Team",
            "Class Target %": clsPct,
            "Class Band/Range": clsBand.band,
            "Class Upper %": clsBand.upper,
            "Class Lower %": clsBand.lower,
            "Subclass Namespace": "Default Team",
            "* Security Set SubModel Name": ssName,
            "* Security Set Target %": 100,
            "Security Set Band/Range": 25,
            "Security Set Upper %": 25,
            "Security Set Lower %": 25,
            "* Dynamic": "NO",
            "* Name Space": "Default Team",
          });

          if (!existing) {
            tks.forEach(t => {
              if (!ssRowsByName.has(ssName+"|"+t.ticker) && t.targetPct > 0) {
                ssRowsByName.set(ssName+"|"+t.ticker, {
                  "Name": ssName, "Symbol": t.ticker, "Allocation %": +(t.targetPct*100).toFixed(2),
                  "Fix Band %": 50, "Dynamic": 0,
                  "Security Set Do Not TLH": "false", "Security Do Not TLH": "false",
                  "Buy Priority": "Default", "Sell Priority": "Default",
                });
              }
            });
          }
        });
      });
    });
  });

  const ssRowsFull = [...ssRowsByName.values()].map(partial => {
    const full = {};
    ACM_SS_TEMPLATE_COLS.forEach(c => full[c] = partial[c] !== undefined ? partial[c] : null);
    return full;
  });

  return { modelRows, ssRows: ssRowsFull };
}

function downloadAcmFinalExport(reimportedFamilies, advisorName) {
  const { modelRows, ssRows } = buildAcmFinalExport(reimportedFamilies);
  downloadXlsxWithHeaders(modelRows, TEMPLATE_COLS, `${advisorName||"Advisor"}_ACM_Models.xlsx`);
  downloadXlsxWithHeaders(ssRows, ACM_SS_TEMPLATE_COLS, `${advisorName||"Advisor"}_ACM_SecuritySets.xlsx`);
}

// ── ACM UI ───────────────────────────────────────────────────────────────────

function AcmFlow({ onBack }) {
  const [stage, setStage] = useState("upload"); // upload (parses, categorizes, computes, exports) | reimport | done
  const [advisorName, setAdvisorName] = useState("");
  const [rawFile, setRawFile] = useState(null);
  const [parsed, setParsed] = useState(null); // {securities, familyModelOrder}
  const [categorized, setCategorized] = useState(null); // securities with .category/.class set
  const [error, setError] = useState(null);
  const [familyTrees, setFamilyTrees] = useState(null);
  const [reimportedFamilies, setReimportedFamilies] = useState(null);
  const [reimportFile, setReimportFile] = useState(null);

  const [exporting, setExporting] = useState(false);

  function handleRawFile(file) {
    if (!file) return;
    setError(null); setRawFile(file);
    const reader = new FileReader();
    reader.onload = async e => {
      try {
        const result = parseAdvisorTemplate(new Uint8Array(e.target.result));
        if (result.securities.length===0) throw new Error("No securities found — check that the sheet has a Ticker column.");
        setParsed(result);

        const lookup = await loadTickerLookup();
        const merged = result.securities.map(s => {
          // The template's own Category/Class (if filled in) is the source
          // of truth; the remembered lookup only fills gaps for tickers the
          // template left blank.
          if (s.category) return s;
          const known = lookup[s.ticker];
          return known ? { ...s, category: known.category, class: known.class ?? null } : s;
        });
        const missing = merged.filter(s => !s.category);
        if (missing.length > 0) {
          setCategorized(merged);
          throw new Error(`${missing.length} ticker${missing.length!==1?"s":""} missing a Category (not filled in on the template, and not seen before): ${missing.map(s=>s.ticker).join(", ")}. Add Category/Class to the template and re-upload.`);
        }
        setCategorized(merged);

        // Everything's categorized — go straight to compute + export, no
        // review screen needed since Category/Class already came from the
        // template itself. Build all families' trees together so Security
        // Sets sharing identical tickers across families can pool their
        // sample for the z-score/average computation.
        const familyData = {};
        Object.entries(result.familyModelOrder).forEach(([fam, models]) => {
          familyData[fam] = { securities: merged.map(s => ({ ...s, raw: s.raw[fam] })), models };
        });
        const trees = buildAcmFamilyTrees(familyData);
        setFamilyTrees(trees);
        setExporting(true);
        try {
          merged.forEach(s => { lookup[s.ticker] = { category: s.category, class: s.class }; });
          await saveTickerLookup(lookup);
          await downloadAcmDigest(trees, advisorName);
        } finally {
          setExporting(false);
        }
        setStage("reimport");
      } catch (err) { setError(err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleReimportFile(file) {
    if (!file) return;
    setError(null); setReimportFile(file);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const families = parseAcmDigestForReimport(new Uint8Array(e.target.result));
        if (Object.keys(families).length===0) throw new Error("Couldn't find recognizable Category/Class/Ticker/Target % columns in this file.");
        setReimportedFamilies(families);
      } catch (err) { setError(err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function finalize() {
    downloadAcmFinalExport(reimportedFamilies, advisorName);
    setStage("done");
  }

  if (stage === "upload") {
    return (
      <div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Advisor / model family name</label>
          <input value={advisorName} onChange={e=>setAdvisorName(e.target.value)} placeholder="e.g. Fortify Wealth"
            style={{width:"100%",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 10px",fontSize:13}} />
        </div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <label style={{fontSize:12,color:"#6b7280"}}>Advisor model, in the standard template</label>
          <button onClick={downloadAcmTemplate} style={{fontSize:11,color:"#7c3aed",background:"none",border:"none",cursor:"pointer",padding:0,textDecoration:"underline"}}>
            Download template ↓
          </button>
        </div>
        <FilePickBox hint="Ticker + Category/Class + model columns, reformatted from whatever the advisor sent (.xlsx)"
          file={rawFile} onFile={handleRawFile} accentColor="#7c3aed" />
        {exporting && <div style={{marginTop:14,fontSize:13,color:"#7c3aed"}}>Computing targets & exporting digest…</div>}
        {error && <div style={{marginTop:14,background:"#fee2e2",border:"0.5px solid #fca5a5",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#991b1b"}}><strong>Error:</strong> {error}</div>}
        <div style={{marginTop:16,background:"#f5f3ff",border:"0.5px solid #ddd6fe",borderRadius:8,padding:"12px 16px",fontSize:12,color:"#4c1d95",lineHeight:1.6}}>
          Category/Class are read directly from the template — remembers any ticker's assignment after the first time, so future imports (any advisor) auto-fill anything already seen. Goes straight from upload to a computed, exported digest — no separate review step needed. Computes a proportional target weight per group, excluding zero and statistically extreme models per your judgment call rather than strict stats.
        </div>
        <div style={{marginTop:16}}>
          <button onClick={onBack} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 16px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back</button>
          <button onClick={()=>{setError(null);setStage("reimport");}} style={{marginLeft:12,background:"none",border:"none",fontSize:12,color:"#7c3aed",cursor:"pointer",textDecoration:"underline"}}>
            Already have a digest file? Skip to upload it →
          </button>
        </div>
      </div>
    );
  }

  if (stage === "reimport") {
    return (
      <div>
        <div style={{fontSize:13,color:"#374151",marginBottom:16}}>
          Upload the digest file — edited by the advisor or not — to build the final Model and Security Set import files.
        </div>
        <FilePickBox label="Digest file" hint="The exported digest (.xlsx), possibly with Target % adjusted"
          file={reimportFile} onFile={handleReimportFile} accentColor="#7c3aed" />
        {error && <div style={{marginTop:14,background:"#fee2e2",border:"0.5px solid #fca5a5",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#991b1b"}}><strong>Error:</strong> {error}</div>}
        {reimportedFamilies && (
          <div style={{marginTop:14,background:"#f0fdf4",border:"0.5px solid #bbf7d0",borderRadius:8,padding:"12px 16px",fontSize:12,color:"#166534"}}>
            Loaded {Object.values(reimportedFamilies).reduce((s,f)=>s+f.tickers.length,0)} ticker rows across {Object.keys(reimportedFamilies).length} famil{Object.keys(reimportedFamilies).length===1?"y":"ies"}: {Object.keys(reimportedFamilies).join(", ")}.
          </div>
        )}
        <div style={{display:"flex",justifyContent:"space-between",marginTop:16}}>
          <button onClick={()=>setStage("upload")} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 16px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back</button>
          <div style={{display:"flex",gap:8}}>
            {familyTrees && <button onClick={()=>downloadAcmDigest(familyTrees, advisorName)} style={{background:"none",border:"0.5px solid #7c3aed",borderRadius:6,padding:"8px 16px",fontSize:13,color:"#7c3aed",cursor:"pointer"}}>Download digest again</button>}
            <button onClick={finalize} disabled={!reimportedFamilies}
              style={{background:reimportedFamilies?"#7c3aed":"#c4b5fd",border:"none",borderRadius:6,padding:"8px 20px",fontSize:13,fontWeight:600,color:"#fff",cursor:reimportedFamilies?"pointer":"default"}}>
              Export final Model + Security Set files ↓
            </button>
          </div>
        </div>
      </div>
    );
  }

  // done
  return (
    <div style={{textAlign:"center",padding:"48px 24px"}}>
      <div style={{width:56,height:56,borderRadius:"50%",background:"#dcfce7",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}>
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
      </div>
      <div style={{fontSize:18,fontWeight:700,color:"#111827",marginBottom:6}}>Files downloaded</div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:28}}>Model import + Security Set import — ready to bring into Orion Eclipse</div>
      <div style={{display:"flex",gap:10,justifyContent:"center"}}>
        <button onClick={()=>setStage("reimport")} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back</button>
        <button onClick={()=>downloadAcmFinalExport(reimportedFamilies, advisorName)} style={{background:"none",border:"0.5px solid #7c3aed",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#7c3aed",cursor:"pointer"}}>Download again</button>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────

export default function OrionImportBuilder() {
  const [mode, setMode] = useState(null); // null | "import" | "update"
  const [stage, setStage] = useState("upload");
  const [allRows, setAllRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState(null);
  const [dragging, setDragging] = useState(false);
  const fileRef = useRef();

  const processFile = useCallback((file) => {
    if (!file) return;
    setError(null);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseExcel(new Uint8Array(e.target.result));
        if (parsed.length===0) throw new Error("No data rows found.");
        setAllRows(parsed);
        setStage("edit");
      } catch(err) { setError(err.message); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false);
    processFile(e.dataTransfer.files[0]);
  }, [processFile]);

  const handleDownload = () => {
    const out = buildOutputRows(allRows);
    const baseName = fileName.replace(/\.[^.]+$/,"");
    downloadXlsx(out, `${baseName}_ImportReady.xlsx`);
    setStage("done");
  };

  const reset = () => { setStage("upload"); setAllRows([]); setFileName(""); setError(null); };

  const modelGroups = groupByModel(allRows);
  const modelCount = Object.keys(modelGroups).length;

  // Validation across all groups
  function hasWarnings() {
    // check every category's children sum to 100, etc — quick check
    return false; // tree components handle their own validation visually
  }

  const STEPS = [["1","Upload"],["2","Edit"],["3","Download"]];
  const stageIdx = {upload:0,edit:1,done:2}[stage];

  return (
    <div style={{fontFamily:"system-ui,sans-serif",maxWidth:900,margin:"0 auto",padding:"24px 16px"}}>

      {/* Header */}
      <div style={{marginBottom:24}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"linear-gradient(135deg,#1a56db,#0ea5e9)",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
              <line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/>
            </svg>
          </div>
          <div>
            <div style={{fontSize:18,fontWeight:700,color:"#111827"}}>Orion Import Builder</div>
            <div style={{fontSize:12,color:"#6b7280"}}>
              {mode===null ? "Upload · visualize · edit · export" : mode==="import" ? "New import file" : mode==="update" ? "Update existing models" : "Advisor Custom Model builder"}
            </div>
          </div>
        </div>
        {mode!==null && (
          <button onClick={()=>{setMode(null);setStage("upload");setAllRows([]);setFileName("");setError(null);}}
            style={{fontSize:11,color:"#6b7280",background:"none",border:"none",cursor:"pointer",padding:0,textDecoration:"underline",marginBottom:4}}>
            ← Choose a different workflow
          </button>
        )}
        {mode==="import" && (
        <div style={{display:"flex",alignItems:"center",gap:0}}>
          {STEPS.map(([num,lbl],i)=>(
            <div key={num} style={{display:"flex",alignItems:"center"}}>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                <div style={{width:22,height:22,borderRadius:"50%",background:i<stageIdx?"#1a56db":i===stageIdx?"#1a56db":"#e5e7eb",color:i<=stageIdx?"#fff":"#9ca3af",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
                  {i<stageIdx?"✓":num}
                </div>
                <span style={{fontSize:12,fontWeight:i===stageIdx?600:400,color:i===stageIdx?"#111827":i<stageIdx?"#6b7280":"#9ca3af"}}>{lbl}</span>
              </div>
              {i<2&&<div style={{width:32,height:1,background:"#e5e7eb",margin:"0 8px"}}/>}
            </div>
          ))}
        </div>
        )}
      </div>

      {/* ── Mode selection ── */}
      {mode===null && (
        <div style={{display:"flex",gap:16}}>
          <div onClick={()=>setMode("update")}
            style={{flex:1,border:"0.5px solid #e5e7eb",borderRadius:12,padding:"28px 20px",cursor:"pointer",background:"#fff",transition:"all 0.15s"}}
            onMouseEnter={e=>e.currentTarget.style.borderColor="#0aa89c"} onMouseLeave={e=>e.currentTarget.style.borderColor="#e5e7eb"}>
            <div style={{width:36,height:36,borderRadius:9,background:"#f0fdfa",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:14}}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0aa89c" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
                <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
              </svg>
            </div>
            <div style={{fontSize:15,fontWeight:700,color:"#111827",marginBottom:6}}>Update existing models</div>
            <div style={{fontSize:12,color:"#6b7280",lineHeight:1.6}}>Upload your current model export plus a model library of new targets — get back an updated file with just the changed values.</div>
          </div>
          <div onClick={()=>setMode("acm")}
            style={{flex:1,border:"0.5px solid #e5e7eb",borderRadius:12,padding:"28px 20px",cursor:"pointer",background:"#fff",transition:"all 0.15s"}}
            onMouseEnter={e=>e.currentTarget.style.borderColor="#7c3aed"} onMouseLeave={e=>e.currentTarget.style.borderColor="#e5e7eb"}>
            <div style={{width:36,height:36,borderRadius:9,background:"#f5f3ff",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:14}}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
              </svg>
            </div>
            <div style={{fontSize:15,fontWeight:700,color:"#111827",marginBottom:6}}>Build Advisor Custom Model</div>
            <div style={{fontSize:12,color:"#6b7280",lineHeight:1.6}}>Turn an advisor's raw, non-proportional model file into proportional targets, then export Model + Security Set import files.</div>
          </div>
        </div>
      )}

      {/* ── Update Existing Models workflow ── */}
      {mode==="update" && <UpdateModelsFlow onBack={()=>setMode(null)} />}

      {/* ── Advisor Custom Model workflow ── */}
      {mode==="acm" && <AcmFlow onBack={()=>setMode(null)} />}

      {/* ── Import workflow (upload/edit/done) ── */}
      {mode==="import" && (
      <>
      {/* ── Upload ── */}
      {stage==="upload" && (
        <div>
          <div
            onDragOver={e=>{e.preventDefault();setDragging(true);}}
            onDragLeave={()=>setDragging(false)}
            onDrop={onDrop}
            onClick={()=>fileRef.current.click()}
            style={{border:`2px dashed ${dragging?"#1a56db":"#d1d5db"}`,borderRadius:12,background:dragging?"#eff6ff":"#f9fafb",padding:"48px 24px",textAlign:"center",cursor:"pointer",transition:"all 0.15s"}}
          >
            <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke={dragging?"#1a56db":"#9ca3af"} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{margin:"0 auto 12px",display:"block"}}>
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            <div style={{fontSize:15,fontWeight:600,color:"#374151",marginBottom:4}}>Drop your Excel file here</div>
            <div style={{fontSize:13,color:"#6b7280",marginBottom:16}}>or click to browse — any Orion model export (.xlsx)</div>
            <div style={{display:"inline-block",background:"#1a56db",color:"#fff",padding:"8px 20px",borderRadius:6,fontSize:13,fontWeight:500}}>Choose file</div>
          </div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={e=>{processFile(e.target.files[0]);e.target.value="";}} style={{display:"none"}}/>
          {error&&<div style={{marginTop:14,background:"#fee2e2",border:"0.5px solid #fca5a5",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#991b1b"}}><strong>Error:</strong> {error}</div>}
          <div style={{marginTop:20,background:"#f0f9ff",border:"0.5px solid #bae6fd",borderRadius:8,padding:"12px 16px",fontSize:12,color:"#0c4a6e",lineHeight:1.6}}>
            <strong style={{color:"#0369a1"}}>What this tool does:</strong> Upload any Orion model Excel export. The tool parses all rows, shows an interactive node tree for each model, lets you edit targets and bands, then exports a ready-to-import .xlsx.
          </div>
        </div>
      )}

      {/* ── Edit / Tree View ── */}
      {stage==="edit" && (
        <div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
            <div>
              <span style={{fontSize:14,fontWeight:600,color:"#111827"}}>{modelCount} model{modelCount!==1?"s":""} · {allRows.length} rows</span>
              <span style={{fontSize:12,color:"#6b7280",marginLeft:8}}>from {fileName}</span>
            </div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={reset} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"7px 14px",fontSize:13,color:"#374151",cursor:"pointer"}}>
                ← New file
              </button>
              <button onClick={handleDownload} style={{background:"#1a56db",border:"none",borderRadius:6,padding:"7px 18px",fontSize:13,fontWeight:600,color:"#fff",cursor:"pointer"}}>
                Export import file ↓
              </button>
            </div>
          </div>

          {Object.entries(modelGroups).map(([name, rows]) => (
            <div key={name} style={{marginBottom:36}}>
              <ModelTree
                modelName={name}
                rows={rows}
                onRowsChange={(updatedRows) => {
                  // Replace this model's rows in allRows
                  setAllRows(prev => {
                    const out = [...prev];
                    let ri = 0;
                    prev.forEach((r,i)=>{
                      if (r["* Model Name"]===name) {
                        out[i]=updatedRows[ri++];
                      }
                    });
                    return out;
                  });
                }}
              />
            </div>
          ))}

          <div style={{borderTop:"0.5px solid #e5e7eb",paddingTop:16,display:"flex",justifyContent:"flex-end"}}>
            <button onClick={handleDownload} style={{background:"#1a56db",border:"none",borderRadius:6,padding:"9px 24px",fontSize:14,fontWeight:600,color:"#fff",cursor:"pointer"}}>
              Export import file ↓
            </button>
          </div>
        </div>
      )}

      {/* ── Done ── */}
      {stage==="done" && (
        <div style={{textAlign:"center",padding:"48px 24px"}}>
          <div style={{width:56,height:56,borderRadius:"50%",background:"#dcfce7",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}>
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
          </div>
          <div style={{fontSize:18,fontWeight:700,color:"#111827",marginBottom:6}}>File downloaded</div>
          <div style={{fontSize:13,color:"#6b7280",marginBottom:28}}>{modelCount} model{modelCount!==1?"s":""} · {allRows.length} rows — ready to import into Orion Eclipse</div>
          <div style={{display:"flex",gap:10,justifyContent:"center"}}>
            <button onClick={()=>setStage("edit")} style={{background:"none",border:"0.5px solid #d1d5db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#374151",cursor:"pointer"}}>← Back to editor</button>
            <button onClick={handleDownload} style={{background:"none",border:"0.5px solid #1a56db",borderRadius:6,padding:"8px 18px",fontSize:13,color:"#1a56db",cursor:"pointer"}}>Download again</button>
            <button onClick={reset} style={{background:"#1a56db",border:"none",borderRadius:6,padding:"8px 18px",fontSize:13,fontWeight:600,color:"#fff",cursor:"pointer"}}>Process another file</button>
          </div>
        </div>
      )}
      </>
      )}
    </div>
  );
}
