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
              {mode===null ? "Upload · visualize · edit · export" : mode==="import" ? "New import file" : "Update existing models"}
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
          <div onClick={()=>setMode("import")}
            style={{flex:1,border:"0.5px solid #e5e7eb",borderRadius:12,padding:"28px 20px",cursor:"pointer",background:"#fff",transition:"all 0.15s"}}
            onMouseEnter={e=>e.currentTarget.style.borderColor="#1a56db"} onMouseLeave={e=>e.currentTarget.style.borderColor="#e5e7eb"}>
            <div style={{width:36,height:36,borderRadius:9,background:"#eff6ff",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:14}}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1a56db" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
              </svg>
            </div>
            <div style={{fontSize:15,fontWeight:700,color:"#111827",marginBottom:6}}>Build new import file</div>
            <div style={{fontSize:12,color:"#6b7280",lineHeight:1.6}}>Upload an Orion export, edit targets and bands on an interactive tree, and export a ready-to-import .xlsx.</div>
          </div>
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
        </div>
      )}

      {/* ── Update Existing Models workflow ── */}
      {mode==="update" && <UpdateModelsFlow onBack={()=>setMode(null)} />}

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
