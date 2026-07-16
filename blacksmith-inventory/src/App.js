import { useState, useEffect } from "react";
import { initializeApp } from "firebase/app";
import {
  getFirestore, collection, doc, onSnapshot,
  updateDoc, deleteDoc, setDoc, getDocs, addDoc, serverTimestamp, query, orderBy
} from "firebase/firestore";
import * as XLSX from "xlsx";

// ── Firebase ──────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: process.env.REACT_APP_FIREBASE_API_KEY,
  authDomain: "blacksmith-ordering.firebaseapp.com",
  projectId: process.env.REACT_APP_FIREBASE_PROJECT_ID || "blacksmith-ordering",
  storageBucket: "blacksmith-ordering.firebasestorage.app",
  messagingSenderId: process.env.REACT_APP_FIREBASE_SENDER_ID || "627729632784",
  appId: process.env.REACT_APP_FIREBASE_APP_ID,
};
const firebaseApp = initializeApp(firebaseConfig, "inventory");
const db = getFirestore(firebaseApp);

// ── Config ────────────────────────────────────────────────────────────────────
const LOCATIONS = ["Bountiful", "Millcreek", "Midvale"];
const ROLES = [
  { id: "kitchen",   label: "Kitchen",        icon: "🍳" },
  { id: "bountiful", label: "Bountiful",      icon: "📍", location: "Bountiful" },
  { id: "millcreek", label: "Millcreek",      icon: "📍", location: "Millcreek" },
  { id: "midvale",   label: "Midvale",        icon: "📍", location: "Midvale"   },
];

const SECTIONS = ["Dry Storage","Fridge/Freezer","Paper Products","Tools","SWAG","Office","Employee","Other"];
const VENDORS  = ["Farr's","Shamrock","US Foods","Costco","Carpenter","Webstaurant","Amazon","Store","Orson Gygi","Blacksmith","Custom"];

// ── Helpers ───────────────────────────────────────────────────────────────────
function genId() { return Math.random().toString(36).slice(2,9).toUpperCase(); }
function fmtDate(ts) {
  if (!ts) return "";
  const d = ts.toDate ? ts.toDate() : new Date(ts);
  return d.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}) + " " +
    d.toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit"});
}
function getDocId(o) { return o._docId || o.id; }

// ── Styles ────────────────────────────────────────────────────────────────────
const S = {
  label: { display:"block", fontSize:11, fontWeight:700, color:"#6B7280", marginBottom:5, letterSpacing:"0.06em", textTransform:"uppercase" },
  input: { width:"100%", padding:"10px 12px", borderRadius:8, border:"1.5px solid #E5E7EB", fontSize:14, color:"#1C1917", background:"#fff", boxSizing:"border-box" },
  card:  { background:"#fff", borderRadius:12, padding:"14px 16px", marginBottom:8, boxShadow:"0 1px 5px rgba(0,0,0,.07)" },
};

// ── Shared components ─────────────────────────────────────────────────────────
function Toast({ msg, onClose }) {
  useEffect(()=>{ const t=setTimeout(onClose,2800); return ()=>clearTimeout(t); },[onClose]);
  return <div style={{ position:"fixed", bottom:28, left:"50%", transform:"translateX(-50%)",
    background:"#1C1917", color:"#fff", padding:"12px 24px", borderRadius:10,
    fontWeight:600, fontSize:14, zIndex:9999, boxShadow:"0 4px 24px rgba(0,0,0,.25)" }}>{msg}</div>;
}

function Spinner() {
  return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:"60vh", gap:16 }}>
      <div style={{ width:36, height:36, border:"3px solid #E5E7EB", borderTop:"3px solid #1C1917", borderRadius:"50%", animation:"spin 0.8s linear infinite" }}/>
      <div style={{ color:"#9CA3AF", fontSize:14 }}>Loading…</div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

function QtyInput({ value, onChange, min=0 }) {
  return (
    <div style={{ display:"flex", alignItems:"center", border:"1.5px solid #E5E7EB", borderRadius:8, overflow:"hidden", background:"#fff", flexShrink:0 }}>
      <button onClick={()=>onChange(Math.max(min, Number(value)-1))} style={{ width:32, height:38, border:"none", background:"#F3F4F6", color:"#374151", fontSize:18, cursor:"pointer", fontWeight:700, lineHeight:1 }}>−</button>
      <input type="number" value={value} min={min}
        onChange={e=>onChange(Math.max(min, Number(e.target.value)))}
        style={{ width:54, border:"none", textAlign:"center", fontSize:14, fontWeight:600, color:"#1C1917", padding:"8px 4px", outline:"none" }}/>
      <button onClick={()=>onChange(Number(value)+1)} style={{ width:32, height:38, border:"none", background:"#F3F4F6", color:"#374151", fontSize:18, cursor:"pointer", fontWeight:700, lineHeight:1 }}>+</button>
    </div>
  );
}

// ── Login ─────────────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [role, setRole] = useState("kitchen");
  return (
    <div style={{ minHeight:"100vh", background:"#F8F7F4", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'Inter',system-ui,sans-serif" }}>
      <div style={{ background:"#fff", borderRadius:18, padding:"44px 36px", width:340, boxShadow:"0 2px 32px rgba(0,0,0,.09)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:36 }}>
          <div style={{ width:40, height:40, background:"#1C1917", borderRadius:11, display:"flex", alignItems:"center", justifyContent:"center", fontSize:22 }}>🔨</div>
          <div>
            <div style={{ fontWeight:800, fontSize:19, color:"#1C1917" }}>Blacksmith</div>
            <div style={{ fontSize:12, color:"#9CA3AF" }}>Inventory</div>
          </div>
        </div>
        <label style={S.label}>Sign in as</label>
        <div style={{ display:"flex", flexDirection:"column", gap:8, marginBottom:24 }}>
          {ROLES.map(r=>(
            <button key={r.id} onClick={()=>setRole(r.id)} style={{
              padding:"13px 16px", borderRadius:10, border:"2px solid",
              borderColor:role===r.id?"#1C1917":"#E5E7EB",
              background:role===r.id?"#1C1917":"#fff",
              color:role===r.id?"#fff":"#374151",
              fontWeight:600, fontSize:14, cursor:"pointer",
              display:"flex", alignItems:"center", gap:10 }}>
              <span style={{ fontSize:18 }}>{r.icon}</span>{r.label}
              {role===r.id&&<span style={{ marginLeft:"auto" }}>✓</span>}
            </button>
          ))}
        </div>
        <button onClick={()=>onLogin(ROLES.find(r=>r.id===role))} style={{
          width:"100%", padding:"13px 0", background:"#1C1917", color:"#fff",
          border:"none", borderRadius:10, fontWeight:700, fontSize:15, cursor:"pointer" }}>
          Enter
        </button>
      </div>
    </div>
  );
}

// ── Kitchen: Inventory Item Editor ────────────────────────────────────────────
function ItemEditor({ item, onSave, onCancel }) {
  const [vals, setVals] = useState({
    name: item?.name||"",
    unit: item?.unit||"",
    section: item?.section||SECTIONS[0],
    purchaseLocation: item?.purchaseLocation||VENDORS[0],
    backupLocation: item?.backupLocation||"",
    notes: item?.notes||"",
  });
  const set = (k,v) => setVals(p=>({...p,[k]:v}));

  return (
    <div style={{ background:"#EFF6FF", border:"1.5px solid #BFDBFE", borderRadius:12, padding:"16px", marginBottom:8 }}>
      <div style={{ fontWeight:700, fontSize:14, color:"#1C1917", marginBottom:12 }}>{item ? "Edit Item" : "New Item"}</div>
      <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:8 }}>
        <div style={{ flex:2, minWidth:140 }}>
          <label style={S.label}>Item Name</label>
          <input value={vals.name} onChange={e=>set("name",e.target.value)} style={S.input} placeholder="e.g. Heavy Cream"/>
        </div>
        <div style={{ flex:1, minWidth:80 }}>
          <label style={S.label}>Unit</label>
          <input value={vals.unit} onChange={e=>set("unit",e.target.value)} style={S.input} placeholder="qt, lb…"/>
        </div>
      </div>
      <div style={{ marginBottom:8 }}>
        <label style={S.label}>Section</label>
        <select value={vals.section} onChange={e=>set("section",e.target.value)} style={S.input}>
          {SECTIONS.map(s=><option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:8 }}>
        <div style={{ flex:1 }}>
          <label style={S.label}>Purchase Location</label>
          <select value={vals.purchaseLocation} onChange={e=>set("purchaseLocation",e.target.value)} style={S.input}>
            {VENDORS.map(v=><option key={v} value={v}>{v}</option>)}
          </select>
        </div>
        <div style={{ flex:1 }}>
          <label style={S.label}>Backup Location</label>
          <select value={vals.backupLocation} onChange={e=>set("backupLocation",e.target.value)} style={S.input}>
            <option value="">None</option>
            {VENDORS.map(v=><option key={v} value={v}>{v}</option>)}
          </select>
        </div>
      </div>
      <div style={{ marginBottom:12 }}>
        <label style={S.label}>Notes (optional)</label>
        <input value={vals.notes} onChange={e=>set("notes",e.target.value)} style={S.input} placeholder="Any notes about this item…"/>
      </div>
      <div style={{ display:"flex", gap:8 }}>
        <button onClick={()=>onSave(vals)} style={{ flex:2, padding:"10px 0", borderRadius:8, border:"none", background:"#1C1917", color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer" }}>Save Item</button>
        <button onClick={onCancel} style={{ flex:1, padding:"10px 0", borderRadius:8, border:"1.5px solid #E5E7EB", background:"#fff", color:"#374151", fontWeight:600, fontSize:13, cursor:"pointer" }}>Cancel</button>
      </div>
    </div>
  );
}

// ── Kitchen: Master Inventory Catalog ─────────────────────────────────────────
function KitchenCatalog({ items, shopItems, onAddItem, onUpdateItem, onDeleteItem }) {
  const [editingId, setEditingId] = useState(null);
  const [addingNew, setAddingNew] = useState(false);
  const [activeTab, setActiveTab] = useState("master");

  const masterItems = items.filter(i=>!i.shopSpecific);
  const bountifulItems = shopItems.filter(i=>i.location==="Bountiful");
  const millcreekItems  = shopItems.filter(i=>i.location==="Millcreek");
  const midvaleItems    = shopItems.filter(i=>i.location==="Midvale");
  const kitchenItems    = shopItems.filter(i=>i.location==="Kitchen");

  const tabs = [
    { key:"master",    label:"General Items",   list:masterItems },
    { key:"bountiful", label:"Bountiful",        list:bountifulItems },
    { key:"millcreek", label:"Millcreek",        list:millcreekItems },
    { key:"midvale",   label:"Midvale",          list:midvaleItems },
    { key:"kitchen",   label:"Kitchen",          list:kitchenItems },
  ];

  const currentList = tabs.find(t=>t.key===activeTab)?.list||[];
  const bySection = currentList.reduce((acc,i)=>{ const s=i.section||"Other"; if(!acc[s])acc[s]=[]; acc[s].push(i); return acc; },{});
  const sections = SECTIONS.filter(s=>bySection[s]);

  const locationMap = { master:"", bountiful:"Bountiful", millcreek:"Millcreek", midvale:"Midvale", kitchen:"Kitchen" };
  const handleSave = async (vals) => {
    if (editingId) {
      await onUpdateItem(editingId, vals);
      setEditingId(null);
    } else {
      await onAddItem({ ...vals, shopSpecific: activeTab!=="master", location: locationMap[activeTab]||"" });
      setAddingNew(false);
    }
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <h2 style={{ fontSize:18, fontWeight:800, color:"#1C1917" }}>Inventory Catalog</h2>
        <button onClick={()=>setAddingNew(true)} style={{ padding:"8px 14px", borderRadius:8, border:"none", background:"#1C1917", color:"#fff", fontWeight:700, fontSize:12, cursor:"pointer" }}>+ Add Item</button>
      </div>

      {/* Section tabs */}
      <div style={{ display:"flex", gap:0, background:"#fff", borderRadius:"10px 10px 0 0", overflow:"hidden", border:"1px solid #F3F4F6", marginBottom:12, overflowX:"auto" }}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setActiveTab(t.key)} style={{
            padding:"10px 12px", border:"none", background:"none", whiteSpace:"nowrap",
            fontWeight:activeTab===t.key?700:500, fontSize:12,
            color:activeTab===t.key?"#1C1917":"#9CA3AF",
            borderBottom:activeTab===t.key?"2.5px solid #1C1917":"2.5px solid transparent",
            cursor:"pointer" }}>
            {t.label}
            {t.list.length>0&&<span style={{ marginLeft:4, fontSize:11, color:activeTab===t.key?"#1C1917":"#9CA3AF" }}>({t.list.length})</span>}
          </button>
        ))}
      </div>

      {addingNew && <ItemEditor onSave={handleSave} onCancel={()=>setAddingNew(false)}/>}

      {sections.length===0&&!addingNew&&(
        <div style={{ textAlign:"center", color:"#9CA3AF", padding:"40px 0" }}>No items yet. Add one above.</div>
      )}

      {sections.map(sec=>(
        <div key={sec} style={{ marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#9CA3AF", textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:6 }}>{sec}</div>
          {bySection[sec].map(item=>(
            <div key={getDocId(item)}>
              {editingId===getDocId(item)
                ? <ItemEditor item={item} onSave={handleSave} onCancel={()=>setEditingId(null)}/>
                : (
                  <div style={{ ...S.card, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <div>
                      <div style={{ fontWeight:600, fontSize:14, color:"#1C1917" }}>{item.name}</div>
                      <div style={{ fontSize:12, color:"#9CA3AF", marginTop:2 }}>
                        {item.unit} · {item.purchaseLocation}
                        {item.backupLocation&&` · ${item.backupLocation}`}
                      </div>
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <button onClick={()=>setEditingId(getDocId(item))} style={{ background:"none", border:"none", color:"#9CA3AF", fontSize:12, fontWeight:600, cursor:"pointer" }}>Edit</button>
                      {item.shopSpecific&&(
                        <button onClick={()=>onDeleteItem(getDocId(item))} style={{ background:"none", border:"none", color:"#EF4444", fontSize:12, fontWeight:600, cursor:"pointer" }}>Delete</button>
                      )}
                    </div>
                  </div>
                )
              }
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

// ── Inventory Count Form ──────────────────────────────────────────────────────
function InventoryCount({ user, masterItems, shopItems, thresholds, onSubmit, onUpdateThreshold }) {
  const location = user.location;

  // Combine master items + this shop's specific items
  const myItems = [
    ...masterItems,
    ...shopItems.filter(i=>i.location===location),
  ];

  const [counts, setCounts] = useState(()=>{
    const c={};
    myItems.forEach(it=>{ c[getDocId(it)]=""; });
    return c;
  });
  const [infoOpen, setInfoOpen] = useState(null);
  const [editThreshold, setEditThreshold] = useState(null); // itemId being edited
  const [thresholdEdit, setThresholdEdit] = useState({low:"",critical:""});
  const [submitting, setSubmitting] = useState(false);
  const [sessionNote, setSessionNote] = useState("");

  const getThreshold = (itemId) => thresholds[`${location}:${itemId}`] || { low:"", critical:"" };

  const getHighlight = (itemId, value) => {
    if(value===""||value===null) return null;
    const t = getThreshold(itemId);
    const v = Number(value);
    if(t.critical!==""&&v<=Number(t.critical)) return "critical";
    if(t.low!==""&&v<=Number(t.low)) return "low";
    return "ok";
  };

  const highlightStyle = (level) => {
    if(level==="critical") return { background:"#FEE2E2", border:"1.5px solid #FECACA" };
    if(level==="low")      return { background:"#FEF3C7", border:"1.5px solid #FDE68A" };
    return { background:"#F0FDF4", border:"1.5px solid #BBF7D0" };
  };

  const bySection = myItems.reduce((acc,i)=>{ const s=i.section||"Other"; if(!acc[s])acc[s]=[]; acc[s].push(i); return acc; },{});
  const sections = SECTIONS.filter(s=>bySection[s]);

  const filledCount = Object.values(counts).filter(v=>v!=="").length;
  const totalCount  = myItems.length;

  const handleSubmit = async () => {
    setSubmitting(true);
    const snapshot = myItems.map(it=>{
      const id=getDocId(it);
      const t=getThreshold(id);
      return {
        itemId:id, name:it.name, unit:it.unit, section:it.section,
        qty:counts[id]===""?null:Number(counts[id]),
        low:t.low, critical:t.critical,
        highlight:getHighlight(id,counts[id]),
      };
    });
    await onSubmit({ location, items:snapshot, note:sessionNote, filledCount, totalCount });
    setSubmitting(false);
  };

  return (
    <div style={{ paddingBottom:80 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div>
          <h2 style={{ fontSize:18, fontWeight:800, color:"#1C1917" }}>Inventory Count</h2>
          <div style={{ fontSize:12, color:"#9CA3AF", marginTop:2 }}>{filledCount}/{totalCount} items filled</div>
        </div>
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>
          <div style={{ width:80, height:6, background:"#E5E7EB", borderRadius:10, overflow:"hidden" }}>
            <div style={{ height:"100%", width:`${totalCount?Math.round(filledCount/totalCount*100):0}%`, background:"#1C1917", borderRadius:10, transition:"width 0.3s" }}/>
          </div>
          <span style={{ fontSize:11, color:"#6B7280" }}>{totalCount?Math.round(filledCount/totalCount*100):0}%</span>
        </div>
      </div>

      {/* Legend */}
      <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
        {[["#FEE2E2","#FECACA","🔴 Critical"],["#FEF3C7","#FDE68A","🟡 Low"],["#F0FDF4","#BBF7D0","🟢 OK"]].map(([bg,border,label])=>(
          <div key={label} style={{ background:bg, border:`1px solid ${border}`, borderRadius:6, padding:"3px 10px", fontSize:11, fontWeight:600, color:"#374151" }}>{label}</div>
        ))}
      </div>

      {sections.map(sec=>(
        <div key={sec} style={{ marginBottom:20 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#9CA3AF", textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:8 }}>{sec}</div>
          {bySection[sec].map(item=>{
            const id=getDocId(item);
            const val=counts[id];
            const level=getHighlight(id,val);
            const hs=val!==""?highlightStyle(level):{background:"#F8F7F4",border:"1.5px solid #E5E7EB"};
            const t=getThreshold(id);
            const isInfoOpen=infoOpen===id;
            const isEditingThreshold=editThreshold===id;

            return (
              <div key={id} style={{ marginBottom:6 }}>
                <div style={{ ...hs, borderRadius:10, padding:"10px 12px", display:"flex", alignItems:"center", gap:8 }}>
                  {/* Item name */}
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:14, fontWeight:600, color:"#1C1917", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{item.name}</div>
                    {item.shopSpecific&&<span style={{ fontSize:10, color:"#9CA3AF" }}>shop item</span>}
                  </div>

                  {/* Qty input or N/A placeholder */}
                  {val===""
                    ? (
                      <button onClick={()=>setCounts(p=>({...p,[id]:0}))} style={{
                        padding:"6px 14px", borderRadius:8, border:"1.5px dashed #D1D5DB",
                        background:"none", color:"#9CA3AF", fontSize:13, cursor:"pointer", fontWeight:500 }}>
                        N/A → Enter
                      </button>
                    )
                    : <QtyInput value={val} onChange={v=>setCounts(p=>({...p,[id]:v}))} min={0}/>
                  }

                  {/* Unit + info */}
                  <button onClick={()=>setInfoOpen(isInfoOpen?null:id)} style={{
                    width:22, height:22, borderRadius:"50%", border:"1.5px solid",
                    borderColor:isInfoOpen?"#1C1917":"#9CA3AF",
                    background:isInfoOpen?"#1C1917":"none",
                    color:isInfoOpen?"#fff":"#9CA3AF",
                    fontSize:11, fontWeight:700, cursor:"pointer",
                    display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                    i
                  </button>
                </div>

                {/* Info panel */}
                {isInfoOpen&&(
                  <div style={{ background:"#F8F7F4", borderRadius:"0 0 10px 10px", padding:"10px 14px", border:"1.5px solid #E5E7EB", borderTop:"none" }}>
                    <div style={{ display:"flex", gap:16, flexWrap:"wrap", marginBottom:8 }}>
                      <div><span style={{ fontSize:11, color:"#9CA3AF" }}>UNIT </span><span style={{ fontSize:13, fontWeight:600 }}>{item.unit}</span></div>
                      <div><span style={{ fontSize:11, color:"#9CA3AF" }}>BUY FROM </span><span style={{ fontSize:13, fontWeight:600 }}>{item.purchaseLocation}</span></div>
                      {item.backupLocation&&<div><span style={{ fontSize:11, color:"#9CA3AF" }}>BACKUP </span><span style={{ fontSize:13, fontWeight:600 }}>{item.backupLocation}</span></div>}
                    </div>
                    {/* Threshold editor */}
                    {!isEditingThreshold
                      ? (
                        <div style={{ display:"flex", gap:12, alignItems:"center" }}>
                          <div style={{ fontSize:12, color:"#6B7280" }}>
                            🟡 Low: <b>{t.low===""?"—":t.low}</b> &nbsp; 🔴 Critical: <b>{t.critical===""?"—":t.critical}</b>
                          </div>
                          <button onClick={()=>{ setEditThreshold(id); setThresholdEdit({low:t.low,critical:t.critical}); }} style={{ fontSize:12, color:"#2563EB", background:"none", border:"none", cursor:"pointer", fontWeight:600 }}>Set thresholds</button>
                        </div>
                      )
                      : (
                        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <span style={{ fontSize:12, color:"#6B7280" }}>🟡 Low:</span>
                            <input type="number" value={thresholdEdit.low} onChange={e=>setThresholdEdit(p=>({...p,low:e.target.value}))}
                              style={{ width:60, padding:"4px 8px", borderRadius:6, border:"1.5px solid #E5E7EB", fontSize:13 }} placeholder="—"/>
                          </div>
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <span style={{ fontSize:12, color:"#6B7280" }}>🔴 Critical:</span>
                            <input type="number" value={thresholdEdit.critical} onChange={e=>setThresholdEdit(p=>({...p,critical:e.target.value}))}
                              style={{ width:60, padding:"4px 8px", borderRadius:6, border:"1.5px solid #E5E7EB", fontSize:13 }} placeholder="—"/>
                          </div>
                          <button onClick={async()=>{ await onUpdateThreshold(`${location}:${id}`,thresholdEdit); setEditThreshold(null); }} style={{ padding:"5px 12px", borderRadius:6, border:"none", background:"#1C1917", color:"#fff", fontWeight:700, fontSize:12, cursor:"pointer" }}>Save</button>
                          <button onClick={()=>setEditThreshold(null)} style={{ padding:"5px 8px", borderRadius:6, border:"1px solid #E5E7EB", background:"#fff", color:"#374151", fontSize:12, cursor:"pointer" }}>Cancel</button>
                        </div>
                      )
                    }
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ))}

      {/* Session note */}
      <div style={{ marginBottom:16 }}>
        <label style={S.label}>Session note (optional)</label>
        <input value={sessionNote} onChange={e=>setSessionNote(e.target.value)}
          placeholder="e.g. End of week count, post-event…"
          style={S.input}/>
      </div>

      {/* Summary row */}
      {filledCount>0&&(()=>{
        const flagged=myItems.filter(it=>{ const l=getHighlight(getDocId(it),counts[getDocId(it)]); return l==="low"||l==="critical"; });
        return flagged.length>0&&(
          <div style={{ background:"#FEF3C7", borderRadius:10, padding:"10px 14px", marginBottom:16 }}>
            <div style={{ fontSize:12, fontWeight:700, color:"#92400E", marginBottom:4 }}>⚠️ {flagged.length} item{flagged.length!==1?"s":""} flagged</div>
            {flagged.map(it=>{
              const l=getHighlight(getDocId(it),counts[getDocId(it)]);
              return <div key={getDocId(it)} style={{ fontSize:12, color:l==="critical"?"#DC2626":"#92400E" }}>{l==="critical"?"🔴":"🟡"} {it.name}: {counts[getDocId(it)]} {it.unit}</div>;
            })}
          </div>
        );
      })()}

      <button onClick={handleSubmit} disabled={submitting||filledCount===0} style={{
        width:"100%", padding:"14px 0", borderRadius:12, border:"none",
        background:filledCount>0?"#1C1917":"#E5E7EB", color:filledCount>0?"#fff":"#9CA3AF",
        fontWeight:700, fontSize:15, cursor:filledCount>0?"pointer":"default",
        opacity:submitting?0.7:1 }}>
        {submitting?"Submitting…":`Submit Inventory (${filledCount}/${totalCount} items)`}
      </button>
    </div>
  );
}

// ── Submissions View ──────────────────────────────────────────────────────────
function SubmissionsView({ user, submissions }) {
  const [expanded, setExpanded] = useState(null);

  // Filter: shops see only their own; kitchen sees all
  const visible = user.role==="kitchen"
    ? submissions
    : submissions.filter(s=>s.location===user.location);

  if(!visible.length) return (
    <div style={{ textAlign:"center", color:"#9CA3AF", padding:"60px 0" }}>No submissions yet.</div>
  );

  return (
    <div>
      {visible.map(sub=>{
        const id=getDocId(sub);
        const isOpen=expanded===id;
        const flagged=sub.items?.filter(i=>i.highlight==="low"||i.highlight==="critical")||[];
        const critical=sub.items?.filter(i=>i.highlight==="critical")||[];
        return (
          <div key={id} style={{ ...S.card, cursor:"pointer" }} onClick={()=>setExpanded(isOpen?null:id)}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
              <div>
                <div style={{ fontWeight:700, fontSize:14, color:"#1C1917" }}>{sub.location}</div>
                <div style={{ fontSize:12, color:"#9CA3AF", marginTop:2 }}>{fmtDate(sub.createdAt)} · {sub.filledCount}/{sub.totalCount} items</div>
                {sub.note&&<div style={{ fontSize:12, color:"#6B7280", fontStyle:"italic", marginTop:2 }}>"{sub.note}"</div>}
              </div>
              <div style={{ textAlign:"right" }}>
                {critical.length>0&&<div style={{ fontSize:11, fontWeight:700, color:"#DC2626" }}>🔴 {critical.length} critical</div>}
                {flagged.length>critical.length&&<div style={{ fontSize:11, fontWeight:700, color:"#D97706" }}>🟡 {flagged.length-critical.length} low</div>}
                <div style={{ fontSize:12, color:"#9CA3AF", marginTop:2 }}>{isOpen?"▲":"▼"}</div>
              </div>
            </div>

            {isOpen&&(
              <div style={{ marginTop:12, borderTop:"1px solid #F3F4F6", paddingTop:12 }}>
                {SECTIONS.filter(sec=>sub.items?.some(i=>i.section===sec)).map(sec=>(
                  <div key={sec} style={{ marginBottom:10 }}>
                    <div style={{ fontSize:10, fontWeight:700, color:"#9CA3AF", textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:4 }}>{sec}</div>
                    {sub.items.filter(i=>i.section===sec).map((it,idx)=>{
                      const hl=it.highlight;
                      return (
                        <div key={idx} style={{
                          display:"flex", justifyContent:"space-between", alignItems:"center",
                          padding:"5px 8px", borderRadius:6, marginBottom:3,
                          background:hl==="critical"?"#FEE2E2":hl==="low"?"#FEF3C7":hl==="ok"?"#F0FDF4":"#F8F7F4",
                        }}>
                          <span style={{ fontSize:13, color:"#1C1917" }}>{it.name}</span>
                          <span style={{ fontSize:13, fontWeight:600, color:hl==="critical"?"#DC2626":hl==="low"?"#D97706":"#374151" }}>
                            {it.qty===null?"—":`${it.qty} ${it.unit}`}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── Export View ───────────────────────────────────────────────────────────────
function ExportView({ user, submissions }) {
  const visible = user.role==="kitchen"
    ? submissions
    : submissions.filter(s=>s.location===user.location);

  const exportSubmission = (sub) => {
    const rows = (sub.items||[]).map(it=>({
      "Section": it.section||"",
      "Item": it.name,
      "Qty": it.qty===null?"N/A":it.qty,
      "Unit": it.unit,
      "Status": it.highlight==="critical"?"Critical":it.highlight==="low"?"Low":it.highlight==="ok"?"OK":"—",
      "Low Threshold": it.low||"—",
      "Critical Threshold": it.critical||"—",
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[14,22,8,8,10,14,18].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Inventory");
    const date=sub.createdAt?.toDate?sub.createdAt.toDate():new Date();
    const dateStr=date.toLocaleDateString("en-US",{month:"2-digit",day:"2-digit",year:"2-digit"}).replace(/\//g,"-");
    XLSX.writeFile(wb,`${sub.location}_Inventory_${dateStr}.xlsx`);
  };

  const exportAll = () => {
    const rows=[];
    visible.forEach(sub=>{
      (sub.items||[]).forEach(it=>{
        rows.push({
          "Date": fmtDate(sub.createdAt),
          "Location": sub.location,
          "Section": it.section||"",
          "Item": it.name,
          "Qty": it.qty===null?"N/A":it.qty,
          "Unit": it.unit,
          "Status": it.highlight==="critical"?"Critical":it.highlight==="low"?"Low":it.highlight==="ok"?"OK":"—",
          "Note": sub.note||"",
        });
      });
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=[20,12,14,22,8,8,10,28].map(w=>({wch:w}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"All Inventory");
    XLSX.writeFile(wb,`Blacksmith_Inventory_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  return (
    <div>
      <h2 style={{ fontSize:18, fontWeight:800, color:"#1C1917", marginBottom:6 }}>Export</h2>
      <p style={{ fontSize:14, color:"#6B7280", marginBottom:20 }}>Download inventory submissions as spreadsheets.</p>

      {visible.length>1&&(
        <button onClick={exportAll} style={{
          width:"100%", padding:"13px", borderRadius:12, border:"1.5px solid #E5E7EB",
          background:"#1C1917", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer", marginBottom:12,
          display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}>
          📊 Export All Submissions
        </button>
      )}

      {!visible.length&&<div style={{ textAlign:"center", color:"#9CA3AF", padding:"40px 0" }}>No submissions to export yet.</div>}

      {visible.map(sub=>(
        <div key={getDocId(sub)} style={{ ...S.card, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontWeight:600, fontSize:14, color:"#1C1917" }}>{sub.location}</div>
            <div style={{ fontSize:12, color:"#9CA3AF", marginTop:2 }}>{fmtDate(sub.createdAt)} · {sub.filledCount}/{sub.totalCount} items</div>
            {sub.note&&<div style={{ fontSize:12, color:"#6B7280", fontStyle:"italic" }}>"{sub.note}"</div>}
          </div>
          <button onClick={()=>exportSubmission(sub)} style={{
            padding:"8px 14px", borderRadius:8, border:"none", background:"#F3F4F6",
            color:"#1C1917", fontWeight:700, fontSize:12, cursor:"pointer" }}>
            XLSX ↓
          </button>
        </div>
      ))}
    </div>
  );
}

// ── Shop Inventory Item Editor (shop-specific items) ──────────────────────────
function ShopItemEditor({ item, location, onSave, onCancel }) {
  const [vals, setVals] = useState({
    name: item?.name||"",
    unit: item?.unit||"",
    section: item?.section||SECTIONS[0],
    purchaseLocation: item?.purchaseLocation||VENDORS[0],
    backupLocation: item?.backupLocation||"",
    notes: item?.notes||"",
  });
  const set=(k,v)=>setVals(p=>({...p,[k]:v}));
  return (
    <div style={{ background:"#FFFBEB", border:"1.5px dashed #FDE68A", borderRadius:12, padding:"16px", marginBottom:8 }}>
      <div style={{ fontWeight:700, fontSize:13, color:"#92400E", marginBottom:10 }}>
        {item?"Edit Shop Item":"New Shop-Specific Item"}
        <span style={{ fontSize:11, fontWeight:500, marginLeft:6 }}>(visible to Kitchen too)</span>
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:8 }}>
        <div style={{ flex:2 }}><label style={S.label}>Name</label><input value={vals.name} onChange={e=>set("name",e.target.value)} style={S.input}/></div>
        <div style={{ flex:1 }}><label style={S.label}>Unit</label><input value={vals.unit} onChange={e=>set("unit",e.target.value)} style={S.input}/></div>
      </div>
      <div style={{ marginBottom:8 }}><label style={S.label}>Section</label>
        <select value={vals.section} onChange={e=>set("section",e.target.value)} style={S.input}>
          {SECTIONS.map(s=><option key={s}>{s}</option>)}
        </select>
      </div>
      <div style={{ display:"flex", gap:8, marginBottom:12 }}>
        <div style={{ flex:1 }}><label style={S.label}>Purchase Location</label>
          <select value={vals.purchaseLocation} onChange={e=>set("purchaseLocation",e.target.value)} style={S.input}>
            {VENDORS.map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
        <div style={{ flex:1 }}><label style={S.label}>Backup</label>
          <select value={vals.backupLocation} onChange={e=>set("backupLocation",e.target.value)} style={S.input}>
            <option value="">None</option>
            {VENDORS.map(v=><option key={v}>{v}</option>)}
          </select>
        </div>
      </div>
      <div style={{ display:"flex", gap:8 }}>
        <button onClick={async()=>{ await onSave({...vals, shopSpecific:true, location}); }} style={{ flex:2, padding:"9px 0", borderRadius:8, border:"none", background:"#1C1917", color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer" }}>Save</button>
        <button onClick={onCancel} style={{ flex:1, padding:"9px 0", borderRadius:8, border:"1.5px solid #E5E7EB", background:"#fff", color:"#374151", fontWeight:600, fontSize:13, cursor:"pointer" }}>Cancel</button>
      </div>
    </div>
  );
}

// ── Shop Inventory Management (manage shop-specific items + thresholds) ───────
function ShopInventoryManage({ user, masterItems, shopItems, thresholds, onAddShopItem, onUpdateShopItem, onDeleteShopItem, onUpdateThreshold }) {
  const location = user.location;
  const myShopItems = shopItems.filter(i=>i.location===location);
  const [addingNew, setAddingNew] = useState(false);
  const [editingId, setEditingId] = useState(null);

  const handleSave = async(vals) => {
    if(editingId){ await onUpdateShopItem(editingId,vals); setEditingId(null); }
    else { await onAddShopItem(vals); setAddingNew(false); }
  };

  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <h2 style={{ fontSize:18, fontWeight:800, color:"#1C1917" }}>Manage Items</h2>
        <button onClick={()=>setAddingNew(true)} style={{ padding:"8px 14px", borderRadius:8, border:"none", background:"#1C1917", color:"#fff", fontWeight:700, fontSize:12, cursor:"pointer" }}>+ Shop Item</button>
      </div>

      {addingNew&&<ShopItemEditor location={location} onSave={handleSave} onCancel={()=>setAddingNew(false)}/>}

      <div style={{ fontSize:12, color:"#6B7280", marginBottom:12 }}>
        Master items are managed by Kitchen. You can set thresholds for each item in the Inventory Count view.
      </div>

      <div style={{ fontSize:11, fontWeight:700, color:"#9CA3AF", textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:8 }}>
        {location} Shop Items ({myShopItems.length})
      </div>

      {myShopItems.length===0&&!addingNew&&(
        <div style={{ textAlign:"center", color:"#9CA3AF", padding:"24px 0", fontSize:13 }}>No shop-specific items yet.</div>
      )}

      {myShopItems.map(item=>(
        <div key={getDocId(item)}>
          {editingId===getDocId(item)
            ? <ShopItemEditor item={item} location={location} onSave={handleSave} onCancel={()=>setEditingId(null)}/>
            : (
              <div style={{ ...S.card, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <div>
                  <div style={{ fontWeight:600, fontSize:14, color:"#1C1917" }}>{item.name}</div>
                  <div style={{ fontSize:12, color:"#9CA3AF" }}>{item.unit} · {item.section} · {item.purchaseLocation}</div>
                </div>
                <div style={{ display:"flex", gap:8 }}>
                  <button onClick={()=>setEditingId(getDocId(item))} style={{ background:"none", border:"none", color:"#9CA3AF", fontSize:12, fontWeight:600, cursor:"pointer" }}>Edit</button>
                  <button onClick={()=>onDeleteShopItem(getDocId(item))} style={{ background:"none", border:"none", color:"#EF4444", fontSize:12, fontWeight:600, cursor:"pointer" }}>Delete</button>
                </div>
              </div>
            )
          }
        </div>
      ))}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);
  const [view, setView] = useState("count");
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);

  const [masterItems, setMasterItems] = useState([]);    // kitchen-managed shared items
  const [shopItems, setShopItems] = useState([]);        // shop-specific items
  const [thresholds, setThresholds] = useState({});      // {location:itemId -> {low,critical}}
  const [submissions, setSubmissions] = useState([]);

  const showToast = msg => setToast(msg);

  // ── Firestore listeners ──────────────────────────────────────────────────────
  useEffect(()=>{
    const unsubs=[];
    unsubs.push(onSnapshot(collection(db,"inv_master"),snap=>{
      setMasterItems(snap.docs.map(d=>({...d.data(),_docId:d.id})));
      setLoading(false);
    }));
    unsubs.push(onSnapshot(collection(db,"inv_shop"),snap=>{
      setShopItems(snap.docs.map(d=>({...d.data(),_docId:d.id})));
    }));
    unsubs.push(onSnapshot(collection(db,"inv_thresholds"),snap=>{
      const t={};
      snap.docs.forEach(d=>{ t[d.id]={...d.data()}; });
      setThresholds(t);
    }));
    unsubs.push(onSnapshot(query(collection(db,"inv_submissions"),orderBy("createdAt","desc")),snap=>{
      setSubmissions(snap.docs.map(d=>({...d.data(),_docId:d.id})));
    }));
    return ()=>unsubs.forEach(u=>u());
  },[]);

  // ── CRUD ────────────────────────────────────────────────────────────────────
  const addMasterItem = async(vals) => {
    const id=genId();
    await setDoc(doc(db,"inv_master",id),{id,...vals,createdAt:serverTimestamp()});
    showToast("Item added");
  };
  const updateMasterItem = async(docId,vals) => {
    await updateDoc(doc(db,"inv_master",docId),vals);
    showToast("Item updated");
  };
  const deleteMasterItem = async(docId) => {
    await deleteDoc(doc(db,"inv_master",docId));
    showToast("Item deleted");
  };

  const addShopItem = async(vals) => {
    const id=genId();
    await setDoc(doc(db,"inv_shop",id),{id,...vals,createdAt:serverTimestamp()});
    showToast("Shop item added");
  };
  const updateShopItem = async(docId,vals) => {
    await updateDoc(doc(db,"inv_shop",docId),vals);
    showToast("Item updated");
  };
  const deleteShopItem = async(docId) => {
    await deleteDoc(doc(db,"inv_shop",docId));
    showToast("Item deleted");
  };

  const updateThreshold = async(key, vals) => {
    await setDoc(doc(db,"inv_thresholds",key),{...vals},{ merge:true });
    showToast("Thresholds saved");
  };

  const submitInventory = async(data) => {
    await addDoc(collection(db,"inv_submissions"),{...data,createdAt:serverTimestamp()});
    showToast("Inventory submitted!");
    setView("submissions");
  };

  if(!user) return <LoginScreen onLogin={u=>{
    setUser(u);
    setView(u.role==="kitchen"?"catalog":"count");
  }}/>;

  const isKitchen = user.role==="kitchen";

  const navTabs = isKitchen
    ? [["count","New Count"],["catalog","Item Catalog"],["submissions","Submissions"],["export","Export"]]
    : [["count","New Count"],["manage","My Items"],["submissions","Submissions"],["export","Export"]];

  return (
    <div style={{ minHeight:"100vh", background:"#F8F7F4", fontFamily:"'Inter',system-ui,sans-serif" }}>
      {/* Header */}
      <div style={{ background:"#1C1917", color:"#fff", padding:"13px 20px",
        display:"flex", justifyContent:"space-between", alignItems:"center",
        position:"sticky", top:0, zIndex:100 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span style={{ fontSize:22 }}>🔨</span>
          <div>
            <div style={{ fontWeight:800, fontSize:16 }}>Blacksmith Inventory</div>
            <div style={{ fontSize:11, color:"#A8A29E" }}>
              {user.icon} {user.label}{user.location?` · ${user.location}`:""}
            </div>
          </div>
        </div>
        <button onClick={()=>setUser(null)} style={{ background:"rgba(255,255,255,.1)", border:"none", color:"#fff", padding:"6px 12px", borderRadius:8, fontSize:12, cursor:"pointer" }}>Sign out</button>
      </div>

      {/* Nav */}
      <div style={{ background:"#fff", borderBottom:"1px solid #F3F4F6", padding:"0 16px", display:"flex", overflowX:"auto" }}>
        {navTabs.map(([key,lbl])=>(
          <button key={key} onClick={()=>setView(key)} style={{
            padding:"13px 12px", border:"none", background:"none",
            fontWeight:view===key?700:500, fontSize:13,
            color:view===key?"#1C1917":"#9CA3AF",
            borderBottom:view===key?"2.5px solid #1C1917":"2.5px solid transparent",
            cursor:"pointer", whiteSpace:"nowrap" }}>{lbl}</button>
        ))}
      </div>

      {/* Content */}
      <div style={{ padding:"16px 20px", maxWidth:700, margin:"0 auto" }}>
        {loading ? <Spinner/> : (
          <>
            {view==="count"&&(
              <InventoryCount
                user={user}
                masterItems={masterItems}
                shopItems={shopItems}
                thresholds={thresholds}
                onSubmit={submitInventory}
                onUpdateThreshold={updateThreshold}/>
            )}
            {view==="catalog"&&(
              <KitchenCatalog
                items={masterItems}
                shopItems={shopItems}
                onAddItem={addMasterItem}
                onUpdateItem={updateMasterItem}
                onDeleteItem={deleteMasterItem}/>
            )}
            {view==="manage"&&(
              <ShopInventoryManage
                user={user}
                masterItems={masterItems}
                shopItems={shopItems}
                thresholds={thresholds}
                onAddShopItem={addShopItem}
                onUpdateShopItem={updateShopItem}
                onDeleteShopItem={deleteShopItem}
                onUpdateThreshold={updateThreshold}/>
            )}
            {view==="submissions"&&(
              <SubmissionsView user={user} submissions={submissions}/>
            )}
            {view==="export"&&(
              <ExportView user={user} submissions={submissions}/>
            )}
          </>
        )}
      </div>

      {toast&&<Toast msg={toast} onClose={()=>setToast(null)}/>}
    </div>
  );
}
