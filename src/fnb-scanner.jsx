import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ── THEME ─────────────────────────────────────────────────────────────────
const T = {
  bg:       "#0a0f0b",
  panel:    "#111812",
  card:     "#162019",
  border:   "#1e3022",
  green:    "#3ecf6e",
  greenDim: "#1a5c33",
  greenPale:"#0d2e1a",
  gold:     "#f0c060",
  goldDim:  "#7a5c20",
  red:      "#e05555",
  redDim:   "#5c1e1e",
  blue:     "#5599ee",
  txt:      "#d8eedd",
  txt2:     "#5a8a65",
  txt3:     "#2e5038",
};

// ── STORAGE ───────────────────────────────────────────────────────────────
const db = {
  get: async k => { try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : null; } catch { return null; } },
  set: async (k, v) => { try { await window.storage.set(k, JSON.stringify(v)); } catch {} },
};

// ── GOOGLE SHEETS SYNC ────────────────────────────────────────────────────
async function syncSheet(url, action, data) {
  if (!url) return false;
  try {
    await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify({ action, ...data }),
    });
    return true;
  } catch { return false; }
}

// ── CLAUDE AI EXTRACTION ──────────────────────────────────────────────────
async function extractInvoice(b64, mime) {
  const isPdf = mime === "application/pdf";
  const fileBlock = isPdf
    ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }
    : { type: "image",    source: { type: "base64", media_type: mime, data: b64 } };

  const prompt = isPdf
    ? `Αυτό το PDF μπορεί να περιέχει ΠΟΛΛΑ τιμολόγια σε πολλές σελίδες. Εξήγαγε ΟΛΕΣ τις σελίδες.
Επέστρεψε ΜΟΝΟ JSON array (χωρίς markdown), ένα object ανά τιμολόγιο:
[{"supplier":"...","invoice_number":"...","date":"YYYY-MM-DD","items":[{"name":"...","category":"κρέας|ψάρι|λαχανικά|γαλακτοκομικά|ποτά|αναλώσιμα|άλλο","quantity":0,"unit":"kg|τεμ|lt|κιβ","unit_price":0,"total":0}],"subtotal":0,"vat":0,"grand_total":0}]
Αν υπάρχει 1 τιμολόγιο, επέστρεψε array με 1 στοιχείο. ΜΟΝΟ JSON array.`
    : `Ανάλυσε αυτό το τιμολόγιο F&B. Επέστρεψε ΜΟΝΟ JSON array με 1 στοιχείο (χωρίς markdown):
[{"supplier":"...","invoice_number":"...","date":"YYYY-MM-DD","items":[{"name":"...","category":"κρέας|ψάρι|λαχανικά|γαλακτοκομικά|ποτά|αναλώσιμα|άλλο","quantity":0,"unit":"kg|τεμ|lt|κιβ","unit_price":0,"total":0}],"subtotal":0,"vat":0,"grand_total":0}]
ΜΟΝΟ JSON array.`;

  const res = await fetch("/api/extract", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-5",
      max_tokens: 8000,
      messages: [{ role: "user", content: [fileBlock, { type: "text", text: prompt }] }]
    })
  });
  const d = await res.json();
  if (d.error) throw new Error(d.error.message || "API error");
  const txt = d.content?.find(b => b.type === "text")?.text || "[]";
  const cleaned = txt.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(cleaned);
  return Array.isArray(parsed) ? parsed : [parsed];
}

// ── EXCEL EXPORT ──────────────────────────────────────────────────────────
function exportXLSX(invoices, offers) {
  const wb = XLSX.utils.book_new();

  // Sheet 1 – Αναλυτικά
  const rows1 = [["Ημερομηνία","Προμηθευτής","Τιμολόγιο #","Είδος","Κατηγορία","Ποσότητα","Μονάδα","Τιμή/Μον.€","Σύνολο €","Γεν.Σύνολο €"]];
  invoices.forEach(inv => (inv.items||[]).forEach(it => rows1.push([
    inv.date||"", inv.supplier||"", inv.invoice_number||"",
    it.name||"", it.category||"", it.quantity??"", it.unit||"",
    it.unit_price??"", it.total??"", inv.grand_total??""
  ])));
  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"] = [12,20,12,24,14,10,8,12,12,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, "Τιμολόγια");

  // Sheet 2 – Σύνοψη
  const rows2 = [["Ημερομηνία","Προμηθευτής","Τιμολόγιο #","Είδη","Καθαρή €","ΦΠΑ €","Σύνολο €"]];
  invoices.forEach(inv => rows2.push([inv.date||"",inv.supplier||"",inv.invoice_number||"",inv.items?.length||0,inv.subtotal??"",inv.vat??"",inv.grand_total??""]));
  const n = rows2.length;
  rows2.push(["ΣΥΝΟΛΟ","","",`=SUM(D2:D${n})`,`=SUM(E2:E${n})`,`=SUM(F2:F${n})`,`=SUM(G2:G${n})`]);
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [12,20,12,8,14,12,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws2, "Σύνοψη");

  // Sheet 3 – Σύγκριση
  const rows3 = [["Ημερομηνία","Προμηθευτής","Είδος","Τιμή Τιμολ.€","Τιμή Προσφ.€","Διαφορά €","% Απόκλιση","Επίπτωση €","Κατάσταση"]];
  invoices.forEach(inv => (inv.items||[]).forEach(it => {
    if (it.unit_price==null) return;
    const off = offers.find(o => o.name.toLowerCase().includes(it.name.toLowerCase().slice(0,5)) || it.name.toLowerCase().includes(o.name.toLowerCase().slice(0,5)));
    if (!off) return;
    const diff = it.unit_price - off.offer_price;
    const pct  = off.offer_price ? diff/off.offer_price*100 : 0;
    rows3.push([inv.date||"",inv.supplier||"",it.name,it.unit_price,off.offer_price,diff,pct/100,(it.quantity||1)*diff,
      Math.abs(pct)<=1?"✓ OK":diff>0?pct>5?"⚠ ΥΠΕΡΒΑΣΗ":"↑ Ακριβότερο":"↓ Φθηνότερο"]);
  }));
  const ws3 = XLSX.utils.aoa_to_sheet(rows3);
  ws3["!cols"] = [12,18,22,14,14,12,12,14,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws3, "Σύγκριση Τιμών");

  // Sheet 4 – Προσφορές
  const rows4 = [["Προμηθευτής","Είδος","Κατηγορία","Μονάδα","Τιμή Προσφοράς €"]];
  offers.forEach(o => rows4.push([o.supplier||"",o.name,o.category,o.unit,o.offer_price]));
  const ws4 = XLSX.utils.aoa_to_sheet(rows4);
  ws4["!cols"] = [18,24,14,10,18].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws4, "Τιμές Προσφορών");

  // Sheet 5 – Ανά Κατηγορία (Pivot)
  const catMap = {};
  invoices.forEach(inv => (inv.items||[]).forEach(it => {
    const cat = it.category || "άλλο";
    if (!catMap[cat]) catMap[cat] = { items: {}, total: 0, qty: 0 };
    const name = it.name || "—";
    if (!catMap[cat].items[name]) catMap[cat].items[name] = { qty: 0, total: 0, count: 0 };
    catMap[cat].items[name].qty   += it.quantity || 0;
    catMap[cat].items[name].total += it.total || 0;
    catMap[cat].items[name].count += 1;
    catMap[cat].total += it.total || 0;
    catMap[cat].qty   += it.quantity || 0;
  }));
  const rows5 = [["Κατηγορία","Είδος","Συν. Ποσότητα","Φορές","Συν. Αξία €","% επί Συνόλου"]];
  const grandTotalCat = Object.values(catMap).reduce((s,c)=>s+c.total,0);
  Object.entries(catMap).sort((a,b)=>b[1].total-a[1].total).forEach(([cat,data])=>{
    rows5.push([cat,"— ΣΥΝΟΛΟ ΚΑΤΗΓΟΡΙΑΣ —","",Object.values(data.items).reduce((s,i)=>s+i.count,0),data.total, grandTotalCat ? data.total/grandTotalCat : 0]);
    Object.entries(data.items).sort((a,b)=>b[1].total-a[1].total).forEach(([name,d])=>{
      rows5.push(["", name, d.qty, d.count, d.total, grandTotalCat ? d.total/grandTotalCat : 0]);
    });
    rows5.push(["","","","","",""]);
  });
  const ws5 = XLSX.utils.aoa_to_sheet(rows5);
  ws5["!cols"] = [18,28,14,10,14,14].map(w=>({wch:w}));
  const pctColCat = rows5.length;
  for (let r=1; r<rows5.length; r++) {
    const cell = XLSX.utils.encode_cell({r, c:5});
    if (ws5[cell] && typeof ws5[cell].v === "number") ws5[cell].z = "0.0%";
    const valCell = XLSX.utils.encode_cell({r, c:4});
    if (ws5[valCell] && typeof ws5[valCell].v === "number") ws5[valCell].z = "#,##0.00";
  }
  XLSX.utils.book_append_sheet(wb, ws5, "Ανά Κατηγορία");

  // Sheet 6 – Ανά Προμηθευτή (Pivot)
  const supplMap = {};
  invoices.forEach(inv => {
    const s = inv.supplier || "Άγνωστος";
    if (!supplMap[s]) supplMap[s] = { invoices: {}, items: {}, total: 0 };
    const invKey = inv.invoice_number || inv.date || inv.id;
    supplMap[s].invoices[invKey] = true;
    (inv.items||[]).forEach(it => {
      const name = it.name || "—";
      if (!supplMap[s].items[name]) supplMap[s].items[name] = { qty:0, total:0, count:0, cat: it.category||"" };
      supplMap[s].items[name].qty   += it.quantity || 0;
      supplMap[s].items[name].total += it.total || 0;
      supplMap[s].items[name].count += 1;
      supplMap[s].total += it.total || 0;
    });
  });
  const rows6 = [["Προμηθευτής","Είδος","Κατηγορία","Συν. Ποσότητα","Φορές","Συν. Αξία €","% επί Συνόλου"]];
  const grandTotalSuppl = Object.values(supplMap).reduce((s,v)=>s+v.total,0);
  Object.entries(supplMap).sort((a,b)=>b[1].total-a[1].total).forEach(([supp,data])=>{
    const invCount = Object.keys(data.invoices).length;
    rows6.push([supp, `— ${invCount} τιμολόγι${invCount===1?"ο":"α"} —`,"","",
      Object.values(data.items).reduce((s,i)=>s+i.count,0),
      data.total, grandTotalSuppl ? data.total/grandTotalSuppl : 0]);
    Object.entries(data.items).sort((a,b)=>b[1].total-a[1].total).forEach(([name,d])=>{
      rows6.push(["", name, d.cat, d.qty, d.count, d.total, grandTotalSuppl ? d.total/grandTotalSuppl : 0]);
    });
    rows6.push(["","","","","","",""]);
  });
  const ws6 = XLSX.utils.aoa_to_sheet(rows6);
  ws6["!cols"] = [20,28,14,14,10,14,14].map(w=>({wch:w}));
  for (let r=1; r<rows6.length; r++) {
    const cell = XLSX.utils.encode_cell({r, c:6});
    if (ws6[cell] && typeof ws6[cell].v === "number") ws6[cell].z = "0.0%";
    const valCell = XLSX.utils.encode_cell({r, c:5});
    if (ws6[valCell] && typeof ws6[valCell].v === "number") ws6[valCell].z = "#,##0.00";
  }
  XLSX.utils.book_append_sheet(wb, ws6, "Ανά Προμηθευτή");

  XLSX.writeFile(wb, `FnB_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ── HELPERS ───────────────────────────────────────────────────────────────
const toB64 = f => new Promise((res,rej) => {
  const r = new FileReader();
  r.onload = ()=>res(r.result.split(",")[1]);
  r.onerror = rej;
  r.readAsDataURL(f);
});

const catColor = c => ({
  κρέας:"#e05555", ψάρι:"#5599ee", λαχανικά:"#3ecf6e",
  γαλακτοκομικά:"#f0c060", ποτά:"#a855f7", αναλώσιμα:"#22d3ee", άλλο:"#5a8a65"
}[c?.split(" ")[0]] || "#5a8a65");

const fmt = n => n != null ? Number(n).toLocaleString("el-GR",{minimumFractionDigits:2,maximumFractionDigits:2})+" €" : "—";

// ── ICONS ─────────────────────────────────────────────────────────────────
const Ic = ({p, s=16, c="currentColor"}) => (
  <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth={1.6} strokeLinecap="round" strokeLinejoin="round"><path d={p}/></svg>
);

// ── SCAN STATUS PILL ─────────────────────────────────────────────────────
function Pill({status}) {
  const map = {
    scanning: [T.gold,   "⏳ Σάρωση AI…"],
    syncing:  [T.blue,   "⟳ Sync Sheets…"],
    done:     [T.green,  "✓ Καταχωρήθηκε"],
    error:    [T.red,    "✗ Σφάλμα"],
    waiting:  [T.txt3,   "⏸ Αναμονή"],
  };
  const [col, lbl] = map[status] || map.waiting;
  return (
    <span style={{
      background: col+"22", color: col, border:`1px solid ${col}44`,
      borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:600,
      letterSpacing:"0.04em", whiteSpace:"nowrap",
      animation: status==="scanning"||status==="syncing" ? "pulse 1.2s infinite" : "none"
    }}>{lbl}</span>
  );
}

// ── UPLOAD ZONE ──────────────────────────────────────────────────────────
function UploadZone({ onFiles }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  const handle = files => onFiles(Array.from(files).filter(f=>f.type.startsWith("image/") || f.type==="application/pdf"));
  return (
    <div
      onDragOver={e=>{e.preventDefault();setDrag(true)}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);handle(e.dataTransfer.files)}}
      onClick={()=>ref.current.click()}
      style={{
        border:`2px dashed ${drag?T.green:T.border}`,
        borderRadius:16, padding:"52px 32px", textAlign:"center",
        cursor:"pointer", background: drag ? T.greenPale : T.panel,
        transition:"all 0.2s", userSelect:"none"
      }}
    >
      <div style={{marginBottom:16}}>
        <Ic p="M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2zM12 17a4 4 0 100-8 4 4 0 000 8z" s={44} c={drag?T.green:T.greenDim}/>
      </div>
      <div style={{color:T.txt, fontSize:17, fontWeight:700, marginBottom:6, letterSpacing:"-0.01em"}}>
        Φωτογράφισε ή ανέβασε τιμολόγια
      </div>
      <div style={{color:T.txt2, fontSize:13}}>JPG · PNG · HEIC · PDF — πολλαπλά αρχεία ταυτόχρονα</div>
      <div style={{marginTop:20, display:"inline-flex", alignItems:"center", gap:8,
        background:T.green, color:"#0a0f0b", borderRadius:10, padding:"10px 24px",
        fontSize:13, fontWeight:700, letterSpacing:"0.02em"}}>
        <Ic p="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" s={16} c="#0a0f0b"/>
        Επιλογή Αρχείων
      </div>
      <input ref={ref} type="file" accept="image/*,application/pdf" multiple style={{display:"none"}} onChange={e=>handle(e.target.files)}/>
    </div>
  );
}

// ── INVOICE CARD ──────────────────────────────────────────────────────────
function InvoiceCard({ inv, status, onDelete }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{
      background:T.card, border:`1px solid ${status==="done"?T.greenDim:T.border}`,
      borderRadius:14, overflow:"hidden", transition:"border 0.3s",
      boxShadow: status==="done" ? `0 0 0 1px ${T.greenDim}33` : "none"
    }}>
      {/* Header row */}
      <div style={{display:"flex", alignItems:"center", gap:12, padding:"14px 18px", cursor:"pointer"}}
        onClick={()=>setOpen(o=>!o)}>
        <Ic p="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6" s={18} c={T.greenDim}/>
        <div style={{flex:1, minWidth:0}}>
          <div style={{color:T.txt, fontWeight:600, fontSize:14, letterSpacing:"-0.01em",
            whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis"}}>
            {inv.supplier || "Άγνωστος Προμηθευτής"}
          </div>
          <div style={{color:T.txt2, fontSize:11, marginTop:2}}>
            {inv.date||"—"} {inv.invoice_number ? `· #${inv.invoice_number}` : ""}
            {" · "}{inv.items?.length||0} είδη
          </div>
        </div>
        <div style={{display:"flex", alignItems:"center", gap:10, flexShrink:0}}>
          <span style={{color:T.gold, fontFamily:"monospace", fontSize:14, fontWeight:600}}>
            {fmt(inv.grand_total)}
          </span>
          <Pill status={status||"done"}/>
          <Ic p={open?"M18 15l-6-6-6 6":"M6 9l6 6 6-6"} s={14} c={T.txt3}/>
        </div>
      </div>

      {/* Items table */}
      {open && (
        <div style={{borderTop:`1px solid ${T.border}`}}>
          <table style={{width:"100%", borderCollapse:"collapse"}}>
            <thead>
              <tr style={{background:T.panel}}>
                {["Είδος","Κατηγ.","Ποσ.","Μον.","Τιμή/Μον.","Σύνολο"].map(h=>(
                  <th key={h} style={{padding:"8px 12px", textAlign:"left", color:T.txt3,
                    fontSize:10, letterSpacing:"0.07em", fontWeight:600}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {(inv.items||[]).map((it,i)=>(
                <tr key={i} style={{borderTop:`1px solid ${T.border}`, background: i%2===0?T.card:T.panel+"cc"}}>
                  <td style={{padding:"9px 12px", color:T.txt, fontSize:13, fontWeight:500}}>{it.name}</td>
                  <td style={{padding:"9px 12px"}}>
                    <span style={{background:catColor(it.category)+"22", color:catColor(it.category),
                      border:`1px solid ${catColor(it.category)}44`, borderRadius:4,
                      padding:"2px 7px", fontSize:10, fontWeight:600}}>{it.category||"άλλο"}</span>
                  </td>
                  <td style={{padding:"9px 12px", color:T.txt2, fontFamily:"monospace", fontSize:12}}>{it.quantity}</td>
                  <td style={{padding:"9px 12px", color:T.txt3, fontSize:12}}>{it.unit}</td>
                  <td style={{padding:"9px 12px", color:T.txt2, fontFamily:"monospace", fontSize:12}}>{fmt(it.unit_price)}</td>
                  <td style={{padding:"9px 12px", color:T.gold, fontFamily:"monospace", fontSize:13, fontWeight:600}}>{fmt(it.total)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center",
            padding:"10px 18px", borderTop:`1px solid ${T.border}`}}>
            <div style={{color:T.txt3, fontSize:11}}>
              {inv.vat != null ? `ΦΠΑ: ${fmt(inv.vat)}` : ""}
            </div>
            <button onClick={()=>onDelete(inv.id)} style={{
              background:"transparent", color:T.txt3, border:`1px solid ${T.border}`,
              borderRadius:7, padding:"5px 12px", fontSize:12, cursor:"pointer",
              display:"flex", alignItems:"center", gap:5
            }}>
              <Ic p="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6" s={13} c={T.red}/>
              Διαγραφή
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── OFFERS MINI PANEL ─────────────────────────────────────────────────────
function OffersPanel({ offers, onSave }) {
  const [rows, setRows] = useState(offers);
  const [n, setN] = useState({supplier:"",name:"",category:"λαχανικά",unit:"kg",offer_price:""});
  useEffect(()=>setRows(offers),[offers]);
  const inp = s => ({
    background:T.panel, border:`1px solid ${T.border}`, borderRadius:7,
    padding:"8px 10px", color:T.txt, fontSize:12, outline:"none", width:"100%", boxSizing:"border-box"
  });
  const add = () => {
    if (!n.name||!n.offer_price) return;
    const r = [...rows,{...n,id:Date.now(),offer_price:parseFloat(n.offer_price)}];
    setRows(r); onSave(r);
    setN({supplier:"",name:"",category:"λαχανικά",unit:"kg",offer_price:""});
  };
  const cats = ["κρέας","ψάρι","λαχανικά","γαλακτοκομικά","ποτά","αναλώσιμα","άλλο"];
  return (
    <div style={{background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:20}}>
      <div style={{color:T.green, fontSize:12, fontWeight:700, letterSpacing:"0.06em", marginBottom:14}}>
        ΤΙΜΕΣ ΠΡΟΣΦΟΡΩΝ ΠΡΟΜΗΘΕΥΤΩΝ
      </div>
      {/* Add row */}
      <div style={{display:"grid", gridTemplateColumns:"1fr 2fr 1fr 1fr 1fr auto", gap:8, marginBottom:12}}>
        {[
          [n.supplier,"supplier","Προμηθευτής"],
          [n.name,"name","Είδος *"],
          [n.unit,"unit","Μονάδα"],
          [n.offer_price,"offer_price","Τιμή € *"],
        ].map(([v,f,ph])=>(
          <input key={f} placeholder={ph} value={v} onChange={e=>setN(p=>({...p,[f]:e.target.value}))} style={inp()}/>
        ))}
        <select value={n.category} onChange={e=>setN(p=>({...p,category:e.target.value}))} style={{...inp(),cursor:"pointer"}}>
          {cats.map(c=><option key={c}>{c}</option>)}
        </select>
        <button onClick={add} style={{
          background:T.green, color:"#0a0f0b", border:"none", borderRadius:7,
          padding:"8px 14px", fontWeight:700, fontSize:12, cursor:"pointer", whiteSpace:"nowrap"
        }}>+ Προσθ.</button>
      </div>
      {/* Rows */}
      {rows.length>0 && (
        <table style={{width:"100%", borderCollapse:"collapse"}}>
          <thead><tr>
            {["Προμηθευτής","Είδος","Κατηγ.","Μον.","Τιμή €",""].map(h=>(
              <th key={h} style={{padding:"6px 8px", textAlign:"left", color:T.txt3, fontSize:10, letterSpacing:"0.06em"}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>{rows.map((r,i)=>(
            <tr key={r.id} style={{borderTop:`1px solid ${T.border}`, background:i%2===0?"transparent":T.panel+"66"}}>
              <td style={{padding:"7px 8px", color:T.txt2, fontSize:12}}>{r.supplier||"—"}</td>
              <td style={{padding:"7px 8px", color:T.txt, fontSize:12, fontWeight:500}}>{r.name}</td>
              <td style={{padding:"7px 8px"}}>
                <span style={{color:catColor(r.category), fontSize:10}}>{r.category}</span>
              </td>
              <td style={{padding:"7px 8px", color:T.txt3, fontSize:11}}>{r.unit}</td>
              <td style={{padding:"7px 8px", color:T.gold, fontFamily:"monospace", fontSize:12, fontWeight:600}}>
                {Number(r.offer_price).toFixed(2)} €
              </td>
              <td style={{padding:"7px 8px"}}>
                <button onClick={()=>{const u=rows.filter(x=>x.id!==r.id);setRows(u);onSave(u);}}
                  style={{background:"transparent",border:"none",cursor:"pointer",color:T.txt3}}>✕</button>
              </td>
            </tr>
          ))}</tbody>
        </table>
      )}
    </div>
  );
}

// ── SETTINGS PANEL ────────────────────────────────────────────────────────
function SettingsPanel({ url, onSave, onFullSync, lastSync }) {
  const [v, setV] = useState(url||"");
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState(null);
  const test = async () => {
    setTesting(true); setResult(null);
    const ok = await syncSheet(v,"full_sync",{invoices:[],offers:[]});
    setResult(ok?"ok":"err"); setTesting(false);
  };
  return (
    <div style={{background:T.card, border:`1px solid ${T.border}`, borderRadius:14, padding:24}}>
      <div style={{color:T.green, fontSize:12, fontWeight:700, letterSpacing:"0.06em", marginBottom:14}}>
        GOOGLE SHEETS ΣΥΝΔΕΣΗ
      </div>
      <div style={{color:T.txt2, fontSize:13, marginBottom:16, lineHeight:1.6}}>
        Κάθε νέο τιμολόγιο στέλνεται αυτόματα στο Google Sheet σου.
        {lastSync && <span style={{color:T.txt3, marginLeft:8}}>Τελ. sync: {lastSync}</span>}
      </div>
      <input
        placeholder="https://script.google.com/macros/s/…/exec"
        value={v} onChange={e=>setV(e.target.value)}
        style={{width:"100%", background:T.panel, border:`1px solid ${T.border}`, borderRadius:8,
          padding:"11px 14px", color:T.txt, fontSize:12, fontFamily:"monospace",
          outline:"none", marginBottom:10, boxSizing:"border-box"}}
      />
      <div style={{display:"flex", gap:8, flexWrap:"wrap"}}>
        <button onClick={()=>onSave(v)} style={{
          background:T.green, color:"#0a0f0b", border:"none", borderRadius:8,
          padding:"9px 18px", fontSize:12, fontWeight:700, cursor:"pointer"}}>Αποθήκευση</button>
        <button onClick={test} disabled={testing||!v} style={{
          background:T.panel, color:T.txt2, border:`1px solid ${T.border}`,
          borderRadius:8, padding:"9px 18px", fontSize:12, cursor:v?"pointer":"not-allowed"}}>
          {testing?"Έλεγχος…":"Test Σύνδεση"}</button>
        {url && <button onClick={onFullSync} style={{
          background:T.gold+"18", color:T.gold, border:`1px solid ${T.goldDim}`,
          borderRadius:8, padding:"9px 18px", fontSize:12, fontWeight:600, cursor:"pointer"}}>
          ⟳ Full Sync</button>}
      </div>
      {result && (
        <div style={{marginTop:12, padding:"9px 14px", borderRadius:8, fontSize:12,
          background:result==="ok"?T.green+"15":T.red+"15",
          color:result==="ok"?T.green:T.red}}>
          {result==="ok"?"✓ Η σύνδεση λειτουργεί — το Sheet ενημερώθηκε!":"✗ Αποτυχία σύνδεσης. Έλεγξε το URL."}
        </div>
      )}
      {/* Instructions */}
      <div style={{marginTop:20, padding:16, background:T.panel, borderRadius:10, border:`1px solid ${T.border}`}}>
        <div style={{color:T.txt, fontSize:12, fontWeight:600, marginBottom:12}}>Οδηγίες (μία φορά):</div>
        {[
          ["1","Άνοιξε Google Sheets","Νέο κενό φύλλο"],
          ["2","Extensions → Apps Script","Επικόλλησε το αρχείο GoogleAppsScript.js"],
          ["3","Deploy → New deployment","Web app · Execute as: Me · Access: Anyone"],
          ["4","Αντέγραψε το Web App URL","Επικόλλησέ το εδώ πάνω"],
        ].map(([n,t,d])=>(
          <div key={n} style={{display:"flex", gap:10, marginBottom:10}}>
            <div style={{width:22, height:22, borderRadius:"50%", background:T.greenDim+"44",
              color:T.green, fontSize:11, fontWeight:700, display:"flex", alignItems:"center",
              justifyContent:"center", flexShrink:0}}>{n}</div>
            <div>
              <div style={{color:T.txt, fontSize:12, fontWeight:600}}>{t}</div>
              <div style={{color:T.txt2, fontSize:11}}>{d}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── STATS BAR ─────────────────────────────────────────────────────────────
function StatsBar({ invoices }) {
  const total = invoices.reduce((s,i)=>s+(i.grand_total||0),0);
  const items = invoices.reduce((s,i)=>s+(i.items?.length||0),0);
  const suppliers = new Set(invoices.map(i=>i.supplier).filter(Boolean)).size;
  const today = invoices.filter(i=>i.uploaded_at?.startsWith(new Date().toISOString().slice(0,10))).length;
  return (
    <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:10}}>
      {[
        ["Σύνολο Αγορών", total>0?`${total.toLocaleString("el-GR",{minimumFractionDigits:2})} €`:"—", T.gold],
        ["Τιμολόγια",      invoices.length, T.green],
        ["Προμηθευτές",    suppliers, T.blue],
        ["Σήμερα",         today, T.txt2],
      ].map(([l,v,c])=>(
        <div key={l} style={{background:T.card, border:`1px solid ${T.border}`, borderRadius:12, padding:"16px 18px"}}>
          <div style={{color:T.txt3, fontSize:10, letterSpacing:"0.08em", marginBottom:6}}>{l}</div>
          <div style={{color:c, fontSize:24, fontFamily:"monospace", fontWeight:500}}>{v}</div>
        </div>
      ))}
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────
export default function App() {
  const [tab,        setTab]        = useState("scan");
  const [invoices,   setInvoices]   = useState([]);
  const [offers,     setOffers]     = useState([]);
  const [webhook,    setWebhook]    = useState("");
  const [statuses,   setStatuses]   = useState({}); // id → status
  const [syncStatus, setSyncStatus] = useState(null);
  const [sheetsBadge,setSheetsBadge]= useState(null);
  const [ready,      setReady]      = useState(false);

  // Load persisted data
  useEffect(()=>{
    (async()=>{
      const [inv,off,wh] = await Promise.all([db.get("inv"),db.get("off"),db.get("wh")]);
      if (inv) setInvoices(inv);
      if (off) setOffers(off);
      if (wh)  setWebhook(wh);
      setReady(true);
    })();
  },[]);

  // Sheets notification flash
  const flashSheets = ok => {
    setSheetsBadge(ok?"ok":"err");
    setSyncStatus(new Date().toLocaleTimeString("el-GR")+(ok?" ✓":" ✗"));
    setTimeout(()=>setSheetsBadge(null),4000);
  };

  // Process uploaded files
  const handleFiles = useCallback(async files => {
    for (const f of files) {
      const id = Date.now() + Math.random();
      const stub = { id, file_name: f.name, items:[], uploaded_at: new Date().toISOString() };

      setStatuses(p=>({...p,[id]:"scanning"}));
      setInvoices(p=>[stub,...p]);

      try {
        const b64 = await toB64(f);
        const dataArr = await extractInvoice(b64, f.type); // always array

        // Remove placeholder
        setInvoices(p => p.filter(x => x.id !== id));

        // Add each invoice separately
        const newInvs = dataArr.map((data, i) => ({
          ...data,
          id: id + i,
          file_name: f.name,
          uploaded_at: new Date().toISOString()
        }));

        setStatuses(p => {
          const next = {...p};
          delete next[id];
          newInvs.forEach(inv => { next[inv.id] = "syncing"; });
          return next;
        });

        setInvoices(p => {
          const updated = [...newInvs, ...p];
          db.set("inv", updated);
          return updated;
        });

        // Sync each to Sheets
        for (const inv of newInvs) {
          await syncSheet(webhook, "add_invoice", {invoice: inv});
          setStatuses(p=>({...p,[inv.id]:"done"}));
        }
        flashSheets(true);

      } catch(e) {
        setStatuses(p=>({...p,[id]:"error"}));
      }
    }
  }, [webhook]);

  const deleteInvoice = async id => {
    setInvoices(p=>{const u=p.filter(x=>x.id!==id); db.set("inv",u); return u;});
    await syncSheet(webhook,"delete_invoice",{invoice_id:id});
  };

  const saveOffers = async rows => {
    setOffers(rows); await db.set("off",rows);
    await syncSheet(webhook,"save_offers",{offers:rows});
  };

  const saveWebhook = async url => {
    setWebhook(url); await db.set("wh",url);
  };

  const fullSync = async () => {
    setSheetsBadge("syncing");
    const ok = await syncSheet(webhook,"full_sync",{invoices,offers});
    flashSheets(ok);
  };

  if (!ready) return (
    <div style={{background:T.bg, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center"}}>
      <div style={{color:T.green}}>Φόρτωση…</div>
    </div>
  );

  const TABS = [
    {id:"scan",    label:"Σάρωση",         icon:"M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2zM12 17a4 4 0 100-8 4 4 0 000 8z"},
    {id:"invoices",label:`Τιμολόγια (${invoices.length})`, icon:"M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6"},
    {id:"offers",  label:"Προσφορές",      icon:"M20.59 13.41l-7.17 7.17a2 2 0 01-2.83 0L2 12V2h10l8.59 8.59a2 2 0 010 2.82zM7 7h.01"},
    {id:"settings",label:"Sheets",         icon:"M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065zM15 12a3 3 0 11-6 0 3 3 0 016 0z"},
  ];

  return (
    <div style={{background:T.bg, minHeight:"100vh", fontFamily:"'DM Sans','Segoe UI',sans-serif", color:T.txt}}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&display=swap" rel="stylesheet"/>
      <style>{`
        * { box-sizing:border-box; margin:0; padding:0; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.35} }
        @keyframes slideIn { from{opacity:0;transform:translateY(-8px)} to{opacity:1;transform:translateY(0)} }
        input,select { font-family:inherit; }
        ::-webkit-scrollbar { width:4px; } ::-webkit-scrollbar-track { background:${T.bg}; }
        ::-webkit-scrollbar-thumb { background:${T.border}; border-radius:4px; }
      `}</style>

      {/* NAV */}
      <nav style={{borderBottom:`1px solid ${T.border}`, padding:"0 28px",
        display:"flex", alignItems:"center", gap:0, background:T.panel, position:"sticky", top:0, zIndex:10}}>
        {/* Logo */}
        <div style={{display:"flex", alignItems:"center", gap:10, padding:"14px 0", marginRight:24}}>
          <div style={{width:32, height:32, borderRadius:8, background:T.greenDim,
            display:"flex", alignItems:"center", justifyContent:"center"}}>
            <Ic p="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2M9 12l2 2 4-4" s={16} c={T.green}/>
          </div>
          <div>
            <div style={{fontSize:13, fontWeight:700, color:T.txt, lineHeight:1}}>F&B Invoices</div>
            <div style={{fontSize:10, color:T.txt3}}>Adama Hotel</div>
          </div>
        </div>

        {/* Tabs */}
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{
            display:"flex", alignItems:"center", gap:6, padding:"18px 14px",
            background:"transparent", border:"none", cursor:"pointer",
            color: tab===t.id ? T.green : T.txt2,
            borderBottom: `2px solid ${tab===t.id ? T.green : "transparent"}`,
            fontSize:13, fontWeight: tab===t.id ? 600 : 400,
            transition:"all 0.15s", whiteSpace:"nowrap"
          }}>
            <Ic p={t.icon} s={13} c={tab===t.id?T.green:T.txt3}/>
            {t.label}
            {t.id==="settings" && webhook && (
              <span style={{width:6,height:6,borderRadius:"50%",background:T.green,display:"inline-block"}}/>
            )}
          </button>
        ))}

        {/* Right: sync indicator + export */}
        <div style={{marginLeft:"auto", display:"flex", alignItems:"center", gap:10}}>
          {sheetsBadge && (
            <div style={{
              display:"flex", alignItems:"center", gap:6, padding:"5px 12px",
              background: sheetsBadge==="ok"?T.green+"18":sheetsBadge==="err"?T.red+"18":T.gold+"18",
              border:`1px solid ${sheetsBadge==="ok"?T.greenDim:sheetsBadge==="err"?T.redDim:T.goldDim}44`,
              borderRadius:20, fontSize:11, fontWeight:600,
              color: sheetsBadge==="ok"?T.green:sheetsBadge==="err"?T.red:T.gold,
              animation:"slideIn 0.2s ease"
            }}>
              <span style={{width:6,height:6,borderRadius:"50%",
                background:sheetsBadge==="ok"?T.green:sheetsBadge==="err"?T.red:T.gold,
                animation: sheetsBadge==="syncing"?"pulse 1s infinite":"none"}}/>
              {sheetsBadge==="ok"?"Sheet ενημερώθηκε ✓":sheetsBadge==="err"?"Αποτυχία Sheets":"Sync…"}
            </div>
          )}
          <button
            onClick={()=>exportXLSX(invoices,offers)}
            disabled={!invoices.length}
            style={{
              display:"flex", alignItems:"center", gap:6, padding:"7px 14px",
              background:invoices.length?T.gold+"18":"transparent",
              color:invoices.length?T.gold:T.txt3,
              border:`1px solid ${invoices.length?T.goldDim:T.border}`,
              borderRadius:8, fontSize:12, fontWeight:600, cursor:invoices.length?"pointer":"not-allowed"
            }}>
            <Ic p="M12 15V3m0 12l-4-4m4 4l4-4M2 17l.621 2.485A2 2 0 004.561 21h14.878a2 2 0 001.94-1.515L22 17" s={14} c={invoices.length?T.gold:T.txt3}/>
            Excel
          </button>
        </div>
      </nav>

      {/* CONTENT */}
      <div style={{padding:"24px 28px", maxWidth:1060, margin:"0 auto"}}>

        {tab==="scan" && (
          <div style={{display:"flex", flexDirection:"column", gap:16}}>
            <UploadZone onFiles={handleFiles}/>
            {invoices.length>0 && (
              <>
                <div style={{color:T.txt3, fontSize:11, letterSpacing:"0.07em", marginTop:4}}>
                  ΤΕΛΕΥΤΑΙΑ ΤΙΜΟΛΟΓΙΑ
                </div>
                {invoices.slice(0,5).map(inv=>(
                  <div key={inv.id} style={{animation:"slideIn 0.25s ease"}}>
                    <InvoiceCard inv={inv} status={statuses[inv.id]||"done"} onDelete={deleteInvoice}/>
                  </div>
                ))}
                {invoices.length>5 && (
                  <button onClick={()=>setTab("invoices")} style={{
                    background:"transparent", color:T.green, border:`1px solid ${T.greenDim}`,
                    borderRadius:8, padding:"10px", fontSize:12, cursor:"pointer"
                  }}>Προβολή όλων ({invoices.length}) →</button>
                )}
              </>
            )}
          </div>
        )}

        {tab==="invoices" && (
          <div style={{display:"flex", flexDirection:"column", gap:12}}>
            <StatsBar invoices={invoices}/>
            {invoices.length===0 ? (
              <div style={{textAlign:"center", color:T.txt3, padding:60, fontSize:14}}>
                Δεν υπάρχουν τιμολόγια. Ανέβασε από τη Σάρωση.
              </div>
            ) : invoices.map(inv=>(
              <InvoiceCard key={inv.id} inv={inv} status={statuses[inv.id]||"done"} onDelete={deleteInvoice}/>
            ))}
          </div>
        )}

        {tab==="offers" && (
          <OffersPanel offers={offers} onSave={saveOffers}/>
        )}

        {tab==="settings" && (
          <SettingsPanel url={webhook} onSave={saveWebhook} onFullSync={fullSync} lastSync={syncStatus}/>
        )}

      </div>
    </div>
  );
}
