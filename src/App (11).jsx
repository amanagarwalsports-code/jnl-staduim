import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ─── CONSTANTS ─────────────────────────────────────────────── */
const K = {
  master:"jln_master_v6",      // all vehicles ever visited (auto-updated)
  officials:"jln_officials_v1", // department officials
  visitors:"jln_visitors_v1",   // one-time visitors
  cnp:"jln_cnp_v2",            // come & play
  users:"jln_users_v5",
  banned:"jln_banned_v5",
  passes:"jln_passes_v5",
  log:"jln_log_v6"             // entry/exit log
};
const DEFAULT_USERS = [
  { id:"admin_default", username:"admin", password:"admin123", role:"admin", createdOn:"System" },
  { id:"staff_default", username:"staff", password:"staff123", role:"staff", createdOn:"System" }
];
const DEPARTMENTS = ["Administration","Protocol & VIP","Security","Operations","Media & Press","Events Management","Technical/IT","Finance","Sports Management","General"];
const fmt = d => new Date(d).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",hour12:true});
const fmtDate = d => new Date(d).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
const todayStr = () => new Date().toDateString();
const norm = p => p.replace(/\s+/g,"").toUpperCase();
const uid = () => Math.random().toString(36).slice(2,9);
async function gs(k){try{const r=await window.storage.get(k);return r?JSON.parse(r.value):null}catch{return null}}
async function ss(k,v){try{await window.storage.set(k,JSON.stringify(v))}catch{}}

/* ─── DESIGN TOKENS ──────────────────────────────────────────── */
const T = {
  navy:"#080f1e", mid:"#0d1a30", light:"#162842",
  gold:"#e8a020", goldL:"#fbbf24", goldD:"#92560a",
  green:"#16a34a", greenBg:"#052e16", greenT:"#86efac",
  red:"#dc2626", redBg:"#1c0707", redT:"#fca5a5",
  amber:"#d97706", amberBg:"#1c0f00", amberT:"#fcd34d",
  blue:"#2563eb", blueBg:"#0d1b3e", blueT:"#93c5fd",
  purple:"#7c3aed", purpleBg:"#1e0a3c", purpleT:"#c4b5fd",
  cyan:"#0891b2", cyanBg:"#041f2a", cyanT:"#67e8f9",
  border:"rgba(255,255,255,0.08)", card:"rgba(255,255,255,0.035)", muted:"rgba(255,255,255,0.4)"
};

/* ─── ATOMS ──────────────────────────────────────────────────── */
function Pill({color,children}){
  const m={green:{bg:"rgba(22,163,74,0.18)",c:T.greenT,b:"rgba(22,163,74,0.35)"},red:{bg:"rgba(220,38,38,0.18)",c:T.redT,b:"rgba(220,38,38,0.35)"},amber:{bg:"rgba(217,119,6,0.18)",c:T.amberT,b:"rgba(217,119,6,0.35)"},blue:{bg:"rgba(37,99,235,0.18)",c:T.blueT,b:"rgba(37,99,235,0.35)"},gold:{bg:"rgba(232,160,32,0.15)",c:T.goldL,b:"rgba(232,160,32,0.35)"},purple:{bg:"rgba(124,58,237,0.18)",c:T.purpleT,b:"rgba(124,58,237,0.35)"},cyan:{bg:"rgba(8,145,178,0.18)",c:T.cyanT,b:"rgba(8,145,178,0.35)"}};
  const s=m[color]||m.blue;
  return <span style={{background:s.bg,color:s.c,border:`1px solid ${s.b}`,padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,whiteSpace:"nowrap",textTransform:"uppercase"}}>{children}</span>;
}
const Card=({children,style})=><div style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:14,padding:"16px 15px",marginBottom:12,...style}}>{children}</div>;
const SecTitle=({children})=><div style={{fontSize:9,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",color:"rgba(255,255,255,0.28)",margin:"0 0 10px",display:"flex",alignItems:"center",gap:8}}><span style={{flex:1,height:"1px",background:"rgba(255,255,255,0.06)"}}></span><span>{children}</span><span style={{flex:1,height:"1px",background:"rgba(255,255,255,0.06)"}}></span></div>;
const BigBtn=({onClick,color,children,disabled,style})=>{
  const cfg={green:{bg:T.green,shadow:"rgba(22,163,74,0.3)"},red:{bg:T.red,shadow:"rgba(220,38,38,0.3)"},gold:{bg:"linear-gradient(135deg,#e8a020,#f5c842)",shadow:"rgba(232,160,32,0.35)"},navy:{bg:T.light,shadow:"transparent"},amber:{bg:T.amber,shadow:"rgba(217,119,6,0.3)"},purple:{bg:T.purple,shadow:"rgba(124,58,237,0.3)"},blue:{bg:T.blue,shadow:"rgba(37,99,235,0.3)"},cyan:{bg:T.cyan,shadow:"rgba(8,145,178,0.3)"}}[color]||{bg:T.light,shadow:"transparent"};
  return <button onClick={onClick} disabled={disabled} style={{width:"100%",padding:"13px 0",borderRadius:10,border:"none",background:disabled?"rgba(255,255,255,0.06)":cfg.bg,color:disabled?"rgba(255,255,255,0.22)":(color==="gold"?T.navy:"white"),fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,letterSpacing:1.2,cursor:disabled?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8,boxShadow:disabled?"none":`0 4px 16px ${cfg.shadow}`,transition:"opacity .15s",...style}}>{children}</button>;
};
const Toast=({msg,type})=>{
  const cfg={green:{bg:"rgba(22,163,74,0.95)",b:"rgba(22,163,74,0.5)"},red:{bg:"rgba(220,38,38,0.95)",b:"rgba(220,38,38,0.5)"},amber:{bg:"rgba(217,119,6,0.95)",b:"rgba(217,119,6,0.5)"},purple:{bg:"rgba(124,58,237,0.95)",b:"rgba(124,58,237,0.5)"},blue:{bg:"rgba(37,99,235,0.95)",b:"rgba(37,99,235,0.5)"},cyan:{bg:"rgba(8,145,178,0.95)",b:"rgba(8,145,178,0.5)"}}[type]||{bg:"rgba(13,26,48,0.97)",b:T.border};
  return msg?<div style={{position:"absolute",bottom:20,left:"50%",transform:"translateX(-50%)",background:cfg.bg,border:`1px solid ${cfg.b}`,color:"white",padding:"10px 20px",borderRadius:40,fontSize:12,fontWeight:600,zIndex:100,whiteSpace:"nowrap",pointerEvents:"none"}}>{msg}</div>:null;
};
function Hdr({role,onLogout,title,sub,onBack}){
  return <div style={{background:`linear-gradient(135deg,${T.mid},${T.light})`,padding:"13px 15px 11px",borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
      <div style={{display:"flex",alignItems:"center",gap:9}}>
        {onBack&&<button onClick={onBack} style={{background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.6)",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:600}}>← Back</button>}
        <div><div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,letterSpacing:1,color:T.goldL,lineHeight:1.1}}>{title}</div>
        {sub&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>{sub}</div>}</div>
      </div>
      <div style={{display:"flex",gap:6,alignItems:"center"}}>
        {role&&<div style={{background:role==="admin"?"rgba(232,160,32,0.12)":"rgba(37,99,235,0.12)",border:`1px solid ${role==="admin"?"rgba(232,160,32,0.3)":"rgba(37,99,235,0.3)"}`,borderRadius:6,padding:"3px 9px",fontSize:10,fontWeight:700,color:role==="admin"?T.goldL:T.blueT}}>{role==="admin"?"ADMIN":"STAFF"}</div>}
        {onLogout&&<button onClick={onLogout} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:600,cursor:"pointer"}}>Logout</button>}
      </div>
    </div>
  </div>;
}
function FormField({label,children,required}){
  return <div style={{marginBottom:12}}>
    <div style={{fontSize:10,fontWeight:600,letterSpacing:.8,textTransform:"uppercase",color:"rgba(255,255,255,0.38)",marginBottom:6}}>{label}{required&&<span style={{color:T.red,marginLeft:2}}>*</span>}</div>
    {children}
  </div>;
}
function ConfirmModal({msg,onYes,onNo}){
  return <div style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:60,padding:"0 20px",backdropFilter:"blur(4px)"}}>
    <div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:16,padding:"22px 20px",width:"100%"}}>
      <div style={{fontSize:13,lineHeight:1.65,marginBottom:18,color:"rgba(255,255,255,0.8)"}}>{msg}</div>
      <div style={{display:"flex",gap:9}}>
        <button onClick={onNo} style={{flex:1,padding:"11px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,fontSize:13,cursor:"pointer"}}>Cancel</button>
        <button onClick={onYes} style={{flex:1,padding:"11px 0",borderRadius:9,border:"none",background:T.red,color:"white",fontSize:13,fontWeight:700,cursor:"pointer"}}>Confirm</button>
      </div>
    </div>
  </div>;
}

/* ─── SHARED DB MANAGER ──────────────────────────────────────── */
// Generic database view/add/import/export component
function DbManager({title,color,icon,storageKey,fields,onBack,onLogout,t_,role}){
  const [records,setRecords]=useState([]);
  const [tab,setTab]=useState("list");
  const [search,setSearch]=useState("");
  const [form,setForm]=useState({});
  const [confirm,setConfirm]=useState(null);
  const [sel,setSel]=useState(null);
  const importRef=useRef();

  useEffect(()=>{gs(storageKey).then(d=>d&&setRecords(d));},[storageKey]);

  const saveRecord=async()=>{
    const required=fields.filter(f=>f.required);
    for(const f of required){if(!form[f.key]?.trim()){t_(`${f.label} is required`,"red");return;}}
    const rec={id:uid(),addedOn:fmtDate(Date.now()),...Object.fromEntries(fields.map(f=>[f.key,form[f.key]?.trim()||""]))};
    const u=[...records,rec];setRecords(u);await ss(storageKey,u);
    setForm({});t_("Record added ✓","green");setTab("list");
  };

  const removeRecord=async(id)=>{
    const u=records.filter(r=>r.id!==id);setRecords(u);await ss(storageKey,u);
    t_("Record removed","amber");setConfirm(null);setSel(null);setTab("list");
  };

  const exportXL=()=>{
    if(!records.length){t_("No data to export","amber");return;}
    const ws=XLSX.utils.json_to_sheet(records.map(r=>{
      const row={};
      fields.forEach(f=>row[f.label]=r[f.key]||"");
      row["Added On"]=r.addedOn||"";
      return row;
    }));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,title);
    XLSX.writeFile(wb,`JLN_${title.replace(/\s+/g,"_")}.xlsx`);
    t_("Downloaded ✓","green");
  };

  const importXL=ev=>{
    const file=ev.target.files[0];if(!file)return;
    if(importRef.current)importRef.current.value="";
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
        const getCol=(row,aliases)=>{for(const a of aliases){const found=Object.keys(row).find(k=>k.trim().toLowerCase()===a.toLowerCase());if(found&&row[found]!=="")return row[found].toString().trim();}return "";};
        let added=0;const cur=[...records];
        rows.forEach(row=>{
          const rec={id:uid(),addedOn:fmtDate(Date.now())};
          fields.forEach(f=>rec[f.key]=getCol(row,[f.label,...(f.aliases||[])]));
          const keyField=fields.find(f=>f.isKey);
          if(keyField&&!rec[keyField.key])return;
          if(keyField&&cur.find(r=>norm(r[keyField.key])===norm(rec[keyField.key])))return;
          cur.push(rec);added++;
        });
        setRecords(cur);ss(storageKey,cur);
        t_(`${added} record${added!==1?"s":""} imported ✓`,"green");
        setTab("list");
      }catch{t_("Error reading file","red");}
    };
    reader.readAsArrayBuffer(file);
  };

  const filtered=search.trim()?records.filter(r=>fields.some(f=>r[f.key]?.toLowerCase().includes(search.toLowerCase()))):records;
  const accentColor=T[color]||T.blue;

  if(tab==="detail"&&sel) return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="RECORD DETAILS" onBack={()=>{setTab("list");setSel(null);}}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        {fields.map(f=><div key={f.key} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${T.border}`,alignItems:"flex-start",gap:10}}>
          <span style={{fontSize:10,color:T.muted,fontWeight:600,letterSpacing:.5,textTransform:"uppercase",flexShrink:0}}>{f.label}</span>
          <span style={{fontSize:12,fontWeight:600,textAlign:"right",maxWidth:"60%"}}>{sel[f.key]||"—"}</span>
        </div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"9px 0",alignItems:"center"}}>
          <span style={{fontSize:10,color:T.muted,fontWeight:600,letterSpacing:.5,textTransform:"uppercase"}}>Added On</span>
          <span style={{fontSize:12,fontWeight:600}}>{sel.addedOn||"—"}</span>
        </div>
      </Card>
      <button onClick={()=>setConfirm(sel.id)} style={{width:"100%",padding:"11px 0",borderRadius:9,border:"1px solid rgba(220,38,38,0.3)",background:"rgba(220,38,38,0.08)",color:"#fca5a5",fontSize:13,cursor:"pointer",fontWeight:600,marginBottom:20}}>🗑 Remove This Record</button>
    </div>
    {confirm&&<ConfirmModal msg="Remove this record? This cannot be undone." onYes={()=>removeRecord(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;

  if(tab==="add") return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title={`ADD TO ${title.toUpperCase()}`} onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        {fields.filter(f=>f.type!=="readonly").map(f=><FormField key={f.key} label={f.label} required={f.required}>
          {f.type==="select"
            ?<select value={form[f.key]||""} onChange={e=>setForm(p=>({...p,[f.key]:e.target.value}))} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%"}}>
              <option value="">-- Select --</option>
              {(f.options||[]).map(o=><option key={o} value={o}>{o}</option>)}
            </select>
            :<input placeholder={f.placeholder||""} value={form[f.key]||""} onChange={e=>setForm(p=>({...p,[f.key]:e.target.value}))}
              style={f.big?{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center"}:{}}
              type={f.inputType||"text"} inputMode={f.inputMode||"text"} maxLength={f.maxLength}/>}
        </FormField>)}
        <BigBtn onClick={saveRecord} color={color} style={{marginTop:4}}>+ ADD RECORD</BigBtn>
      </Card>
    </div>
  </div>;

  if(tab==="import") return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="IMPORT FROM EXCEL" onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(37,99,235,0.07)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"12px 14px",marginBottom:14,fontSize:11,color:T.blueT,lineHeight:1.8}}>
        <b>Expected Excel columns:</b><br/>
        {fields.filter(f=>f.type!=="readonly").map(f=><span key={f.key} style={{fontFamily:"monospace",color:T.goldL,marginRight:8}}>{f.label}</span>)}
      </div>
      <div onClick={()=>importRef.current?.click()} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:10,padding:"32px",border:"2px dashed rgba(37,99,235,0.4)",borderRadius:12,cursor:"pointer",background:"rgba(37,99,235,0.04)"}}>
        <span style={{fontSize:36}}>📂</span>
        <div style={{fontSize:14,fontWeight:600,color:"rgba(255,255,255,0.85)"}}>Tap to Select Excel File</div>
        <div style={{fontSize:11,color:T.muted}}>Supports .xlsx and .xls</div>
      </div>
      <input ref={importRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importXL}/>
    </div>
  </div>;

  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title={title.toUpperCase()} sub={`${records.length} records`} onBack={onBack}/>
    <div style={{display:"flex",borderBottom:`1px solid ${T.border}`,background:T.mid}}>
      {[["list","📋 View"],["add","➕ Add"],["import","📥 Import"]].map(([v,l])=>(
        <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"10px 0",border:"none",borderBottom:`2px solid ${tab===v?accentColor:"transparent"}`,background:"transparent",color:tab===v?accentColor:T.muted,fontSize:11,fontWeight:tab===v?700:400,cursor:"pointer"}}>{l}</button>
      ))}
    </div>
    <div style={{padding:"14px 15px 0"}}>
      <input placeholder={`🔍 Search ${title.toLowerCase()}…`} value={search} onChange={e=>setSearch(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?accentColor:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:10,transition:"border-color .2s"}}/>
      <button onClick={exportXL} style={{width:"100%",padding:"9px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.65)",fontSize:12,cursor:"pointer",fontWeight:600,marginBottom:12}}>⬇ Download Excel</button>
      {search&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""}</div>}
      {!filtered.length?<div style={{textAlign:"center",color:T.muted,padding:"40px 0",fontSize:12}}>{search?"No records match":"No records yet"}</div>:
        filtered.slice().reverse().map(r=><div key={r.id} onClick={()=>{setSel(r);setTab("detail");}} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${accentColor}`,borderRadius:11,padding:"11px 13px",marginBottom:7,cursor:"pointer",transition:"background .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.055)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.03)"}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:800,letterSpacing:1}}>{r[fields[0].key]||"—"}</div>
          <div style={{fontSize:10,color:T.muted,marginTop:2}}>{fields.slice(1,3).map(f=>r[f.key]).filter(Boolean).join(" · ")}</div>
          {fields[3]&&r[fields[3].key]&&<div style={{fontSize:10,color:T.muted,marginTop:1}}>{r[fields[3].key]}</div>}
        </div>)
      }
    </div>
    {confirm&&<ConfirmModal msg="Remove this record?" onYes={()=>removeRecord(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── MASTER DATABASE VIEW ───────────────────────────────────── */
function MasterDbV({onBack,onLogout,master,t_,role}){
  const [search,setSearch]=useState("");
  const exportXL=()=>{
    if(!master.length){t_("No data","amber");return;}
    const ws=XLSX.utils.json_to_sheet(master.map(r=>({
      "Plate Number":r.plate,"Type":r.type||"","Name/Officer":r.name||r.officer||"","Department/Division":r.dept||r.division||"","Contact":r.contact||"","First Seen":r.firstSeen||"","Last Seen":r.lastSeen||""
    })));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Master");
    XLSX.writeFile(wb,"JLN_Master_Database.xlsx");t_("Downloaded ✓","green");
  };
  const filtered=search.trim()?master.filter(r=>norm(r.plate).includes(norm(search))||r.name?.toLowerCase().includes(search.toLowerCase())||r.officer?.toLowerCase().includes(search.toLowerCase())||r.dept?.toLowerCase().includes(search.toLowerCase())):master;
  const typeColor={official:"gold",visitor:"cyan","come&play":"purple",exception:"amber"};
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="MASTER DATABASE" sub={`${master.length} unique vehicles`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(232,160,32,0.07)",border:"1px solid rgba(232,160,32,0.2)",borderRadius:9,padding:"10px 14px",marginBottom:12,fontSize:11,color:T.amberT,lineHeight:1.6}}>
        📋 Auto-updated every time a vehicle enters. Contains all vehicles that have ever visited JLN Stadium.
      </div>
      <input placeholder="🔍 Search plate, name, department…" value={search} onChange={e=>setSearch(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?T.gold:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:10,transition:"border-color .2s"}}/>
      <button onClick={exportXL} style={{width:"100%",padding:"9px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.65)",fontSize:12,cursor:"pointer",fontWeight:600,marginBottom:12}}>⬇ Download Master Excel</button>
      {search&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""}</div>}
      {!filtered.length?<div style={{textAlign:"center",color:T.muted,padding:"40px 0",fontSize:12}}>{search?"No vehicles match":"No vehicles yet"}</div>:
        filtered.slice().reverse().map((r,i)=><div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${T[typeColor[r.type]||"green"]||T.green}`,borderRadius:11,padding:"11px 13px",marginBottom:7}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:2.5}}>{r.plate}</div>
            <Pill color={typeColor[r.type]||"green"}>{r.type||"official"}</Pill>
          </div>
          <div style={{fontSize:10,color:T.muted}}>{r.name||r.officer||"—"}{(r.dept||r.division)?` · ${r.dept||r.division}`:""}</div>
          <div style={{fontSize:9,color:"rgba(255,255,255,0.2)",marginTop:3}}>First: {r.firstSeen||"—"} · Last: {r.lastSeen||"—"}</div>
        </div>)
      }
    </div>
  </div>;
}

/* ─── HOME SCREEN ────────────────────────────────────────────── */
function HomeScreen({onSelectRole,onRequestPass}){
  return <div style={{minHeight:600,background:T.navy}}>
    <div style={{background:`linear-gradient(160deg,${T.mid},${T.light})`,padding:"36px 20px 28px"}}>
      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:4}}>
        <div style={{width:46,height:46,borderRadius:13,background:"rgba(232,160,32,0.1)",border:"1px solid rgba(232,160,32,0.25)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>🏟</div>
        <div>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:900,letterSpacing:1.5,color:T.goldL,lineHeight:1}}>JLN STADIUM</div>
          <div style={{fontSize:9,color:"rgba(255,255,255,0.3)",letterSpacing:2.5,marginTop:4,textTransform:"uppercase"}}>Vehicle Access Control</div>
        </div>
      </div>
    </div>
    <div style={{padding:"22px 16px 24px"}}>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Security Personnel</div>
      {[{role:"admin",icon:"◆",label:"Super Admin",desc:"Dashboard · Databases · Reports",border:"rgba(232,160,32,0.3)",clr:T.goldL,bg:"rgba(232,160,32,0.05)"},
        {role:"staff",icon:"◈",label:"Security Staff",desc:"Entry / Exit · Verification · Exemption",border:"rgba(37,99,235,0.35)",clr:T.blueT,bg:"rgba(37,99,235,0.05)"}
      ].map(({role,icon,label,desc,border,clr,bg})=>(
        <div key={role} onClick={()=>onSelectRole(role)} style={{background:bg,border:`1px solid ${border}`,borderRadius:13,padding:"16px 15px",marginBottom:9,cursor:"pointer",display:"flex",alignItems:"center",gap:12}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.04)"}
          onMouseLeave={e=>e.currentTarget.style.background=bg}>
          <div style={{width:36,height:36,borderRadius:9,background:"rgba(255,255,255,0.06)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:clr,flexShrink:0}}>{icon}</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,color:clr}}>{label}</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.38)",marginTop:2}}>{desc}</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:16}}>›</div>
        </div>
      ))}
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10,marginTop:20}}>Visitor / Public</div>
      <div onClick={onRequestPass} style={{background:"rgba(8,145,178,0.05)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:13,padding:"16px 15px",cursor:"pointer",display:"flex",alignItems:"center",gap:12}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(8,145,178,0.09)"}
        onMouseLeave={e=>e.currentTarget.style.background="rgba(8,145,178,0.05)"}>
        <div style={{width:36,height:36,borderRadius:9,background:"rgba(8,145,178,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:T.cyanT,flexShrink:0}}>🎫</div>
        <div style={{flex:1}}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,color:T.cyanT}}>Request Gate Pass</div>
          <div style={{fontSize:11,color:"rgba(255,255,255,0.38)",marginTop:2}}>Submit visit request · Admin reviews & approves</div>
        </div>
        <div style={{color:"rgba(255,255,255,0.2)",fontSize:16}}>›</div>
      </div>
      <div style={{textAlign:"center",fontSize:10,color:"rgba(255,255,255,0.15)",marginTop:28}}>
        Default — Admin: <span style={{color:"rgba(255,255,255,0.28)"}}>admin123</span> · Staff: <span style={{color:"rgba(255,255,255,0.28)"}}>staff123</span>
      </div>
    </div>
  </div>;
}

/* ─── LOGIN ──────────────────────────────────────────────────── */
function LoginScreen({role,users,onLogin,onBack}){
  const [selUser,setSelUser]=useState("");
  const [pwd,setPwd]=useState(""); const [err,setErr]=useState("");
  const roleUsers=users.filter(u=>u.role===role);
  useEffect(()=>{if(roleUsers.length===1)setSelUser(roleUsers[0].id);},[]);
  const try_=()=>{const u=users.find(x=>x.id===selUser&&x.password===pwd);if(u)onLogin(u);else{setErr("Incorrect password");setPwd("");}};
  const isAdmin=role==="admin";
  return <div style={{minHeight:600,background:T.navy}}>
    <div style={{background:`linear-gradient(160deg,${T.mid},${T.light})`,padding:"22px 16px 20px"}}>
      <button onClick={onBack} style={{background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.55)",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",marginBottom:16,fontWeight:600}}>← Back</button>
      <div style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,color:isAdmin?T.goldL:T.blueT}}>{isAdmin?"Super Admin":"Security Staff"}</div>
    </div>
    <div style={{padding:"22px 16px"}}>
      {roleUsers.length>1&&<FormField label="Account" required>
        <select value={selUser} onChange={e=>setSelUser(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%"}}>
          <option value="">— Select account —</option>
          {roleUsers.map(u=><option key={u.id} value={u.id}>{u.username}</option>)}
        </select>
      </FormField>}
      {roleUsers.length===1&&<div style={{background:"rgba(255,255,255,0.04)",borderRadius:9,padding:"11px 13px",marginBottom:14,fontSize:13,color:"rgba(255,255,255,0.6)",border:`1px solid ${T.border}`}}>Logging in as <b style={{color:"white"}}>{roleUsers[0].username}</b></div>}
      <FormField label="Password" required>
        <input type="password" autoFocus placeholder="Enter password" value={pwd} onChange={e=>{setPwd(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&try_()}/>
      </FormField>
      {err&&<div style={{color:"#fca5a5",fontSize:12,marginBottom:12}}>⚠ {err}</div>}
      <BigBtn onClick={try_} color={isAdmin?"gold":"blue"} style={{marginTop:6}}>Sign In →</BigBtn>
    </div>
  </div>;
}

/* ─── REQUEST GATE PASS ──────────────────────────────────────── */
function RequestPassScreen({onBack,passes,setPasses}){
  const [name,setName]=useState(""); const [car,setCar]=useState("");
  const [phone,setPhone]=useState(""); const [purpose,setPurpose]=useState("");
  const [dept,setDept]=useState(""); const [refName,setRefName]=useState("");
  const [submitted,setSubmitted]=useState(false); const [refId,setRefId]=useState("");

  const submit=()=>{
    if(!name.trim()||!car.trim()||!phone.trim()||!purpose.trim()){alert("Please fill all required fields.");return;}
    const id="GP-"+Math.random().toString(36).slice(2,8).toUpperCase();
    const entry={id,name:name.trim(),car:car.toUpperCase().trim(),phone:phone.trim(),purpose:purpose.trim(),dept:dept.trim(),refName:refName.trim(),status:"pending",requestedOn:Date.now(),date:todayStr()};
    const np=[...(passes||[]),entry];setPasses(np);ss(K.passes,np);
    setRefId(id);setSubmitted(true);
  };

  if(submitted) return <div style={{minHeight:600,background:T.navy}}>
    <Hdr title="REQUEST SUBMITTED" onBack={onBack}/>
    <div style={{padding:"32px 18px",textAlign:"center"}}>
      <div style={{width:64,height:64,borderRadius:20,background:"rgba(22,163,74,0.15)",border:"1px solid rgba(22,163,74,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,margin:"0 auto 18px"}}>✅</div>
      <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,color:T.greenT,marginBottom:8}}>Request Submitted</div>
      <div style={{background:"rgba(255,255,255,0.04)",border:"1px dashed rgba(232,160,32,0.35)",borderRadius:12,padding:"18px",marginBottom:24}}>
        <div style={{fontSize:9,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",color:"rgba(255,255,255,0.3)",marginBottom:8}}>Reference ID</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:30,fontWeight:900,letterSpacing:5,color:T.goldL}}>{refId}</div>
      </div>
      <BigBtn onClick={onBack} color="navy">← Back to Home</BigBtn>
    </div>
  </div>;

  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr title="REQUEST GATE PASS" onBack={onBack}/>
    <div style={{padding:"14px 14px 0"}}>
      <Card>
        <SecTitle>VISITOR DETAILS</SecTitle>
        <FormField label="Full Name" required><input placeholder="e.g. Priya Sharma" value={name} onChange={e=>setName(e.target.value)}/></FormField>
        <FormField label="Vehicle Number" required><input placeholder="DL 01 AB 1234" value={car} onChange={e=>setCar(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:700,letterSpacing:4,textAlign:"center"}}/></FormField>
        <FormField label="Phone Number" required><input placeholder="10-digit mobile" value={phone} onChange={e=>setPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
        <FormField label="Purpose of Visit" required><input placeholder="e.g. Meeting, Event, Delivery" value={purpose} onChange={e=>setPurpose(e.target.value)}/></FormField>
        <FormField label="Department to Visit">
          <select value={dept} onChange={e=>setDept(e.target.value)} style={{background:"rgba(255,255,255,0.07)",border:`1px solid ${T.border}`,borderRadius:8,color:dept?"white":"rgba(255,255,255,0.3)",padding:"10px 13px",fontSize:14,width:"100%"}}>
            <option value="">-- Select Department --</option>
            {DEPARTMENTS.map(d=><option key={d} value={d}>{d}</option>)}
          </select>
        </FormField>
        <FormField label="Reference Officer Name (Optional)"><input placeholder="Name of JLN officer who referred you" value={refName} onChange={e=>setRefName(e.target.value)}/></FormField>
      </Card>
      <BigBtn onClick={submit} color="cyan" style={{marginBottom:20}}>🎫 SUBMIT GATE PASS REQUEST</BigBtn>
    </div>
  </div>;
}

/* ─── GATE PASS MGMT (Admin) ─────────────────────────────────── */
function GatePassMgmtV({onBack,passes,setPasses,master,setMaster,visitors,setVisitors,onLogout,t_}){
  const [filter,setFilter]=useState("pending");
  const filtered=passes.filter(p=>filter==="all"||p.status===filter).slice().reverse();

  const approve=id=>{
    const pass=passes.find(p=>p.id===id);
    const u=passes.map(p=>p.id===id?{...p,status:"approved"}:p);
    setPasses(u);ss(K.passes,u);
    if(pass){
      // Add to visitors DB if not already there
      if(!visitors.find(v=>norm(v.carNumber)===norm(pass.car))){
        const nv=[...visitors,{id:uid(),visitorName:pass.name,carNumber:pass.car.toUpperCase(),purpose:pass.purpose,officerReference:pass.refName||"",contactNumber:pass.phone,entryTime:"",exitTime:"",addedOn:fmtDate(Date.now()),passId:pass.id}];
        setVisitors(nv);ss(K.visitors,nv);
      }
      // Auto-add to master
      if(!master.find(e=>norm(e.plate)===norm(pass.car))){
        const nm=[...master,{plate:pass.car.toUpperCase(),type:"visitor",name:pass.name,dept:pass.dept,contact:pass.phone,firstSeen:fmtDate(Date.now()),lastSeen:fmtDate(Date.now())}];
        setMaster(nm);ss(K.master,nm);
      }
    }
    t_("Pass approved + added to Visitors DB ✓","green");
  };

  const reject=id=>{const u=passes.map(p=>p.id===id?{...p,status:"rejected"}:p);setPasses(u);ss(K.passes,u);t_("Pass rejected","amber");};
  const statusColor={pending:"amber",approved:"green",rejected:"red",used:"cyan"};

  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="GATE PASS REQUESTS" sub={`${passes.filter(p=>p.status==="pending").length} pending`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{display:"flex",gap:6,marginBottom:13}}>
        {["pending","approved","rejected","used","all"].map(f=><button key={f} onClick={()=>setFilter(f)} style={{flex:1,padding:"7px 0",borderRadius:9,border:`1px solid ${filter===f?"rgba(232,160,32,0.5)":T.border}`,background:filter===f?"rgba(232,160,32,0.1)":"transparent",color:filter===f?T.goldL:T.muted,fontSize:11,cursor:"pointer",fontWeight:filter===f?700:400,textTransform:"capitalize"}}>{f}</button>)}
      </div>
      {!filtered.length?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>No {filter} requests</div>:
        filtered.map(p=><div key={p.id} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:12,padding:"13px 14px",marginBottom:9}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:7}}>
            <div>
              <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,letterSpacing:3}}>{p.car}</div>
              <div style={{fontSize:12,color:"rgba(255,255,255,0.75)",marginTop:2}}>{p.name}</div>
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
              <Pill color={statusColor[p.status]||"blue"}>{p.status}</Pill>
              <div style={{fontSize:9,color:T.muted}}>{p.id}</div>
            </div>
          </div>
          <div style={{fontSize:10,color:T.muted,marginBottom:2}}>📞 {p.phone} · 🎯 {p.purpose}</div>
          {p.dept&&<div style={{fontSize:10,color:T.muted,marginBottom:6}}>🏛 {p.dept}</div>}
          {p.refName&&<div style={{fontSize:10,color:T.muted,marginBottom:6}}>👤 Ref: {p.refName}</div>}
          <div style={{display:"flex",gap:7}}>
            {p.status==="pending"&&<>
              <button onClick={()=>approve(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.green,color:"white",fontSize:11,fontWeight:700,cursor:"pointer"}}>✓ Approve</button>
              <button onClick={()=>reject(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.red,color:"white",fontSize:11,fontWeight:700,cursor:"pointer"}}>✕ Reject</button>
            </>}
          </div>
        </div>)}
    </div>
  </div>;
}

/* ─── STAFF APP ──────────────────────────────────────────────── */
function StaffApp({onLogout,master,setMaster,officials,visitors,setVisitors,cnp,log,setLog,banned,passes,setPasses,user}){
  const [view,setView]=useState("verify");
  const [plate,setPlate]=useState(""); const [result,setResult]=useState(null);
  const [last4,setLast4]=useState(""); const [suggestions,setSuggestions]=useState([]);
  const [guestName,setGuestName]=useState(""); const [guestReason,setGuestReason]=useState("");
  const [toast,setToast]=useState(null);
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2400);};

  const allVehicles=[
    ...officials.map(o=>({plate:norm(o.carNumber),label:o.officialName,sub:o.department,badge:"Official",badgeColor:"gold",type:"official",raw:o})),
    ...visitors.map(v=>({plate:norm(v.carNumber),label:v.visitorName,sub:v.purpose,badge:"Visitor",badgeColor:"cyan",type:"visitor",raw:v})),
    ...cnp.map(c=>({plate:norm(c.carNumber),label:c.memberName,sub:"Come & Play",badge:"Come & Play",badgeColor:"purple",type:"cnp",raw:c})),
    ...master.filter(m=>!officials.find(o=>norm(o.carNumber)===norm(m.plate))&&!visitors.find(v=>norm(v.carNumber)===norm(m.plate))&&!cnp.find(c=>norm(c.carNumber)===norm(m.plate))).map(m=>({plate:norm(m.plate),label:m.name||m.officer||"",sub:m.dept||m.division||"",badge:"Master",badgeColor:"green",type:"master",raw:m}))
  ];

  const onLast4Change=v=>{
    setLast4(v);setPlate("");setResult(null);setSuggestions([]);
    if(v.length>=1){
      const n=norm(v);
      const matches=allVehicles.filter(e=>e.plate.endsWith(n));
      // Deduplicate by plate
      const seen=new Set();
      setSuggestions(matches.filter(m=>{if(seen.has(m.plate))return false;seen.add(m.plate);return true;}));
    }
  };
  const selectSuggestion=p=>{setPlate(p);setLast4("");setSuggestions([]);setResult(null);};

  const doAction=action=>{
    const raw=plate.trim();if(!raw){t_("Enter a vehicle number","red");return;}
    const isBanned=banned.find(b=>norm(b.plate)===norm(raw));
    if(isBanned&&action==="entry"){setResult({status:"banned",plate:raw.toUpperCase(),detail:`Reason: ${isBanned.reason||"Banned"}`});return;}
    const now=Date.now();
    const vehicle=allVehicles.find(v=>v.plate===norm(raw));
    if(action==="entry"){
      const alreadyIn=log.find(e=>norm(e.plate)===norm(raw)&&!e.exitTime&&e.date===todayStr());
      if(alreadyIn){t_("Vehicle already inside","amber");return;}
      const approvedPass=passes.find(p=>norm(p.car)===norm(raw)&&p.status==="approved");
      const usedPass=passes.find(p=>norm(p.car)===norm(raw)&&p.status==="used");
      if(usedPass&&!vehicle&&!approvedPass){setResult({status:"denied",plate:raw.toUpperCase(),detail:"Gate pass already used — single entry only"});return;}
      if(vehicle&&vehicle.type==="visitor"){
        const prev=log.find(e=>norm(e.plate)===norm(raw));
        if(prev){setResult({status:"denied",plate:raw.toUpperCase(),detail:"Single entry already used for this vehicle"});return;}
      }
      const entry={id:now,plate:raw.toUpperCase(),type:vehicle?.type||"guest",entryTime:now,exitTime:null,date:todayStr(),name:vehicle?.label||"",division:vehicle?.sub||"",loggedBy:user.username};
      if(approvedPass){entry.passId=approvedPass.id;entry.name=approvedPass.name;}
      const nl=[...log,entry];setLog(nl);ss(K.log,nl);
      if(approvedPass){const np=passes.map(p=>p.id===approvedPass.id?{...p,status:"used",entryTime:now,entryLoggedBy:user.username}:p);setPasses(np);ss(K.passes,np);}
      // Update master lastSeen
      if(master.find(m=>norm(m.plate)===norm(raw))){
        const nm=master.map(m=>norm(m.plate)===norm(raw)?{...m,lastSeen:fmtDate(now)}:m);setMaster(nm);ss(K.master,nm);
      } else {
        const nm=[...master,{plate:raw.toUpperCase(),type:vehicle?.type||"guest",name:vehicle?.label||"",dept:vehicle?.sub||"",firstSeen:fmtDate(now),lastSeen:fmtDate(now)}];
        setMaster(nm);ss(K.master,nm);
      }
      const statusLabel=vehicle?"allowed":approvedPass?"allowed_guest":"allowed";
      setResult({status:statusLabel,plate:raw.toUpperCase(),detail:vehicle?`${vehicle.badge}: ${vehicle.label||"—"}`:approvedPass?`Gate Pass: ${approvedPass.id}`:"Entry logged"});
      t_("Entry logged ✓","green");
    } else {
      const entry=[...log].reverse().find(e=>norm(e.plate)===norm(raw)&&!e.exitTime);
      if(!entry){t_("No active entry found","amber");return;}
      const nl=log.map(e=>e.id===entry.id?{...e,exitTime:now}:e);setLog(nl);ss(K.log,nl);
      if(entry.passId){const np=passes.map(p=>p.id===entry.passId?{...p,exitTime:now}:p);setPasses(np);ss(K.passes,np);}
      // Update visitor exit time
      if(entry.type==="visitor"){const nv=visitors.map(v=>norm(v.carNumber)===norm(raw)?{...v,exitTime:fmt(now)}:v);setVisitors(nv);ss(K.visitors,nv);}
      setResult({status:"exit",plate:raw.toUpperCase(),detail:`Exited at ${fmt(now)}`});t_("Exit logged ✓","amber");
    }
  };

  const addException=()=>{
    if(!plate.trim()){t_("Enter vehicle number","red");return;}
    if(!guestName.trim()){t_("Enter visitor name","red");return;}
    if(!guestReason.trim()){t_("Enter purpose","red");return;}
    const now=Date.now();
    const entry={id:now,plate:plate.toUpperCase(),type:"exception",entryTime:now,exitTime:null,date:todayStr(),name:guestName.trim(),division:guestReason.trim(),loggedBy:user.username};
    const nl=[...log,entry];setLog(nl);ss(K.log,nl);
    // Add to visitors DB
    const nv=[...visitors,{id:uid(),visitorName:guestName.trim(),carNumber:plate.toUpperCase(),purpose:guestReason.trim(),officerReference:"",contactNumber:"",entryTime:fmt(now),exitTime:"",addedOn:fmtDate(now)}];
    setVisitors(nv);ss(K.visitors,nv);
    // Add to master
    if(!master.find(m=>norm(m.plate)===norm(plate))){const nm=[...master,{plate:plate.toUpperCase(),type:"exception",name:guestName.trim(),dept:guestReason.trim(),firstSeen:fmtDate(now),lastSeen:fmtDate(now)}];setMaster(nm);ss(K.master,nm);}
    setGuestName("");setGuestReason("");
    t_("Exception entry logged ✓","amber");setView("verify");
    setResult({status:"allowed_guest",plate:plate.toUpperCase(),detail:`Exception: ${guestName.trim()}`});
  };

  const approvedPasses=passes.filter(p=>p.status==="approved");
  const resultColors={allowed:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},allowed_guest:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},exit:{bg:"rgba(217,119,6,0.12)",border:T.amber,tc:T.amberT},denied:{bg:"rgba(220,38,38,0.12)",border:T.red,tc:T.redT},banned:{bg:"rgba(127,29,29,0.3)",border:"#7f1d1d",tc:"#fca5a5"}};
  const rc=result?resultColors[result.status]||resultColors.denied:null;

  if(view==="exception") return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="GRANT EXEMPTION" sub="One-time exception entry" onBack={()=>{setView("verify");setResult(null);}}/>
    <div style={{padding:"15px 15px 0"}}>
      <div style={{background:"rgba(217,119,6,0.08)",border:"1px solid rgba(217,119,6,0.3)",borderRadius:12,padding:"14px 15px",marginBottom:14}}>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,color:T.amberT}}>Exception / Exemption Entry</div>
        <div style={{fontSize:11,color:"rgba(253,211,77,0.6)",marginTop:4}}>Fill 3 fields to grant one-time entry</div>
      </div>
      <Card>
        <FormField label="Vehicle Registration Number" required><input placeholder="DL 01 AB 1234" value={plate} onChange={e=>setPlate(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center"}}/></FormField>
        <FormField label="Name of Visitor" required><input placeholder="Full name" value={guestName} onChange={e=>setGuestName(e.target.value)}/></FormField>
        <FormField label="Purpose of Visit" required><input placeholder="e.g. Meeting, Delivery, Event" value={guestReason} onChange={e=>setGuestReason(e.target.value)}/></FormField>
        <BigBtn onClick={addException} color="amber" style={{marginTop:6}}>⚠ ALLOW EXCEPTION ENTRY</BigBtn>
      </Card>
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;

  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="SECURITY STAFF" sub={user.username}/>
    <div style={{padding:"16px 15px 0"}}>
      <Card style={{marginBottom:14}}>
        <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.28)",marginBottom:8}}>Vehicle Registration</div>
        {plate?<div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:"rgba(232,160,32,0.08)",border:`1px solid ${T.gold}`,borderRadius:10,padding:"12px 14px",marginBottom:10,cursor:"pointer"}} onClick={()=>{setPlate("");setLast4("");setSuggestions([]);setResult(null);}}>
          <span style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,letterSpacing:5,color:T.goldL}}>{plate}</span>
          <span style={{fontSize:11,color:T.muted,background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,borderRadius:6,padding:"3px 8px"}}>✕ Clear</span>
        </div>:
        <div style={{position:"relative",marginBottom:suggestions.length?0:10}}>
          <input placeholder="Type last 4 digits  e.g. 1234" value={last4} onChange={e=>onLast4Change(e.target.value.toUpperCase())} maxLength={6} style={{fontFamily:"'Barlow Condensed',monospace",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center",background:"rgba(255,255,255,0.05)",border:`1px solid ${last4?T.gold:T.border}`,borderRadius:10,color:"white",padding:"13px",width:"100%",outline:"none",transition:"border-color .2s",marginBottom:0}}/>
          {suggestions.length>0&&<div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:10,marginTop:4,overflow:"hidden",marginBottom:10}}>
            {suggestions.map((s,i)=><div key={i} onClick={()=>selectSuggestion(s.plate)} style={{padding:"11px 14px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${T.border}`,transition:"background .12s"}}
              onMouseEnter={e=>e.currentTarget.style.background="rgba(232,160,32,0.08)"}
              onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <div><span style={{fontFamily:"'Barlow Condensed'",fontSize:17,fontWeight:800,letterSpacing:3,color:T.goldL}}>{s.plate}</span>
              <div style={{fontSize:10,color:T.muted,marginTop:2}}>{s.label}{s.sub?` · ${s.sub}`:""}</div></div>
              <Pill color={s.badgeColor||"green"}>{s.badge}</Pill>
            </div>)}
            <div style={{padding:"9px 14px",fontSize:10,color:T.muted,textAlign:"center"}}>Tap to select</div>
          </div>}
          {last4&&!suggestions.length&&<div style={{padding:"12px 14px",fontSize:12,color:T.muted,textAlign:"center",marginBottom:10}}>No match — <span style={{color:T.amberT,cursor:"pointer"}} onClick={()=>{setPlate(last4);setSuggestions([]);setLast4("");}}>use "{last4}" directly</span></div>}
        </div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <BigBtn onClick={()=>doAction("entry")} color="green" disabled={!plate}>▲ ENTRY</BigBtn>
          <BigBtn onClick={()=>doAction("exit")} color="red" disabled={!plate}>▼ EXIT</BigBtn>
        </div>
      </Card>

      {result&&<div style={{borderRadius:14,padding:"20px 16px",textAlign:"center",marginBottom:14,background:rc.bg,border:`1px solid ${rc.border}`}}>
        <div style={{fontSize:32,marginBottom:6}}>{result.status==="allowed"||result.status==="allowed_guest"?"✅":result.status==="exit"?"🔄":result.status==="banned"?"⛔":"🚫"}</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:800,letterSpacing:1.5,color:rc.tc}}>
          {result.status==="allowed"||result.status==="allowed_guest"?"ENTRY ALLOWED":result.status==="exit"?"EXIT RECORDED":result.status==="banned"?"BANNED VEHICLE":"NOT ALLOWED"}
        </div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:700,letterSpacing:5,marginTop:4,opacity:.6,color:rc.tc}}>{result.plate}</div>
        {result.detail&&<div style={{fontSize:11,marginTop:5,opacity:.7,color:rc.tc}}>{result.detail}</div>}
        {result.status==="denied"&&<button onClick={()=>setView("exception")} style={{marginTop:12,background:"rgba(217,119,6,0.15)",border:"1px solid rgba(217,119,6,0.4)",color:T.amberT,borderRadius:8,padding:"8px 20px",fontSize:12,cursor:"pointer",fontWeight:700}}>⚠ Grant Exemption</button>}
      </div>}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
        <div onClick={()=>setView("exception")} style={{background:"rgba(217,119,6,0.06)",border:"1px solid rgba(217,119,6,0.25)",borderRadius:12,padding:"12px 14px",cursor:"pointer",display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"rgba(217,119,6,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>⚠</div>
          <div><div style={{fontFamily:"'Barlow Condensed'",fontSize:13,fontWeight:700,color:T.amberT}}>Grant Exemption</div>
          <div style={{fontSize:9,color:T.muted,marginTop:1}}>One-time entry</div></div>
        </div>
        <div style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.25)",borderRadius:12,padding:"12px 14px",display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"rgba(8,145,178,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🎫</div>
          <div><div style={{fontFamily:"'Barlow Condensed'",fontSize:13,fontWeight:700,color:T.cyanT}}>Gate Passes</div>
          <div style={{fontSize:9,color:T.muted,marginTop:1}}>{approvedPasses.length} approved</div></div>
        </div>
      </div>
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;
}

/* ─── ADMIN APP ──────────────────────────────────────────────── */
function AdminApp({onLogout,master,setMaster,officials,setOfficials,visitors,setVisitors,cnp,setCnp,log,setLog,users,setUsers,banned,setBanned,passes,setPasses,user}){
  const [view,setView]=useState("dashboard");
  const [toast,setToast]=useState(null);
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2400);};
  const props={onLogout,master,setMaster,officials,setOfficials,visitors,setVisitors,cnp,setCnp,log,setLog,users,setUsers,banned,setBanned,passes,setPasses,t_,user};

  const OFFICIAL_FIELDS=[
    {key:"officialName",label:"Name of Official",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate"]},
    {key:"department",label:"Department",required:true,type:"select",options:DEPARTMENTS},
  ];
  const VISITOR_FIELDS=[
    {key:"visitorName",label:"Name of Visitor",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate"]},
    {key:"purpose",label:"Purpose",required:true},
    {key:"officerReference",label:"Officer Reference"},
    {key:"contactNumber",label:"Contact Number",inputType:"tel",inputMode:"numeric",maxLength:15},
    {key:"entryTime",label:"Entry Time"},
    {key:"exitTime",label:"Exit Time"},
  ];
  const CNP_FIELDS=[
    {key:"memberName",label:"Name of Member",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate"]},
    {key:"cardNumber",label:"SAI Card Number",required:true},
    {key:"contactNumber",label:"Contact Number",required:true,inputType:"tel",inputMode:"numeric",maxLength:15},
  ];

  const downloadAllExcel=()=>{
    const wb=XLSX.utils.book_new();
    if(master.length){const ws=XLSX.utils.json_to_sheet(master.map(r=>({"Plate":r.plate,"Type":r.type||"","Name":r.name||"","Department":r.dept||"","Contact":r.contact||"","First Seen":r.firstSeen||"","Last Seen":r.lastSeen||""})));XLSX.utils.book_append_sheet(wb,ws,"Master DB");}
    if(officials.length){const ws=XLSX.utils.json_to_sheet(officials.map(r=>({"Name of Official":r.officialName,"Car Number":r.carNumber,"Department":r.department})));XLSX.utils.book_append_sheet(wb,ws,"Officials");}
    if(visitors.length){const ws=XLSX.utils.json_to_sheet(visitors.map(r=>({"Name":r.visitorName,"Car Number":r.carNumber,"Purpose":r.purpose,"Officer Reference":r.officerReference||"","Contact":r.contactNumber||"","Entry Time":r.entryTime||"","Exit Time":r.exitTime||""})));XLSX.utils.book_append_sheet(wb,ws,"Visitors");}
    if(cnp.length){const ws=XLSX.utils.json_to_sheet(cnp.map(r=>({"Name":r.memberName,"Car Number":r.carNumber,"SAI Card Number":r.cardNumber,"Contact":r.contactNumber})));XLSX.utils.book_append_sheet(wb,ws,"Come & Play");}
    XLSX.writeFile(wb,"JLN_All_Databases.xlsx");t_("All databases downloaded ✓","green");
  };

  if(view==="master") return <MasterDbV onBack={()=>setView("dashboard")} onLogout={onLogout} master={master} t_={t_} role="admin"/>;
  if(view==="officials") return <DbManager title="Officials Database" color="gold" icon="◆" storageKey={K.officials} fields={OFFICIAL_FIELDS} onBack={()=>setView("dashboard")} onLogout={onLogout} t_={t_} role="admin"/>;
  if(view==="visitors") return <DbManager title="Visitors Database" color="cyan" icon="🎫" storageKey={K.visitors} fields={VISITOR_FIELDS} onBack={()=>setView("dashboard")} onLogout={onLogout} t_={t_} role="admin"/>;
  if(view==="cnp") return <DbManager title="Come & Play Database" color="purple" icon="🏸" storageKey={K.cnp} fields={CNP_FIELDS} onBack={()=>setView("dashboard")} onLogout={onLogout} t_={t_} role="admin"/>;
  if(view==="passes") return <GatePassMgmtV onBack={()=>setView("dashboard")} passes={passes} setPasses={setPasses} master={master} setMaster={setMaster} visitors={visitors} setVisitors={setVisitors} onLogout={onLogout} t_={t_}/>;
  if(view==="banned") return <DbManager title="Banned Vehicles" color="red" icon="⛔" storageKey={K.banned} fields={[{key:"plate",label:"Plate Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number"]},{key:"reason",label:"Reason"}]} onBack={()=>setView("dashboard")} onLogout={onLogout} t_={t_} role="admin"/>;
  if(view==="users"){
    return <div style={{minHeight:600,position:"relative",background:T.navy}}>
      <Hdr role="admin" onLogout={onLogout} title="USER MANAGEMENT" sub={`${users.length} accounts`} onBack={()=>setView("dashboard")}/>
      <div style={{padding:"14px 15px 0"}}>
        {users.map(u=><div key={u.id} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${u.role==="admin"?T.gold:T.blue}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700}}>{u.username}{u.id===user.id?<span style={{fontSize:10,color:T.muted,fontWeight:400}}> · you</span>:null}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:1}}>{u.role==="admin"?"Super Admin":"Staff"}</div>
          </div>
          <Pill color={u.role==="admin"?"gold":"blue"}>{u.role}</Pill>
        </div>)}
      </div>
    </div>;
  }

  const todayLog=log.filter(e=>e.date===todayStr());
  const inside=log.filter(e=>!e.exitTime&&e.date===todayStr()).length;
  const pendingPasses=passes.filter(p=>p.status==="pending");

  return <div style={{minHeight:600,background:T.navy,position:"relative"}}>
    <Hdr role="admin" onLogout={onLogout} title="ADMIN DASHBOARD" sub={`${user.username} · ${fmtDate(Date.now())}`}/>
    <div style={{padding:"16px 15px 0"}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9,marginBottom:16}}>
        {[{label:"Inside Now",value:inside,icon:"🚗",color:T.goldL},{label:"Today's Entries",value:todayLog.length,icon:"📋",color:T.blueT},{label:"Officials",value:officials.length,icon:"◆",color:T.amberT},{label:"Visitors",value:visitors.length,icon:"🎫",color:T.cyanT}].map(s=><div key={s.label} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:12,padding:"14px 13px"}}>
          <div style={{fontSize:16,marginBottom:5,opacity:.7}}>{s.icon}</div>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:28,fontWeight:800,color:s.color,lineHeight:1}}>{s.value}</div>
          <div style={{fontSize:10,color:T.muted,marginTop:4}}>{s.label}</div>
        </div>)}
      </div>

      {pendingPasses.length>0&&<div onClick={()=>setView("passes")} style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:11,padding:"13px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:11,cursor:"pointer"}}>
        <div style={{width:32,height:32,borderRadius:8,background:"rgba(8,145,178,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🎫</div>
        <div style={{flex:1}}><div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700,color:T.cyanT}}>{pendingPasses.length} Pass Request{pendingPasses.length>1?"s":""} Pending</div>
        <div style={{fontSize:10,color:T.muted,marginTop:1}}>Tap to review</div></div>
        <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
      </div>}

      <div style={{marginBottom:10}}>
        <button onClick={downloadAllExcel} style={{width:"100%",padding:"11px 0",borderRadius:9,border:`1px solid rgba(22,163,74,0.35)`,background:"rgba(22,163,74,0.08)",color:T.greenT,fontSize:13,cursor:"pointer",fontWeight:700}}>📥 Download All Databases (Excel)</button>
      </div>

      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Databases</div>
      {[
        {icon:"🗂",label:"Master Database",sub:`${master.length} unique vehicles · auto-updated`,action:"master",color:T.gold},
        {icon:"◆",label:"Officials Database",sub:`${officials.length} officials registered`,action:"officials",color:T.amber},
        {icon:"🎫",label:"Visitors Database",sub:`${visitors.length} visitor records`,action:"visitors",color:T.cyan},
        {icon:"🏸",label:"Come & Play Database",sub:`${cnp.length} members`,action:"cnp",color:T.purple},
        {icon:"🎫",label:"Gate Pass Requests",sub:`${passes.length} total · ${pendingPasses.length} pending`,action:"passes",color:T.cyan},
        {icon:"⛔",label:"Banned Vehicles",sub:`${banned.length} banned`,action:"banned",color:T.red},
        {icon:"👥",label:"User Management",sub:`${users.length} accounts`,action:"users",color:T.purple},
      ].map(({icon,label,sub,action,color})=>(
        <div key={action} onClick={()=>setView(action)} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${color}`,borderRadius:11,padding:"12px 14px",marginBottom:7,display:"flex",alignItems:"center",gap:11,cursor:"pointer",transition:"background .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.055)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.03)"}>
          <div style={{fontSize:16,opacity:.8}}>{icon}</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700}}>{label}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:1}}>{sub}</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
        </div>
      ))}
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;
}

/* ─── ROOT ──────────────────────────────────────────────────────── */
export default function App(){
  useEffect(()=>{
    const s=document.createElement("style");
    s.textContent=`
      @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800;900&display=swap');
      *{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}
      body{background:#080f1e;color:white;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:0;}
      input,select,textarea{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:9px;color:white;padding:11px 13px;font-size:14px;width:100%;outline:none;transition:border-color .2s,background .2s;font-family:inherit;}
      input::placeholder{color:rgba(255,255,255,0.25);}
      input:focus,select:focus{border-color:rgba(232,160,32,0.5);background:rgba(255,255,255,0.07);}
      select option{background:#0d1a30;color:white;}
      input[type="date"]{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:9px;color:white;padding:11px 13px;width:100%;outline:none;font-size:13px;}
      ::-webkit-scrollbar{width:4px;}::-webkit-scrollbar-track{background:transparent;}::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.1);border-radius:4px;}
    `;
    document.head.appendChild(s);return()=>document.head.removeChild(s);
  },[]);

  const [screen,setScreen]=useState(()=>{try{const s=localStorage.getItem("jln_session");return s?"app":"home";}catch{return "home";}});
  const [session,setSession]=useState(()=>{try{const s=localStorage.getItem("jln_session");return s?JSON.parse(s):null;}catch{return null;}});
  const [master,setMaster]=useState([]);
  const [officials,setOfficials]=useState([]);
  const [visitors,setVisitors]=useState([]);
  const [cnp,setCnp]=useState([]);
  const [log,setLog]=useState([]);
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [banned,setBanned]=useState([]);
  const [passes,setPasses]=useState([]);

  useEffect(()=>{
    gs(K.master).then(d=>d&&setMaster(d));
    gs(K.officials).then(d=>d&&setOfficials(d));
    gs(K.visitors).then(d=>d&&setVisitors(d));
    gs(K.cnp).then(d=>d&&setCnp(d));
    gs(K.log).then(d=>d&&setLog(d));
    gs(K.banned).then(d=>d&&setBanned(d));
    gs(K.passes).then(d=>d&&setPasses(d));
    gs(K.users).then(d=>d&&d.length&&setUsers(d));
    const interval=setInterval(()=>{
      gs(K.passes).then(d=>d&&setPasses(d));
      gs(K.master).then(d=>d&&setMaster(d));
      gs(K.officials).then(d=>d&&setOfficials(d));
      gs(K.visitors).then(d=>d&&setVisitors(d));
      gs(K.cnp).then(d=>d&&setCnp(d));
      gs(K.log).then(d=>d&&setLog(d));
    },10000);
    return()=>clearInterval(interval);
  },[]);

  const doLogin=u=>{setSession(u);setScreen("app");try{localStorage.setItem("jln_session",JSON.stringify(u));}catch{}};
  const doLogout=()=>{setSession(null);setScreen("home");try{localStorage.removeItem("jln_session");}catch{}};
  const shared={master,setMaster,officials,setOfficials,visitors,setVisitors,cnp,setCnp,log,setLog,users,setUsers,banned,setBanned,passes,setPasses};

  return <>
    {screen==="home"&&<HomeScreen onSelectRole={r=>setScreen(r==="admin"?"loginAdmin":"loginStaff")} onRequestPass={()=>setScreen("requestPass")}/>}
    {screen==="loginAdmin"&&<LoginScreen role="admin" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="loginStaff"&&<LoginScreen role="staff" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="requestPass"&&<RequestPassScreen onBack={()=>setScreen("home")} passes={passes} setPasses={setPasses}/>}
    {screen==="app"&&session?.role==="staff"&&<StaffApp onLogout={doLogout} {...shared} user={session}/>}
    {screen==="app"&&session?.role==="admin"&&<AdminApp onLogout={doLogout} {...shared} user={session}/>}
  </>;
}
