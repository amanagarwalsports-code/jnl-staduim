import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

/* ─── CONSTANTS ─────────────────────────────────────────────── */
const K = {
  master:"jln_master_v6",
  officials:"jln_officials_v1",
  visitors:"jln_visitors_v1",
  cnp:"jln_cnp_v2",
  users:"jln_users_v5",
  banned:"jln_banned_v5",
  passes:"jln_passes_v5",
  log:"jln_log_v6"
};
const DEFAULT_USERS = [
  { id:"admin_default", username:"admin", password:"admin123", role:"admin", createdOn:"System" },
  { id:"staff_default", username:"staff", password:"staff123", role:"staff", createdOn:"System" }
];
const DEPARTMENTS = ["Administration","Protocol & VIP","Security","Operations","Media & Press","Events Management","Technical/IT","Finance","Sports Management","General"];
const fmt = d => { try{ return new Date(d).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",hour12:true}); }catch{ return ""; } };
const fmtDate = d => { try{ return new Date(d).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}); }catch{ return ""; } };
const todayStr = () => new Date().toDateString();
const norm = p => { try{ return (p||"").replace(/\s+/g,"").toUpperCase(); }catch{ return ""; } };
const uid = () => Math.random().toString(36).slice(2,9);

async function gs(k){
  try{ const r=await window.storage.get(k); return r?JSON.parse(r.value):null; }
  catch{ return null; }
}
async function ss(k,v){
  try{ await window.storage.set(k,JSON.stringify(v)); }
  catch(e){ console.warn("Storage save failed:",k,e); }
}

/* ─── DESIGN TOKENS ──────────────────────────────────────────── */
const T = {
  navy:"#080f1e", mid:"#0d1a30", light:"#162842",
  gold:"#e8a020", goldL:"#fbbf24", goldD:"#92560a",
  green:"#16a34a", greenT:"#86efac",
  red:"#dc2626", redT:"#fca5a5",
  amber:"#d97706", amberT:"#fcd34d",
  blue:"#2563eb", blueT:"#93c5fd",
  purple:"#7c3aed", purpleT:"#c4b5fd",
  cyan:"#0891b2", cyanT:"#67e8f9",
  border:"rgba(255,255,255,0.08)", muted:"rgba(255,255,255,0.4)"
};

/* ─── ATOMS ──────────────────────────────────────────────────── */
function Pill({color,children}){
  const m={
    green:{bg:"rgba(22,163,74,0.18)",c:T.greenT,b:"rgba(22,163,74,0.35)"},
    red:{bg:"rgba(220,38,38,0.18)",c:T.redT,b:"rgba(220,38,38,0.35)"},
    amber:{bg:"rgba(217,119,6,0.18)",c:T.amberT,b:"rgba(217,119,6,0.35)"},
    blue:{bg:"rgba(37,99,235,0.18)",c:T.blueT,b:"rgba(37,99,235,0.35)"},
    gold:{bg:"rgba(232,160,32,0.15)",c:T.goldL,b:"rgba(232,160,32,0.35)"},
    purple:{bg:"rgba(124,58,237,0.18)",c:T.purpleT,b:"rgba(124,58,237,0.35)"},
    cyan:{bg:"rgba(8,145,178,0.18)",c:T.cyanT,b:"rgba(8,145,178,0.35)"}
  };
  const s=m[color]||m.blue;
  return <span style={{background:s.bg,color:s.c,border:`1px solid ${s.b}`,padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,whiteSpace:"nowrap",textTransform:"uppercase"}}>{children}</span>;
}
const Card=({children,style})=><div style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:14,padding:"16px 15px",marginBottom:12,...style}}>{children}</div>;
const SecTitle=({children})=><div style={{fontSize:9,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",color:"rgba(255,255,255,0.28)",margin:"0 0 10px",display:"flex",alignItems:"center",gap:8}}><span style={{flex:1,height:"1px",background:"rgba(255,255,255,0.06)"}}></span><span>{children}</span><span style={{flex:1,height:"1px",background:"rgba(255,255,255,0.06)"}}></span></div>;
const BigBtn=({onClick,color,children,disabled,style})=>{
  const cfg={
    green:{bg:T.green,shadow:"rgba(22,163,74,0.3)"},
    red:{bg:T.red,shadow:"rgba(220,38,38,0.3)"},
    gold:{bg:"linear-gradient(135deg,#e8a020,#f5c842)",shadow:"rgba(232,160,32,0.35)"},
    navy:{bg:T.light,shadow:"transparent"},
    amber:{bg:T.amber,shadow:"rgba(217,119,6,0.3)"},
    purple:{bg:T.purple,shadow:"rgba(124,58,237,0.3)"},
    blue:{bg:T.blue,shadow:"rgba(37,99,235,0.3)"},
    cyan:{bg:T.cyan,shadow:"rgba(8,145,178,0.3)"}
  }[color]||{bg:T.light,shadow:"transparent"};
  return <button onClick={onClick} disabled={disabled} style={{width:"100%",padding:"13px 0",borderRadius:10,border:"none",background:disabled?"rgba(255,255,255,0.06)":cfg.bg,color:disabled?"rgba(255,255,255,0.22)":(color==="gold"?T.navy:"white"),fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,letterSpacing:1.2,cursor:disabled?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8,boxShadow:disabled?"none":`0 4px 16px ${cfg.shadow}`,transition:"opacity .15s",...style}}>{children}</button>;
};
function Toast({msg,type}){
  const cfg={
    green:{bg:"rgba(22,163,74,0.95)",b:"rgba(22,163,74,0.5)"},
    red:{bg:"rgba(220,38,38,0.95)",b:"rgba(220,38,38,0.5)"},
    amber:{bg:"rgba(217,119,6,0.95)",b:"rgba(217,119,6,0.5)"},
    purple:{bg:"rgba(124,58,237,0.95)",b:"rgba(124,58,237,0.5)"},
    blue:{bg:"rgba(37,99,235,0.95)",b:"rgba(37,99,235,0.5)"},
    cyan:{bg:"rgba(8,145,178,0.95)",b:"rgba(8,145,178,0.5)"}
  }[type]||{bg:"rgba(13,26,48,0.97)",b:T.border};
  if(!msg) return null;
  return <div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:cfg.bg,border:`1px solid ${cfg.b}`,color:"white",padding:"10px 20px",borderRadius:40,fontSize:12,fontWeight:600,zIndex:9999,whiteSpace:"nowrap",pointerEvents:"none",boxShadow:"0 4px 20px rgba(0,0,0,0.4)"}}>{msg}</div>;
}
function Hdr({role,onLogout,title,sub,onBack}){
  return <div style={{background:`linear-gradient(135deg,${T.mid},${T.light})`,padding:"13px 15px 11px",borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:50}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
      <div style={{display:"flex",alignItems:"center",gap:9}}>
        {onBack&&<button onClick={onBack} style={{background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.6)",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:600,flexShrink:0}}>← Back</button>}
        <div>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,letterSpacing:1,color:T.goldL,lineHeight:1.1}}>{title}</div>
          {sub&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>{sub}</div>}
        </div>
      </div>
      <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}>
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
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200,padding:"0 20px",backdropFilter:"blur(4px)"}}>
    <div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:16,padding:"22px 20px",width:"100%",maxWidth:400}}>
      <div style={{fontSize:13,lineHeight:1.65,marginBottom:18,color:"rgba(255,255,255,0.8)"}}>{msg}</div>
      <div style={{display:"flex",gap:9}}>
        <button onClick={onNo} style={{flex:1,padding:"11px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,fontSize:13,cursor:"pointer"}}>Cancel</button>
        <button onClick={onYes} style={{flex:1,padding:"11px 0",borderRadius:9,border:"none",background:T.red,color:"white",fontSize:13,fontWeight:700,cursor:"pointer"}}>Confirm</button>
      </div>
    </div>
  </div>;
}

/* ─── SHARED DB MANAGER ──────────────────────────────────────── */
function DbManager({title,color,storageKey,fields,onBack,onLogout,t_,role,records,setRecords}){
  const [tab,setTab]=useState("list");
  const [search,setSearch]=useState("");
  const [form,setForm]=useState({});
  const [confirm,setConfirm]=useState(null);
  const [sel,setSel]=useState(null);
  const importRef=useRef();

  const saveRecord=async()=>{
    const required=fields.filter(f=>f.required);
    for(const f of required){
      if(!(form[f.key]||"").trim()){t_(`${f.label} is required`,"red");return;}
    }
    const rec={id:uid(),addedOn:fmtDate(Date.now()),...Object.fromEntries(fields.map(f=>[f.key,(form[f.key]||"").trim()]))};
    const u=[...records,rec];
    setRecords(u);
    await ss(storageKey,u);
    setForm({});
    t_("Record added ✓","green");
    setTab("list");
  };

  const removeRecord=async(id)=>{
    const u=records.filter(r=>r.id!==id);
    setRecords(u);
    await ss(storageKey,u);
    t_("Record removed","amber");
    setConfirm(null);setSel(null);setTab("list");
  };

  const exportXL=()=>{
    if(!records.length){t_("No data to export","amber");return;}
    const ws=XLSX.utils.json_to_sheet(records.map(r=>{
      const row={};
      fields.forEach(f=>row[f.label]=r[f.key]||"");
      row["Added On"]=r.addedOn||"";
      return row;
    }));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,title.slice(0,31));
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
        const getCol=(row,aliases)=>{
          for(const a of aliases){
            const found=Object.keys(row).find(k=>k.trim().toLowerCase()===a.toLowerCase());
            if(found&&row[found]!=="")return row[found].toString().trim();
          }
          return "";
        };
        let added=0;
        const cur=[...records];
        rows.forEach(row=>{
          const rec={id:uid(),addedOn:fmtDate(Date.now())};
          fields.forEach(f=>rec[f.key]=getCol(row,[f.label,...(f.aliases||[])]));
          const keyField=fields.find(f=>f.isKey);
          if(keyField&&!rec[keyField.key])return;
          if(keyField&&cur.find(r=>norm(r[keyField.key])===norm(rec[keyField.key])))return;
          cur.push(rec);added++;
        });
        setRecords(cur);
        ss(storageKey,cur);
        t_(`${added} record${added!==1?"s":""} imported ✓`,"green");
        setTab("list");
      }catch(e){
        console.error("Import error:",e);
        t_("Error reading file — check format","red");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const filtered=search.trim()
    ?records.filter(r=>fields.some(f=>(r[f.key]||"").toLowerCase().includes(search.toLowerCase())))
    :records;
  const accentColor=T[color]||T.blue;

  if(tab==="detail"&&sel) return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="RECORD DETAILS" onBack={()=>{setTab("list");setSel(null);}}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        {fields.map(f=><div key={f.key} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${T.border}`,alignItems:"flex-start",gap:10}}>
          <span style={{fontSize:10,color:T.muted,fontWeight:600,letterSpacing:.5,textTransform:"uppercase",flexShrink:0}}>{f.label}</span>
          <span style={{fontSize:12,fontWeight:600,textAlign:"right",maxWidth:"65%",wordBreak:"break-word"}}>{sel[f.key]||"—"}</span>
        </div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"9px 0",alignItems:"center"}}>
          <span style={{fontSize:10,color:T.muted,fontWeight:600,textTransform:"uppercase"}}>Added On</span>
          <span style={{fontSize:12,fontWeight:600}}>{sel.addedOn||"—"}</span>
        </div>
      </Card>
      <button onClick={()=>setConfirm(sel.id)} style={{width:"100%",padding:"11px 0",borderRadius:9,border:"1px solid rgba(220,38,38,0.3)",background:"rgba(220,38,38,0.08)",color:"#fca5a5",fontSize:13,cursor:"pointer",fontWeight:600,marginBottom:20}}>🗑 Remove This Record</button>
    </div>
    {confirm&&<ConfirmModal msg="Remove this record? This cannot be undone." onYes={()=>removeRecord(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;

  if(tab==="add") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title={`ADD — ${title.toUpperCase()}`} onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        {fields.filter(f=>f.type!=="readonly").map(f=><FormField key={f.key} label={f.label} required={f.required}>
          {f.type==="select"
            ?<select value={form[f.key]||""} onChange={e=>setForm(p=>({...p,[f.key]:e.target.value}))} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%"}}>
              <option value="">-- Select --</option>
              {(f.options||[]).map(o=><option key={o} value={o}>{o}</option>)}
            </select>
            :<input
              placeholder={f.placeholder||""}
              value={form[f.key]||""}
              onChange={e=>setForm(p=>({...p,[f.key]:e.target.value}))}
              style={f.big?{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center"}:{}}
              type={f.inputType||"text"}
              inputMode={f.inputMode||"text"}
              maxLength={f.maxLength}
            />
          }
        </FormField>)}
        <BigBtn onClick={saveRecord} color={color} style={{marginTop:4}}>+ ADD RECORD</BigBtn>
      </Card>
    </div>
  </div>;

  if(tab==="import") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="IMPORT FROM EXCEL" onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(37,99,235,0.07)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"12px 14px",marginBottom:14,fontSize:11,color:T.blueT,lineHeight:1.9}}>
        <b>Expected Excel columns:</b><br/>
        {fields.filter(f=>f.type!=="readonly").map(f=><span key={f.key} style={{fontFamily:"monospace",color:T.goldL,marginRight:8}}>{f.label}</span>)}
      </div>
      <div onClick={()=>importRef.current?.click()} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:10,padding:"40px 32px",border:"2px dashed rgba(37,99,235,0.4)",borderRadius:12,cursor:"pointer",background:"rgba(37,99,235,0.04)"}}>
        <span style={{fontSize:40}}>📂</span>
        <div style={{fontSize:14,fontWeight:600,color:"rgba(255,255,255,0.85)"}}>Tap to Select Excel File</div>
        <div style={{fontSize:11,color:T.muted}}>Supports .xlsx and .xls</div>
      </div>
      <input ref={importRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importXL}/>
    </div>
  </div>;

  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title={title.toUpperCase()} sub={`${records.length} records`} onBack={onBack}/>
    <div style={{display:"flex",borderBottom:`1px solid ${T.border}`,background:T.mid,position:"sticky",top:58,zIndex:40}}>
      {[["list","📋 View"],["add","➕ Add"],["import","📥 Import"]].map(([v,l])=>(
        <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"10px 0",border:"none",borderBottom:`2px solid ${tab===v?accentColor:"transparent"}`,background:"transparent",color:tab===v?accentColor:T.muted,fontSize:11,fontWeight:tab===v?700:400,cursor:"pointer"}}>{l}</button>
      ))}
    </div>
    <div style={{padding:"14px 15px 0"}}>
      <input
        placeholder={`🔍 Search ${title.toLowerCase()}…`}
        value={search}
        onChange={e=>setSearch(e.target.value)}
        style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?accentColor:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:10,transition:"border-color .2s"}}
      />
      <button onClick={exportXL} style={{width:"100%",padding:"9px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.65)",fontSize:12,cursor:"pointer",fontWeight:600,marginBottom:12}}>⬇ Download Excel</button>
      {search&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""}</div>}
      {!filtered.length
        ?<div style={{textAlign:"center",color:T.muted,padding:"40px 0",fontSize:12}}>{search?"No records match":"No records yet — tap ➕ Add to start"}</div>
        :filtered.slice().reverse().map(r=>(
          <div key={r.id} onClick={()=>{setSel(r);setTab("detail");}} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${accentColor}`,borderRadius:11,padding:"11px 13px",marginBottom:7,cursor:"pointer",transition:"background .15s"}}
            onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.055)"}
            onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.03)"}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:800,letterSpacing:1}}>{r[fields[0].key]||"—"}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:2}}>{fields.slice(1,3).map(f=>r[f.key]).filter(Boolean).join(" · ")}</div>
          </div>
        ))
      }
    </div>
    {confirm&&<ConfirmModal msg="Remove this record?" onYes={()=>removeRecord(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── MASTER DATABASE VIEW ───────────────────────────────────── */
function MasterDbV({onBack,onLogout,master,setMaster,t_,role}){
  const [search,setSearch]=useState("");
  const [tab,setTab]=useState("list");
  const [confirm,setConfirm]=useState(null);
  const [form,setForm]=useState({plate:"",name:"",division:"",dept:"",contact:""});
  const importRef=useRef();

  const exportXL=()=>{
    if(!master.length){t_("No data","amber");return;}
    const ws=XLSX.utils.json_to_sheet(master.map(r=>({
      "Car Number":r.plate||"",
      "Division":r.division||"",
      "Name":r.name||r.officer||"",
      "Department":r.dept||"",
      "Contact Number":r.contact||"",
      "First Seen":r.firstSeen||"",
      "Last Seen":r.lastSeen||""
    })));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Master");
    XLSX.writeFile(wb,"JLN_Master_Database.xlsx");
    t_("Downloaded ✓","green");
  };

  const addRecord=async()=>{
    if(!form.plate.trim()){t_("Car number is required","red");return;}
    const p=norm(form.plate);
    if((master||[]).find(m=>norm(m.plate||"")===p)){t_("Car number already exists in master","amber");return;}
    const rec={plate:p,name:form.name.trim(),division:form.division.trim(),dept:form.dept.trim(),contact:form.contact.trim(),firstSeen:fmtDate(Date.now()),lastSeen:fmtDate(Date.now())};
    const nm=[...(master||[]),rec];setMaster(nm);await ss(K.master,nm);
    setForm({plate:"",name:"",division:"",dept:"",contact:""});
    t_("Record added ✓","green");setTab("list");
  };

  const removeRecord=async(plate)=>{
    const nm=(master||[]).filter(r=>norm(r.plate||"")!==norm(plate||""));
    setMaster(nm);await ss(K.master,nm);
    t_("Record removed","amber");setConfirm(null);
  };

  const importXL=ev=>{
    const file=ev.target.files[0];if(!file)return;
    if(importRef.current)importRef.current.value="";
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
        const getCol=(row,keys)=>{for(const k of keys){const f=Object.keys(row).find(c=>c.trim().toLowerCase()===k.toLowerCase());if(f&&row[f]!=="")return row[f].toString().trim();}return "";};
        let added=0;const cur=[...(master||[])];
        rows.forEach(row=>{
          const plate=getCol(row,["car number","plate number","plate","vehicle number","reg number","registration"]);
          if(!plate)return;
          const p=norm(plate);
          if(cur.find(r=>norm(r.plate||"")===p))return;
          cur.push({
            plate:p,
            name:getCol(row,["name","officer name","official name","visitor name"]),
            division:getCol(row,["division"]),
            dept:getCol(row,["department","dept"]),
            contact:getCol(row,["contact","phone","contact number","phone number"]),
            firstSeen:getCol(row,["first seen"])||fmtDate(Date.now()),
            lastSeen:getCol(row,["last seen"])||fmtDate(Date.now()),
          });
          added++;
        });
        setMaster(cur);ss(K.master,cur);
        t_(`${added} record${added!==1?"s":""} imported ✓`,"green");
        setTab("list");
      }catch(e){console.error(e);t_("Error reading file","red");}
    };
    reader.readAsArrayBuffer(file);
  };

  const safeSearch=search.trim().toLowerCase();
  const filtered=safeSearch
    ?(master||[]).filter(r=>norm(r.plate||"").includes(norm(search))||(r.name||"").toLowerCase().includes(safeSearch)||(r.dept||"").toLowerCase().includes(safeSearch)||(r.type||"").toLowerCase().includes(safeSearch))
    :(master||[]);
  const typeColor={official:"gold",visitor:"cyan","come&play":"purple",cnp:"purple",exception:"amber",guest:"amber",master:"green"};

  if(tab==="add") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="ADD TO MASTER DB" onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <FormField label="Car Number" required>
          <input placeholder="DL01AB1234" value={form.plate} onChange={e=>setForm(p=>({...p,plate:e.target.value.toUpperCase()}))} style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center"}}/>
        </FormField>
        <FormField label="Division"><input placeholder="e.g. Security, Admin" value={form.division} onChange={e=>setForm(p=>({...p,division:e.target.value}))}/></FormField>
        <FormField label="Name"><input placeholder="Full name" value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))}/></FormField>
        <FormField label="Department"><input placeholder="e.g. Administration, Sports" value={form.dept} onChange={e=>setForm(p=>({...p,dept:e.target.value}))}/></FormField>
        <FormField label="Contact Number"><input placeholder="Phone number" value={form.contact} onChange={e=>setForm(p=>({...p,contact:e.target.value}))} type="tel" inputMode="numeric" maxLength={15}/></FormField>
        <BigBtn onClick={addRecord} color="gold" style={{marginTop:4}}>+ ADD TO MASTER DATABASE</BigBtn>
      </Card>
    </div>
  </div>;

  if(tab==="import") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="IMPORT MASTER DB" onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(37,99,235,0.07)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"12px 14px",marginBottom:14,fontSize:11,color:T.blueT,lineHeight:1.9}}>
        <b>Expected Excel columns:</b><br/>
        <span style={{fontFamily:"monospace",color:T.goldL}}>Car Number</span> · <span style={{fontFamily:"monospace",color:T.goldL}}>Division</span> · <span style={{fontFamily:"monospace",color:T.goldL}}>Name</span> · <span style={{fontFamily:"monospace",color:T.goldL}}>Department</span> · <span style={{fontFamily:"monospace",color:T.goldL}}>Contact</span><br/>
        <span style={{color:"rgba(255,255,255,0.4)"}}>Duplicate car numbers are automatically skipped.</span>
      </div>
      <div onClick={()=>importRef.current?.click()} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:10,padding:"40px 32px",border:"2px dashed rgba(37,99,235,0.4)",borderRadius:12,cursor:"pointer",background:"rgba(37,99,235,0.04)"}}>
        <span style={{fontSize:40}}>📂</span>
        <div style={{fontSize:14,fontWeight:600,color:"rgba(255,255,255,0.85)"}}>Tap to Select Excel File</div>
        <div style={{fontSize:11,color:T.muted}}>Supports .xlsx and .xls</div>
      </div>
      <input ref={importRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importXL}/>
    </div>
  </div>;

  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role={role} onLogout={onLogout} title="MASTER DATABASE" sub={`${(master||[]).length} unique vehicles`} onBack={onBack}/>
    <div style={{display:"flex",borderBottom:`1px solid ${T.border}`,background:T.mid,position:"sticky",top:58,zIndex:40}}>
      {[["list","📋 View"],["add","➕ Add"],["import","📥 Import"]].map(([v,l])=>(
        <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"10px 0",border:"none",borderBottom:`2px solid ${tab===v?T.gold:"transparent"}`,background:"transparent",color:tab===v?T.gold:T.muted,fontSize:11,fontWeight:tab===v?700:400,cursor:"pointer"}}>{l}</button>
      ))}
    </div>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(232,160,32,0.07)",border:"1px solid rgba(232,160,32,0.2)",borderRadius:9,padding:"10px 14px",marginBottom:12,fontSize:11,color:T.amberT,lineHeight:1.6}}>
        📋 Auto-updated on every vehicle entry. Admin can also add, import or remove vehicles manually.
      </div>
      <input placeholder="🔍 Search plate, name, department, type…" value={search} onChange={e=>setSearch(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?T.gold:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:10,transition:"border-color .2s"}}/>
      <button onClick={exportXL} style={{width:"100%",padding:"9px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.65)",fontSize:12,cursor:"pointer",fontWeight:600,marginBottom:12}}>⬇ Download Master Excel</button>
      {search&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""}</div>}
      {!filtered.length
        ?<div style={{textAlign:"center",color:T.muted,padding:"40px 0",fontSize:12}}>{search?"No vehicles match":"No vehicles yet — tap ➕ Add or 📥 Import"}</div>
        :filtered.slice().reverse().map((r,i)=>(
          <div key={r.plate||i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${T.gold}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
            <div style={{flex:1}}>
              <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:2.5,marginBottom:3}}>{r.plate}</div>
              <div style={{fontSize:10,color:T.muted}}>{r.name||"—"}{r.division?` · ${r.division}`:""}{r.dept?` · ${r.dept}`:""}</div>
              {r.contact&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>📞 {r.contact}</div>}
              <div style={{fontSize:9,color:"rgba(255,255,255,0.2)",marginTop:3}}>First: {r.firstSeen||"—"} · Last: {r.lastSeen||"—"}</div>
            </div>
            <button onClick={()=>setConfirm(r.plate)} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:600,flexShrink:0}}>Remove</button>
          </div>
        ))
      }
    </div>
    {confirm&&<ConfirmModal msg={`Remove ${confirm} from master database?`} onYes={()=>removeRecord(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── HOME SCREEN ────────────────────────────────────────────── */
function HomeScreen({onSelectRole,onRequestPass}){
  return <div style={{minHeight:"100vh",background:T.navy}}>
    <div style={{background:`linear-gradient(160deg,${T.mid},${T.light})`,padding:"36px 20px 28px"}}>
      <div style={{display:"flex",alignItems:"center",gap:13}}>
        <div style={{width:46,height:46,borderRadius:13,background:"rgba(232,160,32,0.1)",border:"1px solid rgba(232,160,32,0.25)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>🏟</div>
        <div>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:900,letterSpacing:1.5,color:T.goldL,lineHeight:1}}>JLN STADIUM</div>
          <div style={{fontSize:9,color:"rgba(255,255,255,0.3)",letterSpacing:2.5,marginTop:4,textTransform:"uppercase"}}>Vehicle Access Control</div>
        </div>
      </div>
    </div>
    <div style={{padding:"22px 16px 24px"}}>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Security Personnel</div>
      {[
        {role:"admin",icon:"◆",label:"Super Admin",desc:"Dashboard · Databases · Reports",border:"rgba(232,160,32,0.3)",clr:T.goldL,bg:"rgba(232,160,32,0.05)"},
        {role:"staff",icon:"◈",label:"Security Staff",desc:"Entry / Exit · Verification · Exemption",border:"rgba(37,99,235,0.35)",clr:T.blueT,bg:"rgba(37,99,235,0.05)"}
      ].map(({role,icon,label,desc,border,clr,bg})=>(
        <div key={role} onClick={()=>onSelectRole(role)} style={{background:bg,border:`1px solid ${border}`,borderRadius:13,padding:"16px 15px",marginBottom:9,cursor:"pointer",display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:36,height:36,borderRadius:9,background:"rgba(255,255,255,0.06)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:clr,flexShrink:0}}>{icon}</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,color:clr}}>{label}</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.38)",marginTop:2}}>{desc}</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:16}}>›</div>
        </div>
      ))}
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10,marginTop:20}}>Visitor / Public</div>
      <div onClick={onRequestPass} style={{background:"rgba(8,145,178,0.05)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:13,padding:"16px 15px",cursor:"pointer",display:"flex",alignItems:"center",gap:12}}>
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
  const [pwd,setPwd]=useState("");
  const [err,setErr]=useState("");
  const roleUsers=users.filter(u=>u.role===role);
  useEffect(()=>{
    if(roleUsers.length===1) setSelUser(roleUsers[0].id);
  },[users,role]);
  const try_=()=>{
    if(!selUser){setErr("Please select an account");return;}
    const u=users.find(x=>x.id===selUser&&x.password===pwd);
    if(u) onLogin(u);
    else { setErr("Incorrect password");setPwd(""); }
  };
  const isAdmin=role==="admin";
  return <div style={{minHeight:"100vh",background:T.navy}}>
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
      {roleUsers.length===1&&<div style={{background:"rgba(255,255,255,0.04)",borderRadius:9,padding:"11px 13px",marginBottom:14,fontSize:13,color:"rgba(255,255,255,0.6)",border:`1px solid ${T.border}`}}>Logging in as <b style={{color:"white"}}>{roleUsers[0]?.username}</b></div>}
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
    if(!name.trim()||!car.trim()||!phone.trim()||!purpose.trim()){
      alert("Please fill all required fields (Name, Vehicle Number, Phone, Purpose).");return;
    }
    const id="GP-"+Math.random().toString(36).slice(2,8).toUpperCase();
    const entry={id,name:name.trim(),car:norm(car),phone:phone.trim(),purpose:purpose.trim(),dept:dept.trim(),refName:refName.trim(),status:"pending",requestedOn:Date.now(),date:todayStr()};
    const np=[...(passes||[]),entry];
    setPasses(np);ss(K.passes,np);
    setRefId(id);setSubmitted(true);
  };

  if(submitted) return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr title="REQUEST SUBMITTED" onBack={onBack}/>
    <div style={{padding:"32px 18px",textAlign:"center"}}>
      <div style={{width:64,height:64,borderRadius:20,background:"rgba(22,163,74,0.15)",border:"1px solid rgba(22,163,74,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,margin:"0 auto 18px"}}>✅</div>
      <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,color:T.greenT,marginBottom:8}}>Request Submitted!</div>
      <div style={{background:"rgba(255,255,255,0.04)",border:"1px dashed rgba(232,160,32,0.35)",borderRadius:12,padding:"18px",marginBottom:24}}>
        <div style={{fontSize:9,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",color:"rgba(255,255,255,0.3)",marginBottom:8}}>Your Reference ID</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:30,fontWeight:900,letterSpacing:5,color:T.goldL}}>{refId}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:8}}>Save this ID. The admin will review your request.</div>
      </div>
      <BigBtn onClick={onBack} color="navy">← Back to Home</BigBtn>
    </div>
  </div>;

  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr title="REQUEST GATE PASS" onBack={onBack}/>
    <div style={{padding:"14px 14px 0"}}>
      <Card>
        <SecTitle>VISITOR DETAILS</SecTitle>
        <FormField label="Full Name" required><input placeholder="e.g. Priya Sharma" value={name} onChange={e=>setName(e.target.value)}/></FormField>
        <FormField label="Vehicle Number" required><input placeholder="DL01AB1234" value={car} onChange={e=>setCar(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:700,letterSpacing:4,textAlign:"center"}}/></FormField>
        <FormField label="Phone Number" required><input placeholder="10-digit mobile" value={phone} onChange={e=>setPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
        <FormField label="Purpose of Visit" required><input placeholder="e.g. Meeting, Event, Delivery" value={purpose} onChange={e=>setPurpose(e.target.value)}/></FormField>
        <FormField label="Department to Visit">
          <select value={dept} onChange={e=>setDept(e.target.value)} style={{background:"rgba(255,255,255,0.07)",border:`1px solid ${T.border}`,borderRadius:8,color:dept?"white":"rgba(255,255,255,0.3)",padding:"10px 13px",fontSize:14,width:"100%"}}>
            <option value="">-- Select Department --</option>
            {DEPARTMENTS.map(d=><option key={d} value={d}>{d}</option>)}
          </select>
        </FormField>
        <FormField label="Reference Officer (Optional)"><input placeholder="Name of JLN officer who referred you" value={refName} onChange={e=>setRefName(e.target.value)}/></FormField>
      </Card>
      <BigBtn onClick={submit} color="cyan" style={{marginBottom:20}}>🎫 SUBMIT REQUEST</BigBtn>
    </div>
  </div>;
}

/* ─── GATE PASS MGMT ─────────────────────────────────────────── */
function GatePassMgmtV({onBack,passes,setPasses,master,setMaster,visitors,setVisitors,onLogout,t_}){
  const [filter,setFilter]=useState("pending");
  const filtered=(passes||[]).filter(p=>filter==="all"||p.status===filter).slice().reverse();

  const approve=id=>{
    const pass=(passes||[]).find(p=>p.id===id);
    if(!pass)return;
    const u=(passes||[]).map(p=>p.id===id?{...p,status:"approved",approvedOn:fmtDate(Date.now())}:p);
    setPasses(u);ss(K.passes,u);
    // Add to visitors DB
    const curVisitors=visitors||[];
    if(!curVisitors.find(v=>norm(v.carNumber)===norm(pass.car))){
      const nv=[...curVisitors,{id:uid(),visitorName:pass.name,carNumber:norm(pass.car),purpose:pass.purpose,officerReference:pass.refName||"",contactNumber:pass.phone,entryTime:"",exitTime:"",addedOn:fmtDate(Date.now()),passId:pass.id}];
      setVisitors(nv);ss(K.visitors,nv);
    }
    // Add to master
    const curMaster=master||[];
    if(!curMaster.find(e=>norm(e.plate)===norm(pass.car))){
      const nm=[...curMaster,{plate:norm(pass.car),type:"visitor",name:pass.name,dept:pass.dept||"",contact:pass.phone,firstSeen:fmtDate(Date.now()),lastSeen:fmtDate(Date.now())}];
      setMaster(nm);ss(K.master,nm);
    }
    t_("Pass approved ✓","green");
  };

  const reject=id=>{
    const u=(passes||[]).map(p=>p.id===id?{...p,status:"rejected"}:p);
    setPasses(u);ss(K.passes,u);t_("Pass rejected","amber");
  };

  const statusColor={pending:"amber",approved:"green",rejected:"red",used:"cyan"};
  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="GATE PASS REQUESTS" sub={`${(passes||[]).filter(p=>p.status==="pending").length} pending`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{display:"flex",gap:5,marginBottom:13,flexWrap:"wrap"}}>
        {["pending","approved","rejected","used","all"].map(f=>(
          <button key={f} onClick={()=>setFilter(f)} style={{flex:1,minWidth:50,padding:"7px 0",borderRadius:9,border:`1px solid ${filter===f?"rgba(232,160,32,0.5)":T.border}`,background:filter===f?"rgba(232,160,32,0.1)":"transparent",color:filter===f?T.goldL:T.muted,fontSize:10,cursor:"pointer",fontWeight:filter===f?700:400,textTransform:"capitalize"}}>{f}</button>
        ))}
      </div>
      {!filtered.length
        ?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>No {filter} requests</div>
        :filtered.map(p=>(
          <div key={p.id} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:12,padding:"13px 14px",marginBottom:9}}>
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
            {p.dept&&<div style={{fontSize:10,color:T.muted,marginBottom:2}}>🏛 {p.dept}</div>}
            {p.refName&&<div style={{fontSize:10,color:T.muted,marginBottom:6}}>👤 Ref: {p.refName}</div>}
            {p.status==="pending"&&(
              <div style={{display:"flex",gap:7,marginTop:8}}>
                <button onClick={()=>approve(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.green,color:"white",fontSize:12,fontWeight:700,cursor:"pointer"}}>✓ Approve</button>
                <button onClick={()=>reject(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.red,color:"white",fontSize:12,fontWeight:700,cursor:"pointer"}}>✕ Reject</button>
              </div>
            )}
          </div>
        ))
      }
    </div>
  </div>;
}

/* ─── STAFF APP ──────────────────────────────────────────────── */
function StaffApp({onLogout,master,setMaster,officials,visitors,setVisitors,cnp,log,setLog,banned,passes,setPasses,user}){
  const [view,setView]=useState("verify");
  const [plate,setPlate]=useState("");
  const [result,setResult]=useState(null);
  const [last4,setLast4]=useState("");
  const [suggestions,setSuggestions]=useState([]);
  const [guestName,setGuestName]=useState("");
  const [guestReason,setGuestReason]=useState("");
  const [toast,setToast]=useState(null);
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2600);};

  // Build vehicle list safely
  const allVehicles=useCallback(()=>{
    const safeOfficials=officials||[];
    const safeVisitors=visitors||[];
    const safeCnp=cnp||[];
    const safeMaster=master||[];
    const list=[
      ...safeOfficials.map(o=>({plate:norm(o.carNumber||""),label:o.officialName||"",sub:o.department||"",badge:"Official",badgeColor:"gold",type:"official"})),
      ...safeVisitors.map(v=>({plate:norm(v.carNumber||""),label:v.visitorName||"",sub:v.purpose||"",badge:"Visitor",badgeColor:"cyan",type:"visitor"})),
      ...safeCnp.map(c=>({plate:norm(c.carNumber||""),label:c.memberName||"",sub:"Come & Play",badge:"Come & Play",badgeColor:"purple",type:"cnp"})),
    ];
    const platesInList=new Set(list.map(v=>v.plate));
    safeMaster.forEach(m=>{
      const p=norm(m.plate||"");
      if(p&&!platesInList.has(p)){
        list.push({plate:p,label:m.name||m.officer||"",sub:m.dept||m.division||"",badge:"Master",badgeColor:"green",type:"master"});
        platesInList.add(p);
      }
    });
    return list;
  },[officials,visitors,cnp,master]);

  const onLast4Change=v=>{
    setLast4(v);setPlate("");setResult(null);setSuggestions([]);
    if(v.length>=1){
      const n=norm(v);
      const vehicles=allVehicles();
      const seen=new Set();
      const matches=vehicles.filter(e=>{
        if(!e.plate||!e.plate.endsWith(n))return false;
        if(seen.has(e.plate))return false;
        seen.add(e.plate);return true;
      });
      setSuggestions(matches);
    }
  };

  const selectSuggestion=p=>{setPlate(p);setLast4("");setSuggestions([]);setResult(null);};

  const doAction=action=>{
    const raw=(plate||"").trim();
    if(!raw){t_("Enter a vehicle number","red");return;}
    const normalizedRaw=norm(raw);
    const safeBanned=banned||[];
    const safeLog=log||[];
    const safePasses=passes||[];

    const isBanned=safeBanned.find(b=>norm(b.plate||"")===normalizedRaw);
    if(isBanned&&action==="entry"){
      setResult({status:"banned",plate:normalizedRaw,detail:`Reason: ${isBanned.reason||"Banned vehicle"}`});
      return;
    }

    const now=Date.now();
    const vehicles=allVehicles();
    const vehicle=vehicles.find(v=>v.plate===normalizedRaw);

    if(action==="entry"){
      const alreadyIn=safeLog.find(e=>norm(e.plate||"")===normalizedRaw&&!e.exitTime&&e.date===todayStr());
      if(alreadyIn){t_("Vehicle is already inside","amber");return;}

      const approvedPass=safePasses.find(p=>norm(p.car||"")===normalizedRaw&&p.status==="approved");
      const usedPass=safePasses.find(p=>norm(p.car||"")===normalizedRaw&&p.status==="used");

      // Block if gate pass was used and vehicle not in permanent db
      if(usedPass&&!approvedPass&&(!vehicle||vehicle.type==="visitor")){
        setResult({status:"denied",plate:normalizedRaw,detail:"Gate pass already used — single entry only"});
        return;
      }
      // Block visitors from re-entering (not cnp/official)
      if(vehicle&&vehicle.type==="visitor"){
        const prev=safeLog.find(e=>norm(e.plate||"")===normalizedRaw);
        if(prev){setResult({status:"denied",plate:normalizedRaw,detail:"Visitor single entry already used"});return;}
      }

      const entry={
        id:now,
        plate:normalizedRaw,
        type:vehicle?.type||"guest",
        entryTime:now,exitTime:null,
        date:todayStr(),
        name:vehicle?.label||"",
        division:vehicle?.sub||"",
        loggedBy:user?.username||"staff"
      };
      if(approvedPass){entry.passId=approvedPass.id;entry.name=approvedPass.name||entry.name;}

      const nl=[...safeLog,entry];setLog(nl);ss(K.log,nl);

      if(approvedPass){
        const np=safePasses.map(p=>p.id===approvedPass.id?{...p,status:"used",entryTime:now,entryLoggedBy:user?.username}:p);
        setPasses(np);ss(K.passes,np);
      }

      // Update master
      const safeMaster=master||[];
      if(safeMaster.find(m=>norm(m.plate||"")===normalizedRaw)){
        const nm=safeMaster.map(m=>norm(m.plate||"")===normalizedRaw?{...m,lastSeen:fmtDate(now)}:m);
        setMaster(nm);ss(K.master,nm);
      } else {
        const nm=[...safeMaster,{plate:normalizedRaw,type:vehicle?.type||"guest",name:vehicle?.label||"",dept:vehicle?.sub||"",firstSeen:fmtDate(now),lastSeen:fmtDate(now)}];
        setMaster(nm);ss(K.master,nm);
      }

      setResult({
        status:approvedPass?"allowed_guest":(vehicle?"allowed":"allowed_guest"),
        plate:normalizedRaw,
        detail:vehicle?`${vehicle.badge}: ${vehicle.label||"—"}`:approvedPass?`Gate Pass: ${approvedPass.id}`:"Entry logged"
      });
      t_("Entry logged ✓","green");

    } else {
      // EXIT
      const entry=[...safeLog].reverse().find(e=>norm(e.plate||"")===normalizedRaw&&!e.exitTime);
      if(!entry){t_("No active entry found for this vehicle","amber");return;}
      const nl=safeLog.map(e=>e.id===entry.id?{...e,exitTime:now}:e);
      setLog(nl);ss(K.log,nl);
      if(entry.passId){
        const np=safePasses.map(p=>p.id===entry.passId?{...p,exitTime:now,exitLoggedBy:user?.username}:p);
        setPasses(np);ss(K.passes,np);
      }
      // Update visitor exit time
      if(entry.type==="visitor"){
        const safeVisitors=visitors||[];
        const nv=safeVisitors.map(v=>norm(v.carNumber||"")===normalizedRaw?{...v,exitTime:fmt(now)}:v);
        setVisitors(nv);ss(K.visitors,nv);
      }
      setResult({status:"exit",plate:normalizedRaw,detail:`Exited at ${fmt(now)}`});
      t_("Exit logged ✓","amber");
    }
  };

  const addException=()=>{
    if(!(plate||"").trim()){t_("Enter vehicle number","red");return;}
    if(!guestName.trim()){t_("Enter visitor name","red");return;}
    if(!guestReason.trim()){t_("Enter purpose of visit","red");return;}
    const now=Date.now();
    const normalizedPlate=norm(plate);
    const safeLog=log||[];
    const safeVisitors=visitors||[];
    const safeMaster=master||[];

    const entry={id:now,plate:normalizedPlate,type:"exception",entryTime:now,exitTime:null,date:todayStr(),name:guestName.trim(),division:guestReason.trim(),loggedBy:user?.username||"staff"};
    const nl=[...safeLog,entry];setLog(nl);ss(K.log,nl);

    // Add to visitors
    const nv=[...safeVisitors,{id:uid(),visitorName:guestName.trim(),carNumber:normalizedPlate,purpose:guestReason.trim(),officerReference:"",contactNumber:"",entryTime:fmt(now),exitTime:"",addedOn:fmtDate(now)}];
    setVisitors(nv);ss(K.visitors,nv);

    // Add to master
    if(!safeMaster.find(m=>norm(m.plate||"")===normalizedPlate)){
      const nm=[...safeMaster,{plate:normalizedPlate,type:"exception",name:guestName.trim(),dept:guestReason.trim(),firstSeen:fmtDate(now),lastSeen:fmtDate(now)}];
      setMaster(nm);ss(K.master,nm);
    }

    setGuestName("");setGuestReason("");
    t_("Exception entry logged ✓","amber");
    setView("verify");
    setResult({status:"allowed_guest",plate:normalizedPlate,detail:`Exception: ${guestName.trim()} · ${guestReason.trim()}`});
  };

  const approvedPasses=(passes||[]).filter(p=>p.status==="approved");
  const resultColors={
    allowed:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},
    allowed_guest:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},
    exit:{bg:"rgba(217,119,6,0.12)",border:T.amber,tc:T.amberT},
    denied:{bg:"rgba(220,38,38,0.12)",border:T.red,tc:T.redT},
    banned:{bg:"rgba(127,29,29,0.3)",border:"#7f1d1d",tc:"#fca5a5"}
  };
  const rc=result?resultColors[result.status]||resultColors.denied:null;

  if(view==="exception") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="GRANT EXEMPTION" sub="One-time exception entry" onBack={()=>{setView("verify");setResult(null);}}/>
    <div style={{padding:"15px 15px 0"}}>
      <div style={{background:"rgba(217,119,6,0.08)",border:"1px solid rgba(217,119,6,0.3)",borderRadius:12,padding:"14px 15px",marginBottom:14}}>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,color:T.amberT}}>Exception / Exemption Entry</div>
        <div style={{fontSize:11,color:"rgba(253,211,77,0.6)",marginTop:4}}>Fill 3 required fields to grant one-time entry</div>
      </div>
      <Card>
        <FormField label="Vehicle Registration Number" required>
          <input placeholder="DL01AB1234" value={plate||""} onChange={e=>setPlate(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center"}}/>
        </FormField>
        <FormField label="Name of Visitor" required>
          <input placeholder="Full name" value={guestName} onChange={e=>setGuestName(e.target.value)}/>
        </FormField>
        <FormField label="Purpose of Visit" required>
          <input placeholder="e.g. Meeting, Delivery, Event" value={guestReason} onChange={e=>setGuestReason(e.target.value)}/>
        </FormField>
        <BigBtn onClick={addException} color="amber" style={{marginTop:6}}>⚠ ALLOW EXCEPTION ENTRY</BigBtn>
      </Card>
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;

  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="SECURITY STAFF" sub={user?.username||"staff"}/>
    <div style={{padding:"16px 15px 0"}}>
      <Card style={{marginBottom:14}}>
        <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.28)",marginBottom:8}}>Vehicle Registration</div>
        {plate
          ?<div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:"rgba(232,160,32,0.08)",border:`1px solid ${T.gold}`,borderRadius:10,padding:"12px 14px",marginBottom:10,cursor:"pointer"}} onClick={()=>{setPlate("");setLast4("");setSuggestions([]);setResult(null);}}>
            <span style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,letterSpacing:5,color:T.goldL}}>{plate}</span>
            <span style={{fontSize:11,color:T.muted,background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,borderRadius:6,padding:"3px 8px"}}>✕ Clear</span>
          </div>
          :<div style={{position:"relative",marginBottom:suggestions.length?0:10}}>
            <input
              placeholder="Type last 4 digits  e.g. 1234"
              value={last4}
              onChange={e=>onLast4Change(e.target.value.toUpperCase())}
              maxLength={6}
              style={{fontFamily:"'Barlow Condensed',monospace",fontSize:22,fontWeight:800,letterSpacing:4,textAlign:"center",background:"rgba(255,255,255,0.05)",border:`1px solid ${last4?T.gold:T.border}`,borderRadius:10,color:"white",padding:"13px",width:"100%",outline:"none",transition:"border-color .2s",marginBottom:0}}
            />
            {suggestions.length>0&&<div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:10,marginTop:4,overflow:"hidden",marginBottom:10}}>
              {suggestions.map((s,i)=>(
                <div key={i} onClick={()=>selectSuggestion(s.plate)} style={{padding:"11px 14px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:`1px solid ${T.border}`,transition:"background .12s"}}
                  onMouseEnter={e=>e.currentTarget.style.background="rgba(232,160,32,0.08)"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <div>
                    <span style={{fontFamily:"'Barlow Condensed'",fontSize:17,fontWeight:800,letterSpacing:3,color:T.goldL}}>{s.plate}</span>
                    <div style={{fontSize:10,color:T.muted,marginTop:2}}>{s.label}{s.sub?` · ${s.sub}`:""}</div>
                  </div>
                  <Pill color={s.badgeColor||"green"}>{s.badge||"DB"}</Pill>
                </div>
              ))}
              <div style={{padding:"9px 14px",fontSize:10,color:T.muted,textAlign:"center"}}>Tap to select</div>
            </div>}
            {last4&&!suggestions.length&&<div style={{padding:"12px 14px",fontSize:12,color:T.muted,textAlign:"center",marginBottom:10}}>No match found — <span style={{color:T.amberT,cursor:"pointer"}} onClick={()=>{setPlate(norm(last4));setSuggestions([]);setLast4("");}}>use "{last4}" directly</span></div>}
          </div>
        }
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <BigBtn onClick={()=>doAction("entry")} color="green" disabled={!plate}>▲ ENTRY</BigBtn>
          <BigBtn onClick={()=>doAction("exit")} color="red" disabled={!plate}>▼ EXIT</BigBtn>
        </div>
      </Card>

      {result&&rc&&<div style={{borderRadius:14,padding:"20px 16px",textAlign:"center",marginBottom:14,background:rc.bg,border:`1px solid ${rc.border}`}}>
        <div style={{fontSize:32,marginBottom:6}}>
          {result.status==="allowed"||result.status==="allowed_guest"?"✅":result.status==="exit"?"🔄":result.status==="banned"?"⛔":"🚫"}
        </div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:800,letterSpacing:1.5,color:rc.tc}}>
          {result.status==="allowed"||result.status==="allowed_guest"?"ENTRY ALLOWED":result.status==="exit"?"EXIT RECORDED":result.status==="banned"?"BANNED VEHICLE":"NOT ALLOWED"}
        </div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:700,letterSpacing:5,marginTop:4,opacity:.6,color:rc.tc}}>{result.plate}</div>
        {result.detail&&<div style={{fontSize:11,marginTop:5,opacity:.7,color:rc.tc}}>{result.detail}</div>}
        {result.status==="denied"&&<button onClick={()=>{setView("exception");}} style={{marginTop:12,background:"rgba(217,119,6,0.15)",border:"1px solid rgba(217,119,6,0.4)",color:T.amberT,borderRadius:8,padding:"8px 20px",fontSize:12,cursor:"pointer",fontWeight:700}}>⚠ Grant Exemption</button>}
      </div>}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
        <div onClick={()=>setView("exception")} style={{background:"rgba(217,119,6,0.06)",border:"1px solid rgba(217,119,6,0.25)",borderRadius:12,padding:"12px 14px",cursor:"pointer",display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"rgba(217,119,6,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>⚠</div>
          <div>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:13,fontWeight:700,color:T.amberT}}>Grant Exemption</div>
            <div style={{fontSize:9,color:T.muted,marginTop:1}}>One-time entry</div>
          </div>
        </div>
        <div style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.25)",borderRadius:12,padding:"12px 14px",display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:8,background:"rgba(8,145,178,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🎫</div>
          <div>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:13,fontWeight:700,color:T.cyanT}}>Gate Passes</div>
            <div style={{fontSize:9,color:T.muted,marginTop:1}}>{approvedPasses.length} approved</div>
          </div>
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
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2600);};

  const OFFICIAL_FIELDS=[
    {key:"officialName",label:"Name of Official",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate","car number"]},
    {key:"department",label:"Department",required:true,type:"select",options:DEPARTMENTS},
  ];
  const VISITOR_FIELDS=[
    {key:"visitorName",label:"Name of Visitor",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate","car number"]},
    {key:"purpose",label:"Purpose",required:true},
    {key:"officerReference",label:"Officer Reference"},
    {key:"contactNumber",label:"Contact Number",inputType:"tel",inputMode:"numeric",maxLength:15},
    {key:"entryTime",label:"Entry Time"},
    {key:"exitTime",label:"Exit Time"},
  ];
  const CNP_FIELDS=[
    {key:"memberName",label:"Name of Member",required:true},
    {key:"carNumber",label:"Car Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","plate","car number"]},
    {key:"cardNumber",label:"SAI Card Number",required:true},
    {key:"contactNumber",label:"Contact Number",required:true,inputType:"tel",inputMode:"numeric",maxLength:15},
  ];
  const BANNED_FIELDS=[
    {key:"plate",label:"Plate Number",required:true,isKey:true,big:true,aliases:["car no","vehicle number","reg number","car number","plate number"]},
    {key:"reason",label:"Reason"},
  ];

  const downloadAllExcel=()=>{
    try{
      const wb=XLSX.utils.book_new();
      if((master||[]).length){
        const ws=XLSX.utils.json_to_sheet((master||[]).map(r=>({"Plate":r.plate||"","Type":r.type||"","Name":r.name||"","Department":r.dept||"","Contact":r.contact||"","First Seen":r.firstSeen||"","Last Seen":r.lastSeen||""})));
        XLSX.utils.book_append_sheet(wb,ws,"Master DB");
      }
      if((officials||[]).length){
        const ws=XLSX.utils.json_to_sheet((officials||[]).map(r=>({"Name of Official":r.officialName||"","Car Number":r.carNumber||"","Department":r.department||""})));
        XLSX.utils.book_append_sheet(wb,ws,"Officials");
      }
      if((visitors||[]).length){
        const ws=XLSX.utils.json_to_sheet((visitors||[]).map(r=>({"Name":r.visitorName||"","Car Number":r.carNumber||"","Purpose":r.purpose||"","Officer Reference":r.officerReference||"","Contact":r.contactNumber||"","Entry Time":r.entryTime||"","Exit Time":r.exitTime||""})));
        XLSX.utils.book_append_sheet(wb,ws,"Visitors");
      }
      if((cnp||[]).length){
        const ws=XLSX.utils.json_to_sheet((cnp||[]).map(r=>({"Name":r.memberName||"","Car Number":r.carNumber||"","SAI Card Number":r.cardNumber||"","Contact":r.contactNumber||""})));
        XLSX.utils.book_append_sheet(wb,ws,"Come & Play");
      }
      if(wb.SheetNames.length===0){t_("No data to export","amber");return;}
      XLSX.writeFile(wb,"JLN_All_Databases.xlsx");
      t_("All databases downloaded ✓","green");
    }catch(e){
      console.error("Excel export error:",e);
      t_("Export failed — try again","red");
    }
  };

  const safeProps={onLogout,t_,role:"admin"};
  if(view==="master") return <MasterDbV onBack={()=>setView("dashboard")} onLogout={onLogout} master={master||[]} setMaster={setMaster} t_={t_} role="admin"/>;
  if(view==="officials") return <DbManager {...safeProps} title="Officials Database" color="gold" storageKey={K.officials} fields={OFFICIAL_FIELDS} onBack={()=>setView("dashboard")} records={officials||[]} setRecords={setOfficials}/>;
  if(view==="visitors") return <DbManager {...safeProps} title="Visitors Database" color="cyan" storageKey={K.visitors} fields={VISITOR_FIELDS} onBack={()=>setView("dashboard")} records={visitors||[]} setRecords={setVisitors}/>;
  if(view==="cnp") return <DbManager {...safeProps} title="Come & Play Database" color="purple" storageKey={K.cnp} fields={CNP_FIELDS} onBack={()=>setView("dashboard")} records={cnp||[]} setRecords={setCnp}/>;
  if(view==="banned") return <DbManager {...safeProps} title="Banned Vehicles" color="red" storageKey={K.banned} fields={BANNED_FIELDS} onBack={()=>setView("dashboard")} records={banned||[]} setRecords={setBanned}/>;
  if(view==="passes") return <GatePassMgmtV onBack={()=>setView("dashboard")} passes={passes||[]} setPasses={setPasses} master={master||[]} setMaster={setMaster} visitors={visitors||[]} setVisitors={setVisitors} onLogout={onLogout} t_={t_}/>;
  if(view==="users") return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="USER MANAGEMENT" sub={`${(users||[]).length} accounts`} onBack={()=>setView("dashboard")}/>
    <div style={{padding:"14px 15px 0"}}>
      {(users||[]).map(u=>(
        <div key={u.id} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${u.role==="admin"?T.gold:T.blue}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700}}>{u.username}{u.id===user?.id?<span style={{fontSize:10,color:T.muted,fontWeight:400}}> · you</span>:null}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:1}}>{u.role==="admin"?"Super Admin":"Staff"}</div>
          </div>
          <Pill color={u.role==="admin"?"gold":"blue"}>{u.role}</Pill>
        </div>
      ))}
    </div>
  </div>;

  const safeMaster=master||[];
  const safeOfficials=officials||[];
  const safeVisitors=visitors||[];
  const safeCnp=cnp||[];
  const safePasses=passes||[];
  const safeLog=log||[];
  const todayLog=safeLog.filter(e=>e.date===todayStr());
  const inside=safeLog.filter(e=>!e.exitTime&&e.date===todayStr()).length;
  const pendingPasses=safePasses.filter(p=>p.status==="pending");

  return <div style={{minHeight:"100vh",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="ADMIN DASHBOARD" sub={`${user?.username||"admin"} · ${fmtDate(Date.now())}`}/>
    <div style={{padding:"16px 15px 0"}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9,marginBottom:16}}>
        {[
          {label:"Inside Now",value:inside,icon:"🚗",color:T.goldL},
          {label:"Today's Entries",value:todayLog.length,icon:"📋",color:T.blueT},
          {label:"Officials",value:safeOfficials.length,icon:"◆",color:T.amberT},
          {label:"Visitors",value:safeVisitors.length,icon:"🎫",color:T.cyanT}
        ].map(s=>(
          <div key={s.label} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:12,padding:"14px 13px"}}>
            <div style={{fontSize:16,marginBottom:5,opacity:.7}}>{s.icon}</div>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:28,fontWeight:800,color:s.color,lineHeight:1}}>{s.value}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:4}}>{s.label}</div>
          </div>
        ))}
      </div>

      {pendingPasses.length>0&&(
        <div onClick={()=>setView("passes")} style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:11,padding:"13px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:11,cursor:"pointer"}}>
          <div style={{width:32,height:32,borderRadius:8,background:"rgba(8,145,178,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🎫</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700,color:T.cyanT}}>{pendingPasses.length} Gate Pass Request{pendingPasses.length>1?"s":""} Pending</div>
            <div style={{fontSize:10,color:T.muted,marginTop:1}}>Tap to review & approve</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
        </div>
      )}

      <button onClick={downloadAllExcel} style={{width:"100%",padding:"11px 0",borderRadius:9,border:`1px solid rgba(22,163,74,0.35)`,background:"rgba(22,163,74,0.08)",color:T.greenT,fontSize:13,cursor:"pointer",fontWeight:700,marginBottom:14}}>📥 Download All Databases (Excel)</button>

      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Databases</div>
      {[
        {icon:"🗂",label:"Master Database",sub:`${safeMaster.length} unique vehicles · auto-updated`,action:"master",color:T.gold},
        {icon:"◆",label:"Officials Database",sub:`${safeOfficials.length} officials registered`,action:"officials",color:T.amber},
        {icon:"🎫",label:"Visitors Database",sub:`${safeVisitors.length} visitor records`,action:"visitors",color:T.cyan},
        {icon:"🏸",label:"Come & Play Database",sub:`${safeCnp.length} members`,action:"cnp",color:T.purple},
        {icon:"📨",label:"Gate Pass Requests",sub:`${safePasses.length} total · ${pendingPasses.length} pending`,action:"passes",color:T.cyan},
        {icon:"⛔",label:"Banned Vehicles",sub:`${(banned||[]).length} banned`,action:"banned",color:T.red},
        {icon:"👥",label:"User Management",sub:`${(users||[]).length} accounts`,action:"users",color:T.purple},
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
      body{background:#080f1e;color:white;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:0;overscroll-behavior:none;}
      input,select,textarea{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:9px;color:white;padding:11px 13px;font-size:14px;width:100%;outline:none;transition:border-color .2s,background .2s;font-family:inherit;}
      input::placeholder{color:rgba(255,255,255,0.25);}
      input:focus,select:focus{border-color:rgba(232,160,32,0.5);background:rgba(255,255,255,0.07);}
      select option{background:#0d1a30;color:white;}
      button:focus{outline:none;}
      ::-webkit-scrollbar{width:3px;}
      ::-webkit-scrollbar-track{background:transparent;}
      ::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.1);border-radius:4px;}
    `;
    document.head.appendChild(s);
    return()=>{ try{document.head.removeChild(s);}catch{} };
  },[]);

  const [screen,setScreen]=useState(()=>{
    try{const s=localStorage.getItem("jln_session");return s?"app":"home";}catch{return "home";}
  });
  const [session,setSession]=useState(()=>{
    try{const s=localStorage.getItem("jln_session");return s?JSON.parse(s):null;}catch{return null;}
  });
  const [master,setMaster]=useState([]);
  const [officials,setOfficials]=useState([]);
  const [visitors,setVisitors]=useState([]);
  const [cnp,setCnp]=useState([]);
  const [log,setLog]=useState([]);
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [banned,setBanned]=useState([]);
  const [passes,setPasses]=useState([]);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    Promise.all([
      gs(K.master).then(d=>d&&setMaster(d)),
      gs(K.officials).then(d=>d&&setOfficials(d)),
      gs(K.visitors).then(d=>d&&setVisitors(d)),
      gs(K.cnp).then(d=>d&&setCnp(d)),
      gs(K.log).then(d=>d&&setLog(d)),
      gs(K.banned).then(d=>d&&setBanned(d)),
      gs(K.passes).then(d=>d&&setPasses(d)),
      gs(K.users).then(d=>{ if(d&&d.length) setUsers(d); }),
    ]).finally(()=>setLoading(false));

    // Poll Firebase every 12 seconds for live sync
    const interval=setInterval(()=>{
      gs(K.passes).then(d=>{ if(d) setPasses(d); });
      gs(K.master).then(d=>{ if(d) setMaster(d); });
      gs(K.officials).then(d=>{ if(d) setOfficials(d); });
      gs(K.visitors).then(d=>{ if(d) setVisitors(d); });
      gs(K.cnp).then(d=>{ if(d) setCnp(d); });
      gs(K.log).then(d=>{ if(d) setLog(d); });
      gs(K.banned).then(d=>{ if(d) setBanned(d); });
    },12000);
    return()=>clearInterval(interval);
  },[]);

  const doLogin=u=>{
    setSession(u);setScreen("app");
    try{localStorage.setItem("jln_session",JSON.stringify(u));}catch{}
  };
  const doLogout=()=>{
    setSession(null);setScreen("home");
    try{localStorage.removeItem("jln_session");}catch{}
  };

  const shared={master,setMaster,officials,setOfficials,visitors,setVisitors,cnp,setCnp,log,setLog,users,setUsers,banned,setBanned,passes,setPasses};

  if(loading) return <div style={{minHeight:"100vh",background:T.navy,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
    <div style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,color:T.goldL,letterSpacing:2}}>JLN STADIUM</div>
    <div style={{fontSize:11,color:T.muted}}>Loading data…</div>
  </div>;

  return <>
    {screen==="home"&&<HomeScreen onSelectRole={r=>setScreen(r==="admin"?"loginAdmin":"loginStaff")} onRequestPass={()=>setScreen("requestPass")}/>}
    {screen==="loginAdmin"&&<LoginScreen role="admin" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="loginStaff"&&<LoginScreen role="staff" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="requestPass"&&<RequestPassScreen onBack={()=>setScreen("home")} passes={passes} setPasses={setPasses}/>}
    {screen==="app"&&session?.role==="staff"&&<StaffApp onLogout={doLogout} {...shared} user={session}/>}
    {screen==="app"&&session?.role==="admin"&&<AdminApp onLogout={doLogout} {...shared} user={session}/>}
    {screen==="app"&&!session&&<HomeScreen onSelectRole={r=>setScreen(r==="admin"?"loginAdmin":"loginStaff")} onRequestPass={()=>setScreen("requestPass")}/>}
  </>;
}
