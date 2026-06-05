import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ─── CONSTANTS ─────────────────────────────────────────────── */
const K = { master:"jln_master_v5", log:"jln_log_v5", users:"jln_users_v5", banned:"jln_banned_v5", passes:"jln_passes_v5", cnp:"jln_cnp_v1" };
const DEFAULT_USERS = [
  { id:"admin_default", username:"admin", password:"admin123", role:"admin", createdOn:"System" },
  { id:"staff_default", username:"staff", password:"staff123", role:"staff", createdOn:"System" }
];
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

/* ─── SMALL ATOMS ────────────────────────────────────────────── */
function Pill({color,children}){
  const m={green:{bg:"rgba(22,163,74,0.18)",c:T.greenT,border:"rgba(22,163,74,0.35)"},
    red:{bg:"rgba(220,38,38,0.18)",c:T.redT,border:"rgba(220,38,38,0.35)"},
    amber:{bg:"rgba(217,119,6,0.18)",c:T.amberT,border:"rgba(217,119,6,0.35)"},
    blue:{bg:"rgba(37,99,235,0.18)",c:T.blueT,border:"rgba(37,99,235,0.35)"},
    gold:{bg:"rgba(232,160,32,0.15)",c:T.goldL,border:"rgba(232,160,32,0.35)"},
    purple:{bg:"rgba(124,58,237,0.18)",c:T.purpleT,border:"rgba(124,58,237,0.35)"},
    cyan:{bg:"rgba(8,145,178,0.18)",c:T.cyanT,border:"rgba(8,145,178,0.35)"}};
  const s=m[color]||m.blue;
  return <span style={{background:s.bg,color:s.c,border:`1px solid ${s.border}`,padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,whiteSpace:"nowrap",textTransform:"uppercase"}}>{children}</span>;
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
  return <button onClick={onClick} disabled={disabled} style={{width:"100%",padding:"13px 0",borderRadius:10,border:"none",background:disabled?"rgba(255,255,255,0.06)":cfg.bg,color:disabled?"rgba(255,255,255,0.22)":(color==="gold"?T.navy:"white"),fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,letterSpacing:1.2,cursor:disabled?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8,boxShadow:disabled?"none":`0 4px 16px ${cfg.shadow}`,transition:"opacity .15s,transform .1s",...style}} onMouseEnter={e=>{if(!disabled)e.currentTarget.style.opacity=".9"}} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>{children}</button>;
};
const StatCard=({label,value,icon,color})=><div style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:12,padding:"14px 13px",flex:1,minWidth:0}}>
  <div style={{fontSize:16,marginBottom:5,opacity:.7}}>{icon}</div>
  <div style={{fontFamily:"'Barlow Condensed'",fontSize:28,fontWeight:800,color:color||T.goldL,lineHeight:1,letterSpacing:-0.5}}>{value}</div>
  <div style={{fontSize:10,color:T.muted,marginTop:4,letterSpacing:.3}}>{label}</div>
</div>;
const Toast=({msg,type})=>{
  const cfg={green:{bg:"rgba(22,163,74,0.95)",border:"rgba(22,163,74,0.5)"},red:{bg:"rgba(220,38,38,0.95)",border:"rgba(220,38,38,0.5)"},amber:{bg:"rgba(217,119,6,0.95)",border:"rgba(217,119,6,0.5)"},purple:{bg:"rgba(124,58,237,0.95)",border:"rgba(124,58,237,0.5)"},blue:{bg:"rgba(37,99,235,0.95)",border:"rgba(37,99,235,0.5)"},cyan:{bg:"rgba(8,145,178,0.95)",border:"rgba(8,145,178,0.5)"}}[type]||{bg:"rgba(13,26,48,0.97)",border:T.border};
  return msg?<div style={{position:"absolute",bottom:20,left:"50%",transform:"translateX(-50%)",background:cfg.bg,border:`1px solid ${cfg.border}`,backdropFilter:"blur(8px)",color:"white",padding:"10px 20px",borderRadius:40,fontSize:12,fontWeight:600,letterSpacing:.3,zIndex:100,whiteSpace:"nowrap",pointerEvents:"none",boxShadow:"0 8px 24px rgba(0,0,0,0.4)"}}>{msg}</div>:null;
};
function Hdr({role,onLogout,title,sub,onBack}){
  return <div style={{background:`linear-gradient(135deg,${T.mid} 0%,${T.light} 100%)`,padding:"13px 15px 11px",borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10,backdropFilter:"blur(10px)"}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
      <div style={{display:"flex",alignItems:"center",gap:9}}>
        {onBack&&<button onClick={onBack} style={{background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.6)",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:600,letterSpacing:.5,transition:"background .15s"}} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.1)"} onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.06)"}>← Back</button>}
        <div><div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,letterSpacing:1,color:T.goldL,lineHeight:1.1}}>{title}</div>
        {sub&&<div style={{fontSize:10,color:T.muted,marginTop:2,letterSpacing:.3}}>{sub}</div>}</div>
      </div>
      <div style={{display:"flex",gap:6,alignItems:"center"}}>
        {role&&<div style={{background:role==="admin"?"rgba(232,160,32,0.12)":"rgba(37,99,235,0.12)",border:`1px solid ${role==="admin"?"rgba(232,160,32,0.3)":"rgba(37,99,235,0.3)"}`,borderRadius:6,padding:"3px 9px",fontSize:10,fontWeight:700,letterSpacing:.5,color:role==="admin"?T.goldL:T.blueT}}>{role==="admin"?"ADMIN":"STAFF"}</div>}
        {onLogout&&<button onClick={onLogout} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:600,cursor:"pointer",letterSpacing:.3}}>Logout</button>}
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
    <div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:16,padding:"22px 20px",width:"100%",boxShadow:"0 24px 60px rgba(0,0,0,0.6)"}}>
      <div style={{fontSize:13,lineHeight:1.65,marginBottom:18,color:"rgba(255,255,255,0.8)"}}>{msg}</div>
      <div style={{display:"flex",gap:9}}>
        <button onClick={onNo} style={{flex:1,padding:"11px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,fontSize:13,cursor:"pointer",fontWeight:500}}>Cancel</button>
        <button onClick={onYes} style={{flex:1,padding:"11px 0",borderRadius:9,border:"none",background:T.red,color:"white",fontSize:13,fontWeight:700,cursor:"pointer",letterSpacing:.3}}>Confirm</button>
      </div>
    </div>
  </div>;
}

/* ─── HOME SCREEN ────────────────────────────────────────────── */
function HomeScreen({onSelectRole,onRequestPass}){
  return <div style={{minHeight:600,background:T.navy}}>
    <div style={{background:`linear-gradient(160deg,${T.mid} 0%,${T.light} 100%)`,padding:"36px 20px 28px",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",inset:0,background:"radial-gradient(ellipse at top right,rgba(232,160,32,0.06) 0%,transparent 60%)",pointerEvents:"none"}}/>
      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:4}}>
        <div style={{width:46,height:46,borderRadius:13,background:"rgba(232,160,32,0.1)",border:`1px solid rgba(232,160,32,0.25)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>🏟</div>
        <div>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:900,letterSpacing:1.5,color:T.goldL,lineHeight:1}}>JLN STADIUM</div>
          <div style={{fontSize:9,color:"rgba(255,255,255,0.3)",letterSpacing:2.5,marginTop:4,textTransform:"uppercase"}}>Vehicle Access Control</div>
        </div>
      </div>
    </div>

    <div style={{padding:"22px 16px 24px"}}>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Security Personnel</div>
      {[{role:"admin",icon:"◆",label:"Super Admin",desc:"Dashboard · Users · Database · Reports",border:"rgba(232,160,32,0.3)",clr:T.goldL,bg:"rgba(232,160,32,0.05)"},
        {role:"staff",icon:"◈",label:"Security Staff",desc:"Entry / Exit · Verification · Visitor Passes",border:"rgba(37,99,235,0.35)",clr:T.blueT,bg:"rgba(37,99,235,0.05)"}
      ].map(({role,icon,label,desc,border,clr,bg})=>(
        <div key={role} onClick={()=>onSelectRole(role)} style={{background:bg,border:`1px solid ${border}`,borderRadius:13,padding:"16px 15px",marginBottom:9,cursor:"pointer",display:"flex",alignItems:"center",gap:12,transition:"border-color .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.04)"}
          onMouseLeave={e=>e.currentTarget.style.background=bg}>
          <div style={{width:36,height:36,borderRadius:9,background:"rgba(255,255,255,0.06)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:clr,flexShrink:0}}>{icon}</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,color:clr,letterSpacing:.3}}>{label}</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.38)",marginTop:2,lineHeight:1.4}}>{desc}</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:16,flexShrink:0}}>›</div>
        </div>
      ))}

      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10,marginTop:20}}>Visitor / Public</div>
      <div onClick={onRequestPass} style={{background:"rgba(8,145,178,0.05)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:13,padding:"16px 15px",cursor:"pointer",display:"flex",alignItems:"center",gap:12}}
        onMouseEnter={e=>e.currentTarget.style.background="rgba(8,145,178,0.09)"}
        onMouseLeave={e=>e.currentTarget.style.background="rgba(8,145,178,0.05)"}>
        <div style={{width:36,height:36,borderRadius:9,background:"rgba(8,145,178,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:T.cyanT,flexShrink:0}}>🎫</div>
        <div style={{flex:1}}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:700,color:T.cyanT,letterSpacing:.3}}>Request Gate Pass</div>
          <div style={{fontSize:11,color:"rgba(255,255,255,0.38)",marginTop:2,lineHeight:1.4}}>Submit visit request · Admin reviews & approves</div>
        </div>
        <div style={{color:"rgba(255,255,255,0.2)",fontSize:16}}>›</div>
      </div>

      <div style={{textAlign:"center",fontSize:10,color:"rgba(255,255,255,0.15)",marginTop:28,letterSpacing:.3}}>
        Default — Admin: <span style={{color:"rgba(255,255,255,0.28)"}}>admin123</span> · Staff: <span style={{color:"rgba(255,255,255,0.28)"}}>staff123</span>
      </div>
    </div>
  </div>;
}

/* ─── LOGIN MODAL SCREEN ─────────────────────────────────────── */
function LoginScreen({role,users,onLogin,onBack}){
  const [selUser,setSelUser]=useState("");
  const [pwd,setPwd]=useState(""); const [err,setErr]=useState("");
  const roleUsers=users.filter(u=>u.role===role);
  useEffect(()=>{if(roleUsers.length===1)setSelUser(roleUsers[0].id);},[]);
  const try_=()=>{
    const u=users.find(x=>x.id===selUser&&x.password===pwd);
    if(u)onLogin(u); else{setErr("Incorrect password");setPwd("");}
  };
  const isAdmin=role==="admin";
  return <div style={{minHeight:600,background:T.navy}}>
    <div style={{background:`linear-gradient(160deg,${T.mid},${T.light})`,padding:"22px 16px 20px"}}>
      <button onClick={onBack} style={{background:"rgba(255,255,255,0.06)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.55)",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",marginBottom:16,fontWeight:600}}>← Back</button>
      <div style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,letterSpacing:.8,color:isAdmin?T.goldL:T.blueT,lineHeight:1}}>
        {isAdmin?"Super Admin":"Security Staff"}
      </div>
      <div style={{fontSize:11,color:T.muted,marginTop:3}}>Sign in to continue</div>
    </div>
    <div style={{padding:"22px 16px"}}>
      {roleUsers.length>1&&<FormField label="Account" required>
        <select value={selUser} onChange={e=>setSelUser(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%",marginBottom:4}}>
          <option value="">— Select account —</option>
          {roleUsers.map(u=><option key={u.id} value={u.id}>{u.username}</option>)}
        </select>
      </FormField>}
      {roleUsers.length===1&&<div style={{background:"rgba(255,255,255,0.04)",borderRadius:9,padding:"11px 13px",marginBottom:14,fontSize:13,color:"rgba(255,255,255,0.6)",border:`1px solid ${T.border}`}}>Logging in as <b style={{color:"white"}}>{roleUsers[0].username}</b></div>}
      <FormField label="Password" required>
        <input type="password" autoFocus placeholder="Enter password" value={pwd}
          onChange={e=>{setPwd(e.target.value);setErr("");}}
          onKeyDown={e=>e.key==="Enter"&&try_()}/>
      </FormField>
      {err&&<div style={{color:"#fca5a5",fontSize:12,marginBottom:12,display:"flex",alignItems:"center",gap:5}}><span>⚠</span>{err}</div>}
      <BigBtn onClick={try_} color={isAdmin?"gold":"blue"} style={{marginTop:6}}>Sign In →</BigBtn>
    </div>
  </div>;
}

/* ─── REQUEST GATE PASS ──────────────────────────────────────── */
function RequestPassScreen({onBack,passes,setPasses}){
  const [name,setName]=useState(""); const [car,setCar]=useState("");
  const [phone,setPhone]=useState(""); const [purpose,setPurpose]=useState("");
  const [dept,setDept]=useState(""); const [idFile,setIdFile]=useState(null);
  const [idPreview,setIdPreview]=useState(null);
  const [submitted,setSubmitted]=useState(false);
  const [refId,setRefId]=useState("");
  const [refOfficerName,setRefOfficerName]=useState("");
  const [refOfficerDesig,setRefOfficerDesig]=useState("");
  const [refOfficerDept,setRefOfficerDept]=useState("");
  const [refOfficerPhone,setRefOfficerPhone]=useState("");
  const fileRef=useRef();

  const handleFile=e=>{
    const f=e.target.files[0]; if(!f)return;
    setIdFile(f);
    const reader=new FileReader();
    reader.onload=ev=>setIdPreview(ev.target.result);
    reader.readAsDataURL(f);
  };

  const submit=()=>{
    if(!name.trim()||!car.trim()||!phone.trim()||!purpose.trim()||!dept.trim()){
      alert("Please fill all required fields.");return;
    }
    if(!/^\d{10}$/.test(phone.replace(/[\s-]/g,""))){
      alert("Enter a valid 10-digit phone number.");return;
    }
    const id="GP-"+Math.random().toString(36).slice(2,8).toUpperCase();
    const entry={id,name:name.trim(),car:car.toUpperCase().trim(),phone:phone.trim(),
      purpose:purpose.trim(),dept:dept.trim(),idAttached:!!idFile,
      idPreview:idPreview||null,status:"pending",requestedOn:Date.now(),date:todayStr(),
      refOfficerName:refOfficerName.trim(),refOfficerDesig:refOfficerDesig.trim(),
      refOfficerDept:refOfficerDept.trim(),refOfficerPhone:refOfficerPhone.trim()};
    const np=[...(passes||[]),entry];
    setPasses(np);ss(K.passes,np);
    setRefId(id);setSubmitted(true);
  };

  if(submitted) return <div style={{minHeight:600,background:T.navy}}>
    <Hdr title="REQUEST SUBMITTED" sub="Gate Pass" onBack={onBack}/>
    <div style={{padding:"32px 18px",textAlign:"center"}}>
      <div style={{width:64,height:64,borderRadius:20,background:"rgba(22,163,74,0.15)",border:"1px solid rgba(22,163,74,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,margin:"0 auto 18px"}}>✅</div>
      <div style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,color:T.greenT,marginBottom:8,letterSpacing:.5}}>Request Submitted</div>
      <div style={{fontSize:12,color:T.muted,lineHeight:1.75,marginBottom:24,maxWidth:280,margin:"0 auto 24px"}}>Your gate pass request has been sent to admin for review. Note your reference ID below.</div>
      <div style={{background:"rgba(255,255,255,0.04)",border:`1px dashed rgba(232,160,32,0.35)`,borderRadius:12,padding:"18px",marginBottom:24}}>
        <div style={{fontSize:9,fontWeight:700,letterSpacing:2.5,textTransform:"uppercase",color:"rgba(255,255,255,0.3)",marginBottom:8}}>Reference ID</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:30,fontWeight:900,letterSpacing:5,color:T.goldL}}>{refId}</div>
      </div>
      <div style={{fontSize:11,color:T.muted,lineHeight:1.7,marginBottom:28}}>Show this ID at the gate. Staff can verify once admin approves.</div>
      <BigBtn onClick={onBack} color="navy">← Back to Home</BigBtn>
    </div>
  </div>;

  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr title="REQUEST GATE PASS" sub="Fill all required fields" onBack={onBack}/>
    <div style={{padding:"14px 14px 0"}}>
      <div style={{background:"rgba(8,145,178,0.08)",border:`1px solid rgba(8,145,178,0.3)`,borderRadius:9,padding:"10px 13px",marginBottom:13,fontSize:12,color:"#67e8f9",lineHeight:1.6}}>
        🎫 Submit your details below. Admin will review and approve your gate pass. Carry your reference ID on arrival.
      </div>
      <Card>
        <SecTitle>VISITOR DETAILS</SecTitle>
        <FormField label="Full Name" required><input placeholder="e.g. Priya Sharma" value={name} onChange={e=>setName(e.target.value)}/></FormField>
        <FormField label="Vehicle Registration Number" required>
          <input placeholder="DL 01 AB 1234" value={car} onChange={e=>setCar(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:700,letterSpacing:4,textAlign:"center"}}/>
        </FormField>
        <FormField label="Phone Number" required><input placeholder="10-digit mobile number" value={phone} onChange={e=>setPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
      </Card>
      <Card>
        <SecTitle>VISIT DETAILS</SecTitle>
        <FormField label="Purpose of Visit" required><input placeholder="e.g. Meeting, Event, Delivery" value={purpose} onChange={e=>setPurpose(e.target.value)}/></FormField>
        <FormField label="Department to Visit" required>
          <select value={dept} onChange={e=>setDept(e.target.value)} style={{background:"rgba(255,255,255,0.07)",border:`1px solid ${T.border}`,borderRadius:8,color:dept?"white":"rgba(255,255,255,0.3)",padding:"10px 13px",fontSize:14,width:"100%"}}>
            <option value="">-- Select Department --</option>
            {["Administration","Protocol & VIP","Security","Operations","Media & Press","Events Management","Technical/IT","Finance","Sports Management","General"].map(d=><option key={d} value={d}>{d}</option>)}
          </select>
        </FormField>
      </Card>
      <Card>
        <SecTitle>REFERENCE OFFICER DETAILS</SecTitle>
        <div style={{background:"rgba(232,160,32,0.07)",border:"1px solid rgba(232,160,32,0.2)",borderRadius:8,padding:"9px 12px",marginBottom:12,fontSize:11,color:"rgba(251,191,36,0.8)",lineHeight:1.6}}>
          📋 Provide details of the JLN Stadium officer who referred / authorised this visit.
        </div>
        <FormField label="Name of Reference Officer"><input placeholder="Full name" value={refOfficerName} onChange={e=>setRefOfficerName(e.target.value)}/></FormField>
        <FormField label="Designation"><input placeholder="e.g. Deputy Director, Manager" value={refOfficerDesig} onChange={e=>setRefOfficerDesig(e.target.value)}/></FormField>
        <FormField label="Department"><input placeholder="e.g. Administration, Sports" value={refOfficerDept} onChange={e=>setRefOfficerDept(e.target.value)}/></FormField>
        <FormField label="Phone Number"><input placeholder="10-digit mobile number" value={refOfficerPhone} onChange={e=>setRefOfficerPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
      </Card>
      <Card>
        <SecTitle>ID PROOF (OPTIONAL)</SecTitle>
        <div style={{fontSize:12,color:T.muted,marginBottom:10,lineHeight:1.6}}>Upload a photo of your ID card (Aadhaar, PAN, Driving Licence etc.)</div>
        {idPreview?<div style={{marginBottom:10}}>
          <img src={idPreview} alt="ID preview" style={{width:"100%",borderRadius:8,maxHeight:160,objectFit:"cover"}}/>
          <button onClick={()=>{setIdFile(null);setIdPreview(null);}} style={{marginTop:8,background:"rgba(220,38,38,0.15)",border:"1px solid rgba(220,38,38,0.3)",color:"#fca5a5",borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",width:"100%"}}>✕ Remove</button>
        </div>:
        <button onClick={()=>fileRef.current.click()} style={{width:"100%",padding:"20px",borderRadius:9,border:`1px dashed ${T.border}`,background:"rgba(255,255,255,0.03)",color:T.muted,fontSize:13,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:6}}>
          <span style={{fontSize:24}}>📎</span><span>Tap to upload ID card</span>
        </button>}
        <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleFile}/>
      </Card>
      <BigBtn onClick={submit} color="cyan" style={{marginBottom:20}}>🎫 SUBMIT GATE PASS REQUEST</BigBtn>
    </div>
  </div>;
}

/* ─── VISITOR PASS PRINT (for Staff) ────────────────────────── */
function VisitorPass({pass,onClose}){
  const printPass=()=>{
    const w=window.open("","_blank","width=400,height=600");
    w.document.write(`<html><head><title>Gate Pass - ${pass.id}</title><style>
      body{font-family:Arial,sans-serif;padding:20px;background:#fff;color:#000}
      .header{text-align:center;border-bottom:2px solid #0a1628;padding-bottom:12px;margin-bottom:12px}
      .logo{font-size:22px;font-weight:bold;color:#0a1628}
      .sub{font-size:11px;color:#555}
      .ref{font-size:28px;font-weight:bold;letter-spacing:3px;color:#0a1628;text-align:center;border:2px dashed #0a1628;padding:10px;margin:12px 0;border-radius:6px}
      table{width:100%;font-size:13px;border-collapse:collapse}
      td{padding:6px 0;border-bottom:1px solid #eee;vertical-align:top}
      td:first-child{color:#777;width:40%}
      .footer{font-size:10px;color:#999;text-align:center;margin-top:16px;border-top:1px solid #eee;padding-top:10px}
      .valid{background:#dcfce7;color:#14532d;text-align:center;padding:8px;border-radius:6px;font-weight:bold;margin-bottom:10px}
    </style></head><body>
    <div class="header"><div class="logo">🏟 JLN STADIUM</div><div class="sub">VISITOR GATE PASS — VEHICLE ACCESS CONTROL</div></div>
    <div class="valid">✅ APPROVED ENTRY PASS</div>
    <div class="ref">${pass.id}</div>
    <table>
      <tr><td>Name</td><td><b>${pass.name}</b></td></tr>
      <tr><td>Vehicle No.</td><td><b>${pass.car}</b></td></tr>
      <tr><td>Phone</td><td>${pass.phone}</td></tr>
      <tr><td>Purpose</td><td>${pass.purpose}</td></tr>
      <tr><td>Department</td><td>${pass.dept}</td></tr>
      <tr><td>Date</td><td>${fmtDate(pass.requestedOn)}</td></tr>
      <tr><td>Approved by</td><td>JLN Stadium Admin</td></tr>
    </table>
    <div class="footer">This pass is valid for single entry on the date of issue only.<br>Carry a valid photo ID. Subject to security checks.</div>
    </body></html>`);
    w.document.close();w.print();
  };
  return <div style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:60,padding:"16px",backdropFilter:"blur(6px)"}}>
    <div style={{background:T.mid,border:`1px solid ${T.border}`,borderRadius:18,padding:"20px",width:"100%",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 32px 80px rgba(0,0,0,0.7)"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,color:T.goldL,letterSpacing:.5}}>Gate Pass</div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.07)",border:`1px solid ${T.border}`,color:"rgba(255,255,255,0.5)",borderRadius:7,width:28,height:28,cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
      </div>
      <div style={{background:"rgba(22,163,74,0.08)",border:"1px solid rgba(22,163,74,0.3)",borderRadius:12,padding:"14px",textAlign:"center",marginBottom:16}}>
        <div style={{fontSize:9,color:"rgba(134,239,172,0.6)",marginBottom:5,fontWeight:700,letterSpacing:2,textTransform:"uppercase"}}>Reference ID</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:26,fontWeight:900,letterSpacing:5,color:T.greenT}}>{pass.id}</div>
      </div>
      {[["Name",pass.name],["Vehicle",pass.car],["Phone",pass.phone],["Purpose",pass.purpose],["Department",pass.dept],["Date",fmtDate(pass.requestedOn)]].map(([l,v])=>(
        <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:`1px solid ${T.border}`}}>
          <span style={{fontSize:10,color:T.muted,fontWeight:600,letterSpacing:.5,textTransform:"uppercase"}}>{l}</span>
          <span style={{fontSize:12,fontWeight:600,textAlign:"right",maxWidth:"60%"}}>{v}</span>
        </div>
      ))}
      {pass.idPreview&&<div style={{marginTop:14}}><div style={{fontSize:9,color:T.muted,fontWeight:700,letterSpacing:2,textTransform:"uppercase",marginBottom:7}}>ID Proof</div><img src={pass.idPreview} alt="ID" style={{width:"100%",borderRadius:9,maxHeight:140,objectFit:"cover"}}/></div>}
      <BigBtn onClick={printPass} color="green" style={{marginTop:16}}>🖨 Print / Download Pass</BigBtn>
    </div>
  </div>;
}

/* ─── STAFF APP ──────────────────────────────────────────────── */
function StaffApp({onLogout,master,log,setLog,banned,passes,setPasses,cnp,user}){
  const [view,setView]=useState("verify");
  const [plate,setPlate]=useState(""); const [result,setResult]=useState(null);
  const [officer,setOfficer]=useState(""); const [division,setDivision]=useState("");
  const [guestName,setGuestName]=useState(""); const [guestReason,setGuestReason]=useState("");
  const [refOfficerName,setRefOfficerName]=useState(""); const [refOfficerDesig,setRefOfficerDesig]=useState("");
  const [refOfficerDept,setRefOfficerDept]=useState(""); const [refOfficerPhone,setRefOfficerPhone]=useState("");
  const [toast,setToast]=useState(null);
  const [passSearch,setPassSearch]=useState(""); const [viewPass,setViewPass]=useState(null);
  const [last4,setLast4]=useState(""); const [suggestions,setSuggestions]=useState([]);
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2400);};

  const onLast4Change=v=>{
    setLast4(v);setPlate("");setResult(null);setSuggestions([]);
    if(v.length>=1){
      const n=norm(v);
      // Search master database
      const masterMatches=master.filter(e=>norm(e.plate).endsWith(n)).map(e=>({plate:e.plate,label:e.officer||"",sublabel:e.division||"",badge:e.isComePlay?"Come & Play":e.isGuestPass?"Gate Pass":"Master DB",badgeColor:e.isComePlay?"purple":e.isGuestPass?"cyan":"green"}));
      // Search approved gate passes not yet in master
      const passMatches=(passes||[]).filter(p=>p.status==="approved"&&norm(p.car).endsWith(n)&&!master.find(e=>norm(e.plate)===norm(p.car))).map(p=>({plate:p.car.toUpperCase(),label:p.name,sublabel:p.dept||"",badge:"Gate Pass",badgeColor:"cyan"}));
      // Search come&play not yet in master
      const cnpMatches=(cnp||[]).filter(c=>norm(c.carNo).endsWith(n)&&!master.find(e=>norm(e.plate)===norm(c.carNo))).map(c=>({plate:c.carNo.toUpperCase(),label:c.name,sublabel:c.bookingDetails||"Come & Play",badge:"Come & Play",badgeColor:"purple"}));
      setSuggestions([...masterMatches,...passMatches,...cnpMatches]);
    }
  };
  const selectSuggestion=p=>{setPlate(p);setLast4("");setSuggestions([]);setResult(null);};

  const doAction=action=>{
    const raw=plate.trim();if(!raw){t_("Enter a vehicle number","red");return;}
    const isBanned=banned.find(b=>norm(b.plate)===norm(raw));
    if(isBanned&&action==="entry"){setResult({status:"banned",plate:raw.toUpperCase(),detail:`Reason: ${isBanned.reason||"Banned by admin"}`});return;}
    const inMaster=master.find(e=>norm(e.plate)===norm(raw));
    const now=Date.now();
    if(action==="entry"){
      const alreadyIn=log.find(e=>norm(e.plate)===norm(raw)&&!e.exitTime&&e.date===todayStr());
      if(alreadyIn){t_("Vehicle already logged as inside","amber");return;}
      // Check for approved gate pass
      const approvedPass=passes.find(p=>norm(p.car)===norm(raw)&&p.status==="approved");
      // Check if gate pass already used
      const usedPass=passes.find(p=>norm(p.car)===norm(raw)&&p.status==="used");
      // Check if this is a gate pass or exception vehicle that already entered before
      const isGuestPassVehicle=inMaster&&(inMaster.isGuestPass||inMaster.isComePlay);
      const prevGuestEntry=isGuestPassVehicle&&log.find(e=>norm(e.plate)===norm(raw)&&e.type==="guest");
      const prevExceptionEntry=!inMaster&&!approvedPass&&log.find(e=>norm(e.plate)===norm(raw));

      // Block if gate pass already used
      if(usedPass&&!approvedPass){
        setResult({status:"denied",plate:raw.toUpperCase(),detail:`Gate pass already used on ${fmtDate(usedPass.entryTime||Date.now())} — single entry only`});
        t_("Gate pass already used — access denied","red");return;
      }
      // Block exception vehicles that already entered before
      if(prevExceptionEntry){
        setResult({status:"denied",plate:raw.toUpperCase(),detail:"Exception/guest entry already used — single entry only"});
        t_("Single use only — access denied","red");return;
      }

      if(inMaster&&!isGuestPassVehicle){
        const entry={id:now,plate:raw.toUpperCase(),type:"regular",entryTime:now,exitTime:null,date:todayStr(),officer:inMaster.officer||"",division:inMaster.division||"",loggedBy:user.username};
        const nl=[...log,entry];setLog(nl);ss(K.log,nl);
        if(approvedPass){
          const np=passes.map(p=>p.id===approvedPass.id?{...p,status:"used",entryLogId:now,entryTime:now,entryLoggedBy:user.username}:p);
          setPasses(np);ss(K.passes,np);
        }
        setResult({status:"allowed",plate:raw.toUpperCase(),detail:`Officer: ${inMaster.officer||"—"}`});t_("Entry logged ✓","green");
      } else if(approvedPass){
        const entry={id:now,plate:raw.toUpperCase(),type:"guest",entryTime:now,exitTime:null,date:todayStr(),guestName:approvedPass.name,guestReason:approvedPass.purpose,officer:"Gate Pass",division:approvedPass.dept,loggedBy:user.username,passId:approvedPass.id};
        const nl=[...log,entry];setLog(nl);ss(K.log,nl);
        const np=passes.map(p=>p.id===approvedPass.id?{...p,status:"used",entryLogId:now,entryTime:now,entryLoggedBy:user.username}:p);
        setPasses(np);ss(K.passes,np);
        setResult({status:"allowed_guest",plate:raw.toUpperCase(),detail:`Gate Pass: ${approvedPass.id} · ${approvedPass.name}`});
        t_("Gate pass entry logged ✓","green");
      } else if(isGuestPassVehicle){
        // Come & Play or gate pass vehicle — allow single entry
        const prevEntry=log.find(e=>norm(e.plate)===norm(raw));
        if(prevEntry){
          setResult({status:"denied",plate:raw.toUpperCase(),detail:"Single entry already used for this vehicle"});
          t_("Single use only — access denied","red");return;
        }
        const entry={id:now,plate:raw.toUpperCase(),type:"guest",entryTime:now,exitTime:null,date:todayStr(),officer:inMaster.officer||"",division:inMaster.division||"",loggedBy:user.username};
        const nl=[...log,entry];setLog(nl);ss(K.log,nl);
        setResult({status:"allowed_guest",plate:raw.toUpperCase(),detail:`${inMaster.division}: ${inMaster.officer||"—"}`});
        t_("Entry logged ✓","green");
      } else {setResult({status:"denied",plate:raw.toUpperCase()});setView("exception");}
    } else {
      const entry=[...log].reverse().find(e=>norm(e.plate)===norm(raw)&&!e.exitTime);
      if(!entry){t_("No active entry found","amber");return;}
      const nl=log.map(e=>e.id===entry.id?{...e,exitTime:now}:e);
      setLog(nl);ss(K.log,nl);
      if(entry.passId){
        const np=passes.map(p=>p.id===entry.passId?{...p,exitTime:now,exitLoggedBy:user.username}:p);
        setPasses(np);ss(K.passes,np);
      }
      setResult({status:"exit",plate:raw.toUpperCase(),detail:`Exited at ${fmt(now)}`});t_("Exit logged ✓","amber");
    }
  };

  const addException=()=>{
    if(!guestName.trim()){t_("Enter guest name","red");return;}
    if(!officer.trim()){t_("Enter authorising officer","red");return;}
    const now=Date.now();
    const entry={id:now,plate:plate.toUpperCase(),type:"guest",entryTime:now,exitTime:null,date:todayStr(),guestName:guestName.trim(),guestReason:guestReason.trim(),officer:officer.trim(),division:division.trim(),loggedBy:user.username,refOfficerName:refOfficerName.trim(),refOfficerDesig:refOfficerDesig.trim(),refOfficerDept:refOfficerDept.trim(),refOfficerPhone:refOfficerPhone.trim()};
    const nl=[...log,entry];setLog(nl);ss(K.log,nl);
    setGuestName("");setOfficer("");setDivision("");setGuestReason("");
    setRefOfficerName("");setRefOfficerDesig("");setRefOfficerDept("");setRefOfficerPhone("");
    t_("Guest entry logged ✓","amber");setView("verify");
    setResult({status:"allowed_guest",plate:plate.toUpperCase(),detail:`Guest: ${entry.guestName}`});
  };

  const approvedPasses=passes.filter(p=>p.status==="approved");
  const filtered=passSearch?approvedPasses.filter(p=>norm(p.car).includes(norm(passSearch))||p.name.toLowerCase().includes(passSearch.toLowerCase())||p.id.toLowerCase().includes(passSearch.toLowerCase())):approvedPasses;

  if(view==="passes") return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="VISITOR PASSES" sub={`${approvedPasses.length} approved`} onBack={()=>setView("verify")}/>
    <div style={{padding:"14px 15px 0"}}>
      <input placeholder="Search name, plate, or pass ID…" value={passSearch} onChange={e=>setPassSearch(e.target.value)} style={{marginBottom:14,background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"10px 13px",fontSize:13,width:"100%",outline:"none"}}/>
      {!filtered.length?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>No approved passes found</div>:
        filtered.map(p=><div key={p.id} style={{background:"rgba(255,255,255,0.03)",border:"1px solid rgba(8,145,178,0.25)",borderRadius:11,padding:"12px 14px",marginBottom:8,cursor:"pointer",transition:"border-color .15s"}} onClick={()=>setViewPass(p)}
          onMouseEnter={e=>e.currentTarget.style.borderColor="rgba(8,145,178,0.5)"}
          onMouseLeave={e=>e.currentTarget.style.borderColor="rgba(8,145,178,0.25)"}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,letterSpacing:3,color:"white"}}>{p.car}</div>
            <Pill color="cyan">Approved</Pill>
          </div>
          <div style={{fontSize:12,color:"rgba(255,255,255,0.75)",marginBottom:2}}>{p.name}</div>
          <div style={{fontSize:10,color:T.muted}}>{p.dept} · {p.id}</div>
        </div>)}
    </div>
    {viewPass&&<VisitorPass pass={viewPass} onClose={()=>setViewPass(null)}/>}
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;

  if(view==="exception") return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="EXCEPTION ENTRY" sub="Guest / one-time vehicle" onBack={()=>{setView("verify");setResult(null);}}/>
    <div style={{padding:"15px 15px 0"}}>
      <div style={{background:"rgba(220,38,38,0.08)",border:`1px solid rgba(220,38,38,0.3)`,borderRadius:12,padding:"14px 15px",marginBottom:14}}>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:16,fontWeight:800,color:T.redT,letterSpacing:.5}}>Not in Master Database</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:20,letterSpacing:5,color:"#fca5a5",marginTop:3,fontWeight:700}}>{plate.toUpperCase()}</div>
        <div style={{fontSize:11,color:"rgba(252,165,165,0.6)",marginTop:4}}>Fill details below to grant exception entry</div>
      </div>
      <Card>
        <SecTitle>GUEST DETAILS</SecTitle>
        <FormField label="Guest / Visitor Name" required><input placeholder="Full name" value={guestName} onChange={e=>setGuestName(e.target.value)}/></FormField>
        <FormField label="Purpose / Reason"><input placeholder="e.g. VIP Guest, Event Staff" value={guestReason} onChange={e=>setGuestReason(e.target.value)}/></FormField>
        <SecTitle>AUTHORISATION</SecTitle>
        <FormField label="Authorising Officer" required><input placeholder="Name & rank" value={officer} onChange={e=>setOfficer(e.target.value)}/></FormField>
        <FormField label="Division / Unit"><input placeholder="e.g. Admin, Security" value={division} onChange={e=>setDivision(e.target.value)}/></FormField>
        <SecTitle>REFERENCE OFFICER DETAILS</SecTitle>
        <FormField label="Name of Reference Officer"><input placeholder="Full name of reference officer" value={refOfficerName} onChange={e=>setRefOfficerName(e.target.value)}/></FormField>
        <FormField label="Designation"><input placeholder="e.g. Deputy Director, Manager" value={refOfficerDesig} onChange={e=>setRefOfficerDesig(e.target.value)}/></FormField>
        <FormField label="Department"><input placeholder="e.g. Administration, Sports" value={refOfficerDept} onChange={e=>setRefOfficerDept(e.target.value)}/></FormField>
        <FormField label="Phone Number"><input placeholder="10-digit mobile number" value={refOfficerPhone} onChange={e=>setRefOfficerPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
        <BigBtn onClick={addException} color="amber" style={{marginTop:6}}>⚠ ALLOW EXCEPTION ENTRY</BigBtn>
      </Card>
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;

  const resultColors={allowed:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},allowed_guest:{bg:"rgba(22,163,74,0.12)",border:T.green,tc:T.greenT},exit:{bg:"rgba(217,119,6,0.12)",border:T.amber,tc:T.amberT},denied:{bg:"rgba(220,38,38,0.12)",border:T.red,tc:T.redT},banned:{bg:"rgba(127,29,29,0.3)",border:"#7f1d1d",tc:"#fca5a5"}};
  const rc=result?resultColors[result.status]||resultColors.denied:null;
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="staff" onLogout={onLogout} title="SECURITY STAFF" sub={`${user.username}`}/>
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
              <div>
                <span style={{fontFamily:"'Barlow Condensed'",fontSize:17,fontWeight:800,letterSpacing:3,color:T.goldL}}>{s.plate}</span>
                <div style={{fontSize:10,color:T.muted,marginTop:2}}>{s.label}{s.sublabel?` · ${s.sublabel}`:""}</div>
              </div>
              <Pill color={s.badgeColor||"green"}>{s.badge||"Master DB"}</Pill>
            </div>)}
            <div style={{padding:"9px 14px",fontSize:10,color:T.muted,textAlign:"center"}}>Tap to select a vehicle</div>
          </div>}
          {last4&&!suggestions.length&&<div style={{padding:"12px 14px",fontSize:12,color:T.muted,textAlign:"center",marginBottom:10}}>No vehicles match — <span style={{color:T.amberT,cursor:"pointer"}} onClick={()=>{setPlate(last4);setSuggestions([]);setLast4("");}}>use "{last4}" as full plate</span></div>}
        </div>}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <BigBtn onClick={()=>doAction("entry")} color="green" disabled={!plate}>▲ ENTRY</BigBtn>
          <BigBtn onClick={()=>doAction("exit")} color="red" disabled={!plate}>▼ EXIT</BigBtn>
        </div>
      </Card>

      {result&&<div style={{borderRadius:14,padding:"20px 16px",textAlign:"center",marginBottom:14,background:rc.bg,border:`1px solid ${rc.border}`}}>
        <div style={{fontSize:32,marginBottom:6}}>{result.status==="allowed"?"✅":result.status==="allowed_guest"?"🎫":result.status==="exit"?"🔄":result.status==="banned"?"⛔":"🚫"}</div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:20,fontWeight:800,letterSpacing:1.5,color:rc.tc}}>
          {result.status==="allowed"?"ENTRY ALLOWED":result.status==="allowed_guest"?"GATE PASS ENTRY":result.status==="exit"?"EXIT RECORDED":result.status==="banned"?"BANNED VEHICLE":"NOT ALLOWED"}
        </div>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:700,letterSpacing:5,marginTop:4,opacity:.6,color:rc.tc}}>{result.plate}</div>
        {result.detail&&<div style={{fontSize:11,marginTop:5,opacity:.7,color:rc.tc}}>{result.detail}</div>}
      </div>}

      <div onClick={()=>setView("passes")} style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.25)",borderRadius:12,padding:"14px 15px",cursor:"pointer",display:"flex",alignItems:"center",gap:11}}>
        <div style={{width:34,height:34,borderRadius:8,background:"rgba(8,145,178,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>🎫</div>
        <div style={{flex:1}}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700,color:T.cyanT,letterSpacing:.3}}>Visitor Gate Passes</div>
          <div style={{fontSize:10,color:T.muted,marginTop:1}}>{approvedPasses.length} approved · View & print</div>
        </div>
        <div style={{color:"rgba(255,255,255,0.2)",fontSize:15}}>›</div>
      </div>
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;
}

/* ─── ADMIN APP ──────────────────────────────────────────────── */
function AdminApp({onLogout,master,setMaster,log,setLog,users,setUsers,banned,setBanned,passes,setPasses,cnp,setCnp,user}){
  const [view,setView]=useState("dashboard");
  const [selEntry,setSelEntry]=useState(null);
  const [toast,setToast]=useState(null);
  const t_=(msg,type)=>{setToast({msg,type});setTimeout(()=>setToast(null),2400);};
  const todayLog=log.filter(e=>e.date===todayStr());
  const inside=log.filter(e=>!e.exitTime&&e.date===todayStr()).length;
  const pendingPasses=passes.filter(p=>p.status==="pending");
  const props={onLogout,master,setMaster,log,setLog,users,setUsers,banned,setBanned,passes,setPasses,cnp,setCnp,t_,user};

  if(view==="addCar") return <AddCarV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="db") return <DbV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="logView") return <LogV {...props} onBack={()=>setView("dashboard")} onSelect={e=>{setSelEntry(e);setView("logDetail");}}/>;
  if(view==="logDetail"&&selEntry) return <LogDetailV entry={selEntry} onBack={()=>setView("logView")} onLogout={onLogout}/>;
  if(view==="userMgmt") return <UserMgmtV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="bannedMgmt") return <BannedV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="report") return <ReportV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="changePwd") return <ChangePwdV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="gatePassMgmt") return <GatePassMgmtV {...props} onBack={()=>setView("dashboard")}/>;
  if(view==="comePlay") return <ComePlayV {...props} onBack={()=>setView("dashboard")}/>;

  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="ADMIN DASHBOARD" sub={`${user.username} · ${fmtDate(Date.now())}`}/>
    <div style={{padding:"16px 15px 0"}}>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Live Status</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9,marginBottom:16}}>
        <StatCard label="Currently Inside" value={inside} icon="🚗" color={T.goldL}/>
        <StatCard label="Total Today" value={todayLog.length} icon="📋" color={T.blueT}/>
        <StatCard label="Official" value={todayLog.filter(e=>e.type==="regular").length} icon="✅" color={T.greenT}/>
        <StatCard label="Guests" value={todayLog.filter(e=>e.type==="guest").length} icon="⚠️" color={T.amberT}/>
      </div>

      {pendingPasses.length>0&&<div style={{background:"rgba(8,145,178,0.06)",border:"1px solid rgba(8,145,178,0.3)",borderRadius:11,padding:"13px 14px",marginBottom:16,display:"flex",alignItems:"center",gap:11,cursor:"pointer"}} onClick={()=>setView("gatePassMgmt")}>
        <div style={{width:32,height:32,borderRadius:8,background:"rgba(8,145,178,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>🎫</div>
        <div style={{flex:1}}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700,color:T.cyanT,letterSpacing:.3}}>{pendingPasses.length} Pass Request{pendingPasses.length>1?"s":""} Pending</div>
          <div style={{fontSize:10,color:T.muted,marginTop:1}}>Tap to review and approve</div>
        </div>
        <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
      </div>}

      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Manage</div>
      {[
        {icon:"📊",label:"Entry / Exit Log",sub:`${todayLog.length} entries today`,action:"logView",color:T.blue},
        {icon:"🗂",label:"Master Database",sub:`${master.length} registered vehicles`,action:"db",color:T.gold},
        {icon:"➕",label:"Add / Import Vehicles",sub:"Register · Bulk import · Search · Remove",action:"addCar",color:T.green},
        {icon:"🎫",label:"Gate Pass Requests",sub:`${passes.length} total · ${pendingPasses.length} pending`,action:"gatePassMgmt",color:T.cyan},
        {icon:"⛔",label:"Banned Vehicles",sub:`${banned.length} banned`,action:"bannedMgmt",color:T.red},
        {icon:"🏸",label:"Come & Play Database",sub:`${cnp.length} members registered`,action:"comePlay",color:T.purple},
        {icon:"👥",label:"User Management",sub:`${users.length} accounts`,action:"userMgmt",color:T.purple},
        {icon:"🔑",label:"Change Passwords",sub:"Update credentials",action:"changePwd",color:T.amber},
        {icon:"📥",label:"Download Report",sub:"Custom date range export",action:"report",color:T.cyan},
      ].map(({icon,label,sub,action,color})=>(
        <div key={action} onClick={()=>setView(action)} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${color}`,borderRadius:11,padding:"12px 14px",marginBottom:7,display:"flex",alignItems:"center",gap:11,cursor:"pointer",transition:"background .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.055)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.03)"}>
          <div style={{fontSize:16,opacity:.8}}>{icon}</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700,letterSpacing:.2}}>{label}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:1}}>{sub}</div>
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
        </div>
      ))}
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type}/>}
  </div>;
}

/* ─── GATE PASS MANAGEMENT (Admin) ──────────────────────────── */
function GatePassMgmtV({onBack,passes,setPasses,master,setMaster,onLogout,t_}){
  const [filter,setFilter]=useState("pending");
  const [viewPass,setViewPass]=useState(null);
  const filtered=passes.filter(p=>filter==="all"||p.status===filter).slice().reverse();
  const approve=id=>{
    const pass=passes.find(p=>p.id===id);
    const u=passes.map(p=>p.id===id?{...p,status:"approved"}:p);
    setPasses(u);ss(K.passes,u);
    // Add car to master database if not already there
    if(pass&&!master.find(e=>norm(e.plate)===norm(pass.car))){
      const nm=[...master,{plate:pass.car.toUpperCase(),officer:pass.name,division:"Gate Pass",addedOn:fmtDate(Date.now()),addedBy:"Gate Pass System",isGuestPass:true}];
      setMaster(nm);ss(K.master,nm);
      t_("Pass approved + car added to master ✓","green");
    } else {
      t_("Pass approved ✓","green");
    }
  };
  const reject=id=>{const u=passes.map(p=>p.id===id?{...p,status:"rejected"}:p);setPasses(u);ss(K.passes,u);t_("Pass rejected","amber");};
  const statusColor={pending:"amber",approved:"green",rejected:"red",used:"cyan"};
  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
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
              <div style={{fontSize:9,color:T.muted,letterSpacing:.5}}>{p.id}</div>
            </div>
          </div>
          <div style={{fontSize:10,color:T.muted,marginBottom:2}}>📞 {p.phone} · 🏛 {p.dept}</div>
          <div style={{fontSize:10,color:T.muted,marginBottom:p.status==="used"?4:10}}>📝 {p.purpose} · {fmtDate(p.requestedOn)}</div>
          {p.status==="used"&&<div style={{background:"rgba(8,145,178,0.08)",border:"1px solid rgba(8,145,178,0.2)",borderRadius:7,padding:"6px 10px",marginBottom:10,fontSize:10,color:T.cyanT}}>
            ✅ Entered: {p.entryTime?fmt(p.entryTime):"—"}{p.exitTime?` · Exited: ${fmt(p.exitTime)}`:" · Still inside"}{p.entryLoggedBy?` · by ${p.entryLoggedBy}`:""}
          </div>}
          <div style={{display:"flex",gap:7}}>
            <button onClick={()=>setViewPass(p)} style={{flex:1,padding:"8px 0",borderRadius:8,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.7)",fontSize:11,cursor:"pointer",fontWeight:600}}>Details</button>
            {p.status==="pending"&&<>
              <button onClick={()=>approve(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.green,color:"white",fontSize:11,fontWeight:700,cursor:"pointer"}}>✓ Approve</button>
              <button onClick={()=>reject(p.id)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.red,color:"white",fontSize:11,fontWeight:700,cursor:"pointer"}}>✕ Reject</button>
            </>}
            {p.status==="approved"&&<button onClick={()=>setViewPass(p)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"none",background:T.cyan,color:"white",fontSize:11,fontWeight:700,cursor:"pointer"}}>🖨 Print</button>}
          </div>
        </div>)}
    </div>
    {viewPass&&<VisitorPass pass={viewPass} onClose={()=>setViewPass(null)}/>}
  </div>;
}

/* ─── ADD CAR ─────────────────────────────────────────────────── */
function AddCarV({onBack,master,setMaster,t_,onLogout,user}){
  const [plate,setPlate]=useState(""); const [officer,setOfficer]=useState(""); const [division,setDivision]=useState("");
  const [search,setSearch]=useState(""); const [confirm,setConfirm]=useState(null);
  const [tab,setTab]=useState("add"); // "add" | "list"
  const importRef=useRef();

  const add=()=>{
    if(!plate.trim()){t_("Enter plate number","red");return;}
    if(master.find(e=>norm(e.plate)===norm(plate))){t_("Already in master database","amber");return;}
    const u=[...master,{plate:plate.toUpperCase(),officer,division,addedOn:fmtDate(Date.now()),addedBy:user.username}];
    setMaster(u);ss(K.master,u);setPlate("");setOfficer("");setDivision("");t_("Vehicle added ✓","green");
  };

  const importXL=ev=>{
    const file=ev.target.files[0];if(!file)return;
    if(importRef.current)importRef.current.value="";
    const reader=new FileReader();
    reader.onerror=()=>t_("Could not read file","red");
    reader.onload=e=>{
      try{
        const data=new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:"array"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
        const getCol=(row,aliases)=>{for(const a of aliases){const found=Object.keys(row).find(k=>k.trim().toLowerCase()===a.toLowerCase());if(found&&row[found]!=="")return row[found].toString().trim();}return "";};
        let added=0;const cur=[...master];
        rows.forEach(row=>{
          const p=getCol(row,["registration number","reg number","reg no","plate","vehicle no","vehicle number","reg","car number","car no"]).toUpperCase().replace(/\s+/g,"");
          if(!p||cur.find(x=>norm(x.plate)===norm(p)))return;
          cur.push({plate:p,officer:getCol(row,["officer name","officer","name","owner"]),division:getCol(row,["division","unit","department","dept"]),addedOn:fmtDate(Date.now()),addedBy:user.username});
          added++;
        });
        setMaster(cur);ss(K.master,cur);
        t_(`${added} vehicle${added!==1?"s":""} imported ✓`,"green");
      }catch(err){console.error(err);t_("Error reading file — check format","red");}
    };
    reader.readAsArrayBuffer(file);
  };

  const removePlate=p=>{
    const u=master.filter(e=>norm(e.plate)!==norm(p));
    setMaster(u);ss(K.master,u);t_("Vehicle removed","amber");setConfirm(null);
  };

  const filtered=search.trim()
    ?master.filter(e=>norm(e.plate).includes(norm(search))||e.officer?.toLowerCase().includes(search.toLowerCase())||e.division?.toLowerCase().includes(search.toLowerCase()))
    :master;

  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="VEHICLES" sub={`${master.length} in master database`} onBack={onBack}/>

    {/* Tab bar */}
    <div style={{display:"flex",gap:0,borderBottom:`1px solid ${T.border}`,background:T.mid}}>
      {[["add","➕ Add / Import"],["list","🗂 Search & Remove"]].map(([v,l])=>(
        <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"11px 0",border:"none",borderBottom:`2px solid ${tab===v?T.goldL:"transparent"}`,background:"transparent",color:tab===v?T.goldL:T.muted,fontSize:12,fontWeight:tab===v?700:400,cursor:"pointer",letterSpacing:.3,transition:"color .15s"}}>{l}</button>
      ))}
    </div>

    <div style={{padding:"15px 15px 0"}}>
      {tab==="add"&&<>
        <Card>
          <SecTitle>ADD SINGLE VEHICLE</SecTitle>
          <FormField label="Registration Number" required>
            <input placeholder="DL 01 AB 1234" value={plate} onChange={e=>setPlate(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed',monospace",fontSize:24,fontWeight:800,letterSpacing:5,textAlign:"center",background:"rgba(255,255,255,0.05)",border:`1px solid ${plate?T.gold:T.border}`,borderRadius:10,color:"white",padding:"13px",width:"100%",outline:"none",transition:"border-color .2s"}}/>
          </FormField>
          <FormField label="Officer Name"><input placeholder="e.g. Supt. Rajesh Kumar" value={officer} onChange={e=>setOfficer(e.target.value)}/></FormField>
          <FormField label="Division / Unit"><input placeholder="e.g. Delhi Traffic Police" value={division} onChange={e=>setDivision(e.target.value)}/></FormField>
          <BigBtn onClick={add} color="gold" style={{marginTop:4}}>+ ADD TO MASTER DATABASE</BigBtn>
        </Card>

        <Card>
          <SecTitle>BULK IMPORT VIA EXCEL</SecTitle>
          <div style={{fontSize:11,color:T.muted,lineHeight:1.75,marginBottom:12}}>
            Upload an <b style={{color:"rgba(255,255,255,0.55)"}}>.xlsx</b> file with columns:<br/>
            <span style={{color:T.goldL,fontFamily:"monospace",fontSize:11}}>Registration Number, Officer Name, Division</span><br/>
            Duplicates are automatically skipped.
          </div>
          <div
            onClick={()=>importRef.current&&importRef.current.click()}
            style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:8,padding:"22px",border:`1px dashed ${T.border}`,borderRadius:10,cursor:"pointer",background:"rgba(255,255,255,0.02)",transition:"background .15s"}}
            onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.06)"}
            onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.02)"}>
            <span style={{fontSize:28,pointerEvents:"none"}}>📂</span>
            <span style={{fontSize:13,color:T.muted,pointerEvents:"none"}}>Tap to select Excel file</span>
            <span style={{fontSize:10,color:"rgba(255,255,255,0.2)",pointerEvents:"none"}}>Supports .xlsx and .xls</span>
          </div>
          <input ref={importRef} type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" style={{position:"absolute",opacity:0,pointerEvents:"none",width:0,height:0}} onChange={importXL}/>
        </Card>
      </>}

      {tab==="list"&&<>
        <div style={{marginBottom:13}}>
          <input
            placeholder="🔍  Search plate, officer, division…"
            value={search}
            onChange={e=>setSearch(e.target.value)}
            style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?T.gold:T.border}`,borderRadius:9,color:"white",padding:"11px 14px",fontSize:13,width:"100%",outline:"none",transition:"border-color .2s"}}
          />
          {search&&<div style={{fontSize:10,color:T.muted,marginTop:6,paddingLeft:2}}>{filtered.length} result{filtered.length!==1?"s":""} for "{search}"</div>}
        </div>

        {!filtered.length
          ?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>{search?"No vehicles match your search":"No vehicles in master database"}</div>
          :filtered.slice().reverse().map((e,i)=>(
            <div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${T.green}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1}}>
                <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:2.5}}>{e.plate}</div>
                <div style={{fontSize:10,color:T.muted,marginTop:2}}>{e.officer||"—"}{e.division?` · ${e.division}`:""}</div>
                {e.addedOn&&<div style={{fontSize:9,color:"rgba(255,255,255,0.18)",marginTop:2}}>Added {e.addedOn}{e.addedBy?` by ${e.addedBy}`:""}</div>}
              </div>
              <button onClick={()=>setConfirm(e.plate)} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:600,flexShrink:0}}>Remove</button>
            </div>
          ))
        }
      </>}
    </div>

    {confirm&&<ConfirmModal msg={`Remove "${confirm}" from master database? Existing log entries will not be affected.`} onYes={()=>removePlate(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── DATABASE VIEW ─────────────────────────────────────────── */
function DbV({onBack,master,setMaster,t_,onLogout}){
  const [confirm,setConfirm]=useState(null);
  const [tab,setTab]=useState("view"); // "view" | "update"
  const [dbSearch,setDbSearch]=useState("");
  // Upload-preview state
  const [fileName,setFileName]=useState("");
  const [preview,setPreview]=useState(null); // {toAdd:[], duplicates:[]}
  const [parseErr,setParseErr]=useState("");
  const fileRef=useRef();

  const removePlate=p=>{const u=master.filter(e=>norm(e.plate)!==norm(p));setMaster(u);ss(K.master,u);t_("Vehicle removed","amber");setConfirm(null);};

  const exportXL=()=>{
    if(!master.length){t_("No data","amber");return;}
    const ws=XLSX.utils.json_to_sheet(master.map(e=>({"Registration Number":e.plate,"Officer Name":e.officer||"","Division":e.division||"","Added On":e.addedOn||"","Added By":e.addedBy||""})));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Master");XLSX.writeFile(wb,"JLN_Master_DB.xlsx");t_("Exported!","green");
  };

  const handleFileSelect=ev=>{
    const file=ev.target.files[0];
    if(!file)return;
    const name=file.name; // capture before async — event object gets recycled
    setFileName(name);
    setPreview(null);
    setParseErr("");
    if(fileRef.current)fileRef.current.value=""; // reset input so same file can be re-selected
    const reader=new FileReader();
    reader.onerror=()=>setParseErr("Could not read the file. Please try again.");
    reader.onload=e=>{
      try{
        const data=new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:"array"});
        if(!wb.SheetNames.length){setParseErr("No sheets found in this file.");return;}
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!rows.length){setParseErr("Sheet is empty — no data rows found.");return;}

        // Flexible column detection: try every known alias, case-insensitive
        const getCol=(row,aliases)=>{
          for(const a of aliases){
            const found=Object.keys(row).find(k=>k.trim().toLowerCase()===a.toLowerCase());
            if(found&&row[found]!==undefined&&row[found]!=="")return row[found].toString().trim();
          }
          return "";
        };

        const toAdd=[];
        const duplicates=[];
        const skippedBlank=[];

        rows.forEach((row,idx)=>{
          const p=getCol(row,["registration number","reg number","reg no","plate","vehicle no","vehicle number","reg","car number","car no"]).toUpperCase().replace(/\s+/g,"");
          if(!p){skippedBlank.push(idx+2);return;} // row number for user reference (1-indexed + header)
          const officer=getCol(row,["officer name","officer","name","owner","owner name"]);
          const division=getCol(row,["division","unit","department","dept","dept."]);
          if(master.find(x=>norm(x.plate)===norm(p))){
            duplicates.push({plate:p,officer,division});
          } else {
            toAdd.push({plate:p,officer,division});
          }
        });

        setPreview({toAdd,duplicates,skippedBlank});
      }catch(err){
        console.error("XLSX parse error:",err);
        setParseErr("Could not parse the file. Make sure it is a valid .xlsx or .xls file and not password-protected.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const submitUpdate=async()=>{
    if(!preview||!preview.toAdd.length){t_("No new vehicles to add","amber");return;}
    const today=fmtDate(Date.now());
    const newEntries=preview.toAdd.map(e=>({
      plate:e.plate,
      officer:e.officer||"",
      division:e.division||"",
      addedOn:today
    }));
    const updated=[...master,...newEntries];
    setMaster(updated);
    await ss(K.master,updated);
    t_(`${newEntries.length} vehicle${newEntries.length!==1?"s":""} added to master database ✓`,"green");
    setPreview(null);setFileName("");setParseErr("");
    setTab("view");
  };

  const resetUpload=()=>{setPreview(null);setFileName("");setParseErr("");};

  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="MASTER DATABASE" sub={`${master.length} vehicles registered`} onBack={onBack}/>

    {/* Tabs */}
    <div style={{display:"flex",borderBottom:`1px solid ${T.border}`,background:T.mid}}>
      {[["view","🗂 View Database"],["update","📥 Update from Excel"]].map(([v,l])=>(
        <button key={v} onClick={()=>{setTab(v);resetUpload();}} style={{flex:1,padding:"11px 0",border:"none",borderBottom:`2px solid ${tab===v?T.goldL:"transparent"}`,background:"transparent",color:tab===v?T.goldL:T.muted,fontSize:12,fontWeight:tab===v?700:400,cursor:"pointer",letterSpacing:.3,transition:"color .15s"}}>{l}</button>
      ))}
    </div>

    <div style={{padding:"14px 15px 0"}}>

      {/* ── VIEW TAB ── */}
      {tab==="view"&&<>
        <input placeholder="🔍  Search by name, plate, division…" value={dbSearch} onChange={e=>setDbSearch(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${dbSearch?T.gold:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:12,transition:"border-color .2s"}}/>
        <button onClick={exportXL} style={{width:"100%",padding:"9px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.65)",fontSize:12,cursor:"pointer",fontWeight:600,marginBottom:14}}>⬇ Export Master Database as Excel</button>
        {(()=>{
          const filtered=dbSearch.trim()?master.filter(e=>norm(e.plate).includes(norm(dbSearch))||e.officer?.toLowerCase().includes(dbSearch.toLowerCase())||e.division?.toLowerCase().includes(dbSearch.toLowerCase())):master;
          return !filtered.length
            ?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>{dbSearch?"No results found":"No vehicles in master database"}</div>
            :<>{dbSearch&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""} for "{dbSearch}"</div>}
            {filtered.slice().reverse().map((e,i)=>(
              <div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${e.isComePlay?T.purple:e.isGuestPass?T.cyan:T.green}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
                <div style={{flex:1}}>
                  <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:2.5}}>{e.plate}</div>
                  <div style={{fontSize:10,color:T.muted,marginTop:2}}>{e.officer||"—"}{e.division?` · ${e.division}`:""}</div>
                  {(e.isComePlay||e.isGuestPass)&&<div style={{marginTop:4}}><Pill color={e.isComePlay?"purple":"cyan"}>{e.isComePlay?"Come & Play":"Gate Pass"}</Pill></div>}
                </div>
                <button onClick={()=>setConfirm(e.plate)} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:600}}>Remove</button>
              </div>
            ))}</>;
        })()}
      </>}

      {/* ── UPDATE TAB ── */}
      {tab==="update"&&<>

        {/* Step 1 — File format hint */}
        <div style={{background:"rgba(37,99,235,0.07)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"11px 14px",marginBottom:14,fontSize:11,color:T.blueT,lineHeight:1.8}}>
          <b style={{fontSize:12,letterSpacing:.3}}>Excel column format expected:</b><br/>
          <span style={{fontFamily:"monospace",color:T.goldL}}>Registration Number</span> &nbsp;·&nbsp; <span style={{fontFamily:"monospace",color:T.goldL}}>Officer Name</span> &nbsp;·&nbsp; <span style={{fontFamily:"monospace",color:T.goldL}}>Division</span><br/>
          <span style={{color:"rgba(255,255,255,0.3)",fontSize:10}}>Only "Registration Number" is mandatory. Duplicates are detected and skipped automatically.</span>
        </div>

        {/* Step 2 — Upload zone */}
        {!fileName&&!preview&&!parseErr&&(
          <div style={{marginBottom:14}}>
            <div
              onClick={()=>fileRef.current&&fileRef.current.click()}
              style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:10,padding:"32px 20px",border:"2px dashed rgba(37,99,235,0.4)",borderRadius:12,cursor:"pointer",background:"rgba(37,99,235,0.04)",transition:"background .15s,border-color .15s"}}
              onMouseEnter={e=>{e.currentTarget.style.background="rgba(37,99,235,0.09)";e.currentTarget.style.borderColor="rgba(37,99,235,0.7)";}}
              onMouseLeave={e=>{e.currentTarget.style.background="rgba(37,99,235,0.04)";e.currentTarget.style.borderColor="rgba(37,99,235,0.4)";}}>
              <span style={{fontSize:36,pointerEvents:"none"}}>📂</span>
              <div style={{textAlign:"center",pointerEvents:"none"}}>
                <div style={{fontSize:14,fontWeight:600,color:"rgba(255,255,255,0.85)",marginBottom:4}}>Tap to Select Excel File</div>
                <div style={{fontSize:11,color:T.muted}}>Supports .xlsx and .xls</div>
              </div>
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              style={{position:"absolute",opacity:0,pointerEvents:"none",width:0,height:0}}
              onChange={handleFileSelect}
            />
          </div>
        )}

        {/* Parse error */}
        {parseErr&&(
          <div style={{background:"rgba(220,38,38,0.08)",border:"1px solid rgba(220,38,38,0.3)",borderRadius:10,padding:"14px",marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:T.redT,marginBottom:4}}>⚠ Could Not Read File</div>
            <div style={{fontSize:11,color:"rgba(252,165,165,0.75)",lineHeight:1.65}}>{parseErr}</div>
            <button onClick={resetUpload} style={{marginTop:12,background:"rgba(220,38,38,0.15)",border:"1px solid rgba(220,38,38,0.3)",color:T.redT,borderRadius:7,padding:"7px 16px",fontSize:12,cursor:"pointer",fontWeight:600}}>Try Again</button>
          </div>
        )}

        {/* Preview panel — shown after parsing */}
        {preview&&(
          <>
            {/* File info bar */}
            <div style={{background:"rgba(255,255,255,0.04)",border:`1px solid ${T.border}`,borderRadius:9,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
              <div style={{display:"flex",alignItems:"center",gap:9}}>
                <span style={{fontSize:18}}>📄</span>
                <div>
                  <div style={{fontSize:12,fontWeight:600,color:"rgba(255,255,255,0.85)"}}>{fileName}</div>
                  <div style={{fontSize:10,color:T.muted,marginTop:1}}>{preview.toAdd.length+preview.duplicates.length} valid rows · {preview.skippedBlank?.length||0} blank skipped</div>
                </div>
              </div>
              <button onClick={resetUpload} style={{background:"transparent",border:`1px solid ${T.border}`,color:T.muted,borderRadius:7,padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:600,flexShrink:0}}>✕ Clear</button>
            </div>

            {/* Summary counters */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9,marginBottom:14}}>
              <div style={{background:"rgba(22,163,74,0.08)",border:"1px solid rgba(22,163,74,0.25)",borderRadius:10,padding:"13px 14px",textAlign:"center"}}>
                <div style={{fontFamily:"'Barlow Condensed'",fontSize:30,fontWeight:900,color:T.greenT,lineHeight:1}}>{preview.toAdd.length}</div>
                <div style={{fontSize:10,color:T.greenT,opacity:.7,marginTop:4,fontWeight:600,letterSpacing:.5,textTransform:"uppercase"}}>New vehicles</div>
                <div style={{fontSize:9,color:T.muted,marginTop:3}}>Will be added</div>
              </div>
              <div style={{background:"rgba(217,119,6,0.08)",border:"1px solid rgba(217,119,6,0.25)",borderRadius:10,padding:"13px 14px",textAlign:"center"}}>
                <div style={{fontFamily:"'Barlow Condensed'",fontSize:30,fontWeight:900,color:T.amberT,lineHeight:1}}>{preview.duplicates.length}</div>
                <div style={{fontSize:10,color:T.amberT,opacity:.7,marginTop:4,fontWeight:600,letterSpacing:.5,textTransform:"uppercase"}}>Duplicates</div>
                <div style={{fontSize:9,color:T.muted,marginTop:3}}>Already in database</div>
              </div>
            </div>

            {/* New entries list */}
            {preview.toAdd.length>0&&<>
              <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:8,marginTop:4}}>New entries to be added ({preview.toAdd.length})</div>
              <div style={{maxHeight:220,overflowY:"auto",marginBottom:14,borderRadius:9,border:`1px solid ${T.border}`}}>
                {preview.toAdd.map((e,i)=>(
                  <div key={i} style={{padding:"9px 13px",borderBottom:i<preview.toAdd.length-1?`1px solid ${T.border}`:"none",display:"flex",alignItems:"center",gap:10,background:"rgba(22,163,74,0.03)"}}>
                    <span style={{width:6,height:6,borderRadius:"50%",background:T.green,flexShrink:0,display:"inline-block"}}/>
                    <div style={{flex:1}}>
                      <div style={{fontFamily:"'Barlow Condensed'",fontSize:13,fontWeight:800,letterSpacing:2,color:"rgba(255,255,255,0.9)"}}>{e.plate}</div>
                      <div style={{fontSize:10,color:T.muted,marginTop:1}}>{e.officer||"—"}{e.division?` · ${e.division}`:""}</div>
                    </div>
                  </div>
                ))}
              </div>
            </>}

            {/* Duplicates list (collapsible) */}
            {preview.duplicates.length>0&&<>
              <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:8}}>Skipped duplicates ({preview.duplicates.length})</div>
              <div style={{maxHeight:130,overflowY:"auto",marginBottom:14,borderRadius:9,border:`1px solid rgba(217,119,6,0.2)`}}>
                {preview.duplicates.map((e,i)=>(
                  <div key={i} style={{padding:"8px 13px",borderBottom:i<preview.duplicates.length-1?`1px solid rgba(217,119,6,0.12)`:"none",display:"flex",alignItems:"center",gap:10,background:"rgba(217,119,6,0.03)"}}>
                    <span style={{width:6,height:6,borderRadius:"50%",background:T.amber,flexShrink:0,display:"inline-block"}}/>
                    <div style={{fontFamily:"'Barlow Condensed'",fontSize:12,fontWeight:700,letterSpacing:1.5,color:"rgba(252,211,77,0.6)"}}>{e.plate}</div>
                    <div style={{fontSize:10,color:"rgba(255,255,255,0.2)",marginLeft:"auto"}}>Already exists</div>
                  </div>
                ))}
              </div>
            </>}

            {/* Action buttons */}
            {preview.toAdd.length>0
              ?<BigBtn onClick={submitUpdate} color="green" style={{marginBottom:10}}>✅ CONFIRM & UPDATE DATABASE ({preview.toAdd.length} vehicles)</BigBtn>
              :<div style={{textAlign:"center",padding:"16px",fontSize:12,color:T.muted,background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderRadius:9,marginBottom:10}}>All vehicles in the file already exist in the database. Nothing new to add.</div>
            }
            <button onClick={resetUpload} style={{width:"100%",padding:"10px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,fontSize:13,cursor:"pointer",fontWeight:500}}>Cancel — Upload Different File</button>
          </>
        )}
      </>}
    </div>
    {confirm&&<ConfirmModal msg={`Remove "${confirm}" from master database? Existing log entries will not be affected.`} onYes={()=>removePlate(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── LOG VIEW ──────────────────────────────────────────────── */
function LogV({onBack,log,t_,onLogout,onSelect}){
  const [filter,setFilter]=useState("today");
  const filtered=filter==="today"?log.filter(e=>e.date===todayStr()):[...log];
  const sorted=[...filtered].reverse();
  const exportXL=()=>{
    if(!sorted.length){t_("No data","amber");return;}
    const ws=XLSX.utils.json_to_sheet(sorted.map(e=>({Plate:e.plate,Type:e.type,"Guest Name":e.guestName||"",Reason:e.guestReason||"",Officer:e.officer||"",Division:e.division||"","Logged By":e.loggedBy||"","Entry Time":fmt(e.entryTime),"Exit Time":e.exitTime?fmt(e.exitTime):"Still inside",Date:e.date})));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Log");XLSX.writeFile(wb,"JLN_Entry_Log.xlsx");t_("Exported!","green");
  };
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="ENTRY / EXIT LOG" sub={`${sorted.length} records`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{display:"flex",gap:7,marginBottom:13}}>
        {["today","all"].map(f=><button key={f} onClick={()=>setFilter(f)} style={{flex:1,padding:"8px 0",borderRadius:9,border:`1px solid ${filter===f?"rgba(232,160,32,0.5)":T.border}`,background:filter===f?"rgba(232,160,32,0.1)":"transparent",color:filter===f?T.goldL:T.muted,fontSize:12,cursor:"pointer",fontWeight:filter===f?700:400}}>{f==="today"?"Today":"All Time"}</button>)}
        <button onClick={exportXL} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${T.border}`,background:"rgba(255,255,255,0.04)",color:"rgba(255,255,255,0.6)",fontSize:11,cursor:"pointer",fontWeight:600}}>⬇ XLS</button>
      </div>
      {!sorted.length?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>No entries for this period</div>:
        sorted.map(e=><div key={e.id} onClick={()=>onSelect(e)} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${e.type==="guest"?T.amber:T.green}`,borderRadius:11,padding:"11px 13px",marginBottom:7,cursor:"pointer",transition:"background .15s"}}
          onMouseEnter={ev=>ev.currentTarget.style.background="rgba(255,255,255,0.055)"}
          onMouseLeave={ev=>ev.currentTarget.style.background="rgba(255,255,255,0.03)"}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:800,letterSpacing:2.5}}>{e.plate}</div>
            <div style={{display:"flex",gap:5}}><Pill color={e.type==="guest"?"amber":"green"}>{e.type==="guest"?"Guest":"Regular"}</Pill>{!e.exitTime&&<Pill color="blue">Inside</Pill>}</div>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:T.muted}}>
            <span>IN: {fmt(e.entryTime)}</span><span>{e.exitTime?`OUT: ${fmt(e.exitTime)}`:"Not exited"}</span>
          </div>
          {e.type==="guest"&&<div style={{fontSize:10,color:T.muted,marginTop:2}}>👤 {e.guestName||"—"}</div>}
        </div>)}
    </div>
  </div>;
}
function LogDetailV({entry,onBack,onLogout}){
  const rows=[["Plate",entry.plate],["Type",entry.type==="guest"?"Guest / Non-official":"Regular / Official"],["Entry Time",fmt(entry.entryTime)],["Exit Time",entry.exitTime?fmt(entry.exitTime):"Still inside"],["Date",entry.date],["Logged By",entry.loggedBy||"—"],...(entry.type==="guest"?[["Guest Name",entry.guestName||"—"],["Purpose",entry.guestReason||"—"]]:[]),(["Officer",entry.officer||"—"]),["Division",entry.division||"—"]];
  const isGuest=entry.type==="guest";
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="ENTRY DETAILS" sub="Log record" onBack={onBack}/>
    <div style={{padding:"15px 15px 0"}}>
      <div style={{background:isGuest?"rgba(217,119,6,0.08)":"rgba(22,163,74,0.08)",border:`1px solid ${isGuest?"rgba(217,119,6,0.3)":"rgba(22,163,74,0.3)"}`,borderRadius:14,padding:"16px",marginBottom:14,textAlign:"center"}}>
        <div style={{fontFamily:"'Barlow Condensed'",fontSize:28,fontWeight:900,letterSpacing:6,color:isGuest?T.amberT:T.greenT}}>{entry.plate}</div>
        <div style={{marginTop:7,display:"flex",gap:6,justifyContent:"center"}}><Pill color={isGuest?"amber":"green"}>{isGuest?"Guest Entry":"Regular Entry"}</Pill>{!entry.exitTime&&<Pill color="blue">Inside</Pill>}</div>
      </div>
      <Card>{rows.map((row,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:i<rows.length-1?`1px solid ${T.border}`:"none"}}><span style={{fontSize:10,color:T.muted,fontWeight:600,letterSpacing:.5,textTransform:"uppercase"}}>{row[0]}</span><span style={{fontSize:12,fontWeight:600,textAlign:"right",maxWidth:"58%",color:"rgba(255,255,255,0.85)"}}>{row[1]}</span></div>)}</Card>
    </div>
  </div>;
}

/* ─── USER MANAGEMENT ─────────────────────────────────────────── */
function UserMgmtV({onBack,users,setUsers,t_,onLogout,user}){
  const [showForm,setShowForm]=useState(false);
  const [username,setUsername]=useState(""); const [password,setPassword]=useState(""); const [role,setRole]=useState("staff");
  const [confirm,setConfirm]=useState(null);
  const addUser=()=>{
    if(!username.trim()){t_("Enter username","red");return;}
    if(password.length<4){t_("Password min 4 characters","red");return;}
    if(users.find(u=>u.username.toLowerCase()===username.toLowerCase())){t_("Username already exists","amber");return;}
    const u=[...users,{id:uid(),username:username.trim(),password,role,createdBy:user.username,createdOn:fmtDate(Date.now())}];
    setUsers(u);ss(K.users,u);setUsername("");setPassword("");setRole("staff");setShowForm(false);t_("User created ✓","green");
  };
  const deleteUser=id=>{
    if(users.find(u=>u.id===id)?.username===user.username){t_("Cannot delete your own account","red");return;}
    const u=users.filter(x=>x.id!==id);setUsers(u);ss(K.users,u);t_("User deleted","amber");setConfirm(null);
  };
  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="USER MANAGEMENT" sub={`${users.length} accounts`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <BigBtn onClick={()=>setShowForm(!showForm)} color="purple" style={{marginBottom:13}}>+ Create New Login</BigBtn>
      {showForm&&<Card>
        <SecTitle>NEW ACCOUNT</SecTitle>
        <FormField label="Username" required><input placeholder="e.g. gate_north" value={username} onChange={e=>setUsername(e.target.value)}/></FormField>
        <FormField label="Password" required><input type="password" placeholder="Min 4 characters" value={password} onChange={e=>setPassword(e.target.value)}/></FormField>
        <FormField label="Role">
          <select value={role} onChange={e=>setRole(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%"}}>
            <option value="staff">Security Staff</option>
            <option value="admin">Super Admin</option>
          </select>
        </FormField>
        <div style={{display:"flex",gap:8,marginTop:4}}>
          <button onClick={()=>setShowForm(false)} style={{flex:1,padding:"11px 0",borderRadius:9,border:`1px solid ${T.border}`,background:"transparent",color:T.muted,fontSize:13,cursor:"pointer",fontWeight:500}}>Cancel</button>
          <button onClick={addUser} style={{flex:1,padding:"11px 0",borderRadius:9,border:"none",background:T.purple,color:"white",fontSize:13,fontWeight:700,cursor:"pointer",letterSpacing:.3}}>Create Account</button>
        </div>
      </Card>}
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>All Accounts</div>
      {users.map(u=><div key={u.id} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${u.role==="admin"?T.gold:T.blue}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:30,height:30,borderRadius:8,background:u.role==="admin"?"rgba(232,160,32,0.1)":"rgba(37,99,235,0.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{u.role==="admin"?"◆":"◈"}</div>
        <div style={{flex:1}}>
          <div style={{fontFamily:"'Barlow Condensed'",fontSize:14,fontWeight:700}}>{u.username}{u.id===user.id?<span style={{fontSize:10,color:T.muted,fontWeight:400}}> · you</span>:null}</div>
          <div style={{fontSize:10,color:T.muted,marginTop:1}}>{u.role==="admin"?"Super Admin":"Staff"} · {u.createdOn||"—"}</div>
        </div>
        {u.id!==user.id&&<button onClick={()=>setConfirm(u)} style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.25)",color:"#fca5a5",borderRadius:7,padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:600}}>Delete</button>}
      </div>)}
    </div>
    {confirm&&<ConfirmModal msg={`Delete user "${confirm.username}"? This cannot be undone.`} onYes={()=>deleteUser(confirm.id)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── BANNED VEHICLES ─────────────────────────────────────────── */
function BannedV({onBack,banned,setBanned,t_,onLogout,user}){
  const [plate,setPlate]=useState(""); const [reason,setReason]=useState(""); const [confirm,setConfirm]=useState(null);
  const addBan=()=>{
    if(!plate.trim()){t_("Enter plate number","red");return;}
    if(banned.find(b=>norm(b.plate)===norm(plate))){t_("Already banned","amber");return;}
    const u=[...banned,{plate:plate.toUpperCase(),reason:reason.trim(),bannedBy:user.username,bannedOn:fmtDate(Date.now())}];
    setBanned(u);ss(K.banned,u);setPlate("");setReason("");t_("Vehicle banned","red");
  };
  const removeBan=p=>{const u=banned.filter(b=>norm(b.plate)!==norm(p));setBanned(u);ss(K.banned,u);t_("Ban lifted","green");setConfirm(null);};
  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="BANNED VEHICLES" sub={`${banned.length} banned`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <SecTitle>BAN A VEHICLE</SecTitle>
        <FormField label="Registration Number" required>
          <input placeholder="DL 01 AB 1234" value={plate} onChange={e=>setPlate(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed',monospace",fontSize:20,fontWeight:800,letterSpacing:4,textAlign:"center",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:10,color:"white",padding:"12px",width:"100%",outline:"none"}}/>
        </FormField>
        <FormField label="Reason for Ban"><input placeholder="e.g. Security threat, Unauthorized entry" value={reason} onChange={e=>setReason(e.target.value)}/></FormField>
        <BigBtn onClick={addBan} color="red" style={{marginTop:4}}>⛔ ADD TO BANNED LIST</BigBtn>
      </Card>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:2,textTransform:"uppercase",color:"rgba(255,255,255,0.25)",marginBottom:10}}>Banned Vehicles ({banned.length})</div>
      {!banned.length?<div style={{textAlign:"center",color:T.muted,padding:"30px 0",fontSize:12}}>No vehicles on banned list</div>:
        banned.slice().reverse().map((b,i)=><div key={i} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${T.red}`,borderRadius:11,padding:"11px 13px",marginBottom:7,display:"flex",alignItems:"center",gap:10}}>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:2.5}}>{b.plate}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:2}}>{b.reason||"No reason given"}{b.bannedOn?` · ${b.bannedOn}`:""}</div>
          </div>
          <button onClick={()=>setConfirm(b.plate)} style={{background:"rgba(22,163,74,0.1)",border:"1px solid rgba(22,163,74,0.25)",color:"#86efac",borderRadius:7,padding:"4px 10px",fontSize:10,cursor:"pointer",fontWeight:600}}>Unban</button>
        </div>)}
    </div>
    {confirm&&<ConfirmModal msg={`Remove ban on "${confirm}"?`} onYes={()=>removeBan(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;
}

/* ─── CHANGE PASSWORD ─────────────────────────────────────────── */
function ChangePwdV({onBack,users,setUsers,t_,onLogout,user}){
  const [selId,setSelId]=useState(user.id);
  const [newPwd,setNewPwd]=useState(""); const [confirmPwd,setConfirmPwd]=useState(""); const [currentPwd,setCurrentPwd]=useState("");
  const isSelf=selId===user.id;
  const save=()=>{
    if(newPwd.length<4){t_("Minimum 4 characters","red");return;}
    if(newPwd!==confirmPwd){t_("Passwords don't match","red");return;}
    if(isSelf){const me=users.find(u=>u.id===user.id);if(me.password!==currentPwd){t_("Current password incorrect","red");return;}}
    const u=users.map(x=>x.id===selId?{...x,password:newPwd}:x);
    setUsers(u);ss(K.users,u);setNewPwd("");setConfirmPwd("");setCurrentPwd("");t_("Password updated ✓","green");
  };
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="CHANGE PASSWORD" sub="Update credentials" onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <SecTitle>SELECT ACCOUNT</SecTitle>
        <FormField label="Account">
          <select value={selId} onChange={e=>setSelId(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:9,color:"white",padding:"11px 13px",fontSize:14,width:"100%",marginBottom:4}}>
            {users.map(u=><option key={u.id} value={u.id}>{u.username} ({u.role}){u.id===user.id?" — You":""}</option>)}
          </select>
        </FormField>
        <SecTitle>NEW PASSWORD</SecTitle>
        {isSelf&&<FormField label="Current Password" required><input type="password" placeholder="Enter current password" value={currentPwd} onChange={e=>setCurrentPwd(e.target.value)}/></FormField>}
        <FormField label="New Password" required><input type="password" placeholder="Min 4 characters" value={newPwd} onChange={e=>setNewPwd(e.target.value)}/></FormField>
        <FormField label="Confirm New Password" required><input type="password" placeholder="Repeat new password" value={confirmPwd} onChange={e=>setConfirmPwd(e.target.value)}/></FormField>
        {!isSelf&&<div style={{background:"rgba(232,160,32,0.07)",border:"1px solid rgba(232,160,32,0.2)",borderRadius:9,padding:"9px 12px",fontSize:11,color:T.goldL,marginBottom:12,lineHeight:1.5}}>⚠ Changing another user's password — no current password required.</div>}
        <BigBtn onClick={save} color="amber">🔑 UPDATE PASSWORD</BigBtn>
      </Card>
    </div>
  </div>;
}

/* ─── REPORT DOWNLOAD ─────────────────────────────────────────── */
function ReportV({onBack,log,t_,onLogout}){
  const today=new Date().toISOString().split("T")[0];
  const [from,setFrom]=useState(today); const [to,setTo]=useState(today); const [type,setType]=useState("all");
  const getFiltered=()=>{
    const f=new Date(from);f.setHours(0,0,0,0);const t2=new Date(to);t2.setHours(23,59,59,999);
    return log.filter(e=>{const d=new Date(e.entryTime);if(d<f||d>t2)return false;if(type==="regular"&&e.type!=="regular")return false;if(type==="guest"&&e.type!=="guest")return false;return true;});
  };
  const preview=getFiltered();
  const downloadReport=()=>{
    const rows=getFiltered();if(!rows.length){t_("No data in selected range","amber");return;}
    const wb=XLSX.utils.book_new();
    const ws1=XLSX.utils.json_to_sheet(rows.map(e=>({Date:e.date,"Reg Number":e.plate,Type:e.type==="guest"?"Guest":"Official","Guest Name":e.guestName||"","Purpose":e.guestReason||"","Officer/Auth By":e.officer||"","Division":e.division||"","Logged By":e.loggedBy||"","Entry Time":fmt(e.entryTime),"Exit Time":e.exitTime?fmt(e.exitTime):"Not recorded"})));
    XLSX.utils.book_append_sheet(wb,ws1,"Entry-Exit Log");
    const total=rows.length,reg=rows.filter(r=>r.type==="regular").length,guest=rows.filter(r=>r.type==="guest").length;
    const ws2=XLSX.utils.json_to_sheet([{Metric:"Total Entries",Value:total},{Metric:"Official Vehicles",Value:reg},{Metric:"Guest Vehicles",Value:guest},{Metric:"Not Exited",Value:rows.filter(r=>!r.exitTime).length},{Metric:"Date From",Value:from},{Metric:"Date To",Value:to}]);
    XLSX.utils.book_append_sheet(wb,ws2,"Summary");
    XLSX.writeFile(wb,`JLN_Report_${from}_to_${to}.xlsx`);t_("Report downloaded ✓","green");
  };
  return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="DOWNLOAD REPORT" sub="Custom date range" onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <SecTitle>DATE RANGE</SecTitle>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
          <div><div style={{fontSize:9,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:"rgba(255,255,255,0.3)",marginBottom:6}}>From</div><input type="date" value={from} onChange={e=>setFrom(e.target.value)} style={{colorScheme:"dark"}}/></div>
          <div><div style={{fontSize:9,fontWeight:700,letterSpacing:1.5,textTransform:"uppercase",color:"rgba(255,255,255,0.3)",marginBottom:6}}>To</div><input type="date" value={to} onChange={e=>setTo(e.target.value)} style={{colorScheme:"dark"}}/></div>
        </div>
        <SecTitle>VEHICLE TYPE</SecTitle>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7,marginBottom:14}}>
          {[["all","All"],["regular","Official"],["guest","Guest"]].map(([v,l])=><button key={v} onClick={()=>setType(v)} style={{padding:"8px 0",borderRadius:9,border:`1px solid ${type===v?"rgba(232,160,32,0.5)":T.border}`,background:type===v?"rgba(232,160,32,0.1)":"transparent",color:type===v?T.goldL:T.muted,fontSize:12,cursor:"pointer",fontWeight:type===v?700:400}}>{l}</button>)}
        </div>
        <div style={{background:"rgba(255,255,255,0.04)",borderRadius:9,padding:"12px 14px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:12,color:T.muted}}>Records in range</span>
          <span style={{fontFamily:"'Barlow Condensed'",fontSize:24,fontWeight:800,color:T.goldL}}>{preview.length}</span>
        </div>
        <BigBtn onClick={downloadReport} color="cyan">📥 DOWNLOAD EXCEL</BigBtn>
      </Card>
      <div style={{fontSize:10,color:T.muted,lineHeight:1.75,padding:"0 2px"}}>Report includes 2 sheets: full entry/exit log with all details, and a summary with totals.</div>
    </div>
  </div>;
}

/* ─── COME & PLAY DATABASE ──────────────────────────────────── */
function ComePlayV({onBack,onLogout,t_,master,setMaster}){
  const [members,setMembers]=useState([]);
  const [tab,setTab]=useState("list"); // "list" | "add" | "detail"
  const [sel,setSel]=useState(null);
  const [search,setSearch]=useState("");
  const [confirm,setConfirm]=useState(null);
  const idRef=useRef(); const cardRef=useRef();

  // form state
  const [name,setName]=useState(""); const [phone,setPhone]=useState("");
  const [carNo,setCarNo]=useState(""); const [bookingDetails,setBookingDetails]=useState("");
  const [idPreview,setIdPreview]=useState(null); const [cardPreview,setCardPreview]=useState(null);

  useEffect(()=>{gs(K.cnp).then(d=>d&&setMembers(d));},[]);

  const readFile=async(file,cb)=>{
    const reader=new FileReader();
    reader.onload=e=>cb(e.target.result);
    reader.readAsDataURL(file);
  };

  const save=async()=>{
    if(!name.trim()||!phone.trim()||!carNo.trim()){t_("Name, phone & car number required","red");return;}
    const entry={id:"CP-"+Math.random().toString(36).slice(2,7).toUpperCase(),name:name.trim(),phone:phone.trim(),carNo:carNo.toUpperCase().trim(),bookingDetails:bookingDetails.trim(),idPreview,cardPreview,addedOn:fmtDate(Date.now())};
    const u=[...members,entry];setMembers(u);await ss(K.cnp,u);
    // Add car to master database if not already there
    if(!master.find(e=>norm(e.plate)===norm(carNo))){
      const nm=[...master,{plate:carNo.toUpperCase().trim(),officer:name.trim(),division:"Come & Play",addedOn:fmtDate(Date.now()),addedBy:"Come & Play System",isComePlay:true}];
      setMaster(nm);ss(K.master,nm);
      t_("Member added + car registered in master ✓","green");
    } else {
      t_("Member added ✓","green");
    }
    setName("");setPhone("");setCarNo("");setBookingDetails("");setIdPreview(null);setCardPreview(null);
    setTab("list");
  };

  const remove=id=>{
    const u=members.filter(m=>m.id!==id);setMembers(u);ss(K.cnp,u);t_("Record removed","amber");setConfirm(null);setSel(null);setTab("list");
  };

  const filtered=search.trim()?members.filter(m=>m.name.toLowerCase().includes(search.toLowerCase())||norm(m.carNo).includes(norm(search))||m.phone.includes(search)):members;

  if(tab==="detail"&&sel) return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="MEMBER DETAILS" sub={sel.id} onBack={()=>{setTab("list");setSel(null);}}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:22,fontWeight:800,color:T.goldL,letterSpacing:.5}}>{sel.name}</div>
            <div style={{fontSize:11,color:T.muted,marginTop:2}}>{sel.id} · Added {sel.addedOn}</div>
          </div>
          <Pill color="purple">Come & Play</Pill>
        </div>
        {[["📞 Phone",sel.phone],["🚗 Car Number",sel.carNo],["📋 Booking Details",sel.bookingDetails||"—"]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${T.border}`,alignItems:"flex-start",gap:10}}>
            <span style={{fontSize:11,color:T.muted,flexShrink:0}}>{l}</span>
            <span style={{fontSize:12,fontWeight:600,textAlign:"right"}}>{v}</span>
          </div>
        ))}
      </Card>
      {sel.idPreview&&<Card><SecTitle>ID PROOF</SecTitle><img src={sel.idPreview} alt="ID" style={{width:"100%",borderRadius:9,maxHeight:180,objectFit:"cover"}}/></Card>}
      {sel.cardPreview&&<Card><SecTitle>SAI BOOKING CARD</SecTitle><img src={sel.cardPreview} alt="SAI Card" style={{width:"100%",borderRadius:9,maxHeight:180,objectFit:"cover"}}/></Card>}
      <button onClick={()=>setConfirm(sel.id)} style={{width:"100%",padding:"11px 0",borderRadius:9,border:"1px solid rgba(220,38,38,0.3)",background:"rgba(220,38,38,0.08)",color:"#fca5a5",fontSize:13,cursor:"pointer",fontWeight:600,marginBottom:20}}>🗑 Remove This Record</button>
    </div>
    {confirm&&<ConfirmModal msg={`Remove ${sel.name} from Come & Play database?`} onYes={()=>remove(confirm)} onNo={()=>setConfirm(null)}/>}
  </div>;

  if(tab==="add") return <div style={{minHeight:600,background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="ADD MEMBER" sub="Come & Play database" onBack={()=>setTab("list")}/>
    <div style={{padding:"14px 15px 0"}}>
      <Card>
        <SecTitle>PERSONAL DETAILS</SecTitle>
        <FormField label="Full Name" required><input placeholder="e.g. Rahul Sharma" value={name} onChange={e=>setName(e.target.value)}/></FormField>
        <FormField label="Phone Number" required><input placeholder="10-digit mobile number" value={phone} onChange={e=>setPhone(e.target.value)} type="tel" inputMode="numeric" maxLength={15}/></FormField>
        <FormField label="Car / Vehicle Number" required><input placeholder="DL 01 AB 1234" value={carNo} onChange={e=>setCarNo(e.target.value.toUpperCase())} style={{fontFamily:"'Barlow Condensed'",fontSize:18,fontWeight:700,letterSpacing:3,textAlign:"center"}}/></FormField>
      </Card>
      <Card>
        <SecTitle>BOOKING DETAILS</SecTitle>
        <FormField label="Booking Details"><input placeholder="e.g. Court 3, Badminton, Mon-Fri 6-8am" value={bookingDetails} onChange={e=>setBookingDetails(e.target.value)}/></FormField>
      </Card>
      <Card>
        <SecTitle>ID PROOF</SecTitle>
        <div style={{fontSize:12,color:T.muted,marginBottom:10,lineHeight:1.6}}>Upload Aadhaar, PAN, Driving Licence etc.</div>
        {idPreview?<div><img src={idPreview} alt="ID" style={{width:"100%",borderRadius:8,maxHeight:150,objectFit:"cover",marginBottom:8}}/>
          <button onClick={()=>setIdPreview(null)} style={{width:"100%",padding:"7px",borderRadius:7,border:"1px solid rgba(220,38,38,0.3)",background:"rgba(220,38,38,0.08)",color:"#fca5a5",fontSize:11,cursor:"pointer"}}>✕ Remove</button>
        </div>:
        <button onClick={()=>idRef.current.click()} style={{width:"100%",padding:"20px",borderRadius:9,border:`1px dashed ${T.border}`,background:"rgba(255,255,255,0.02)",color:T.muted,fontSize:13,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:6}}>
          <span style={{fontSize:24}}>📎</span><span>Tap to upload ID card</span>
        </button>}
        <input ref={idRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(f)readFile(f,setIdPreview);}}/>
      </Card>
      <Card>
        <SecTitle>SAI BOOKING CARD</SecTitle>
        <div style={{fontSize:12,color:T.muted,marginBottom:10,lineHeight:1.6}}>Upload the card issued by SAI for this member.</div>
        {cardPreview?<div><img src={cardPreview} alt="SAI Card" style={{width:"100%",borderRadius:8,maxHeight:150,objectFit:"cover",marginBottom:8}}/>
          <button onClick={()=>setCardPreview(null)} style={{width:"100%",padding:"7px",borderRadius:7,border:"1px solid rgba(220,38,38,0.3)",background:"rgba(220,38,38,0.08)",color:"#fca5a5",fontSize:11,cursor:"pointer"}}>✕ Remove</button>
        </div>:
        <button onClick={()=>cardRef.current.click()} style={{width:"100%",padding:"20px",borderRadius:9,border:`1px dashed ${T.border}`,background:"rgba(255,255,255,0.02)",color:T.muted,fontSize:13,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:6}}>
          <span style={{fontSize:24}}>🪪</span><span>Tap to upload SAI card</span>
        </button>}
        <input ref={cardRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(f)readFile(f,setCardPreview);}}/>
      </Card>
      <BigBtn onClick={save} color="purple" style={{marginBottom:20}}>✅ SAVE TO COME & PLAY DATABASE</BigBtn>
    </div>
  </div>;

  return <div style={{minHeight:600,position:"relative",background:T.navy}}>
    <Hdr role="admin" onLogout={onLogout} title="COME & PLAY" sub={`${members.length} members registered`} onBack={onBack}/>
    <div style={{padding:"14px 15px 0"}}>
      <div style={{background:"rgba(124,58,237,0.07)",border:"1px solid rgba(124,58,237,0.25)",borderRadius:10,padding:"10px 14px",marginBottom:13,fontSize:11,color:T.purpleT,lineHeight:1.6}}>
        🏸 Members who have booked stadium facilities via SAI. Only registered members are listed here.
      </div>
      <input placeholder="🔍  Search by name, car number, phone…" value={search} onChange={e=>setSearch(e.target.value)} style={{background:"rgba(255,255,255,0.05)",border:`1px solid ${search?T.purple:T.border}`,borderRadius:9,color:"white",padding:"10px 14px",fontSize:13,width:"100%",outline:"none",marginBottom:12,transition:"border-color .2s"}}/>
      <BigBtn onClick={()=>setTab("add")} color="purple" style={{marginBottom:14}}>+ ADD NEW MEMBER</BigBtn>
      {search&&<div style={{fontSize:10,color:T.muted,marginBottom:8}}>{filtered.length} result{filtered.length!==1?"s":""}</div>}
      {!filtered.length?<div style={{textAlign:"center",color:T.muted,padding:"50px 0",fontSize:12}}>{search?"No members match your search":"No members yet — add one above"}</div>:
        filtered.slice().reverse().map(m=><div key={m.id} onClick={()=>{setSel(m);setTab("detail");}} style={{background:"rgba(255,255,255,0.03)",border:`1px solid ${T.border}`,borderLeft:`2px solid ${T.purple}`,borderRadius:11,padding:"12px 14px",marginBottom:8,cursor:"pointer",display:"flex",alignItems:"center",gap:11,transition:"background .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(124,58,237,0.07)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.03)"}>
          <div style={{width:36,height:36,borderRadius:9,background:"rgba(124,58,237,0.15)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>🏸</div>
          <div style={{flex:1}}>
            <div style={{fontFamily:"'Barlow Condensed'",fontSize:15,fontWeight:800,letterSpacing:.3,color:"white"}}>{m.name}</div>
            <div style={{fontSize:10,color:T.muted,marginTop:2}}>{m.carNo} · {m.phone}</div>
            {m.bookingDetails&&<div style={{fontSize:10,color:T.purpleT,marginTop:2,opacity:.8}}>{m.bookingDetails}</div>}
          </div>
          <div style={{color:"rgba(255,255,255,0.2)",fontSize:14}}>›</div>
        </div>)
      }
    </div>
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
      input,select,textarea{
        background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);
        border-radius:9px;color:white;padding:11px 13px;font-size:14px;width:100%;
        outline:none;transition:border-color .2s,background .2s;
        font-family:inherit;
      }
      input::placeholder{color:rgba(255,255,255,0.25);}
      input:focus,select:focus{border-color:rgba(232,160,32,0.5);background:rgba(255,255,255,0.07);}
      select option{background:#0d1a30;color:white;}
      input[type="date"]{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:9px;color:white;padding:11px 13px;width:100%;outline:none;font-size:13px;}
      ::-webkit-scrollbar{width:4px;}
      ::-webkit-scrollbar-track{background:transparent;}
      ::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.1);border-radius:4px;}
    `;
    document.head.appendChild(s);
    return()=>document.head.removeChild(s);
  },[]);
  const [screen,setScreen]=useState("home");
  const [session,setSession]=useState(null);
  const [master,setMaster]=useState([]);
  const [log,setLog]=useState([]);
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [banned,setBanned]=useState([]);
  const [passes,setPasses]=useState([]);
  const [cnp,setCnp]=useState([]);

  // ── Real-time sync from Firebase ──────────────────────────────
  useEffect(()=>{
    // Initial load for all keys
    gs(K.master).then(d=>d&&setMaster(d));
    gs(K.log).then(d=>d&&setLog(d));
    gs(K.banned).then(d=>d&&setBanned(d));
    gs(K.passes).then(d=>d&&setPasses(d));
    gs(K.users).then(d=>d&&d.length&&setUsers(d));
    gs(K.cnp).then(d=>d&&setCnp(d));

    // Poll Firebase every 10 seconds for live sync across devices
    const interval=setInterval(()=>{
      gs(K.passes).then(d=>d&&setPasses(d));
      gs(K.master).then(d=>d&&setMaster(d));
      gs(K.log).then(d=>d&&setLog(d));
      gs(K.cnp).then(d=>d&&setCnp(d));
      gs(K.banned).then(d=>d&&setBanned(d));
    },10000);
    return()=>clearInterval(interval);
  },[]);

  const doLogin=u=>{setSession(u);setScreen("app");};
  const doLogout=()=>{setSession(null);setScreen("home");};
  const shared={master,setMaster,log,setLog,users,setUsers,banned,setBanned,passes,setPasses,cnp,setCnp};

  return <>
    {screen==="home"&&<HomeScreen onSelectRole={r=>setScreen(r==="admin"?"loginAdmin":"loginStaff")} onRequestPass={()=>setScreen("requestPass")}/>}
    {screen==="loginAdmin"&&<LoginScreen role="admin" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="loginStaff"&&<LoginScreen role="staff" users={users} onLogin={doLogin} onBack={()=>setScreen("home")}/>}
    {screen==="requestPass"&&<RequestPassScreen onBack={()=>setScreen("home")} passes={passes} setPasses={setPasses}/>}
    {screen==="app"&&session?.role==="staff"&&<StaffApp onLogout={doLogout} master={master} log={log} setLog={setLog} banned={banned} passes={passes} setPasses={setPasses} cnp={cnp} user={session}/>}
    {screen==="app"&&session?.role==="admin"&&<AdminApp onLogout={doLogout} {...shared} user={session}/>}
  </>;
}
