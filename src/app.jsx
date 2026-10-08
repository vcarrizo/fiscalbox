import React,{useState,useMemo,useCallback,useEffect} from"react";import{createRoot}from"react-dom/client";import*as XLSX from"xlsx";import{BarChart,Bar,XAxis,YAxis,Tooltip,Legend,ResponsiveContainer,ReferenceLine}from"recharts";
const MONTHS=["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"],DB_KEY="iva_studio_db";
function FBLogo({size=24}){return(<svg width={size} height={size} viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="2" width="36" height="36" rx="8" fill="#6C9CFF" fillOpacity=".15" stroke="#6C9CFF" strokeWidth="2"/><rect x="8" y="8" width="12" height="12" rx="2" fill="#6C9CFF"/><rect x="22" y="8" width="10" height="5" rx="1.5" fill="#4ADE80"/><rect x="22" y="15" width="10" height="5" rx="1.5" fill="#4ADE80" fillOpacity=".5"/><rect x="8" y="22" width="24" height="4" rx="1.5" fill="#6C9CFF" fillOpacity=".4"/><rect x="8" y="28" width="18" height="4" rx="1.5" fill="#6C9CFF" fillOpacity=".25"/></svg>);}
function TopBar({onHome}){return(<div style={{display:"flex",alignItems:"center",gap:10,marginBottom:24,paddingBottom:16,borderBottom:"1px solid #1a1a30",cursor:onHome?"pointer":"default"}} onClick={onHome}><FBLogo size={32}/><div><div style={{fontSize:18,fontWeight:800,letterSpacing:"-0.5px",lineHeight:1}}><span style={{color:"#6C9CFF"}}>Fiscal</span><span style={{color:"#4ADE80"}}>Box</span></div><div style={{fontSize:10,color:"#555",letterSpacing:"1.5px",textTransform:"uppercase",marginTop:2}}>Estudio Contable</div></div></div>);}
function loadDB(){try{return JSON.parse(localStorage.getItem(DB_KEY))||{clients:{}};}catch{return{clients:{}};}}
function saveDB(db){localStorage.setItem(DB_KEY,JSON.stringify(db));}
function parseDate(v){if(!v)return null;if(v instanceof Date&&!isNaN(v))return v;if(typeof v==="number"){const d=XLSX.SSF.parse_date_code(v);if(d)return new Date(d.y,d.m-1,d.d);}const s=String(v).trim();const m1=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);if(m1)return new Date(+m1[3],+m1[2]-1,+m1[1]);const m2=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);if(m2)return new Date(+m2[1],+m2[2]-1,+m2[3]);const d=new Date(s);return isNaN(d.getTime())?null:d;}
function findCol(H,P){const h=H.map(x=>String(x||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim());for(const p of P){const i=h.findIndex(x=>x.includes(p));if(i>=0)return i;}return-1;}
function parseNum(v){if(v==null||v==="")return 0;if(typeof v==="number")return v;return parseFloat(String(v).replace(/\./g,"").replace(",","."))||0;}
const NC_P=["nota de credito","nota de crédito","nc ","notas de credito","notas de crédito"];
const NC_CODES=new Set(["3","8","13","53","203","208","213"]);
function isNotaCredito(t){if(!t)return false;const s=String(t).trim();if(NC_CODES.has(s))return true;const l=s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");return NC_P.some(n=>l.includes(n));}
function makeDK(r){const tc=String(r.tipo||"").trim().replace(/^(\d+).*$/,"$1");return`${tc}|${r.ptoVenta}|${r.numDesde}|${r.cuit}`.toLowerCase().trim();}
function fmt(n){return new Intl.NumberFormat("es-AR",{style:"currency",currency:"ARS",minimumFractionDigits:2}).format(n);}
function fmtShort(n){const a=Math.abs(n);if(a>=1e6)return`$${(n/1e6).toFixed(1)}M`;if(a>=1e3)return`$${(n/1e3).toFixed(0)}k`;return`$${n.toFixed(0)}`;}
function pKey(y,m){return`${y}-${String(m).padStart(2,"0")}`;}
function pLabel(k){const[y,m]=k.split("-");return`${MONTHS[+m-1]} ${y}`;}

function parseArcaCSV(text){const lines=text.split("\n").map(l=>l.replace(/\$\s*$/,"").replace(/\r$/,"")).filter(l=>l.trim());if(!lines.length)return[];const semi=lines[0].split(";").length,comma=lines[0].split(",").length;const delim=semi>comma?";":",";return lines.map(line=>{const result=[];let cur="",inQ=false;for(let i=0;i<line.length;i++){const ch=line[i];if(ch==='"'){if(inQ&&line[i+1]==='"'){cur+='"';i++;}else inQ=!inQ;}else if(ch===delim&&!inQ){result.push(cur.trim());cur="";}else cur+=ch;}result.push(cur.trim());return result;});}

function processFile(data,fileType,sourceFile){
  let sheets=[];const head=new Uint8Array(data.slice(0,4));const isCSV=!(head[0]===0x50&&head[1]===0x4B)&&!(head[0]===0xD0&&head[1]===0xCF);
  if(isCSV){let text;try{text=new TextDecoder("utf-8").decode(data);}catch{text=new TextDecoder("latin1").decode(data);}const json=parseArcaCSV(text);if(json.length>=2)sheets.push(json);}
  else{const wb=XLSX.read(data,{type:"array",cellDates:true});for(const sn of wb.SheetNames){const s=wb.Sheets[sn];const j=XLSX.utils.sheet_to_json(s,{header:1,defval:""});if(j.length>=2)sheets.push(j);}}
  const allRows=[];for(const json of sheets){if(json.length<2)continue;let hi=0;
    for(let i=0;i<Math.min(json.length,10);i++){const row=json[i].map(x=>String(x||"").toLowerCase());if(row.some(c=>c.includes("fecha"))&&row.some(c=>c.includes("tipo")||c.includes("comprobante"))){hi=i;break;}}
    const H=json[hi],dc=findCol(H,["fecha"]),tc=findCol(H,["tipo"]),pvc=findCol(H,["punto de venta"]),ndc=findCol(H,["numero desde","número desde"]);
    const nc=findCol(H,["neto gravado total","neto grav. total","imp. neto gravado","neto gravado"]),ic=findCol(H,["total iva"]),ngc=findCol(H,["neto no gravado","neto no grav"]),ec=findCol(H,["op. exentas","exentas","imp. op. exentas"]);
    const tlc=findCol(H,["imp. total"]);
    // For emitidos: use receptor fields. For recibidos: use emisor fields.
    const dnc=fileType==="recibidos"?findCol(H,["denominacion emisor","denominacion","razon social"]):findCol(H,["denominacion receptor","denominacion","razon social"]);
    const cc=fileType==="recibidos"?findCol(H,["nro. doc. emisor","nro. doc"]):findCol(H,["nro. doc. receptor","nro. doc"]);
    const tcc=findCol(H,["tipo cambio"]);
    const tdrc=fileType==="emitidos"?findCol(H,["tipo doc. receptor","tipo doc receptor"]):findCol(H,["tipo doc. emisor","tipo doc emisor"]);
    if(dc<0)continue;const i21=findCol(H,["iva 21%"]),i105=findCol(H,["iva 10,5%","iva 10.5%"]),i27=findCol(H,["iva 27%"]);
    for(let i=hi+1;i<json.length;i++){const r=json[i];if(!r||r.length===0||r.every(c=>!c&&c!==0))continue;const fecha=parseDate(r[dc]);if(!fecha)continue;
      const tcVal=tcc>=0?parseNum(r[tcc]):1;const tcM=tcVal>0?tcVal:1;
      const row={fecha:fecha.toISOString(),tipo:tc>=0?String(r[tc]||""):"",ptoVenta:pvc>=0?String(r[pvc]||""):"",numDesde:ndc>=0?String(r[ndc]||""):"",
        neto:(nc>=0?parseNum(r[nc]):0)*tcM,iva:(ic>=0?parseNum(r[ic]):0)*tcM,iva21:(i21>=0?parseNum(r[i21]):0)*tcM,iva105:(i105>=0?parseNum(r[i105]):0)*tcM,iva27:(i27>=0?parseNum(r[i27]):0)*tcM,
        noGrav:(ngc>=0?parseNum(r[ngc]):0)*tcM,exentas:(ec>=0?parseNum(r[ec]):0)*tcM,total:(tlc>=0?parseNum(r[tlc]):0)*tcM,
        denom:dnc>=0?String(r[dnc]||""):"",cuit:cc>=0?String(r[cc]||""):"",tipoDocRec:tdrc>=0?String(r[tdrc]||""):"",moneda:"$",tc:tcM,
        fileType,isNC:isNotaCredito(tc>=0?r[tc]:""),sourceFile,key:pKey(fecha.getFullYear(),fecha.getMonth()+1)};
      row.dedupKey=makeDK(row);allRows.push(row);}}return allRows;}

// Monotributo
const MONO_SCALES=[{cat:"A",tope:10277988.13},{cat:"B",tope:15058447.71},{cat:"C",tope:21113696.52},{cat:"D",tope:26212853.42},{cat:"E",tope:30833964.37},{cat:"F",tope:38642048.36},{cat:"G",tope:46211109.37},{cat:"H",tope:70113407.33},{cat:"I",tope:78479211.62},{cat:"J",tope:89872640.30},{cat:"K",tope:108357084.05}];
function getMonoCat(f){return MONO_SCALES.find(s=>f<=s.tope)||null;}

function MonotributoPanel({client,emitidos,onIIBB,onSelectPeriod}){
  const cat=client.monoCat||"A";const catData=MONO_SCALES.find(s=>s.cat===cat);
  const mb=useMemo(()=>{const map={};(emitidos||[]).forEach(r=>{const s=r.isNC?-1:1;if(!map[r.key])map[r.key]=0;map[r.key]+=r.total*s;});return Object.entries(map).sort(([a],[b])=>a.localeCompare(b)).map(([key,total])=>({key,label:pLabel(key),total}));},[emitidos]);
  const rolling=useMemo(()=>mb.map(m=>{let sum=0;const td=new Date(m.key+"-01");for(const x of mb){const xd=new Date(x.key+"-01");const diff=(td.getFullYear()-xd.getFullYear())*12+(td.getMonth()-xd.getMonth());if(diff>=0&&diff<12)sum+=x.total;}return{...m,rolling12:sum};}),[mb]);
  const latest=rolling.length?rolling[rolling.length-1]:null;const fa=latest?latest.rolling12:0;
  const pct=catData?(fa/catData.tope)*100:0;const rec=getMonoCat(fa);const exceeded=catData&&fa>catData.tope;const near=pct>80&&!exceeded;
  const avg=mb.length?mb.reduce((s,m)=>s+m.total,0)/mb.length:0;
  return(<div style={{background:"#12122a",borderRadius:10,padding:"18px",marginBottom:20,border:exceeded?"1px solid #F8717144":near?"1px solid #FBBF2444":"1px solid transparent"}}>
    <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>📋 Control Monotributo — Cat. {cat}</div>
    <div style={{marginBottom:16}}><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}><span>Fact. 12 meses: <strong style={{color:exceeded?"#F87171":near?"#FBBF24":"#4ADE80"}}>{fmt(fa)}</strong></span><span style={{color:"#777"}}>Tope: {catData?fmt(catData.tope):"—"}</span></div>
      <div style={{background:"#0a0a14",borderRadius:6,height:22,overflow:"hidden",position:"relative"}}><div style={{height:"100%",borderRadius:6,background:exceeded?"#F87171":near?"#FBBF24":"#4ADE80",width:`${Math.min(pct,100)}%`,transition:"width .5s",opacity:.8}}/><div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700}}>{pct.toFixed(1)}%</div></div></div>
    {exceeded&&<div style={{background:"#F871711A",borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13}}><strong style={{color:"#F87171"}}>⚠ Tope superado</strong><span style={{color:"#F8717199",marginLeft:8}}>Exceso: {fmt(fa-catData.tope)}</span>{rec?<span style={{color:"#FBBF24",marginLeft:8}}>→ Cat. {rec.cat}</span>:<span style={{color:"#F87171",marginLeft:8}}>→ Pasar a RI</span>}</div>}
    {near&&!exceeded&&<div style={{background:"#FBBF241A",borderRadius:8,padding:"10px 14px",marginBottom:12,fontSize:13,color:"#FBBF24"}}>⚡ Margen restante: {fmt(catData.tope-fa)}</div>}
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}} className="grid-3">
      <div style={{background:"#0a0a14",borderRadius:8,padding:"12px"}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Promedio mensual</div><div style={{fontSize:16,fontWeight:700,color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(avg)}</div></div>
      <div style={{background:"#0a0a14",borderRadius:8,padding:"12px"}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Proyección 12m</div><div style={{fontSize:16,fontWeight:700,color:"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{fmt(avg*12)}</div></div>
      <div style={{background:"#0a0a14",borderRadius:8,padding:"12px"}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Margen</div><div style={{fontSize:16,fontWeight:700,color:exceeded?"#F87171":"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{catData?fmt(catData.tope-fa):"—"}</div></div>
    </div>
    {mb.length>0&&<details><summary style={{fontSize:12,fontWeight:600,cursor:"pointer",color:"#888",marginBottom:8}}>Detalle mensual ({mb.length} meses)</summary>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr style={{borderBottom:"1px solid #2a2a40"}}>{["Período","Fact. Mes","Acum. 12m","% Tope","",""].map(h=>(<th key={h} style={{padding:"6px 8px",textAlign:h==="Período"||h===""?"left":"right",color:"#666",fontWeight:500,fontSize:10,textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{rolling.map(m=>{const p=catData?(m.rolling12/catData.tope)*100:0;return(<tr key={m.key} onClick={()=>onSelectPeriod(m.key)} style={{borderBottom:"1px solid #1a1a30",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.background="#1e1e40"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}><td style={{padding:"6px 8px",fontWeight:600}}>{m.label}</td><td style={{padding:"6px 8px",textAlign:"right",color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(m.total)}</td><td style={{padding:"6px 8px",textAlign:"right",color:p>100?"#F87171":p>80?"#FBBF24":"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{fmt(m.rolling12)}</td><td style={{padding:"6px 8px",textAlign:"right",color:p>100?"#F87171":"#888",fontVariantNumeric:"tabular-nums"}}>{p.toFixed(1)}%</td>
          <td style={{padding:"6px 4px"}}><button onClick={e=>{e.stopPropagation();onIIBB(m.key);}} style={{background:"#8B5CF622",color:"#A78BFA",border:"1px solid #8B5CF633",borderRadius:4,padding:"3px 8px",fontSize:10,fontWeight:600,cursor:"pointer"}}>IIBB</button></td>
          <td style={{padding:"6px 4px"}}><button onClick={e=>{e.stopPropagation();onSelectPeriod(m.key);}} style={{background:"#6C9CFF15",color:"#6C9CFF",border:"1px solid #6C9CFF33",borderRadius:4,padding:"3px 8px",fontSize:10,fontWeight:600,cursor:"pointer"}}>Ver</button></td>
        </tr>);})}</tbody></table></details>}
  </div>);
}

// IIBB Modal
const IIBB_CATS=["Inscriptos","Consumidor Final","Exento","Exportaciones"];
function classifyComprobante(tipo,tipoDocRec){const t=String(tipo||"").trim();const code=parseInt(t);const doc=String(tipoDocRec||"").trim().toLowerCase();
  if([1,2,3,51,52,53,201,202,203].includes(code)||/factura\s*a|nd\s*a|nc\s*a/i.test(t))return"Inscriptos";
  if([19,20,21].includes(code)||/exporta/i.test(t))return"Exportaciones";
  if([6,7,8,56,57,58,206,207,208].includes(code)||/factura\s*b|nd\s*b|nc\s*b/i.test(t))return"Consumidor Final";
  // Factura C: classify by receptor document type
  if([11,12,13,61,62,63,211,212,213].includes(code)||/factura\s*c|nd\s*c|nc\s*c/i.test(t)){
    if(doc==="cuit"||doc==="80")return"Inscriptos";
    return"Consumidor Final";}
  return"Consumidor Final";}

function IIBBModal({emitidos,period,onClose}){
  const[copied,setCopied]=useState("");
  const rows=emitidos.filter(r=>r.key===period);
  const byCategory={};IIBB_CATS.forEach(c=>byCategory[c]={neto:0,count:0,items:[]});
  rows.forEach(r=>{const cat=classifyComprobante(r.tipo,r.tipoDocRec);const sign=r.isNC?-1:1;const monto=r.neto||r.total;byCategory[cat].neto+=monto*sign;byCategory[cat].count++;byCategory[cat].items.push(r);});
  const grandTotal=IIBB_CATS.reduce((s,c)=>s+byCategory[c].neto,0);
  const copyVal=(cat,val)=>{const txt=val.toFixed(2).replace(".",",");navigator.clipboard.writeText(txt).then(()=>{setCopied(cat);setTimeout(()=>setCopied(""),1500);});};
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}} onClick={onClose}>
    <div style={{background:"#14142e",borderRadius:14,maxWidth:640,width:"100%",maxHeight:"85vh",overflow:"auto",border:"1px solid #2a2a50"}} onClick={e=>e.stopPropagation()}>
      <div style={{padding:"20px 22px 16px",borderBottom:"1px solid #2a2a40",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:16,fontWeight:700}}>IIBB Misiones — {pLabel(period)}</div><div style={{fontSize:12,color:"#888",marginTop:2}}>Neto facturado por tipo de receptor</div></div><button onClick={onClose} style={{background:"none",border:"none",color:"#777",cursor:"pointer",fontSize:22,lineHeight:1}}>×</button></div>
      <div style={{padding:"18px 22px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:18}}>
          {IIBB_CATS.map(c=>{const d=byCategory[c];const ic=copied===c;return(
            <div key={c} style={{background:"#0a0a14",borderRadius:8,padding:"12px 14px",borderLeft:`3px solid ${d.neto>0?"#6C9CFF":"#333"}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{fontSize:10,color:"#888",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{c}</div>
                <button onClick={()=>copyVal(c,d.neto)} style={{background:ic?"#4ADE8022":"#6C9CFF15",border:"1px solid "+(ic?"#4ADE8044":"#6C9CFF33"),borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:600,color:ic?"#4ADE80":"#6C9CFF",cursor:"pointer"}}>{ic?"Copiado ✓":"Copiar"}</button></div>
              <div style={{fontSize:18,fontWeight:700,color:d.neto>0?"#6C9CFF":"#555",fontVariantNumeric:"tabular-nums"}}>{fmt(d.neto)}</div>
              <div style={{fontSize:11,color:"#555",marginTop:2}}>{d.count} comp.</div></div>);})}
        </div>
        <div style={{background:"#6C9CFF11",borderRadius:8,padding:"12px 16px",marginBottom:18,display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontSize:13,fontWeight:600}}>Total neto facturado</span>
          <div style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:18,fontWeight:700,color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(grandTotal)}</span>
            <button onClick={()=>copyVal("total",grandTotal)} style={{background:copied==="total"?"#4ADE8022":"#6C9CFF15",border:"1px solid "+(copied==="total"?"#4ADE8044":"#6C9CFF33"),borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:600,color:copied==="total"?"#4ADE80":"#6C9CFF",cursor:"pointer"}}>{copied==="total"?"Copiado ✓":"Copiar"}</button></div></div>
        {IIBB_CATS.filter(c=>byCategory[c].count>0).map(c=>{const d=byCategory[c];return(<details key={c} style={{marginBottom:8}}><summary style={{fontSize:12,fontWeight:600,cursor:"pointer",color:"#888",marginBottom:6}}>{c} — {d.count} comp. — {fmt(d.neto)}</summary>
          <div style={{maxHeight:200,overflowY:"auto",marginBottom:8}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}><thead><tr style={{borderBottom:"1px solid #2a2a40"}}>{["Fecha","Tipo","PV","Nro","Denominación","Neto"].map((h,i)=>(<th key={h} style={{padding:"5px 7px",textAlign:i<5?"left":"right",color:"#666",fontWeight:500,fontSize:10,textTransform:"uppercase",position:"sticky",top:0,background:"#14142e"}}>{h}</th>))}</tr></thead>
            <tbody>{d.items.map((r,i)=>(<tr key={i} style={{borderBottom:"1px solid #1a1a30",color:r.isNC?"#e88":"inherit"}}><td style={{padding:"4px 7px",whiteSpace:"nowrap"}}>{new Date(r.fecha).toLocaleDateString("es-AR")}</td><td style={{padding:"4px 7px",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.tipo}</td><td style={{padding:"4px 7px"}}>{r.ptoVenta}</td><td style={{padding:"4px 7px"}}>{r.numDesde}</td><td style={{padding:"4px 7px",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.denom}</td><td style={{padding:"4px 7px",textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{fmt((r.neto||r.total)*(r.isNC?-1:1))}</td></tr>))}</tbody></table></div></details>);})}
      </div></div></div>);}

// Dedup Modal
function DedupModal({report,onConfirm,onCancel}){const{newRows,duplicates,fileResults}=report;const tN=newRows.length,tD=duplicates.length;
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}}><div style={{background:"#14142e",borderRadius:14,maxWidth:600,width:"100%",maxHeight:"85vh",overflow:"auto",border:"1px solid #2a2a50"}}>
    <div style={{padding:"20px 22px 16px",borderBottom:"1px solid #2a2a40"}}><div style={{fontSize:16,fontWeight:700,marginBottom:4}}>Resultado de carga</div><div style={{fontSize:13,color:"#888"}}>{fileResults.length} archivo{fileResults.length!==1?"s":""}</div></div>
    <div style={{padding:"16px 22px"}}>{fileResults.map((fr,i)=>(<div key={i} style={{marginBottom:12,padding:"10px 12px",background:"#0a0a14",borderRadius:8,border:"1px solid #1a1a30"}}><div style={{fontSize:12,fontWeight:600,marginBottom:4,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{fr.fileName}</div><div style={{display:"flex",gap:12,fontSize:12}}><span style={{color:fr.type==="emitidos"?"#6C9CFF":"#4ADE80"}}>{fr.type==="emitidos"?"Emitidos":"Recibidos"}</span><span style={{color:"#4ADE80"}}>{fr.newCount} nuevos</span>{fr.dupCount>0&&<span style={{color:"#F87171"}}>{fr.dupCount} dup.</span>}</div></div>))}
      <div style={{display:"flex",gap:10,marginTop:8,marginBottom:16}}><div style={{flex:1,background:"#0a0a14",borderRadius:8,padding:"12px 14px",borderLeft:"3px solid #4ADE80"}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Nuevos</div><div style={{fontSize:22,fontWeight:700,color:"#4ADE80"}}>{tN}</div></div><div style={{flex:1,background:"#0a0a14",borderRadius:8,padding:"12px 14px",borderLeft:`3px solid ${tD?"#F87171":"#333"}`}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Duplicados</div><div style={{fontSize:22,fontWeight:700,color:tD?"#F87171":"#555"}}>{tD}</div></div></div>
      {tN===0&&tD>0&&<div style={{padding:"10px 14px",background:"#F871711A",borderRadius:8,fontSize:13,color:"#F87171"}}>Todos ya están cargados.</div>}</div>
    <div style={{padding:"12px 22px 18px",display:"flex",gap:10,justifyContent:"flex-end",borderTop:"1px solid #2a2a40"}}><button onClick={onCancel} style={{background:"none",border:"1px solid #333",borderRadius:8,padding:"8px 18px",color:"#888",fontSize:13,cursor:"pointer"}}>Cancelar</button>{tN>0&&<button onClick={()=>onConfirm(newRows)} style={{background:"#4ADE80",color:"#000",border:"none",borderRadius:8,padding:"8px 22px",fontSize:13,fontWeight:600,cursor:"pointer"}}>Cargar {tN}</button>}</div></div></div>);}

// Manual CAI
const COMP_TYPES=[{code:"1",label:"Factura A"},{code:"2",label:"ND A"},{code:"3",label:"NC A"},{code:"6",label:"Factura B"},{code:"7",label:"ND B"},{code:"8",label:"NC B"},{code:"11",label:"Factura C"},{code:"12",label:"ND C"},{code:"13",label:"NC C"}];
function ManualEntryForm({onAdd,onClose,existingKeys}){
  const today=new Date().toISOString().slice(0,10);const[ft,setFt]=useState("emitidos");const[tipo,setTipo]=useState("1");const[fecha,setFecha]=useState(today);const[pv,setPv]=useState("");const[num,setNum]=useState("");const[cuit,setCuit]=useState("");const[denom,setDenom]=useState("");const[neto,setNeto]=useState("");const[iva21,setIva21]=useState("");const[iva105,setIva105]=useState("");const[iva27,setIva27]=useState("");const[totalMan,setTotalMan]=useState("");const[error,setError]=useState("");
  const ivaT=(parseFloat(iva21)||0)+(parseFloat(iva105)||0)+(parseFloat(iva27)||0);const netoV=parseFloat(neto)||0;const compT=netoV+ivaT;const totalF=totalMan?parseFloat(totalMan):compT;
  const handleSubmit=()=>{if(!fecha||!pv||!num){setError("Completá fecha, PV y número.");return;}const d=parseDate(fecha);if(!d){setError("Fecha inválida.");return;}const tl=COMP_TYPES.find(t=>t.code===tipo)?.label||tipo;
    const row={fecha:d.toISOString(),tipo:`${tipo} - ${tl}`,ptoVenta:pv.trim(),numDesde:num.trim(),neto:netoV,iva:ivaT,iva21:parseFloat(iva21)||0,iva105:parseFloat(iva105)||0,iva27:parseFloat(iva27)||0,noGrav:0,exentas:0,total:totalF,denom:denom.trim(),cuit:cuit.trim(),moneda:"$",tc:1,fileType:ft,isNC:NC_CODES.has(tipo),sourceFile:"manual",key:pKey(d.getFullYear(),d.getMonth()+1)};
    row.dedupKey=makeDK(row);if(existingKeys.has(row.dedupKey)){setError("Comprobante duplicado.");return;}onAdd(row);setNum("");setNeto("");setIva21("");setIva105("");setIva27("");setTotalMan("");setDenom("");setCuit("");setError("");};
  const is=s=>({background:"#0a0a14",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"7px 10px",fontSize:13,outline:"none",width:"100%",...s});
  return(<div style={{background:"#12122a",borderRadius:10,padding:"18px",marginBottom:20,border:"1px solid #2a2a50"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><div style={{fontSize:14,fontWeight:600}}>✏️ Carga manual — CAI</div><button onClick={onClose} style={{background:"none",border:"none",color:"#777",cursor:"pointer",fontSize:18}}>×</button></div>
    <div style={{display:"grid",gridTemplateColumns:"140px 1fr 150px",gap:10,marginBottom:10}} className="grid-2"><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Dirección</label><select value={ft} onChange={e=>setFt(e.target.value)} style={is()}><option value="emitidos">Emitido</option><option value="recibidos">Recibido</option></select></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Tipo</label><select value={tipo} onChange={e=>setTipo(e.target.value)} style={is()}>{COMP_TYPES.map(t=>(<option key={t.code} value={t.code}>{t.code} - {t.label}</option>))}</select></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Fecha</label><input type="date" value={fecha} onChange={e=>setFecha(e.target.value)} style={is()}/></div></div>
    <div style={{display:"grid",gridTemplateColumns:"80px 120px 150px 1fr",gap:10,marginBottom:10}} className="grid-2"><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>PV</label><input value={pv} onChange={e=>setPv(e.target.value)} placeholder="0001" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Número</label><input value={num} onChange={e=>setNum(e.target.value)} placeholder="00000001" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>CUIT</label><input value={cuit} onChange={e=>setCuit(e.target.value)} placeholder="XX-XXXXXXXX-X" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Denominación</label><input value={denom} onChange={e=>setDenom(e.target.value)} placeholder="Razón social" style={is()}/></div></div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr 1fr",gap:10,marginBottom:12}} className="grid-3"><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Neto Gravado</label><input type="number" value={neto} onChange={e=>setNeto(e.target.value)} placeholder="0.00" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>IVA 21%</label><input type="number" value={iva21} onChange={e=>setIva21(e.target.value)} placeholder="0.00" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>IVA 10.5%</label><input type="number" value={iva105} onChange={e=>setIva105(e.target.value)} placeholder="0.00" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>IVA 27%</label><input type="number" value={iva27} onChange={e=>setIva27(e.target.value)} placeholder="0.00" style={is()}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Total</label><input type="number" value={totalMan} onChange={e=>setTotalMan(e.target.value)} placeholder={compT?compT.toFixed(2):"0.00"} style={is()}/></div></div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10}}><div style={{fontSize:12,color:"#777"}}>{netoV||ivaT?<>IVA: <strong style={{color:"#4ADE80"}}>{fmt(ivaT)}</strong> · Total: <strong>{fmt(totalF)}</strong></>:"Ingresá los importes"}</div><button onClick={handleSubmit} style={{background:"#4ADE80",color:"#000",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:"pointer"}}>Agregar</button></div>
    {error&&<div style={{marginTop:8,fontSize:12,color:"#F87171"}}>{error}</div>}</div>);}

// Ret/Perc
function RetPercEditor({data,onChange,periods}){const[adding,setAdding]=useState(false);const[ek,setEk]=useState(periods.length?periods[periods.length-1]:"");const[rv,setRv]=useState("");const[pv,setPv]=useState("");
  const handleAdd=()=>{if(!ek)return;const r=parseFloat(rv)||0,p=parseFloat(pv)||0;if(r===0&&p===0)return;const prev=data[ek]||{ret:0,perc:0};onChange({...data,[ek]:{ret:prev.ret+r,perc:prev.perc+p}});setRv("");setPv("");setAdding(false);};
  const entries=Object.entries(data).filter(([,v])=>v.ret||v.perc).sort(([a],[b])=>a.localeCompare(b));
  return(<div style={{background:"#12122a",borderRadius:10,padding:"14px 16px",marginBottom:20}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:entries.length||adding?12:0}}><div><div style={{fontSize:13,fontWeight:600}}>Retenciones y Percepciones</div><div style={{fontSize:11,color:"#666",marginTop:2}}>Se restan del débito fiscal</div></div>{!adding&&<button onClick={()=>setAdding(true)} style={{background:"#8B5CF6",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Agregar</button>}</div>
    {adding&&<div style={{background:"#0a0a14",borderRadius:8,padding:"12px",marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end",border:"1px solid #2a2a40"}}><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Período</label><select value={ek} onChange={e=>setEk(e.target.value)} style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,outline:"none"}}>{periods.map(p=>(<option key={p} value={p}>{pLabel(p)}</option>))}</select></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Retenciones</label><input type="number" value={rv} onChange={e=>setRv(e.target.value)} placeholder="0.00" style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,width:120,outline:"none"}}/></div><div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Percepciones</label><input type="number" value={pv} onChange={e=>setPv(e.target.value)} placeholder="0.00" style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,width:120,outline:"none"}}/></div><button onClick={handleAdd} style={{background:"#8B5CF6",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Agregar</button><button onClick={()=>setAdding(false)} style={{background:"none",border:"1px solid #333",borderRadius:6,padding:"6px 12px",color:"#888",fontSize:12,cursor:"pointer"}}>Cancelar</button></div>}
    {entries.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:6}}>{entries.map(([k,v])=>(<div key={k} style={{background:"#0a0a14",borderRadius:6,padding:"6px 10px",fontSize:11,display:"flex",gap:8,alignItems:"center",border:"1px solid #1a1a30"}}><span style={{fontWeight:600}}>{pLabel(k)}</span>{v.ret>0&&<span style={{color:"#C084FC"}}>R: {fmt(v.ret)}</span>}{v.perc>0&&<span style={{color:"#A78BFA"}}>P: {fmt(v.perc)}</span>}<button onClick={()=>{const n={...data};delete n[k];onChange(n);}} style={{background:"none",border:"none",color:"#F8717166",cursor:"pointer",fontSize:12,padding:0}}>✕</button></div>))}</div>}</div>);}

// Month Detail
function MonthDetail({rows,period,type}){const f=rows.filter(r=>r.key===period&&r.fileType===type);if(!f.length)return<div style={{padding:16,color:"#666",fontSize:13,textAlign:"center"}}>Sin comprobantes</div>;
  return(<div style={{maxHeight:300,overflowY:"auto",overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:620}}><thead><tr style={{borderBottom:"1px solid #333"}}>{["Fecha","Tipo","PV","Nro","Denominación","Neto","IVA","Total"].map((h,i)=>(<th key={h} style={{padding:"7px 8px",textAlign:i<5?"left":"right",color:"#777",fontWeight:500,fontSize:10,textTransform:"uppercase",position:"sticky",top:0,background:"#1a1a2e"}}>{h}</th>))}</tr></thead>
    <tbody>{f.map((r,i)=>(<tr key={i} style={{borderBottom:"1px solid #1f1f35",color:r.isNC?"#e88":"inherit"}}><td style={{padding:"5px 8px",whiteSpace:"nowrap"}}>{new Date(r.fecha).toLocaleDateString("es-AR")}</td><td style={{padding:"5px 8px",maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.tipo}</td><td style={{padding:"5px 8px"}}>{r.ptoVenta}</td><td style={{padding:"5px 8px"}}>{r.numDesde}</td><td style={{padding:"5px 8px",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.denom}{r.tc>1?<span style={{color:"#FBBF24",fontSize:10,marginLeft:4}}>USD</span>:null}</td><td style={{padding:"5px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{fmt(r.neto)}</td><td style={{padding:"5px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{fmt(r.iva)}</td><td style={{padding:"5px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{fmt(r.total)}</td></tr>))}</tbody></table></div>);}

// ── Accounting Module ──

const DEFAULT_PLAN = [
  // 1. Activo
  {code:"1",name:"ACTIVO",type:"A",level:0,editable:false},
  {code:"1.1",name:"Activo Corriente",type:"A",level:1},
  {code:"1.1.01",name:"Caja",type:"A",level:2},
  {code:"1.1.02",name:"Bancos",type:"A",level:2},
  {code:"1.1.03",name:"Deudores por Ventas",type:"A",level:2},
  {code:"1.1.04",name:"Documentos a Cobrar",type:"A",level:2},
  {code:"1.1.05",name:"IVA Crédito Fiscal",type:"A",level:2},
  {code:"1.1.06",name:"Bienes de Cambio",type:"A",level:2},
  {code:"1.1.07",name:"Anticipos Impositivos",type:"A",level:2},
  {code:"1.2",name:"Activo No Corriente",type:"A",level:1},
  {code:"1.2.01",name:"Bienes de Uso",type:"A",level:2},
  {code:"1.2.02",name:"Amortizaciones Acumuladas",type:"A",level:2},
  // 2. Pasivo
  {code:"2",name:"PASIVO",type:"P",level:0,editable:false},
  {code:"2.1",name:"Pasivo Corriente",type:"P",level:1},
  {code:"2.1.01",name:"Proveedores",type:"P",level:2},
  {code:"2.1.02",name:"IVA Débito Fiscal",type:"P",level:2},
  {code:"2.1.03",name:"Cargas Sociales a Pagar",type:"P",level:2},
  {code:"2.1.04",name:"Sueldos a Pagar",type:"P",level:2},
  {code:"2.1.05",name:"Retenciones y Percepciones",type:"P",level:2},
  {code:"2.1.06",name:"Impuestos a Pagar",type:"P",level:2},
  {code:"2.2",name:"Pasivo No Corriente",type:"P",level:1},
  {code:"2.2.01",name:"Deudas Bancarias",type:"P",level:2},
  // 3. Patrimonio Neto
  {code:"3",name:"PATRIMONIO NETO",type:"PN",level:0,editable:false},
  {code:"3.1.01",name:"Capital",type:"PN",level:2},
  {code:"3.1.02",name:"Resultados Acumulados",type:"PN",level:2},
  {code:"3.1.03",name:"Resultado del Ejercicio",type:"PN",level:2},
  // 4. Resultados Positivos
  {code:"4",name:"RESULTADOS POSITIVOS",type:"R+",level:0,editable:false},
  {code:"4.1.01",name:"Ventas",type:"R+",level:2},
  {code:"4.1.02",name:"Intereses Ganados",type:"R+",level:2},
  {code:"4.1.03",name:"Otros Ingresos",type:"R+",level:2},
  // 5. Resultados Negativos
  {code:"5",name:"RESULTADOS NEGATIVOS",type:"R-",level:0,editable:false},
  {code:"5.1.01",name:"Costo de Mercaderías Vendidas",type:"R-",level:2},
  {code:"5.1.02",name:"Sueldos y Jornales",type:"R-",level:2},
  {code:"5.1.03",name:"Cargas Sociales",type:"R-",level:2},
  {code:"5.1.04",name:"Servicios",type:"R-",level:2},
  {code:"5.1.05",name:"Alquileres",type:"R-",level:2},
  {code:"5.1.06",name:"Impuestos y Tasas",type:"R-",level:2},
  {code:"5.1.07",name:"Amortizaciones",type:"R-",level:2},
  {code:"5.1.08",name:"Gastos Generales",type:"R-",level:2},
  {code:"5.1.09",name:"Gastos Bancarios",type:"R-",level:2},
];

const TYPE_COLORS={"A":"#6C9CFF","P":"#F87171","PN":"#4ADE80","R+":"#4ADE80","R-":"#F87171"};
const TYPE_LABELS={"A":"Activo","P":"Pasivo","PN":"Pat. Neto","R+":"Res. (+)","R-":"Res. (-)"};

// ── Plan de Cuentas Editor ──
function PlanCuentasEditor({plan:planProp,onChange}){
  const[localPlan,setLocalPlan]=useState(()=>planProp);
  const[adding,setAdding]=useState(false);
  const[newCode,setNewCode]=useState("");
  const[newName,setNewName]=useState("");
  const[newType,setNewType]=useState("R-");
  const[debugMsg,setDebugMsg]=useState("");
  const[editIdx,setEditIdx]=useState(null);
  const[editName,setEditName]=useState("");
  useEffect(()=>{setLocalPlan(planProp);},[planProp]);

  const doSort=(arr)=>[...arr].sort((a,b)=>{const pa=a.code.split(".").map(Number);const pb=b.code.split(".").map(Number);for(let i=0;i<Math.max(pa.length,pb.length);i++){const va=pa[i]||0;const vb=pb[i]||0;if(va!==vb)return va-vb;}return 0;});

  // Auto-suggest next available code when type changes or form opens
  const suggestCode=(type)=>{
    const prefix={"A":"1.1.","P":"2.1.","PN":"3.1.","R+":"4.1.","R-":"5.1."}[type]||"5.1.";
    const existing=localPlan.filter(a=>a.code.startsWith(prefix)).map(a=>{const parts=a.code.split(".");return parts.length>=3?parseInt(parts[2])||0:0;});
    const next=existing.length?Math.max(...existing)+1:1;
    return prefix+String(next).padStart(2,"0");
  };

  const startAdding=()=>{const type="R-";setNewType(type);setNewCode(suggestCode(type));setNewName("");setDebugMsg("");setAdding(true);};
  const handleTypeChange=(type)=>{setNewType(type);setNewCode(suggestCode(type));};

  const handleAdd=()=>{
    const code=newCode.trim(),name=newName.trim();
    if(!code){setDebugMsg("Ingresá un código de cuenta.");return;}
    if(!name){setDebugMsg("Ingresá un nombre de cuenta.");return;}
    if(localPlan.find(a=>a.code===code)){setDebugMsg(`El código ${code} ya existe.`);return;}
    const level=code.split(".").length-1;
    const updated=doSort([...localPlan,{code,name,type:newType,level:Math.min(level,2)}]);
    setLocalPlan(updated);onChange(updated);
    setNewCode("");setNewName("");setDebugMsg("");setAdding(false);
  };
  const handleDelete=(idx)=>{const acc=localPlan[idx];if(acc.editable===false)return;const updated=localPlan.filter((_,i)=>i!==idx);setLocalPlan(updated);onChange(updated);};
  const handleEditSave=(idx)=>{if(!editName.trim())return;const updated=[...localPlan];updated[idx]={...updated[idx],name:editName.trim()};setLocalPlan(updated);onChange(updated);setEditIdx(null);};

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <div style={{fontSize:13,fontWeight:600}}>Plan de Cuentas ({localPlan.length} cuentas)</div>
        <button onClick={startAdding} style={{background:"#6C9CFF",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Nueva cuenta</button>
      </div>

      {adding&&(
        <div style={{background:"#0a0a14",borderRadius:8,padding:"12px",marginBottom:12,border:"1px solid #2a2a40"}}>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Tipo</label>
            <select value={newType} onChange={e=>handleTypeChange(e.target.value)} style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,outline:"none"}}>
              {Object.entries(TYPE_LABELS).map(([k,v])=>(<option key={k} value={k}>{v}</option>))}</select></div>
          <div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Código</label>
            <input value={newCode} onChange={e=>setNewCode(e.target.value)} style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,width:90,outline:"none",fontVariantNumeric:"tabular-nums"}}/></div>
          <div style={{flex:1,minWidth:150}}><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Nombre de cuenta</label>
            <input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="Ej: Honorarios Profesionales" onKeyDown={e=>e.key==="Enter"&&handleAdd()} autoFocus style={{background:"#12122a",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:13,width:"100%",outline:"none"}}/></div>
          <button onClick={handleAdd} style={{background:"#4ADE80",color:"#000",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Agregar</button>
          <button onClick={()=>{setAdding(false);setDebugMsg("");}} style={{background:"none",border:"1px solid #333",borderRadius:6,padding:"6px 12px",color:"#888",fontSize:12,cursor:"pointer"}}>Cancelar</button>
          </div>
          {debugMsg&&<div style={{marginTop:8,fontSize:12,color:"#F87171"}}>{debugMsg}</div>}
        </div>
      )}

      <div style={{maxHeight:500,overflowY:"auto"}}>
        {localPlan.map((acc,i)=>(
          <div key={acc.code} style={{display:"flex",alignItems:"center",gap:8,padding:`${acc.level===0?"10px":"5px"} 8px`,borderBottom:"1px solid #1a1a30",paddingLeft:acc.level*20+8}}>
            <span style={{color:TYPE_COLORS[acc.type],fontSize:11,fontWeight:600,minWidth:60,fontVariantNumeric:"tabular-nums"}}>{acc.code}</span>
            {editIdx===i?(
              <input value={editName} onChange={e=>setEditName(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")handleEditSave(i);if(e.key==="Escape")setEditIdx(null);}} onBlur={()=>handleEditSave(i)} autoFocus
                style={{background:"#0a0a14",border:"1px solid #6C9CFF",borderRadius:4,color:"#e8e8e8",padding:"2px 8px",fontSize:12,flex:1,outline:"none"}}/>
            ):(
              <span onClick={()=>{if(acc.editable!==false){setEditIdx(i);setEditName(acc.name);}}} style={{fontSize:acc.level===0?13:12,fontWeight:acc.level===0?700:acc.level===1?600:400,color:acc.level===0?TYPE_COLORS[acc.type]:"#e8e8e8",flex:1,cursor:acc.editable!==false?"pointer":"default"}}>{acc.name}</span>
            )}
            <span style={{fontSize:10,color:"#555",minWidth:45}}>{TYPE_LABELS[acc.type]}</span>
            {acc.editable!==false&&acc.level===2&&(
              <button onClick={()=>handleDelete(i)} style={{background:"none",border:"none",color:"#F8717144",cursor:"pointer",fontSize:14,padding:"0 4px"}} onMouseEnter={e=>e.currentTarget.style.color="#F87171"} onMouseLeave={e=>e.currentTarget.style.color="#F8717144"}>✕</button>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}


// ── Asiento Form ──
function AsientoForm({plan,onSave,onCancel,initial}){
  const[fecha,setFecha]=useState(initial?.fecha||new Date().toISOString().slice(0,10));
  const[desc,setDesc]=useState(initial?.descripcion||"");
  const[lines,setLines]=useState(initial?.lineas||[{cuenta:"",debe:0,haber:0},{cuenta:"",debe:0,haber:0}]);
  const[error,setError]=useState("");

  const imputableAccounts=plan.filter(a=>a.level===2);
  const totalDebe=lines.reduce((s,l)=>s+(parseFloat(l.debe)||0),0);
  const totalHaber=lines.reduce((s,l)=>s+(parseFloat(l.haber)||0),0);
  const balanced=Math.abs(totalDebe-totalHaber)<0.01;

  const updateLine=(i,field,val)=>{const n=[...lines];n[i]={...n[i],[field]:val};setLines(n);};
  const addLine=()=>setLines([...lines,{cuenta:"",debe:0,haber:0}]);
  const removeLine=(i)=>{if(lines.length<=2)return;setLines(lines.filter((_,j)=>j!==i));};

  const handleSave=()=>{
    if(!fecha){setError("Ingresá la fecha.");return;}
    if(!desc.trim()){setError("Ingresá una descripción.");return;}
    const validLines=lines.filter(l=>l.cuenta&&(parseFloat(l.debe)||parseFloat(l.haber)));
    if(validLines.length<2){setError("Mínimo 2 líneas con cuenta e importe.");return;}
    if(!balanced){setError(`El asiento no balancea. Diferencia: ${fmt(totalDebe-totalHaber)}`);return;}
    onSave({id:initial?.id||`as_${Date.now()}`,fecha,descripcion:desc.trim(),lineas:validLines.map(l=>({cuenta:l.cuenta,debe:parseFloat(l.debe)||0,haber:parseFloat(l.haber)||0}))});
  };

  const is={background:"#0a0a14",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"7px 10px",fontSize:13,outline:"none",width:"100%"};
  return(
    <div style={{background:"#12122a",borderRadius:10,padding:"18px",marginBottom:20,border:"1px solid #2a2a50"}}>
      <div style={{fontSize:14,fontWeight:600,marginBottom:14}}>{initial?"Editar asiento":"Nuevo asiento contable"}</div>
      <div style={{display:"grid",gridTemplateColumns:"150px 1fr",gap:10,marginBottom:14}} className="grid-2">
        <div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Fecha</label><input type="date" value={fecha} onChange={e=>setFecha(e.target.value)} style={is}/></div>
        <div><label style={{fontSize:10,color:"#888",display:"block",marginBottom:3}}>Descripción</label><input value={desc} onChange={e=>setDesc(e.target.value)} placeholder="Concepto del asiento" style={is}/></div>
      </div>

      <div style={{marginBottom:12}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 110px 110px 30px",gap:6,marginBottom:6}}>
          <div style={{fontSize:10,color:"#888",fontWeight:600}}>CUENTA</div>
          <div style={{fontSize:10,color:"#888",fontWeight:600,textAlign:"right"}}>DEBE</div>
          <div style={{fontSize:10,color:"#888",fontWeight:600,textAlign:"right"}}>HABER</div>
          <div></div>
        </div>
        {lines.map((l,i)=>(
          <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 110px 110px 30px",gap:6,marginBottom:4}}>
            <select value={l.cuenta} onChange={e=>updateLine(i,"cuenta",e.target.value)} style={{...is,padding:"6px 8px",fontSize:12}}>
              <option value="">Seleccionar cuenta...</option>
              {imputableAccounts.map(a=>(<option key={a.code} value={a.code}>{a.code} - {a.name}</option>))}
            </select>
            <input type="number" value={l.debe||""} onChange={e=>updateLine(i,"debe",e.target.value)} placeholder="0,00" style={{...is,textAlign:"right",fontSize:12,padding:"6px 8px"}}/>
            <input type="number" value={l.haber||""} onChange={e=>updateLine(i,"haber",e.target.value)} placeholder="0,00" style={{...is,textAlign:"right",fontSize:12,padding:"6px 8px"}}/>
            <button onClick={()=>removeLine(i)} style={{background:"none",border:"none",color:lines.length>2?"#F8717166":"#222",cursor:lines.length>2?"pointer":"default",fontSize:14}}>✕</button>
          </div>
        ))}
        <button onClick={addLine} style={{background:"none",border:"1px dashed #333",borderRadius:6,padding:"6px 14px",color:"#888",fontSize:12,cursor:"pointer",marginTop:4}}>+ Línea</button>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,paddingTop:10,borderTop:"1px solid #2a2a40"}}>
        <div style={{fontSize:12}}>
          <span style={{color:"#6C9CFF"}}>Debe: {fmt(totalDebe)}</span>
          <span style={{color:"#777",margin:"0 8px"}}>|</span>
          <span style={{color:"#4ADE80"}}>Haber: {fmt(totalHaber)}</span>
          <span style={{color:"#777",margin:"0 8px"}}>|</span>
          <span style={{color:balanced?"#4ADE80":"#F87171",fontWeight:600}}>{balanced?"✓ Balanceado":`Diferencia: ${fmt(totalDebe-totalHaber)}`}</span>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={onCancel} style={{background:"none",border:"1px solid #333",borderRadius:8,padding:"8px 16px",color:"#888",fontSize:13,cursor:"pointer"}}>Cancelar</button>
          <button onClick={handleSave} disabled={!balanced} style={{background:balanced?"#4ADE80":"#333",color:balanced?"#000":"#666",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:balanced?"pointer":"default"}}>Guardar asiento</button>
        </div>
      </div>
      {error&&<div style={{marginTop:8,fontSize:12,color:"#F87171"}}>{error}</div>}
    </div>
  );
}

// ── Balance de Sumas y Saldos ──
function BalanceReport({plan,asientos,periodoFilter}){
  const[collapsed,setCollapsed]=useState({});
  const toggle=(code)=>setCollapsed(p=>({...p,[code]:!p[code]}));

  // Calculate saldos for level 2 (detail) accounts
  const saldos=useMemo(()=>{
    const map={};
    plan.filter(a=>a.level===2).forEach(a=>map[a.code]={debe:0,haber:0});
    (asientos||[]).forEach(as=>{
      if(periodoFilter&&as.fecha){const k=as.fecha.slice(0,7);if(k>periodoFilter)return;}
      (as.lineas||[]).forEach(l=>{if(map[l.cuenta]!==undefined){map[l.cuenta].debe+=l.debe||0;map[l.cuenta].haber+=l.haber||0;}});
    });
    return map;
  },[plan,asientos,periodoFilter]);

  // Build hierarchical rows with subtotals
  const rows=useMemo(()=>{
    const result=[];
    const getChildren=(parentCode,level)=>plan.filter(a=>{
      if(a.level!==level)return false;
      if(level===1)return a.code.startsWith(parentCode+".");
      if(level===2){const parts=a.code.split(".");const pparts=parentCode.split(".");return parts[0]===pparts[0]&&parts[1]===pparts[1];}
      return false;
    });

    plan.filter(a=>a.level===0).forEach(l0=>{
      // Sum all level 2 under this level 0
      const l2sUnder=plan.filter(a=>a.level===2&&a.code.startsWith(l0.code+"."));
      const l0sums=l2sUnder.reduce((s,a)=>{const v=saldos[a.code]||{debe:0,haber:0};return{debe:s.debe+v.debe,haber:s.haber+v.haber};},{debe:0,haber:0});
      const l0sd=l0sums.debe>l0sums.haber?l0sums.debe-l0sums.haber:0;
      const l0sa=l0sums.haber>l0sums.debe?l0sums.haber-l0sums.debe:0;
      result.push({...l0,...l0sums,saldoDeudor:l0sd,saldoAcreedor:l0sa,isGroup:true,hasData:l0sums.debe||l0sums.haber});

      if(!collapsed[l0.code]){
        const l1s=getChildren(l0.code,1);
        l1s.forEach(l1=>{
          const l2sUnderL1=plan.filter(a=>a.level===2&&a.code.startsWith(l1.code+"."));
          const l1sums=l2sUnderL1.reduce((s,a)=>{const v=saldos[a.code]||{debe:0,haber:0};return{debe:s.debe+v.debe,haber:s.haber+v.haber};},{debe:0,haber:0});
          const l1sd=l1sums.debe>l1sums.haber?l1sums.debe-l1sums.haber:0;
          const l1sa=l1sums.haber>l1sums.debe?l1sums.haber-l1sums.debe:0;
          result.push({...l1,...l1sums,saldoDeudor:l1sd,saldoAcreedor:l1sa,isGroup:true,hasData:l1sums.debe||l1sums.haber});

          if(!collapsed[l1.code]){
            l2sUnderL1.forEach(l2=>{
              const s=saldos[l2.code]||{debe:0,haber:0};
              const sd=s.debe>s.haber?s.debe-s.haber:0;
              const sa=s.haber>s.debe?s.haber-s.debe:0;
              if(s.debe||s.haber)result.push({...l2,...s,saldoDeudor:sd,saldoAcreedor:sa,isGroup:false});
            });
          }
        });
      }
    });
    return result;
  },[plan,saldos,collapsed]);

  const totals=useMemo(()=>{
    const detail=plan.filter(a=>a.level===2).map(a=>{const s=saldos[a.code]||{debe:0,haber:0};return s;});
    const t=detail.reduce((a,s)=>({debe:a.debe+s.debe,haber:a.haber+s.haber}),{debe:0,haber:0});
    return{...t,sd:t.debe>t.haber?t.debe-t.haber:0,sa:t.haber>t.debe?t.haber-t.debe:0};
  },[plan,saldos]);

  if(!totals.debe&&!totals.haber)return<div style={{padding:20,textAlign:"center",color:"#555",fontSize:13}}>Sin movimientos registrados.</div>;

  const fmtCell=(v,color)=><span style={{fontVariantNumeric:"tabular-nums",color:v?color:"#333"}}>{v?fmt(v):"—"}</span>;

  return(
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:620}}>
        <thead><tr style={{borderBottom:"1px solid #2a2a40"}}>
          {["Código","Cuenta","Sumas Debe","Sumas Haber","Saldo Deudor","Saldo Acreedor"].map(h=>(
            <th key={h} style={{padding:"8px",textAlign:h==="Código"||h==="Cuenta"?"left":"right",color:"#666",fontWeight:600,fontSize:10,textTransform:"uppercase",letterSpacing:.3}}>{h}</th>
          ))}
        </tr></thead>
        <tbody>{rows.map(r=>{
          const isL0=r.level===0,isL1=r.level===1,isL2=r.level===2;
          const canToggle=r.isGroup;
          return(
            <tr key={r.code} onClick={canToggle?()=>toggle(r.code):undefined}
              style={{borderBottom:`1px solid ${isL0?"#2a2a40":"#1a1a30"}`,cursor:canToggle?"pointer":"default",background:isL0&&r.hasData?"#0a0a1a":"transparent"}}
              onMouseEnter={e=>{if(canToggle)e.currentTarget.style.background="#151530";}}
              onMouseLeave={e=>{e.currentTarget.style.background=isL0&&r.hasData?"#0a0a1a":"transparent";}}>
              <td style={{padding:`${isL0?"10px":"6px"} 8px`,color:TYPE_COLORS[r.type],fontWeight:600,fontSize:isL0?12:11,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"}}>
                {canToggle&&<span style={{display:"inline-block",width:14,fontSize:10,color:"#666"}}>{collapsed[r.code]?"▸":"▾"}</span>}
                {r.code}
              </td>
              <td style={{padding:`${isL0?"10px":"6px"} 8px`,fontWeight:isL0?700:isL1?600:400,fontSize:isL0?13:12,paddingLeft:isL2?32:isL1?20:8,color:isL0?TYPE_COLORS[r.type]:"#e8e8e8"}}>
                {r.name}
              </td>
              <td style={{padding:"6px 8px",textAlign:"right",fontWeight:r.isGroup?700:400}}>{fmtCell(r.debe)}</td>
              <td style={{padding:"6px 8px",textAlign:"right",fontWeight:r.isGroup?700:400}}>{fmtCell(r.haber)}</td>
              <td style={{padding:"6px 8px",textAlign:"right",fontWeight:r.isGroup?700:400}}>{fmtCell(r.saldoDeudor,"#6C9CFF")}</td>
              <td style={{padding:"6px 8px",textAlign:"right",fontWeight:r.isGroup?700:400}}>{fmtCell(r.saldoAcreedor,"#4ADE80")}</td>
            </tr>
          );
        })}</tbody>
        <tfoot><tr style={{borderTop:"2px solid #444",background:"#0a0a1a"}}>
          <td colSpan={2} style={{padding:"10px 8px",fontWeight:700,fontSize:13}}>TOTALES</td>
          <td style={{padding:"10px 8px",textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums"}}>{fmt(totals.debe)}</td>
          <td style={{padding:"10px 8px",textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums"}}>{fmt(totals.haber)}</td>
          <td style={{padding:"10px 8px",textAlign:"right",fontWeight:700,color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(totals.sd)}</td>
          <td style={{padding:"10px 8px",textAlign:"right",fontWeight:700,color:"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{fmt(totals.sa)}</td>
        </tr></tfoot>
      </table>
    </div>
  );
}

// ── Generar Asientos Automáticos ──
function GenerarAsientosPanel({client,plan,onGenerate}){
  const reglas=client.reglasCuenta||{};
  const generados=new Set(client.asientosGenerados||[]);
  const isRI=(client.regimen||"RI")==="RI";
  const imputableAccounts=plan.filter(a=>a.level===2);

  // Find pending invoices (not yet generated)
  const emPending=(client.emitidos||[]).filter(r=>!generados.has(makeDK(r)));
  const recPending=(client.recibidos||[]).filter(r=>!generados.has(makeDK(r)));

  // Group recibidos by CUIT+denom for rule assignment
  const recByCuit=useMemo(()=>{
    const map={};
    recPending.forEach(r=>{
      const k=r.cuit||r.denom||"sin-cuit";
      if(!map[k])map[k]={cuit:r.cuit,denom:r.denom,rows:[],total:0};
      map[k].rows.push(r);
      map[k].total+=(r.neto||r.total)*(r.isNC?-1:1);
    });
    return Object.entries(map).sort(([,a],[,b])=>b.total-a.total);
  },[recPending]);

  const[localReglas,setLocalReglas]=useState(reglas);
  const[step,setStep]=useState("config"); // config | preview

  const updateRegla=(cuit,cuenta)=>{setLocalReglas({...localReglas,[cuit]:cuenta});};

  // Count how many recibidos have no rule assigned
  const sinAsignar=recByCuit.filter(([k])=>!localReglas[k]).length;

  // Generate entries
  const generarTodos=()=>{
    const nuevosAsientos=[];
    const nuevosGenerados=[...generados];

    // Emitidos → Ventas
    emPending.forEach(r=>{
      const sign=r.isNC?-1:1;
      const monto=Math.abs(r.total);
      const netoVal=Math.abs(r.neto||r.total);
      const ivaVal=Math.abs(r.iva);
      const otrosVal=monto-netoVal-ivaVal; // no gravado + exentas + otros tributos
      const lineas=[];
      const desc=`${r.isNC?"NC":"Vta"} ${r.tipo.replace(/^\d+\s*-\s*/,"")} ${r.ptoVenta}-${r.numDesde} ${r.denom}`.trim();

      if(r.isNC){
        lineas.push({cuenta:"4.1.01",debe:netoVal,haber:0});
        if(isRI&&ivaVal>0.01)lineas.push({cuenta:"2.1.02",debe:ivaVal,haber:0});
        if(otrosVal>0.01)lineas.push({cuenta:"4.1.03",debe:otrosVal,haber:0});
        lineas.push({cuenta:"1.1.03",debe:0,haber:monto});
      }else{
        lineas.push({cuenta:"1.1.03",debe:monto,haber:0});
        lineas.push({cuenta:"4.1.01",debe:0,haber:netoVal});
        if(isRI&&ivaVal>0.01)lineas.push({cuenta:"2.1.02",debe:0,haber:ivaVal});
        if(otrosVal>0.01)lineas.push({cuenta:"4.1.03",debe:0,haber:otrosVal});
      }
      nuevosAsientos.push({id:`av_${Date.now()}_${Math.random().toString(36).slice(2,5)}`,fecha:r.fecha.slice(0,10),descripcion:desc,lineas,auto:true});
      nuevosGenerados.push(makeDK(r));
    });

    // Recibidos → Compras
    recPending.forEach(r=>{
      const cuitKey=r.cuit||r.denom||"sin-cuit";
      const ctaCompra=localReglas[cuitKey]||"5.1.08"; // default: Gastos Generales
      const monto=Math.abs(r.total);
      const netoVal=Math.abs(r.neto||r.total);
      const ivaVal=Math.abs(r.iva);
      const otrosVal=monto-netoVal-ivaVal; // no gravado + exentas + otros tributos
      const lineas=[];
      const desc=`${r.isNC?"NC":"Cpa"} ${r.tipo.replace(/^\d+\s*-\s*/,"")} ${r.ptoVenta}-${r.numDesde} ${r.denom}`.trim();

      if(r.isNC){
        lineas.push({cuenta:"2.1.01",debe:monto,haber:0});
        lineas.push({cuenta:ctaCompra,debe:0,haber:netoVal});
        if(isRI&&ivaVal>0.01)lineas.push({cuenta:"1.1.05",debe:0,haber:ivaVal});
        if(otrosVal>0.01)lineas.push({cuenta:ctaCompra,debe:0,haber:otrosVal});
      }else{
        lineas.push({cuenta:ctaCompra,debe:netoVal,haber:0});
        if(isRI&&ivaVal>0.01)lineas.push({cuenta:"1.1.05",debe:ivaVal,haber:0});
        if(otrosVal>0.01)lineas.push({cuenta:ctaCompra,debe:otrosVal,haber:0});
        lineas.push({cuenta:"2.1.01",debe:0,haber:monto});
      }
      nuevosAsientos.push({id:`ac_${Date.now()}_${Math.random().toString(36).slice(2,5)}`,fecha:r.fecha.slice(0,10),descripcion:desc,lineas,auto:true});
      nuevosGenerados.push(makeDK(r));
    });

    onGenerate(nuevosAsientos,nuevosGenerados,[...Object.entries(localReglas)].reduce((o,[k,v])=>({...o,[k]:v}),{}));
  };

  const totalPending=emPending.length+recPending.length;
  if(totalPending===0)return(<div style={{padding:20,textAlign:"center",color:"#555",fontSize:13}}>✓ Todos los comprobantes cargados ya tienen asiento generado.</div>);

  return(<div>
    <div style={{fontSize:13,fontWeight:600,marginBottom:14}}>Generar asientos desde comprobantes</div>

    {/* Summary */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}} className="grid-2">
      <div style={{background:"#0a0a14",borderRadius:8,padding:"12px",borderLeft:"3px solid #6C9CFF"}}>
        <div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Ventas pendientes</div>
        <div style={{fontSize:20,fontWeight:700,color:"#6C9CFF"}}>{emPending.length}</div>
      </div>
      <div style={{background:"#0a0a14",borderRadius:8,padding:"12px",borderLeft:"3px solid #4ADE80"}}>
        <div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>Compras pendientes</div>
        <div style={{fontSize:20,fontWeight:700,color:"#4ADE80"}}>{recPending.length}</div>
      </div>
    </div>

    {/* Rules: assign account per supplier */}
    {recByCuit.length>0&&(<>
      <div style={{fontSize:12,fontWeight:600,marginBottom:8,color:"#888"}}>Clasificar compras por proveedor → cuenta contable</div>
      <div style={{maxHeight:300,overflowY:"auto",marginBottom:16}}>
        {recByCuit.map(([k,v])=>(
          <div key={k} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:"1px solid #1a1a30",flexWrap:"wrap"}}>
            <div style={{flex:"1 1 200px",minWidth:0}}>
              <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.denom||"Sin denominación"}</div>
              <div style={{fontSize:10,color:"#777"}}>{v.cuit} · {v.rows.length} comp. · {fmt(v.total)}</div>
            </div>
            <select value={localReglas[k]||""} onChange={e=>updateRegla(k,e.target.value)}
              style={{background:"#0a0a14",border:`1px solid ${localReglas[k]?"#333":"#F8717144"}`,borderRadius:6,color:"#e8e8e8",padding:"5px 8px",fontSize:11,outline:"none",minWidth:200}}>
              <option value="">Seleccionar cuenta...</option>
              {imputableAccounts.filter(a=>a.type==="R-"||a.type==="A").map(a=>(<option key={a.code} value={a.code}>{a.code} - {a.name}</option>))}
            </select>
          </div>
        ))}
      </div>
      {sinAsignar>0&&<div style={{fontSize:12,color:"#FBBF24",marginBottom:12}}>⚠ {sinAsignar} proveedor{sinAsignar!==1?"es":""} sin cuenta asignada — se usará "5.1.08 Gastos Generales" por defecto.</div>}
    </>)}

    {/* Asientos mapping explanation */}
    <div style={{background:"#0a0a14",borderRadius:8,padding:"12px",marginBottom:16,fontSize:11,color:"#777",lineHeight:1.6}}>
      <div style={{fontWeight:600,color:"#888",marginBottom:4}}>Asientos que se generarán:</div>
      <div><span style={{color:"#6C9CFF"}}>Ventas:</span> Deudores por Ventas (D) → Ventas (H){isRI?" + IVA DF (H)":""}</div>
      <div><span style={{color:"#4ADE80"}}>Compras:</span> Cuenta asignada (D){isRI?" + IVA CF (D)":""} → Proveedores (H)</div>
      <div><span style={{color:"#F87171"}}>NC:</span> Se registran con el asiento inverso</div>
    </div>

    <button onClick={generarTodos} style={{background:"#4ADE80",color:"#000",border:"none",borderRadius:8,padding:"10px 24px",fontSize:14,fontWeight:700,cursor:"pointer",width:"100%"}}>
      Generar {totalPending} asiento{totalPending!==1?"s":""}
    </button>
  </div>);
}

// ── Contabilidad Panel ──
function ContabilidadPanel({client,onUpdate}){
  const[tab,setTab]=useState("asientos");
  const[showForm,setShowForm]=useState(false);
  const[editAsiento,setEditAsiento]=useState(null);

  const plan=useMemo(()=>client.planCuentas||JSON.parse(JSON.stringify(DEFAULT_PLAN)),[client.planCuentas]);
  const asientos=client.asientos||[];

  const handlePlanChange=(newPlan)=>{onUpdate({...client,planCuentas:newPlan});};
  const handleSaveAsiento=(asiento)=>{
    const existing=asientos.findIndex(a=>a.id===asiento.id);
    let updated;
    if(existing>=0){updated=[...asientos];updated[existing]=asiento;}
    else{updated=[...asientos,asiento];}
    onUpdate({...client,asientos:updated.sort((a,b)=>a.fecha.localeCompare(b.fecha))});
    setShowForm(false);setEditAsiento(null);
  };
  const handleDeleteAsiento=(id)=>{
    if(confirm("¿Eliminar este asiento?"))onUpdate({...client,asientos:asientos.filter(a=>a.id!==id)});
  };
  const handleGenerate=(nuevos,generados,reglas)=>{
    const allAsientos=[...asientos,...nuevos].sort((a,b)=>a.fecha.localeCompare(b.fecha));
    onUpdate({...client,asientos:allAsientos,asientosGenerados:generados,reglasCuenta:reglas});
    setTab("asientos");
  };

  const tabs=[["asientos","Asientos"],["generar","Generar"],["plan","Plan de Cuentas"],["balance","Balance"]];

  return(
    <div style={{background:"#12122a",borderRadius:10,overflow:"hidden",marginBottom:20}}>
      <div style={{display:"flex",borderBottom:"1px solid #2a2a40"}}>
        {tabs.map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{flex:1,padding:"11px",background:tab===id?"#6C9CFF12":"transparent",border:"none",borderBottom:tab===id?"2px solid #6C9CFF":"2px solid transparent",color:tab===id?"#6C9CFF":"#777",cursor:"pointer",fontSize:13,fontWeight:600}}>{label}</button>
        ))}
      </div>
      <div style={{padding:"18px"}}>
        {tab==="plan"&&<PlanCuentasEditor plan={plan} onChange={handlePlanChange}/>}

        {tab==="generar"&&<GenerarAsientosPanel client={client} plan={plan} onGenerate={handleGenerate}/>}

        {tab==="asientos"&&(<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:600}}>{asientos.length} asiento{asientos.length!==1?"s":""} registrado{asientos.length!==1?"s":""}</div>
            {!showForm&&<button onClick={()=>{setEditAsiento(null);setShowForm(true);}} style={{background:"#6C9CFF",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>+ Nuevo asiento</button>}
          </div>
          {showForm&&<AsientoForm plan={plan} onSave={handleSaveAsiento} onCancel={()=>{setShowForm(false);setEditAsiento(null);}} initial={editAsiento}/>}
          {asientos.length>0&&(
            <div style={{maxHeight:400,overflowY:"auto"}}>
              {asientos.map(as=>{
                const totalD=as.lineas.reduce((s,l)=>s+(l.debe||0),0);
                return(
                  <div key={as.id} style={{background:"#0a0a14",borderRadius:8,padding:"12px",marginBottom:8,border:"1px solid #1a1a30"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                      <div><span style={{fontSize:12,color:"#888",marginRight:8}}>{new Date(as.fecha+"T12:00:00").toLocaleDateString("es-AR")}</span><span style={{fontSize:13,fontWeight:600}}>{as.descripcion}</span></div>
                      <div style={{display:"flex",gap:6}}>
                        <span style={{fontSize:13,fontWeight:700,color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(totalD)}</span>
                        <button onClick={()=>{setEditAsiento(as);setShowForm(true);}} style={{background:"none",border:"none",color:"#6C9CFF88",cursor:"pointer",fontSize:12}}>✎</button>
                        <button onClick={()=>handleDeleteAsiento(as.id)} style={{background:"none",border:"none",color:"#F8717155",cursor:"pointer",fontSize:12}}>✕</button>
                      </div>
                    </div>
                    <div style={{fontSize:11}}>
                      {as.lineas.map((l,i)=>{const acc=plan.find(a=>a.code===l.cuenta);return(
                        <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 90px 90px",gap:4,padding:"2px 0",color:l.debe?"#6C9CFF":l.haber?"#4ADE80":"#888"}}>
                          <span style={{paddingLeft:l.haber?16:0}}>{acc?`${acc.code} ${acc.name}`:l.cuenta}</span>
                          <span style={{textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{l.debe?fmt(l.debe):""}</span>
                          <span style={{textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{l.haber?fmt(l.haber):""}</span>
                        </div>
                      );})}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </>)}

        {tab==="balance"&&<BalanceReport plan={plan} asientos={asientos}/>}
      </div>
    </div>
  );
}

// Client View
function ClientView({client,onUpdate,onBack}){
  const[sp,setSp]=useState(null);const[dt,setDt]=useState("emitidos");const[view,setView]=useState("fiscal");const[sa,setSa]=useState(client.saldoAnterior||0);const[dr,setDr]=useState(null);const[showMan,setShowMan]=useState(false);const[iibbPeriod,setIibbPeriod]=useState(null);
  const allRows=useMemo(()=>[...(client.emitidos||[]),...(client.recibidos||[])],[client]);
  const existingKeys=useMemo(()=>new Set(allRows.map(r=>makeDK(r))),[allRows]);
  const handleManualAdd=useCallback(row=>{onUpdate({...client,[row.fileType]:[...(client[row.fileType]||[]),row]});},[client,onUpdate]);
  const rp=client.retPerc||{};
  const dups=useMemo(()=>{const seen={};const d=[];allRows.forEach(r=>{const dk=makeDK(r);if(seen[dk])d.push(r);else seen[dk]=true;});return d;},[allRows]);
  const handleFiles=useCallback(e=>{e.preventDefault();const files=e.dataTransfer?.files||e.target?.files;if(!files?.length)return;const fl=[];let ld=0;for(const f of files){const r=new FileReader();r.onload=ev=>{fl.push({buf:ev.target.result,name:f.name});ld++;if(ld===files.length)processBatch(fl);};r.readAsArrayBuffer(f);}},[client]);
  const processBatch=useCallback(fileList=>{const allEx=new Set([...(client.emitidos||[]).map(r=>makeDK(r)),...(client.recibidos||[]).map(r=>makeDK(r))]);const allNew=[],allDups=[],fileResults=[],seenB=new Set();
    for(const{buf,name}of fileList){try{const raw=new Uint8Array(buf);let dtt="emitidos";const nl=name.toLowerCase();if(nl.includes("recibido"))dtt="recibidos";else if(!nl.includes("emitido")){const hd=raw.slice(0,4);const isC=!(hd[0]===0x50&&hd[1]===0x4B)&&!(hd[0]===0xD0&&hd[1]===0xCF);let ht="";if(isC)try{ht=new TextDecoder("utf-8").decode(raw.slice(0,500));}catch{}else try{const wb=XLSX.read(raw,{type:"array"});ht=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""}).slice(0,5).flat().join(" ");}catch{}if(ht.toLowerCase().includes("emisor"))dtt="recibidos";}
      const parsed=processFile(raw,dtt,name);let nc=0,dc=0;for(const r of parsed){if(allEx.has(r.dedupKey)||seenB.has(r.dedupKey)){allDups.push(r);dc++;}else{allNew.push(r);seenB.add(r.dedupKey);nc++;}}fileResults.push({fileName:name,type:dtt,total:parsed.length,newCount:nc,dupCount:dc});}catch{fileResults.push({fileName:name,type:"error",total:0,newCount:0,dupCount:0});}}
    setDr({newRows:allNew,duplicates:allDups,fileResults});},[client]);
  const confirmUpload=useCallback(newRows=>{onUpdate({...client,emitidos:[...(client.emitidos||[]),...newRows.filter(r=>r.fileType==="emitidos")],recibidos:[...(client.recibidos||[]),...newRows.filter(r=>r.fileType==="recibidos")]});setDr(null);},[client,onUpdate]);
  const handleSaldo=useCallback(v=>{setSa(v);onUpdate({...client,saldoAnterior:v});},[client,onUpdate]);
  const handleRetPerc=useCallback(r=>{onUpdate({...client,retPerc:r});},[client,onUpdate]);
  const monthly=useMemo(()=>{const map={};allRows.forEach(r=>{if(!map[r.key])map[r.key]={debito:0,credito:0,countEm:0,countRec:0};const m=map[r.key];const s=r.isNC?-1:1;if(r.fileType==="emitidos"){m.debito+=r.iva*s;m.countEm++;}else{m.credito+=r.iva*s;m.countRec++;}});return Object.entries(map).sort(([a],[b])=>a.localeCompare(b)).map(([key,v])=>({key,label:pLabel(key),...v}));},[allRows]);
  const accumulated=useMemo(()=>{let arr=sa<0?sa:0;return monthly.map(m=>{const rpv=rp[m.key]||{ret:0,perc:0};const ret=rpv.ret||0,perc=rpv.perc||0;const sm=m.debito-m.credito-ret-perc;const pos=sm+arr;const next=pos<0?pos:0;const res={...m,ret,perc,saldoMes:sm,arrastre:arr,posicion:pos};arr=next;return res;});},[monthly,sa,rp]);
  const totals=useMemo(()=>accumulated.reduce((a,m)=>({debito:a.debito+m.debito,credito:a.credito+m.credito,ret:a.ret+m.ret,perc:a.perc+m.perc}),{debito:0,credito:0,ret:0,perc:0}),[accumulated]);
  const sel=sp?accumulated.find(m=>m.key===sp):null;const hasData=allRows.length>0;const emC=(client.emitidos||[]).length,recC=(client.recibidos||[]).length;
  const allPeriods=useMemo(()=>[...new Set(monthly.map(m=>m.key))].sort(),[monthly]);
  const loadedP=useMemo(()=>{const p={};allRows.forEach(r=>{if(!p[r.key])p[r.key]={em:0,rec:0};if(r.fileType==="emitidos")p[r.key].em++;else p[r.key].rec++;});return Object.entries(p).sort(([a],[b])=>a.localeCompare(b)).map(([k,v])=>({key:k,label:pLabel(k),...v}));},[allRows]);
  const clearAll=()=>{if(confirm(`¿Eliminar TODOS los datos de ${client.name}?`))onUpdate({...client,emitidos:[],recibidos:[],retPerc:{}});};
  const fixDups=useCallback(()=>{const fix=rows=>{const s=new Set();return rows.filter(r=>{const dk=makeDK(r);if(s.has(dk))return false;s.add(dk);return true;});};if(confirm(`Eliminar ${dups.length} duplicados?`))onUpdate({...client,emitidos:fix(client.emitidos||[]),recibidos:fix(client.recibidos||[])});},[client,onUpdate,dups]);
  const openFD=()=>{const i=document.createElement("input");i.type="file";i.accept=".xls,.xlsx,.csv";i.multiple=true;i.onchange=handleFiles;i.click();};
  const lastAcc=accumulated.length?accumulated[accumulated.length-1]:null;

  return(<div>
    {dr&&<DedupModal report={dr} onConfirm={confirmUpload} onCancel={()=>setDr(null)}/>}
    {iibbPeriod&&<IIBBModal emitidos={client.emitidos||[]} period={iibbPeriod} onClose={()=>setIibbPeriod(null)}/>}
    <div style={{marginBottom:24}}><TopBar onHome={onBack}/><h2 style={{fontSize:20,fontWeight:700,margin:0}}>{client.name}</h2>
      <div style={{display:"flex",gap:12,alignItems:"center",marginTop:6,flexWrap:"wrap"}}><span style={{color:"#777",fontSize:13}}>CUIT: {client.cuit}</span>
        <div style={{display:"flex",gap:4}}>{[["RI","Resp. Inscripto"],["MT","Monotributo"]].map(([r,l])=>(<button key={r} onClick={()=>onUpdate({...client,regimen:r})} style={{padding:"3px 10px",fontSize:11,fontWeight:600,borderRadius:4,cursor:"pointer",border:"none",background:(client.regimen||"RI")===r?(r==="RI"?"#6C9CFF":"#FBBF24")+"22":"transparent",color:(client.regimen||"RI")===r?(r==="RI"?"#6C9CFF":"#FBBF24"):"#555"}}>{l}</button>))}</div>
        {/* View toggle */}
        <div style={{display:"flex",gap:4,marginLeft:8}}>{[["fiscal","Impositivo"],["contab","Contabilidad"]].map(([v,l])=>(<button key={v} onClick={()=>setView(v)} style={{padding:"3px 12px",fontSize:11,fontWeight:600,borderRadius:4,cursor:"pointer",border:"none",background:view===v?"#6C9CFF22":"transparent",color:view===v?"#6C9CFF":"#555"}}>{l}</button>))}</div>
        {client.regimen==="MT"&&<select value={client.monoCat||"A"} onChange={e=>onUpdate({...client,monoCat:e.target.value})} style={{background:"#0a0a14",border:"1px solid #333",borderRadius:4,color:"#FBBF24",padding:"3px 8px",fontSize:11,fontWeight:600,outline:"none"}}>{MONO_SCALES.map(s=>(<option key={s.cat} value={s.cat}>Cat. {s.cat} ({fmtShort(s.tope)})</option>))}</select>}</div></div>
    {dups.length>0&&<div style={{background:"#F871711A",border:"1px solid #F8717133",borderRadius:10,padding:"12px 16px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}><div style={{fontSize:13,fontWeight:600,color:"#F87171"}}>⚠ {dups.length} duplicado{dups.length!==1?"s":""}</div><button onClick={fixDups} style={{background:"#F87171",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Limpiar</button></div>}
    <div onDragOver={e=>{e.preventDefault();e.currentTarget.style.borderColor="#6C9CFF";}} onDragLeave={e=>{e.currentTarget.style.borderColor="#2a2a40";}} onDrop={e=>{e.currentTarget.style.borderColor="#2a2a40";handleFiles(e);}} style={{background:"#12122a",border:"2px dashed #2a2a40",borderRadius:12,padding:"20px",marginBottom:20,transition:"all .2s"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}><div><div style={{fontSize:14,fontWeight:600,marginBottom:4}}>📁 Agregar archivos de ARCA</div><div style={{fontSize:12,color:"#777"}}>Arrastrá o subí archivos. Auto-detecta emitidos/recibidos.</div></div>
        <div style={{display:"flex",gap:8,flexShrink:0}}><button onClick={()=>setShowMan(!showMan)} style={{background:showMan?"#FBBF2422":"transparent",color:showMan?"#FBBF24":"#888",border:"1px solid "+(showMan?"#FBBF2444":"#333"),borderRadius:8,padding:"10px 16px",fontSize:13,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>✏️ CAI</button><button onClick={openFD} style={{background:"#6C9CFF",color:"#fff",border:"none",borderRadius:8,padding:"10px 20px",fontSize:13,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>+ Subir archivos</button></div></div>
      {loadedP.length>0&&<div style={{marginTop:16,paddingTop:14,borderTop:"1px solid #2a2a40"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontSize:12,fontWeight:600,color:"#888"}}>Períodos cargados</div><div style={{display:"flex",gap:8,fontSize:11}}><span style={{color:"#6C9CFF"}}>{emC} em.</span><span style={{color:"#555"}}>·</span><span style={{color:"#4ADE80"}}>{recC} rec.</span><span style={{color:"#555"}}>·</span><button onClick={clearAll} style={{background:"none",border:"none",color:"#F8717166",cursor:"pointer",fontSize:11,padding:0}}>Limpiar todo</button></div></div>
        <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{loadedP.map(p=>(<div key={p.key} style={{background:"#0a0a14",borderRadius:6,padding:"6px 10px",fontSize:11,display:"flex",gap:8,alignItems:"center",border:"1px solid #1a1a30"}}><span style={{fontWeight:600}}>{p.label}</span><span style={{color:"#6C9CFF"}}>{p.em}E</span><span style={{color:"#4ADE80"}}>{p.rec}R</span></div>))}</div></div>}</div>
    {showMan&&<ManualEntryForm onAdd={handleManualAdd} onClose={()=>setShowMan(false)} existingKeys={existingKeys}/>}
    {hasData&&(<>
      {view==="contab"&&<ContabilidadPanel client={client} onUpdate={onUpdate}/>}
      {view==="fiscal"&&(<>
      {client.regimen==="MT"&&<MonotributoPanel client={client} emitidos={client.emitidos||[]} onIIBB={setIibbPeriod} onSelectPeriod={setSp}/>}
      {(client.regimen||"RI")==="RI"&&(<>
        <RetPercEditor data={rp} onChange={handleRetPerc} periods={allPeriods}/>
        <div style={{background:"#12122a",borderRadius:10,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}><label style={{fontSize:13,fontWeight:500}}>Saldo técnico inicial:</label><div style={{display:"flex",alignItems:"center",gap:4}}><span style={{fontSize:13,color:"#888"}}>$</span><input type="number" value={sa} onChange={e=>handleSaldo(parseFloat(e.target.value)||0)} style={{background:"#0a0a14",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"6px 10px",fontSize:14,width:150,fontVariantNumeric:"tabular-nums",outline:"none"}}/></div><span style={{fontSize:11,color:"#555"}}>Negativo = a favor</span></div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:20}} className="grid-3">{[{label:"Débito Fiscal",value:totals.debito,color:"#6C9CFF"},{label:"Crédito Fiscal",value:totals.credito,color:"#4ADE80"},{label:"Ret.+Perc.",value:totals.ret+totals.perc,color:"#8B5CF6"},{label:lastAcc&&lastAcc.posicion>0?"A pagar":"Saldo a favor",value:lastAcc?Math.abs(lastAcc.posicion):0,color:lastAcc&&lastAcc.posicion>0?"#F87171":"#4ADE80"}].map((c,i)=>(<div key={i} style={{background:"#12122a",borderRadius:10,padding:"14px",borderLeft:`3px solid ${c.color}`}}><div style={{fontSize:10,color:"#777",textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>{c.label}</div><div style={{fontSize:18,fontWeight:700,color:c.color,fontVariantNumeric:"tabular-nums"}}>{fmt(c.value)}</div></div>))}</div>
        <div style={{background:"#12122a",borderRadius:10,padding:"16px 12px 6px",marginBottom:20}}><div style={{fontSize:13,fontWeight:600,marginBottom:12,paddingLeft:4}}>Posición IVA mensual</div><ResponsiveContainer width="100%" height={220}><BarChart data={accumulated} onClick={e=>{if(e?.activeLabel){const m=accumulated.find(x=>x.label===e.activeLabel);if(m)setSp(m.key);}}}><XAxis dataKey="label" tick={{fontSize:11,fill:"#777"}} axisLine={{stroke:"#333"}} tickLine={false}/><YAxis tick={{fontSize:10,fill:"#555"}} axisLine={false} tickLine={false} tickFormatter={v=>`$${(v/1000).toFixed(0)}k`}/><Tooltip contentStyle={{background:"#1a1a30",border:"1px solid #333",borderRadius:8,fontSize:12}} formatter={v=>[fmt(v),""]} labelStyle={{color:"#aaa"}}/><Legend wrapperStyle={{fontSize:11,paddingTop:6}}/><ReferenceLine y={0} stroke="#444"/><Bar dataKey="debito" name="Débito" fill="#6C9CFF" radius={[3,3,0,0]} cursor="pointer"/><Bar dataKey="credito" name="Crédito" fill="#4ADE80" radius={[3,3,0,0]} cursor="pointer"/><Bar dataKey="posicion" name="Posición" fill="#F87171" radius={[3,3,0,0]} cursor="pointer"/></BarChart></ResponsiveContainer></div>
        <div style={{background:"#12122a",borderRadius:10,overflow:"hidden",marginBottom:20}}><div style={{padding:"12px 16px 8px",fontSize:13,fontWeight:600}}>Liquidación por período</div>
          <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:860}}><thead><tr style={{borderBottom:"1px solid #2a2a40"}}>{["Período","Em.","Rec.","Débito","Crédito","Ret.","Perc.","Saldo Mes","Arrastre","Posición","Resultado",""].map(h=>(<th key={h} style={{padding:"7px 8px",textAlign:h==="Período"||h===""?"left":"right",color:"#666",fontWeight:500,fontSize:10,textTransform:"uppercase",letterSpacing:.3}}>{h}</th>))}</tr></thead>
            <tbody>{accumulated.map(m=>(<tr key={m.key} onClick={()=>setSp(m.key===sp?null:m.key)} style={{borderBottom:"1px solid #1a1a30",cursor:"pointer",background:sp===m.key?"#1e1e40":"transparent"}}><td style={{padding:"7px 8px",fontWeight:600}}>{m.label}</td><td style={{padding:"7px 8px",textAlign:"right",color:"#777"}}>{m.countEm}</td><td style={{padding:"7px 8px",textAlign:"right",color:"#777"}}>{m.countRec}</td><td style={{padding:"7px 8px",textAlign:"right",color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(m.debito)}</td><td style={{padding:"7px 8px",textAlign:"right",color:"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{fmt(m.credito)}</td><td style={{padding:"7px 8px",textAlign:"right",color:m.ret?"#8B5CF6":"#333",fontVariantNumeric:"tabular-nums"}}>{m.ret?fmt(m.ret):"—"}</td><td style={{padding:"7px 8px",textAlign:"right",color:m.perc?"#A78BFA":"#333",fontVariantNumeric:"tabular-nums"}}>{m.perc?fmt(m.perc):"—"}</td><td style={{padding:"7px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:m.saldoMes>0?"#F87171":"#4ADE80"}}>{fmt(m.saldoMes)}</td><td style={{padding:"7px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:m.arrastre<0?"#FBBF24":"#333"}}>{m.arrastre<0?fmt(m.arrastre):"—"}</td><td style={{padding:"7px 8px",textAlign:"right",fontWeight:600,fontVariantNumeric:"tabular-nums",color:m.posicion>0?"#F87171":"#4ADE80"}}>{fmt(m.posicion)}</td><td style={{padding:"7px 8px",textAlign:"right",fontSize:11,fontWeight:600,color:m.posicion>0?"#F87171":"#4ADE80"}}>{m.posicion>0?"A pagar":"A favor"}</td>
              <td style={{padding:"7px 4px"}}>{m.countEm>0&&<button onClick={e=>{e.stopPropagation();setIibbPeriod(m.key);}} style={{background:"#8B5CF622",color:"#A78BFA",border:"1px solid #8B5CF633",borderRadius:4,padding:"3px 8px",fontSize:10,fontWeight:600,cursor:"pointer"}}>IIBB</button>}</td></tr>))}
              <tr style={{borderTop:"2px solid #333"}}><td style={{padding:"8px",fontWeight:700}}>TOTALES</td><td style={{padding:"8px",textAlign:"right",color:"#777"}}>{accumulated.reduce((s,m)=>s+m.countEm,0)}</td><td style={{padding:"8px",textAlign:"right",color:"#777"}}>{accumulated.reduce((s,m)=>s+m.countRec,0)}</td><td style={{padding:"8px",textAlign:"right",fontWeight:700,color:"#6C9CFF",fontVariantNumeric:"tabular-nums"}}>{fmt(totals.debito)}</td><td style={{padding:"8px",textAlign:"right",fontWeight:700,color:"#4ADE80",fontVariantNumeric:"tabular-nums"}}>{fmt(totals.credito)}</td><td style={{padding:"8px",textAlign:"right",fontWeight:700,color:"#8B5CF6",fontVariantNumeric:"tabular-nums"}}>{totals.ret?fmt(totals.ret):"—"}</td><td style={{padding:"8px",textAlign:"right",fontWeight:700,color:"#A78BFA",fontVariantNumeric:"tabular-nums"}}>{totals.perc?fmt(totals.perc):"—"}</td><td colSpan={2}></td><td style={{padding:"8px",textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums",color:lastAcc&&lastAcc.posicion>0?"#F87171":"#4ADE80"}}>{lastAcc?fmt(lastAcc.posicion):"—"}</td><td colSpan={2}></td></tr></tbody></table></div></div>
      </>)}
      </>)}
      {sel&&view==="fiscal"&&(<div style={{background:"#12122a",borderRadius:10,overflow:"hidden",marginBottom:20}}><div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px",borderBottom:"1px solid #2a2a40"}}><div style={{fontSize:14,fontWeight:600}}>Comprobantes — {sel.label}</div><button onClick={()=>setSp(null)} style={{background:"none",border:"none",color:"#777",cursor:"pointer",fontSize:20,lineHeight:1}}>×</button></div>
        <div style={{display:"flex",borderBottom:"1px solid #2a2a40"}}>{[["emitidos","Emitidos","#6C9CFF"],["recibidos","Recibidos","#4ADE80"]].map(([id,l,c])=>(<button key={id} onClick={()=>setDt(id)} style={{flex:1,padding:"9px",background:dt===id?c+"12":"transparent",border:"none",borderBottom:dt===id?`2px solid ${c}`:"2px solid transparent",color:dt===id?c:"#777",cursor:"pointer",fontSize:13,fontWeight:500}}>{l} ({id==="emitidos"?sel.countEm:sel.countRec})</button>))}</div>
        <MonthDetail rows={allRows} period={sp} type={dt}/></div>)}
    </>)}
    {!hasData&&<div style={{textAlign:"center",padding:"32px 20px",color:"#555",fontSize:13}}>Subí archivos de ARCA para ver la liquidación.</div>}
  </div>);}

// Dashboard
function Dashboard({clients,onSelect,onAdd,onDelete}){const[sf,setSf]=useState(false);const[name,setName]=useState("");const[cuit,setCuit]=useState("");const[search,setSearch]=useState("");
  const handleAdd=()=>{if(!name.trim()||!cuit.trim())return;onAdd({name:name.trim(),cuit:cuit.trim()});setName("");setCuit("");setSf(false);};
  const cl=Object.entries(clients).map(([id,c])=>({id,...c})).filter(c=>{if(!search)return true;const s=search.toLowerCase();return c.name.toLowerCase().includes(s)||c.cuit.includes(s);}).sort((a,b)=>a.name.localeCompare(b.name));
  const summaries=cl.map(c=>{const em=c.emitidos||[],rec=c.recibidos||[];const emP=[...new Set(em.map(r=>r.key))].sort();const recP=[...new Set(rec.map(r=>r.key))].sort();return{...c,totalComp:em.length+rec.length,emCount:em.length,recCount:rec.length,lastEm:emP.length?pLabel(emP[emP.length-1]):null,lastRec:recP.length?pLabel(recP[recP.length-1]):null};});
  return(<div><TopBar/>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:24,flexWrap:"wrap",gap:10}}><div><h1 style={{fontSize:20,fontWeight:700,margin:0}}>Clientes</h1><p style={{color:"#777",fontSize:13,marginTop:3}}>{cl.length} cliente{cl.length!==1?"s":""}</p></div><button onClick={()=>setSf(!sf)} style={{background:"#6C9CFF",color:"#fff",border:"none",borderRadius:8,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>+ Nuevo cliente</button></div>
    {sf&&<div style={{background:"#12122a",borderRadius:10,padding:"18px",marginBottom:20,display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}><div style={{flex:"1 1 200px"}}><label style={{fontSize:11,color:"#888",display:"block",marginBottom:4}}>Razón Social</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="Nombre" onKeyDown={e=>e.key==="Enter"&&handleAdd()} style={{width:"100%",background:"#0a0a14",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"8px 12px",fontSize:14,outline:"none"}}/></div><div style={{flex:"0 0 180px"}}><label style={{fontSize:11,color:"#888",display:"block",marginBottom:4}}>CUIT</label><input value={cuit} onChange={e=>setCuit(e.target.value)} placeholder="XX-XXXXXXXX-X" onKeyDown={e=>e.key==="Enter"&&handleAdd()} style={{width:"100%",background:"#0a0a14",border:"1px solid #333",borderRadius:6,color:"#e8e8e8",padding:"8px 12px",fontSize:14,outline:"none"}}/></div><button onClick={handleAdd} disabled={!name.trim()||!cuit.trim()} style={{background:name.trim()&&cuit.trim()?"#4ADE80":"#333",color:name.trim()&&cuit.trim()?"#000":"#666",border:"none",borderRadius:6,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:name.trim()&&cuit.trim()?"pointer":"default"}}>Agregar</button><button onClick={()=>setSf(false)} style={{background:"none",border:"1px solid #333",borderRadius:6,padding:"8px 14px",color:"#888",fontSize:13,cursor:"pointer"}}>Cancelar</button></div>}
    {cl.length>3&&<div style={{marginBottom:16}}><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por nombre o CUIT..." style={{width:"100%",background:"#12122a",border:"1px solid #222",borderRadius:8,color:"#e8e8e8",padding:"10px 14px",fontSize:14,outline:"none"}}/></div>}
    {cl.length===0?<div style={{textAlign:"center",padding:"56px 20px",color:"#555",border:"1px solid #1a1a30",borderRadius:12}}><div style={{fontSize:40,marginBottom:14}}>🏢</div><div style={{fontWeight:600,color:"#888",fontSize:15}}>Agregá tu primer cliente</div></div>
    :<div style={{display:"grid",gap:8}}>{summaries.map(c=>(<div key={c.id} onClick={()=>onSelect(c.id)} style={{background:"#12122a",borderRadius:10,padding:"14px 18px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,transition:"background .15s",flexWrap:"wrap"}} onMouseEnter={e=>e.currentTarget.style.background="#181840"} onMouseLeave={e=>e.currentTarget.style.background="#12122a"}>
      <div style={{minWidth:0,flex:"1 1 200px"}}><div style={{fontWeight:600,fontSize:14}}>{c.name}</div><div style={{fontSize:12,color:"#777",marginTop:2}}>CUIT: {c.cuit} · <span style={{color:(c.regimen||"RI")==="MT"?"#FBBF24":"#6C9CFF",fontWeight:600}}>{c.regimen==="MT"?`Mono ${c.monoCat||"A"}`:"RI"}</span></div></div>
      <div style={{display:"flex",gap:16,alignItems:"center",flexShrink:0}}>{c.totalComp>0?<><div style={{textAlign:"right"}}><div style={{fontSize:10,color:"#777"}}>Emitidos</div><div style={{fontSize:12,color:"#6C9CFF",fontWeight:600}}>{c.emCount} <span style={{fontWeight:400,color:"#555"}}>hasta</span> {c.lastEm||"—"}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:10,color:"#777"}}>Recibidos</div><div style={{fontSize:12,color:"#4ADE80",fontWeight:600}}>{c.recCount} <span style={{fontWeight:400,color:"#555"}}>hasta</span> {c.lastRec||"—"}</div></div></>:<div style={{fontSize:12,color:"#555"}}>Sin datos</div>}
        <button onClick={e=>{e.stopPropagation();if(confirm(`¿Eliminar a ${c.name}?`))onDelete(c.id);}} style={{background:"none",border:"none",color:"#F8717155",cursor:"pointer",fontSize:16,padding:"4px 8px"}} onMouseEnter={e=>e.currentTarget.style.color="#F87171"} onMouseLeave={e=>e.currentTarget.style.color="#F8717155"}>✕</button></div>
    </div>))}</div>}
    <div style={{textAlign:"center",marginTop:24,fontSize:11,color:"#333"}}>FiscalBox · Los datos se guardan en el navegador.</div></div>);}

function App(){const[db,setDb]=useState(loadDB);const[ac,setAc]=useState(null);useEffect(()=>{saveDB(db);},[db]);
  const add=useCallback(({name,cuit})=>{const id=`c_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;setDb(p=>({...p,clients:{...p.clients,[id]:{name,cuit,regimen:"RI",emitidos:[],recibidos:[],retPerc:{},saldoAnterior:0}}}));},[]);
  const upd=useCallback(u=>{setDb(p=>({...p,clients:{...p.clients,[ac]:u}}));},[ac]);
  const del=useCallback(id=>{setDb(p=>{const c={...p.clients};delete c[id];return{...p,clients:c};});if(ac===id)setAc(null);},[ac]);
  return ac&&db.clients[ac]?<ClientView client={db.clients[ac]} onUpdate={upd} onBack={()=>setAc(null)}/>:<Dashboard clients={db.clients} onSelect={setAc} onAdd={add} onDelete={del}/>;}
createRoot(document.getElementById("root")).render(<App/>);
