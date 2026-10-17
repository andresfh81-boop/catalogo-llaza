(function(){"use strict";window.__appStarted=true;
/* =========================================================
   Catálogo Llaza — cliente
   ========================================================= */
const SUBCATS=["Centralitas","Centralitas iluminación","Emisores","Receptores","Sensores","Motores","Fuentes de alimentación","Cables y alargadores","Conectores","Tiras y neón LED","Focos","Tomas"];
const SECS=[{n:"ELYA",c:SUBCATS},{n:"PERGOHABITAT",c:SUBCATS},{n:"CONSUMIBLES",c:["Cintas y adhesivos","Conectores"]}];
/* Subcarpetas por marca dentro de estas categorías. Sin marca guardada se deduce de la descripción o el modelo. */
const MARCAS=["Somfy","Cherubini","Teleco","Nexta"],MARCA_CATS=["Centralitas","Centralitas iluminación","Motores","Emisores","Receptores","Sensores"],OTRA="Otra",OTRAS="Otras marcas";
/* Stock de seguridad: cantidad mínima en almacén fijada en la ficha (vacío = sin fijar). */
const hasStock=i=>i.stock!==""&&i.stock!=null&&!isNaN(i.stock),stockTxt=i=>hasStock(i)?`Stock seg.: ${i.stock}`:"";
const hasMarca=c=>MARCA_CATS.includes(c);
/* Documentación: manuales, guías rápidas y vídeos (hoja «Manuales»; un documento puede servir para varias referencias). */
const DOC_TIPOS=["Instrucciones","Guía rápida","Vídeo","Programación"],DOC_ICO={"Instrucciones":"📄","Guía rápida":"⚡","Vídeo":"▶","Programación":"⚙"};
const docsOf=code=>manuales.filter(m=>m.codes.includes(code));
const docIco=m=>DOC_ICO[m.tipo]||"📄";
const docBtns=i=>{const L=docsOf(i.code);return L.length?`<div class="docs">${L.map(m=>m.url?`<a href="${esc(m.url)}" target="_blank" rel="noopener">${docIco(m)} ${esc(m.titulo)}</a>`:`<span class="tag" title="Falta guardar en el Drive el PDF ${esc(m.archivo)}">⏳ ${esc(m.titulo)}</span>`).join("")}</div>`:"";};
const isSinRef=c=>/^SINREF-/.test(String(c||""));
const DEFSEC=SECS[0].n, OBS="Obsoletas", INC="Incompletas", STK="Stock seguridad", DOC="Documentación", PROP="Fotos por confirmar";
const faltas=i=>[!i.photos.length&&"foto de etiqueta",!i.prod.length&&"foto de producto",!String(i.desc||"").trim()&&"descripción"].filter(Boolean);
const incompleta=i=>!i.obsoleta&&faltas(i).length>0;
const XLSX_URL="https:\/\/cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
      DOCX_URL="https:\/\/cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js",
      JSZIP_URL="https:\/\/cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

let S=null;               /* datos del servidor (boot) */
let refs=[],fotos=[],items=[],revs=[],manuales=[],moveAsk=null;
let sec=DEFSEC,cur="Todas",marca="Todas",view="cat",notFound=null,recog=null,confirmCode=null;
let selMode=false,expOne=null;const selSet=new Set(),expCats=new Set();let expAddOn=true;
let lbList=[],lbI=0,lbItem=null;
let revId=null,revDraft=null,revT=null,sendState="idle",filling=false;const docBusy={};
const busy=new Set(),fullCache={};

const $=s=>document.querySelector(s);
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
/* Enlaces en las notas: se muestran como «Ver vídeo» o «Ver instrucciones». */
const URL_RE=/https?:\/\/[^\s<>"]+/g;
const linkLabel=u=>/youtu\.?be|vimeo|\.mp4(\?|$)/i.test(u)?"▶ Ver vídeo":"📄 Ver instrucciones";
const linkify=t=>{let out="",k=0;String(t||"").replace(URL_RE,(u,pos)=>{out+=esc(t.slice(k,pos))+`<a class="linkout" href="${esc(u)}" target="_blank" rel="noopener">${linkLabel(u)}</a>`;k=pos+u.length;return u;});return out+esc(String(t||"").slice(k));};
const norm=s=>String(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
const hoy=()=>new Date().toISOString().slice(0,10);
const fmtF=iso=>{try{return iso?new Date(iso).toLocaleDateString("es-ES",{day:"2-digit",month:"2-digit",year:"numeric"}):"";}catch(_){return String(iso||"");}};
/* Tras cambiar referencias, las hojas UMARA/PERGOHABITAT se actualizan detrás, sin hacer esperar. */
const SYNC_FNS=new Set(["saveRef","renameRef","moveRefs","setObsolete"]);
const api=(fn,...a)=>new Promise((res,rej)=>google.script.run.withSuccessHandler(r=>{res(r);if(SYNC_FNS.has(fn))google.script.run.withFailureHandler(()=>{}).syncSheetsBg();}).withFailureHandler(rej)[fn](...a));
/* Cantidades con decimales (metros de cable…): «4,850» = 4,85. Se guardan como número y se muestran con coma. */
const numQ=v=>{if(typeof v==="number")return v;let s=String(v||"").trim().replace(/\s/g,"");if(/,/.test(s))s=s.replace(/\./g,"").replace(",",".");const n=parseFloat(s);return isFinite(n)&&n>0?Math.round(n*1000)/1000:0;};
const fmtQ=n=>(+n||0).toLocaleString("es-ES",{maximumFractionDigits:3});
function toast(t,ms=2800){const el=$("#toast");el.textContent=t;el.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>el.hidden=true,ms);}
function errMsg(e){const m=String((e&&e.message)||e||"");
  if(/SIN_CLAVE/.test(m))return "Falta la clave de Gemini. Ponla en ⚙ Ajustes.";
  if(/GEMINI_4(01|03)/.test(m))return "La clave de Gemini no es válida. Revísala en ⚙ Ajustes.";
  if(/GEMINI_404/.test(m))return "Google no tiene ahora ningún modelo de Gemini disponible para tu clave. Deja vacío el campo Modelo en ⚙ Ajustes y vuelve a probar; si sigue, manda captura a Claude.";
  if(/GEMINI_(503|500|504)/.test(m))return "Los servidores de Google están saturados ahora mismo (no es fallo de la app). He probado todos los modelos de Gemini y ninguno responde. Vuelve a intentarlo en unos minutos.";
  if(/GEMINI_429/.test(m))return "Has llegado al límite gratuito de Gemini por ahora. Prueba dentro de un rato.";
  const gm=m.match(/GEMINI_(\d+):?\s*(.*)/);if(gm)return `Gemini no ha podido responder (error ${gm[1]}${gm[2]?": "+gm[2].slice(0,120):""}). Inténtalo otra vez; si se repite, manda captura a Claude.`;
  if(/Service invoked too many times|Límite|quota/i.test(m))return "Se ha superado el límite diario de Google. Prueba mañana.";
  if(/permission|permiso|access denied|No item with the given ID/i.test(m))return "No tienes acceso a los datos. Pide que te compartan la carpeta APP CATALOGO LLAZA.";
  return m.replace(/^Error:\s*/,"")||"Algo ha fallado. Inténtalo otra vez.";}
const ALL="TODAS",secsOf=i=>[i.sec,...(i.also||[])],inS=(i,s)=>s===ALL||secsOf(i).includes(s);
function catsOf(s){const pre=s===ALL?[...new Set(SECS.flatMap(x=>x.c))]:(SECS.find(x=>x.n===s)||{c:[]}).c;
  const extra=[...new Set(items.filter(i=>inS(i,s)).map(i=>i.cat))].filter(c=>!pre.includes(c)).sort((a,b)=>a.localeCompare(b,"es"));
  return [...pre,...extra];}
function marcaOf(i){if(!hasMarca(i.cat))return "";const m=String(i.marca||"").trim();if(m)return MARCAS.find(x=>norm(x)===norm(m))||OTRAS;
  return guessMarca(i.desc+" "+i.modelo)||OTRAS;}
const guessMarca=t=>MARCAS.find(x=>norm(t).includes(norm(x)))||"";
const marcasIn=L=>[...MARCAS,OTRAS].filter(m=>L.some(i=>marcaOf(i)===m));
/* Tarjetas de una categoría; en las de marca, agrupadas en subcarpetas. */
function catBody(c,L){
  if(!hasMarca(c)||marca!=="Todas")return grid(L);
  return marcasIn(L).map(m=>`<h3 class="msub">${esc(m)} <small>${L.filter(i=>marcaOf(i)===m).length}</small></h3>${grid(L.filter(i=>marcaOf(i)===m))}`).join("");}
const itemOf=code=>code?items.find(i=>i.code===code):null;
const thumbOf=it=>{const p=it&&(it.photos[0]||it.prod[0]);return p?p.src:"";};
function loadLib(url,glob){if(window[glob])return Promise.resolve();loadLib.p=loadLib.p||{};if(loadLib.p[url])return loadLib.p[url];
  return loadLib.p[url]=new Promise((res,rej)=>{const s=document.createElement("script");s.src=url;s.onload=()=>window[glob]?res():rej(new Error(glob));
    s.onerror=()=>{delete loadLib.p[url];rej(new Error(glob));};document.head.appendChild(s);});}
function b64(buf){const u=new Uint8Array(buf);let s="";for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));return btoa(s);}
const blobB64=async b=>b64(await b.arrayBuffer());
async function shrink(file,max=1400,q=0.82){let bmp;try{bmp=await createImageBitmap(file,{imageOrientation:"from-image"});}catch(_){bmp=await createImageBitmap(file);}
  const s=Math.min(1,max/Math.max(bmp.width,bmp.height)),c=document.createElement("canvas");c.width=Math.round(bmp.width*s);c.height=Math.round(bmp.height*s);
  c.getContext("2d").drawImage(bmp,0,0,c.width,c.height);return await new Promise((res,rej)=>c.toBlob(b=>b?res(b):rej(new Error("img")),"image/jpeg",q));}

/* ---------- arranque ---------- */
function bootError(title,detail){$("#bootBox").hidden=false;$("#bootBox").innerHTML=`<h2>${esc(title)}</h2>${detail}<button class="btn-l primary" id="bootRetry">Reintentar</button>`;const r=$("#bootRetry");if(r)r.onclick=reboot;}
function reboot(){S=null;$("#app").hidden=true;$("#bnav").hidden=true;$("#bootBox").hidden=false;$("#bootBox").innerHTML=`<div class="spin"></div><p>Cargando el catálogo…</p>`;start();}
window.addEventListener("error",ev=>{if(!S)bootError("Error al abrir la app",`<p class="note">${esc(ev.message||"")}</p>`);});
let bootTimer=setTimeout(()=>{if(!S)bootError("Google no responde",
  `<p>La app no recibe respuesta de tu cuenta de Google. Suele pasar cuando hay <b>varias cuentas de Google abiertas</b> en el mismo navegador.</p>
   <p class="note">Solución: abre este enlace en una <b>ventana de incógnito</b> y entra solo con tu cuenta, o en un perfil de Chrome con una sola cuenta.</p>`);},window.__pwa?900000:25000);
function start(){api("boot").then(d=>{clearTimeout(bootTimer);S=d;window.__appOK=true;
  if(!d.installed){$("#bootBox").innerHTML=d.owner
    ?`<h2>Primera vez</h2><p>Voy a crear la hoja de datos y conectar las fotos de tu carpeta APP CATALOGO LLAZA.</p><button class="btn-l primary" id="inst">Instalar la app</button><p class="err" id="instErr"></p>`
    :`<h2>La app aún no está instalada</h2><p>Pide a Andrés que la abra una vez para instalarla.</p>`;
    const b=$("#inst");if(b)b.onclick=()=>{b.disabled=true;b.textContent="Instalando… (puede tardar un minuto)";
      api("install").then(()=>reboot()).catch(e=>{b.disabled=false;b.textContent="Instalar la app";$("#instErr").textContent=errMsg(e);});};
    return;}
  refs=d.refs;fotos=d.fotos;revs=d.revs||[];manuales=d.manuales||[];
  try{const s=localStorage.getItem("cl_apartado");if(s===ALL||SECS.some(x=>x.n===s))sec=s;}catch(_){}
  $("#bootBox").hidden=true;$("#app").hidden=false;$("#bnav").hidden=false;
  $("#recL").hidden=true;
  merge();
}).catch(e=>{clearTimeout(bootTimer);const m=String((e&&e.message)||e||"");
  bootError("No se ha podido abrir el catálogo",`<p>${esc(errMsg(e))}</p>${/PERMISSION|autoriz|authoriz|ScriptError|server error/i.test(m)
    ?`<p class="note">Si tienes varias cuentas de Google abiertas, prueba en una ventana de incógnito con solo tu cuenta.</p>`:""}<p class="note">Detalle: ${esc(m.slice(0,200))}</p>`);});}

function merge(){
  const own=f=>/^App/.test(f.origen||"");
  const P=f=>({id:f.id,fileId:f.fileId,src:"data:image/jpeg;base64,"+f.mini,rot:f.rot||0,own:own(f),from:/^Web fabricante/.test(f.origen||"")?"web":"drive",page:String(f.origen||"").replace(/^Web fabricante · /,"")});
  const of=(r,t)=>fotos.filter(f=>f.code===r.code&&f.tipo===t).map(P);
  items=refs.map(r=>({...r,photos:of(r,"Etiqueta"),prod:of(r,"Producto"),ofi:of(r,"Oficial"),prop:of(r,"Propuesta")}));
  chips();render();
  if(!$("#lb").hidden&&lbItem){const n=itemOf(lbItem.code),was=lbList[lbI];
    if(n){lbItem=n;buildLb(n);const j=was?lbList.findIndex(p=>p.id===was.id):-1;lbI=j>=0?j:Math.min(lbI,lbList.length-1);if(!lbList.length)$("#lb").hidden=true;else showLb();}}
}

/* ---------- vistas ---------- */
function updateFab(){const b=$("#addBtn");if(selMode||view==="cfg"){b.hidden=true;return;}
  if(view==="rev"){b.textContent="＋ Nueva revisión";b.hidden=!!revId;}else{b.textContent="＋ Referencia";b.hidden=false;}}
function setView(v){view=v;
  $("#tools").hidden=v!=="cat";$("#list").hidden=v!=="cat";$("#rev").hidden=v!=="rev";$("#cfg").hidden=v!=="cfg";$("#expBtn").hidden=v!=="cat";
  $("#bnav").querySelectorAll("[data-go]").forEach(b=>b.setAttribute("aria-pressed",String(b.dataset.go===v)));updateFab();}
$("#bnav").addEventListener("click",e=>{const b=e.target.closest("[data-go]");if(!b)return;
  if(selMode){selSet.clear();selMode=false;$("#selBar").hidden=true;}
  const g=b.dataset.go;if(view==="rev"&&g!=="rev")flushRev();
  if(g==="rev"){if(view==="rev"&&revId){closeRev();return;}setView("rev");revId=null;revDraft=null;renderRev();}
  else if(g==="cfg"){setView("cfg");renderCfg();}
  else{setView("cat");chips();render();}
  window.scrollTo({top:0});});

function chips(){
  const nOpen=revs.filter(r=>r.estado!=="Enviada").length;$("#bnavN").hidden=!nOpen;$("#bnavN").textContent=nOpen;
  $("#secs").innerHTML=[{n:ALL},...SECS].map(s=>`<button data-s="${esc(s.n)}" aria-pressed="${s.n===sec}">${esc(s.n)}<small>${items.filter(i=>inS(i,s.n)&&!i.obsoleta).length}</small></button>`).join("");
  if(view==="rev"){$("#h1").textContent="Revisión de hojas";$("#count").textContent="Lo que falta y lo que sobra en las hojas de trabajo.";$("#mode").textContent="";return;}
  if(view==="cfg"){$("#h1").textContent="Ajustes";$("#count").textContent="";$("#mode").textContent="";return;}
  $("#h1").textContent=sec===ALL?"Todas las referencias":sec;
  const inSec=items.filter(i=>inS(i,sec)),act=inSec.filter(i=>!i.obsoleta),nObs=inSec.length-act.length;
  const pe=act.filter(i=>i.photos.length).length,pp=act.filter(i=>i.prod.length).length;
  $("#count").textContent=act.length?`${act.length} referencias · ${pe} con foto de etiqueta · ${pp} con foto de producto.`:"Todavía no hay referencias en este apartado.";
  $("#mode").textContent="";
  const cats=["Todas",...catsOf(sec).filter(c=>act.some(i=>i.cat===c))];
  const nInc=act.filter(incompleta).length,nStk=act.filter(hasStock).length,nDoc=act.filter(i=>docsOf(i.code).length).length;
  const nProp=act.filter(i=>i.prop.length).length;
  if(cur!==OBS&&cur!==INC&&cur!==STK&&cur!==DOC&&cur!==PROP&&cur!=="Todas"&&!cats.includes(cur))cur="Todas";if(cur===OBS&&!nObs)cur="Todas";if(cur===INC&&!nInc)cur="Todas";if(cur===PROP&&!nProp)cur="Todas";
  $("#chips").innerHTML=cats.map(c=>`<button class="chip" aria-pressed="${c===cur}" data-c="${esc(c)}">${esc(c)}<small>${c==="Todas"?act.length:act.filter(i=>i.cat===c).length}</small></button>`).join("")+
    `<button class="chip doc" aria-pressed="${cur===DOC}" data-c="${DOC}">📚 ${DOC}<small>${nDoc}</small></button>`+
    `<button class="chip stk" aria-pressed="${cur===STK}" data-c="${STK}">📦 ${STK}<small>${nStk}</small></button>`+
    (nProp?`<button class="chip prop" aria-pressed="${cur===PROP}" data-c="${PROP}">📸 ${PROP}<small>${nProp}</small></button>`:"")+
    (nInc?`<button class="chip inc" aria-pressed="${cur===INC}" data-c="${INC}">⚠ ${INC}<small>${nInc}</small></button>`:"")+
    (nObs?`<button class="chip obs" aria-pressed="${cur===OBS}" data-c="${OBS}">${OBS}<small>${nObs}</small></button>`:"");
  const inCat=hasMarca(cur)?act.filter(i=>i.cat===cur):[],ms=marcasIn(inCat);if(!ms.includes(marca))marca="Todas";
  $("#mchips").hidden=!ms.length;$("#mchips").innerHTML=ms.length?["Todas",...ms].map(m=>`<button class="chip" aria-pressed="${m===marca}" data-m="${esc(m)}">${m==="Todas"?"Todas las marcas":esc(m)}<small>${m==="Todas"?inCat.length:inCat.filter(i=>marcaOf(i)===m).length}</small></button>`).join(""):"";
  ["#chips","#mchips"].forEach(id=>{const r=$(id),b=r.querySelector('[aria-pressed="true"]');if(b&&r.scrollWidth>r.clientWidth)r.scrollLeft=b.offsetLeft-r.offsetLeft-(r.clientWidth-b.offsetWidth)/2;});
  updateFab();}
$("#secs").addEventListener("click",e=>{const b=e.target.closest("button[data-s]");if(!b)return;sec=b.dataset.s;cur="Todas";marca="Todas";moveAsk=null;
  try{localStorage.setItem("cl_apartado",sec);}catch(_){}notFound=null;recog=null;confirmCode=null;$("#q").value="";chips();render();});
$("#chips").addEventListener("click",e=>{const b=e.target.closest(".chip");if(!b)return;cur=b.dataset.c;marca="Todas";confirmCode=null;recog=null;chips();render();});
$("#mchips").addEventListener("click",e=>{const b=e.target.closest(".chip");if(!b)return;marca=b.dataset.m;chips();render();});
$("#q").addEventListener("input",()=>{notFound=null;recog=null;fuzzyQ=false;render();});
/* ---------- búsqueda aproximada ----------
   Con la voz (y al escribir, si no hay ninguna coincidencia exacta) se busca por parecido: palabras sueltas en
   cualquier orden, mal oídas o a medias («oblo» → Hoblo), palabras del día a día («mando» = emisor) y trozos de código.
   Sale primero lo que más se parece. */
let fuzzyQ=false;
const STOP=new Set("el la los las de del un una unos unas para con que y o en al lo le mi me ese esa este esta eso busca buscar buscame quiero dame pon ponme ver coma punto metro metros mts m cm mm".split(" "));
const SYN={mando:"emisor",mandos:"emisor",telemando:"emisor",pulsador:"emisor",central:"centralita",centrales:"centralita",cuadro:"centralita",controlador:"centralita",
  anemometro:"viento",veleta:"viento",enchufe:"toma",schuko:"toma",prolongador:"alargador",extension:"alargador",alargo:"alargador",alargadera:"alargador",
  empalme:"conector",clavija:"conector",spot:"foco",ojo:"foco",bombilla:"foco",tira:"led",neon:"led",luz:"led",luces:"led",transformador:"fuente",trafo:"fuente",
  alimentacion:"fuente",piston:"lineal",actuador:"lineal",estilo:"stl",stilo:"stl",blanco:"bl",blanca:"bl",negro:"ng",negra:"ng",aluminio:"al",calido:"calida",cien:"100",doscientos:"200",
  uno:"1",dos:"2",tres:"3",cuatro:"4",cinco:"5",seis:"6",siete:"7",ocho:"8",nueve:"9",diez:"10",motores:"motor",sensores:"sensor",cables:"cable",conectores:"conector",focos:"foco"};
/* Cómo suena: h muda, b = v, ll = y, c/k/qu, z/s, g/j; sin letras dobles (los números se dejan como están). */
const fon=w=>norm(w).replace(/h/g,"").replace(/v/g,"b").replace(/ll/g,"y").replace(/qu|c(?=[aou])|k/g,"k").replace(/z|c(?=[ei])/g,"s").replace(/g(?=[ei])/g,"j").replace(/([a-z])\1+/g,"$1");
function lev(a,b){if(a===b)return 0;let p=Array.from({length:b.length+1},(_,j)=>j);
  for(let i=1;i<=a.length;i++){const c=[i];for(let j=1;j<=b.length;j++)c[j]=Math.min(p[j]+1,c[j-1]+1,p[j-1]+(a[i-1]===b[j-1]?0:1));p=c;}return p[b.length];}
const sim=(a,b)=>{if(a===b)return 1;if(a.length>=3&&b.startsWith(a))return .95;if(a.length>=4&&b.includes(a))return .85;if(a.length<3)return 0;return 1-lev(a,b)/Math.max(a.length,b.length);};
/* Decimales iguales en la hoja y en lo que se dice: «1.50 m», «1,5» y «uno coma cinco» → 1p5; «7.00» → 7. */
const decs=t=>t.replace(/(\d+)\s*[.,]\s*(\d+)/g,(m,a,b)=>{b=b.replace(/0+$/,"");return b?a+"p"+b:a;});
const NUMW={cero:"0",uno:"1",una:"1",dos:"2",tres:"3",cuatro:"4",cinco:"5",seis:"6",siete:"7",ocho:"8",nueve:"9",diez:"10",cien:"100",doscientos:"200"};
const itemWords=i=>i._w||(i._w=[...new Set(decs(norm([i.desc,i.modelo,i.cat,marcaOf(i),i.medida,i.canal,secsOf(i).join(" ")].join(" "))).split(/[^a-z0-9]+/).filter(w=>w.length>1||/\d/.test(w)).map(fon))]);
/* Palabras de la búsqueda; cada una con sus variantes. La voz a veces parte una palabra («a largo» = alargo): la letra suelta se prueba pegada a la siguiente. */
function qToks(qRaw){let q=norm(qRaw).replace(/\b(cero|uno|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|cien|doscientos)\b/g,m=>NUMW[m]);
  q=decs(q.replace(/(\d)\s+(?:coma|punto|con)\s+(\d)/g,"$1,$2"));
  const raw=q.split(/[^a-z0-9]+/).filter(Boolean),out=[];
  for(let k=0;k<raw.length;k++){const t=raw[k];
    if(/^[a-z]$/.test(t)&&k+1<raw.length&&/^[a-z]{2,}$/.test(raw[k+1])){out.push([raw[k+1],t+raw[k+1]]);k++;continue;}
    if(!STOP.has(t))out.push([t]);}
  return out;}
function fuzzy(qRaw,obs){const toks=qToks(qRaw);if(!toks.length)return [];
  const res=items.filter(i=>obs?i.obsoleta:!i.obsoleta).map(i=>{const W=itemWords(i);let s=0,hit=0;
    for(const vs of toks){let b=0;
      for(const t of vs){const tt=/^\d+$/.test(t)?t:/^\d+$/.test(SYN[t]||"")?SYN[t]:null;
        /* Números: medida o potencia («200» → 200W) y, desde 4 cifras, trozo del código. */
        if(tt)b=Math.max(b,W.includes(tt)?1:W.some(w=>w.startsWith(tt))?.9:tt.length>=4&&i.code.includes(tt)?1:0);
        else for(const x of SYN[t]?[fon(t),fon(SYN[t])]:[fon(t)])for(const w of W)b=Math.max(b,sim(x,w));}
      if(b>=.72){s+=b;hit++;}}
    return {i,s,hit};}).filter(x=>x.hit).sort((a,b)=>b.hit-a.hit||b.s-a.s);
  /* Primero las que encajan en la mitad de las palabras o más; si no hay ninguna, las que encajan en alguna. */
  const need=Math.ceil(toks.length/2),best=res.filter(x=>x.hit>=need);
  return (best.length?best:res).slice(0,12).map(x=>x.i);}


/* ---------- buscar por voz ----------
   Solo en la app instalable: dentro del /exec Google no deja usar el micrófono (la página va en un marco).
   Chrome pasa la voz a texto con los servidores de Google, así que necesita conexión. */
const SR=window.SpeechRecognition||window.webkitSpeechRecognition;let rec=null;
const NUMS={cero:0,uno:1,una:1,dos:2,tres:3,cuatro:4,cinco:5,seis:6,siete:7,ocho:8,nueve:9};
/* «cuatro ocho cero cero…» o «48 009 000 000» → 48009000000; lo demás se busca tal cual. */
function voiceText(t){t=String(t||"").trim().replace(/[.,;]+$/,"");const w=norm(t).split(/\s+/).filter(Boolean);
  if(w.length&&w.every(x=>x in NUMS||/^\d+$/.test(x)))return w.map(x=>x in NUMS?NUMS[x]:x).join("");
  /* Números dichos en medio de una frase («la de nueve cero cero cero»): dos o más seguidos se juntan en cifras. */
  t=t.replace(/(?:\b(?:cero|uno|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|\d+)\b[\s.-]*){2,}/gi,m=>m.trim().split(/[\s.-]+/).map(x=>NUMS[norm(x)]??x).join("")+" ").trim();
  return t.replace(/(\d)[\s.-]+(?=\d)/g,"$1");}
function micState(){const b=$("#micBtn");b.hidden=!SR||!window.__pwa;b.disabled=!navigator.onLine;
  b.title=navigator.onLine?"Buscar por voz":"Sin conexión: buscar por voz necesita internet";}
window.addEventListener("online",micState);window.addEventListener("offline",micState);micState();
$("#micBtn").onclick=()=>{if(rec){rec.stop();return;}
  const b=$("#micBtn");rec=new SR();rec.lang="es-ES";rec.interimResults=true;rec.maxAlternatives=1;
  b.classList.add("listen");b.setAttribute("aria-label","Dejar de escuchar");toast("Te escucho… di un código o el nombre de la pieza",8000);
  rec.onresult=e=>{const r=e.results[e.results.length-1],t=voiceText(r[0].transcript);$("#q").value=t;notFound=null;recog=null;fuzzyQ=true;
    if(!r.isFinal){render();return;}
    const c=pick(t);if(c&&itemOf(c)){showCode(c);return;}
    render();const f=fuzzy(t);if(f.length===1){openFicha(f[0].code);return;}
    toast(f.length?`«${t}»: ${f.length} parecida${f.length>1?"s":""}, la primera es la que más se parece.`:`No encuentro nada parecido a «${t}».`,3500);};
  rec.onerror=e=>toast(e.error==="not-allowed"||e.error==="service-not-allowed"?"Chrome no tiene permiso para usar el micrófono. Toca el candado de la barra de direcciones → Permisos → Micrófono.":
    e.error==="no-speech"?"No te he oído. Toca 🎤 y habla cerca del móvil.":e.error==="network"?"Para buscar por voz hace falta conexión.":`No he podido usar el micrófono (${e.error}).`,5500);
  rec.onend=()=>{rec=null;b.classList.remove("listen");b.setAttribute("aria-label","Buscar por voz");};
  try{rec.start();}catch(_){rec.onend();}};

function render(){renderList();if(fichaCode)renderFicha();}
function renderList(){
  if(recog){renderRecog();return;}
  if(notFound&&$("#q").value===notFound.code){
    $("#list").innerHTML=`<div class="nf"><b>${esc(notFound.code)}</b><p class="sub">Este código no está en ningún apartado.</p><button class="btn-l primary" id="nfAdd">Dar de alta con esta foto</button></div>`;
    $("#nfAdd").onclick=()=>openAdd(notFound.code,notFound.file);return;}
  const qRaw=$("#q").value.trim(),q=norm(qRaw),match=i=>norm([i.code,i.desc,i.modelo,i.medida,i.canal,i.cat,marcaOf(i)].join(" ")).includes(q);
  if(q){let vis=items.filter(i=>(cur===OBS?i.obsoleta:!i.obsoleta)&&match(i));
    if(fuzzyQ||!vis.length){const f=fuzzy(qRaw,cur===OBS);
      if(f.length){$("#list").innerHTML=`<p class="note" style="margin:0 0 12px">${fuzzyQ?"Lo que más se parece":"Ninguna coincide exacta. Lo más parecido"} a «${esc(qRaw)}», de más a menos:</p>${grid(f)}`;return;}}
    if(!vis.length){$("#list").innerHTML=`<p class="empty">Ninguna referencia coincide con «${esc(qRaw)}».</p>`;return;}
    $("#list").innerHTML=SECS.flatMap(s=>catsOf(s.n).map(c=>[s.n,c])).filter(([s,c])=>vis.some(i=>inS(i,s)&&i.cat===c))
      .map(([s,c])=>`<section><h2>${esc(s)} · ${esc(c)}</h2>${grid(vis.filter(i=>inS(i,s)&&i.cat===c))}</section>`).join("");return;}
  const inSec=items.filter(i=>inS(i,sec));
  if(cur===OBS){const own=inSec.filter(i=>i.obsoleta&&i.sec===sec),oth=SECS.filter(x=>x.n!==sec);
    const mv=!own.length?"":moveAsk?`<p class="note" style="margin:0 0 12px"><b>¿Pasar las ${own.length} obsoletas de ${esc(sec)} a ${esc(moveAsk)}?</b> Siguen siendo obsoletas. <button class="btn-s" data-mv-no>No</button> <button class="btn-s yes" data-mv-yes="${esc(moveAsk)}">Sí, pasarlas</button></p>`
      :`<p style="margin:0 0 12px;display:flex;gap:8px;flex-wrap:wrap">${oth.map(x=>`<button class="btn-s" data-mv="${esc(x.n)}">➜ Pasar todas a ${esc(x.n)}</button>`).join("")}</p>`;
    $("#list").innerHTML=mv+`<section><h2>Referencias obsoletas · no salen en el catálogo</h2>${grid(inSec.filter(i=>i.obsoleta))}</section>`;return;}
  if(cur===DOC){const L=inSec.filter(i=>!i.obsoleta&&docsOf(i.code).length);
    $("#list").innerHTML=`<p class="note" style="margin:0 0 12px">${L.length?`<b>${L.length}</b> referencia${L.length===1?"":"s"} con manuales, guías o vídeos.`:"Todavía no hay documentación en este apartado."} Para añadir más: «✏ Editar» en la ficha → 📚 Documentación.</p>`+
      catsOf(sec).filter(c=>L.some(i=>i.cat===c)).map(c=>`<section><h2>${esc(c)}</h2>${grid(L.filter(i=>i.cat===c))}</section>`).join("");return;}
  if(cur===PROP){const L=inSec.filter(i=>!i.obsoleta&&i.prop.length);
    $("#list").innerHTML=`<p class="note" style="margin:0 0 12px"><b>${L.length}</b> referencia${L.length===1?"":"s"} con foto oficial del fabricante por confirmar. Toca cada una y elige «✓ Usar esta foto» o «✕ Descartar».</p>`+
      catsOf(sec).filter(c=>L.some(i=>i.cat===c)).map(c=>`<section><h2>${esc(c)}</h2>${grid(L.filter(i=>i.cat===c))}</section>`).join("");return;}
  if(cur===STK){const L=inSec.filter(i=>!i.obsoleta&&hasStock(i));
    $("#list").innerHTML=`<p class="note" style="margin:0 0 12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><span>${L.length?`<b>${L.length}</b> referencia${L.length===1?"":"s"} con stock de seguridad fijado.`:"Ninguna referencia tiene todavía stock de seguridad. Ponlo en «✏ Editar» de cada ficha."}</span>${L.length?`<button class="btn-s" id="stkExp">⬇ Exportar lista</button>`:""}</p>`+
      catsOf(sec).filter(c=>L.some(i=>i.cat===c)).map(c=>`<section><h2>${esc(c)}</h2>${grid(L.filter(i=>i.cat===c))}</section>`).join("");
    const b=$("#stkExp");if(b)b.onclick=()=>openExp();return;}
  if(cur===INC){const L=inSec.filter(incompleta),n=k=>L.filter(i=>faltas(i).includes(k)).length;
    $("#list").innerHTML=`<p class="note" style="margin:0 0 12px">Referencias incompletas: <b>${n("foto de etiqueta")}</b> sin foto de etiqueta, <b>${n("foto de producto")}</b> sin foto de producto y <b>${n("descripción")}</b> sin descripción. Usa «Hacer foto» o «✏ Editar» para completarlas.</p>`+
      catsOf(sec).filter(c=>L.some(i=>i.cat===c)).map(c=>`<section><h2>${esc(c)}</h2><div class="grid">${L.filter(i=>i.cat===c).map(x=>card(x)).join("")}</div></section>`).join("");return;}
  const vis=inSec.filter(i=>!i.obsoleta&&(cur==="Todas"||i.cat===cur)&&(marca==="Todas"||marcaOf(i)===marca));
  if(!vis.length){$("#list").innerHTML=`<p class="empty">Todavía no hay referencias en ${esc(sec)}. Pulsa «＋ Referencia» para añadir la primera.</p>`;return;}
  $("#list").innerHTML=catsOf(sec).filter(c=>vis.some(i=>i.cat===c)).map(c=>`<section><h2>${esc(c)}${marca!=="Todas"?" · "+esc(marca):""}</h2>${
    catBody(c,vis.filter(i=>i.cat===c).sort((a,b)=>b.nueva-a.nueva))}</section>`).join("");}

function pane(i,kind){
  const list=kind==="e"?i.photos:i.prod,lab=kind==="e"?"Etiqueta":"Producto",key=i.code+kind;
  if(busy.has(key))return `<div class="ph miss busy"><span class="lab">${lab}</span><span>Subiendo foto…</span></div>`;
  if(!list.length){if(!i.obsoleta)return `<label class="ph miss take"><span class="lab">${lab}</span><span class="cam" aria-hidden="true">📷</span><span>Hacer foto</span>
      <input type="file" accept="image/*" capture="environment" data-pick data-multi="1" data-code="${i.code}" data-k="${kind}" aria-label="Hacer foto de ${lab.toLowerCase()} de ${esc(i.desc)}"></label>`;
    return `<div class="ph miss"><span class="lab">${lab}</span><span>Falta foto</span></div>`;}
  const web=kind==="p"&&list[0].from==="web",n=list.length;
  const im=(p,j)=>`<img src="${p.src}" alt="${j?"":`${lab} ${esc(i.desc)} ${i.code}`}" loading="lazy" class="${p.from==="web"?"web ":""}${p.rot?"r"+p.rot:""}">`;
  const inner=n===1?im(list[0],0):list.slice(0,4).map((p,j)=>{const m=Math.min(n,4);
    return `<span class="mc${(m===2||(m===3&&j===0))?" tall":""}">${im(p,j)}${j===3&&n>4?`<span class="more">+${n-4}</span>`:""}</span>`;}).join("");
  const add=!i.obsoleta?`<label class="addmini" title="Añadir otra foto">＋<input type="file" accept="image/*" capture="environment" data-pick data-multi="1" data-code="${i.code}" data-k="${kind}" aria-label="Añadir otra foto de ${lab.toLowerCase()} de ${esc(i.desc)}"></label>`:"";
  return `<div class="cell"><button class="ph${n>1?` mosaic m${Math.min(n,4)}`:" single"}" data-code="${i.code}" data-k="${kind}" aria-label="Ver fotos de ${lab.toLowerCase()} de ${esc(i.desc)}">
    ${inner}<span class="lab">${lab}${web?" · web":""}</span>${n>1&&n<=4?`<span class="n">${n}</span>`:""}</button>${add}</div>`;}
function card(i,ex){
  let foot="";
  if(ex)foot=!i.obsoleta?`<button class="btn-s pick" data-pick="${i.code}">Es este · guardar la foto</button>`:"";
  else if(i.obsoleta)foot=`<button class="btn-s" data-edit="${i.code}">✏ Editar</button><button class="btn-s restore" data-restore="${i.code}">Recuperar al catálogo</button>`;
  else if(confirmCode===i.code)foot=`<span class="q">¿Pasar a obsoletas?</span><button class="btn-s" data-cancel>No</button><button class="btn-s yes" data-obs-yes="${i.code}">Sí, obsoleta</button>`;
  else foot=`<button class="btn-s" data-edit="${i.code}">✏ Editar</button><button class="btn-s obs" data-obs="${i.code}">Marcar obsoleta</button>`;
  const pk=selMode&&!ex,on=pk&&selSet.has(i.code);
  return `<article class="card${i.obsoleta?" is-obs":""}${on?" picked":""}">
    ${pk?`<button class="selbox" data-sel="${esc(i.code)}" aria-pressed="${on}" aria-label="Elegir ${esc(i.code)}">✓</button>`:""}
    <div class="pics">${pane(i,"e")}${pane(i,"p")}</div>
    <div class="body">
      <div class="code"><b>${isSinRef(i.code)?`Sin código <small style="font-weight:400;opacity:.7">${esc(i.code)}</small>`:esc(i.code)}</b><span style="display:flex;gap:6px"><button class="btn-s" data-copy="${esc(i.code)}">Copiar</button><button class="btn-s" data-exp1="${esc(i.code)}" aria-label="Exportar ${esc(i.code)}">⬇</button></span></div>
      <p class="desc">${esc(i.desc)}</p>
      <div class="tags">${cur===INC&&!ex?faltas(i).map(f=>`<span class="tag" style="background:#FBEFD2;color:#8A5A00">Falta ${esc(f)}</span>`).join(""):""}${i.nueva?`<span class="tag new">Nueva</span>`:""}${sec===ALL?`<span class="tag">${esc(secsOf(i).join(" + "))}</span>`:""}${(()=>{if(sec===ALL)return "";const o=secsOf(i).filter(s=>s!==sec);return o.length&&(i.also||[]).length?`<span class="tag">También en ${esc(o.join(", "))}</span>`:"";})()}${i.obsoleta?`<span class="tag obs">Obsoleta · ${esc(i.cat)}</span>`:""}${marcaOf(i)&&marcaOf(i)!==OTRAS?`<span class="tag">${esc(marcaOf(i))}</span>`:""}${i.medida?`<span class="tag">${esc(i.medida)}</span>`:""}${i.canal?`<span class="tag">${esc(i.canal)}</span>`:""}${hasStock(i)?`<span class="tag">Stock seg.: ${i.stock}</span>`:""}</div>
      ${i.modelo?`<p class="mod">${linkify(i.modelo)}</p>`:""}${docBtns(i)}
      ${ex?`<p class="why"><span><span class="tag c-${esc(ex.conf)}">Confianza ${esc(ex.conf)}</span> <span class="tag">${esc(secsOf(i).join(" + "))} · ${esc(i.cat)}</span></span>${ex.motivo?`<span>${esc(ex.motivo)}</span>`:""}</p>`:""}
      <div class="foot">${foot}</div>
    </div></article>`;}

/* ---------- lista (una foto por referencia) y ficha ----------
   En la lista sale la foto oficial del fabricante; si no hay, la de producto (y si no, la de etiqueta).
   Al tocar la referencia se abre la ficha con todas sus fotos. Para elegir y exportar sigue la tarjeta completa. */
const listPic=i=>i.ofi[0]||i.prod.find(p=>p.from==="web")||i.prod[0]||i.photos[0]||null;
const grid=L=>selMode?`<div class="grid">${L.map(i=>card(i)).join("")}</div>`:`<div class="grid gs">${L.map(cardS).join("")}</div>`;
function cardS(i){const p=listPic(i),m=marcaOf(i),nd=docsOf(i.code).length,web=p&&(p===i.ofi[0]||p.from==="web");
  return `<article class="card cs${i.obsoleta?" is-obs":""}" data-open="${esc(i.code)}" tabindex="0" aria-label="Ver la ficha de ${esc(i.desc)}">
    <div class="csph">${p?`<img src="${p.src}" alt="" loading="lazy" class="${web?"web ":""}${p.rot?"r"+p.rot:""}">`:`<span>Sin foto</span>`}${i.prop.length?`<span class="csnew">📸 Por confirmar</span>`:""}</div>
    <div class="body"><b class="cscode">${isSinRef(i.code)?"Sin código":esc(i.code)}</b><p class="desc">${esc(i.desc)}</p>
      <div class="tags">${i.nueva?`<span class="tag new">Nueva</span>`:""}${i.obsoleta?`<span class="tag obs">Obsoleta</span>`:""}${m&&m!==OTRAS?`<span class="tag">${esc(m)}</span>`:""}${i.medida?`<span class="tag">${esc(i.medida)}</span>`:""}${i.canal?`<span class="tag">${esc(i.canal)}</span>`:""}${nd?`<span class="tag">📚 ${nd}</span>`:""}</div></div></article>`;}
let fichaCode=null;
const fichaBox=document.createElement("div");fichaBox.id="ficha";fichaBox.hidden=true;fichaBox.setAttribute("role","dialog");fichaBox.setAttribute("aria-modal","true");document.body.appendChild(fichaBox);
const host=u=>{try{return new URL(u).hostname.replace(/^www\./,"");}catch(_){return "la web del fabricante";}};
/* La ficha entra en el historial: el botón Atrás del móvil la cierra. */
function openFicha(code){const was=!!fichaCode;fichaCode=code;confirmCode=null;renderFicha();fichaBox.hidden=false;fichaBox.scrollTop=0;if(!was)try{history.pushState({ficha:1},"");}catch(_){}}
function closeFicha(){fichaCode=null;fichaBox.hidden=true;fichaBox.innerHTML="";}
window.addEventListener("popstate",()=>{if(fichaCode)closeFicha();});
function renderFicha(){const i=itemOf(fichaCode);if(!i){closeFicha();return;}
  const of=i.ofi[0],pr=i.prop[0],src=x=>x.page?` <a class="linkout" href="${esc(x.page)}" target="_blank" rel="noopener">${esc(host(x.page))}</a>`:"";
  const prop=pr?`<div class="fofi prop"><h3>📸 Foto del fabricante por confirmar</h3><img src="${pr.src}" alt="Foto propuesta para ${esc(i.desc)}">
      <p class="note">Sacada de${src(pr)||" la web del fabricante"}. Si es esta pieza, será la foto que sale en la lista.</p>
      <div class="acts2"><button class="btn-l" data-ofi-no="${esc(pr.id)}">✕ Descartar</button><button class="btn-l primary" data-ofi-ok="${esc(pr.id)}">✓ Usar esta foto</button></div></div>`:"";
  const ofi=of?`<div class="fofi"><img src="${of.src}" alt="Foto oficial de ${esc(i.desc)}"><p class="note">Foto oficial del fabricante${of.page?" ·"+src(of):""}</p></div>`:"";
  fichaBox.innerHTML=`<div class="fbar"><button class="btn-s" data-fclose>‹ Volver</button><b>${esc(secsOf(i).join(" + "))} · ${esc(i.cat)}</b></div><div class="fbody">${prop}${card(i)}${ofi}</div>`;
  /* En la ficha, la foto del fabricante a buen tamaño (en la lista basta la miniatura). */
  const big=pr||of,code=i.code;if(big)fullSrc(big).then(u=>{const im=fichaBox.querySelector(".fofi img");if(im&&fichaCode===code)im.src=u;}).catch(()=>{});}
async function decideOfi(id,ok){const f=fotos.find(x=>x.id===id);if(!f)return;const code=f.code;
  try{await api("setOficial",id,ok);
    if(ok){fotos=fotos.filter(x=>!(x.code===code&&x.tipo==="Oficial"));f.tipo="Oficial";}else fotos=fotos.filter(x=>x.id!==id);
    toast(ok?"Foto oficial guardada. Ya sale en la lista.":"Foto descartada.");}
  catch(e){toast(errMsg(e),5000);return;}
  merge();
  /* Repasando «Fotos por confirmar»: pasa sola a la siguiente. */
  const next=cur===PROP&&items.find(x=>x.prop.length&&!x.obsoleta&&inS(x,sec));if(next){fichaCode=next.code;renderFicha();fichaBox.scrollTop=0;}}
fichaBox.addEventListener("click",e=>{const t=e.target;
  if(t.closest("[data-fclose]")){if(history.state&&history.state.ficha)history.back();else closeFicha();return;}
  const ok=t.closest("[data-ofi-ok]");if(ok){ok.disabled=true;decideOfi(ok.dataset.ofiOk,true);return;}
  const no=t.closest("[data-ofi-no]");if(no){no.disabled=true;decideOfi(no.dataset.ofiNo,false);return;}
  listClick(e);});
fichaBox.addEventListener("change",listChange);

/* ---------- acciones del catálogo ---------- */
const CROP_JS="https://cdnjs.cloudflare.com/ajax/libs/cropperjs/1.6.2/cropper.min.js";
function cropPhoto(file){return new Promise(async res=>{
  try{await loadLib(CROP_JS,"Cropper");}catch(_){res(file);return;}
  const img=$("#cropImg"),box=$("#crop"),url=URL.createObjectURL(file);let cr=null,closed=false;
  const done=async ok=>{if(closed)return;closed=true;let out=null;
    if(ok&&cr){try{const cv=cr.getCroppedCanvas({maxWidth:2400,maxHeight:2400,fillColor:"#fff",imageSmoothingQuality:"high"});out=await new Promise(r=>cv.toBlob(r,"image/jpeg",0.9));}catch(_){}}
    box.hidden=true;if(cr)cr.destroy();cr=null;img.removeAttribute("src");URL.revokeObjectURL(url);res(ok?(out||file):null);};
  $("#cropOk").onclick=()=>done(true);$("#cropCancel").onclick=()=>done(false);
  $("#cropRot").onclick=()=>{if(cr)cr.rotate(90);};
  $("#cropAll").onclick=()=>{if(!cr)return;const d=cr.getCanvasData();cr.setCropBoxData({left:d.left,top:d.top,width:d.width,height:d.height});};
  img.onload=()=>{if(cr)cr.destroy();cr=new window.Cropper(img,{viewMode:1,autoCropArea:0.9,dragMode:"move",background:false,responsive:true,checkOrientation:true,toggleDragModeOnDblclick:false});};
  box.hidden=false;img.src=url;});}
async function addPhoto(code,kind,file,noCrop){if(!noCrop){file=await cropPhoto(file);if(!file)return;}const key=code+kind;busy.add(key);render();
  try{const full=await shrink(file,1400),mini=await shrink(file,360,0.72);
    const f=await api("addPhoto",code,kind==="e"?"Etiqueta":"Producto",await blobB64(full),await blobB64(mini));
    fotos.push(f);toast("Foto guardada.");}
  catch(e){toast(errMsg(e),5000);}
  finally{busy.delete(key);merge();}}
async function listChange(e){const f=e.target;if(f.type!=="file"||!f.files.length)return;const fs=[...f.files],c=f.dataset.code,k=f.dataset.k;f.value="";for(const x of fs)await addPhoto(c,k,x);}
$("#list").addEventListener("change",listChange);
$("#list").addEventListener("click",listClick);
$("#list").addEventListener("keydown",e=>{const o=e.target.closest&&e.target.closest("[data-open]");if(o&&(e.key==="Enter"||e.key===" ")&&e.target===o){e.preventDefault();openFicha(o.dataset.open);}});
async function listClick(e){const t=e.target;
  const op=t.closest("[data-open]");if(op&&!t.closest("a,button,label,input")){openFicha(op.dataset.open);return;}
  const cp=t.closest("[data-copy]");if(cp){const v=cp.dataset.copy;try{await navigator.clipboard.writeText(v);cp.textContent="Copiado";setTimeout(()=>cp.textContent="Copiar",1400);}
    catch(_){const r=document.createRange();r.selectNodeContents(cp.closest(".code").querySelector("b"));const s=getSelection();s.removeAllRanges();s.addRange(r);}return;}
  const sb=t.closest("[data-sel]");if(sb){const c=sb.dataset.sel,on=!selSet.has(c);on?selSet.add(c):selSet.delete(c);sb.setAttribute("aria-pressed",String(on));sb.closest(".card").classList.toggle("picked",on);updSel();return;}
  const e1=t.closest("[data-exp1]");if(e1){openExp(e1.dataset.exp1);return;}
  const ed=t.closest("[data-edit]");if(ed){openEdit(ed.dataset.edit);return;}
  const ob=t.closest("[data-obs]");if(ob){confirmCode=ob.dataset.obs;render();return;}
  if(t.closest("[data-cancel]")){confirmCode=null;render();return;}
  const y=t.closest("[data-obs-yes]");if(y){const c=y.dataset.obsYes;confirmCode=null;
    try{await api("setObsolete",c,true);refs.find(r=>r.code===c).obsoleta=true;merge();toast("Movida a Obsoletas.");}catch(x){toast(errMsg(x),5000);}return;}
  const mv=t.closest("[data-mv]");if(mv){moveAsk=mv.dataset.mv;render();return;}
  if(t.closest("[data-mv-no]")){moveAsk=null;render();return;}
  const my=t.closest("[data-mv-yes]");if(my){const to=my.dataset.mvYes,from=sec,L=items.filter(i=>i.obsoleta&&i.sec===from).map(i=>i.code);moveAsk=null;my.disabled=true;
    try{await api("moveRefs",L,to);refs.forEach(r=>{if(L.includes(r.code)){r.sec=to;r.also=(r.also||[]).filter(x=>x!==to);}});merge();toast(`${L.length} obsoletas pasadas a ${to}.`,4000);}catch(x){toast(errMsg(x),5000);render();}return;}
  const rs=t.closest("[data-restore]");if(rs){const c=rs.dataset.restore;
    try{await api("setObsolete",c,false);refs.find(r=>r.code===c).obsoleta=false;merge();toast("Recuperada en el catálogo.");}catch(x){toast(errMsg(x),5000);}return;}
  const pk=t.closest("[data-pick]");if(pk&&recog&&recog.file){const c=pk.dataset.pick,f=recog.file,it=itemOf(c);recog=null;if(it)sec=it.sec;$("#q").value=c;chips();addPhoto(c,"p",f);window.scrollTo({top:0});return;}
  if(t.closest("[data-rec-close]")){recog=null;render();return;}
  if(t.closest("[data-rec-add]")&&recog){const f=recog.file;recog=null;render();openAdd(undefined,undefined,f);return;}
  const ph=t.closest("button.ph");if(ph)openLb(itemOf(ph.dataset.code),ph.dataset.k);}

/* ---------- nueva referencia ---------- */
/* Dos huecos de foto en la ficha; cada uno lleva su tipo (e = etiqueta, p = producto), que se puede cambiar con su casilla. */
const SLOT_LAB=["fEtqL","fProdL"],slots=[{f:null,k:"e"},{f:null,k:"p"}];
$("#fSec").innerHTML=SECS.map(s=>`<option>${esc(s.n)}</option>`).join("");
$("#fMarca").innerHTML=`<option value="">— Elegir —</option>`+[...MARCAS,OTRA].map(m=>`<option>${esc(m)}</option>`).join("");
function syncMarca(){const on=hasMarca(catsOf($("#fSec").value).find(c=>norm(c)===norm($("#fCat").value.trim()))||"");$("#fMarcaL").hidden=!on;
  if(on&&!$("#fMarca").value)$("#fMarca").value=guessMarca($("#fDesc").value+" "+$("#fMod").value);}
["fCat","fDesc","fMod"].forEach(id=>$("#"+id).addEventListener("input",syncMarca));
/* Categoría: desplegable con todas las del apartado; «Nueva categoría…» muestra el campo para escribirla. #fCat guarda el valor. */
function fillCats(){$("#fCatSel").innerHTML=`<option value="">— Elegir —</option>`+catsOf($("#fSec").value).map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join("")+`<option value="__new">➕ Nueva categoría…</option>`;syncCatSel();fillAlso();}
function syncCatSel(){const v=$("#fCat").value.trim(),k=catsOf($("#fSec").value).find(c=>norm(c)===norm(v)),nw=!!v&&!k;
  $("#fCatSel").value=nw?"__new":(k||"");if(k)$("#fCat").value=k;$("#fCat").hidden=!nw;}
$("#fCatSel").addEventListener("change",e=>{const v=e.target.value;if(v==="__new"){$("#fCat").value="";$("#fCat").hidden=false;$("#fCat").focus();}else{$("#fCat").value=v;$("#fCat").hidden=true;}syncMarca();});
function fillAlso(on){const s=$("#fSec").value,keep=on||[...document.querySelectorAll("#fAlso input:checked")].map(x=>x.value);
  $("#fAlso").innerHTML=SECS.filter(x=>x.n!==s).map(x=>`<label><input type="checkbox" value="${esc(x.n)}"${keep.includes(x.n)?" checked":""}> También en ${esc(x.n)}</label>`).join("");}
$("#fSec").addEventListener("change",()=>{const c=$("#fCat").value.trim();if(c&&!catsOf($("#fSec").value).some(x=>norm(x)===norm(c)))$("#fCat").value="";fillCats();});
function preview(lab,f){const L=$("#"+lab);L.querySelector("img")?.remove();const im=document.createElement("img");im.alt="";im.src=URL.createObjectURL(f);L.appendChild(im);}
let editCode=null;
function openEdit(code){const it=itemOf(code);if(!it)return;openAdd();editCode=code;$("#fTitle").textContent="Editar referencia";const sr=isSinRef(code);$("#fCode").value=sr?"":code;$("#fCode").disabled=!sr;$("#fNoCodeL").hidden=!sr;$("#fNoCode").checked=sr;$("#fCode").placeholder=sr?"Pon aquí el código cuando lo tengas":"48009000000";
  $("#fDesc").value=it.desc||"";$("#fSec").value=it.sec;fillCats();fillAlso(it.also||[]);$("#fCat").value=it.cat||"";syncCatSel();$("#fMed").value=it.medida||"";$("#fCanal").value=it.canal||"";$("#fMod").value=it.modelo||"";$("#fStock").value=hasStock(it)?it.stock:"";
  const mo=marcaOf(it);$("#fMarca").value=mo===OTRAS?(it.marca?OTRA:""):mo;syncMarca();
  $("#fAuto").textContent="Puedes añadir fotos nuevas aquí; se suman a las que ya tiene.";renderDocBox();(it.medida?$("#fMod"):$("#fMed")).focus();}
function openAdd(code,file,prodFile){editCode=null;$("#fTitle").textContent="Nueva referencia";$("#fCode").disabled=false;$("#fCode").placeholder="48009000000";$("#fNoCodeL").hidden=false;$("#addForm").reset();slots[0]={f:null,k:"e"};slots[1]={f:null,k:"p"};SLOT_LAB.forEach(id=>$("#"+id).querySelector("img")?.remove());
  $("#fErr").textContent="";$("#fSec").value=sec===ALL?SECS[0].n:sec;fillAlso([]);fillCats();$("#fCat").value=cur!=="Todas"&&cur!==OBS&&cur!==INC&&cur!==STK&&cur!==DOC&&cur!==PROP?cur:"";syncCatSel();
  $("#fMarca").value=MARCAS.includes(marca)?marca:"";syncMarca();
  secTouched=false;$("#fAuto").textContent="Consejo: haz primero la foto de la etiqueta y se rellenan solos el código y los datos.";
  if(typeof code==="string")$("#fCode").value=code;if(file instanceof Blob){slots[0].f=file;preview("fEtqL",file);fillFromLabel(file);}if(prodFile instanceof Blob){slots[1].f=prodFile;preview("fProdL",prodFile);}
  syncKinds();renderDocBox();$("#addSheet").hidden=false;(typeof code==="string"?$("#fDesc"):$("#fCode")).focus();}
$("#fNoCode").addEventListener("change",e=>{if(!editCode){$("#fCode").disabled=e.target.checked;if(e.target.checked)$("#fCode").value="";}});
$("#addBtn").onclick=()=>{if(view==="rev")newRev();else openAdd();};
$("#fCancel").onclick=()=>{$("#addSheet").hidden=true;};
$("#addSheet").addEventListener("click",e=>{if(e.target.id==="addSheet")$("#addSheet").hidden=true;});
["fEtq","fProd"].forEach((inp,i)=>$("#"+inp).addEventListener("change",async e=>{let f=e.target.files[0];e.target.value="";if(!f)return;f=await cropPhoto(f);if(!f)return;slots[i].f=f;preview(SLOT_LAB[i],f);syncKinds();if(slots[i].k==="e")fillFromLabel(f);}));
/* Casilla Etiqueta / Producto de cada foto: decide con qué tipo se guarda (pueden ser las dos del mismo tipo). */
function syncKinds(){document.querySelectorAll("#addForm .kind").forEach(g=>{const sl=slots[+g.dataset.slot];
  g.querySelectorAll("button").forEach(b=>b.setAttribute("aria-pressed",String(b.dataset.k===sl.k)));
  const sp=$("#"+SLOT_LAB[+g.dataset.slot]).querySelector("span");if(sp)sp.textContent=sl.k==="e"?"📷 Foto etiqueta":"📷 Foto producto";});}
$("#addForm").addEventListener("click",e=>{const b=e.target.closest(".kind button");if(!b)return;const sl=slots[+b.closest(".kind").dataset.slot];
  if(sl.k===b.dataset.k)return;sl.k=b.dataset.k;syncKinds();if(sl.f&&sl.k==="e")fillFromLabel(sl.f);});
let secTouched=false;$("#fSec").addEventListener("change",()=>{secTouched=true;});
async function fillFromLabel(f){const st=$("#fAuto");st.textContent="Leyendo la etiqueta…";let code=null;
  try{code=await readBarcode(f);}catch(_){}
  if(code&&!$("#fCode").value.trim())$("#fCode").value=code;
  if(S.settings.gemini){
    try{const cats=SECS.flatMap(s=>catsOf(s.n).map(c=>s.n+" › "+c));
      const r=await api("aiLabelInfo",await blobB64(await shrink(f,2000,0.88)),cats)||{};if($("#addSheet").hidden)return;
      const set=(id,v)=>{v=String(v||"").trim();const el=$("#"+id);if(v&&!el.value.trim())el.value=v;};
      const c=pick(r.codigo);if(c&&!$("#fCode").value.trim())$("#fCode").value=c;
      set("fDesc",r.descripcion);set("fMed",r.medida);set("fCanal",r.canal);set("fMod",r.modelo);
      if(r.apartado&&!secTouched&&SECS.some(s=>s.n===r.apartado)){$("#fSec").value=r.apartado;fillCats();}
      if(r.categoria&&!$("#fCat").value.trim()){$("#fCat").value=String(r.categoria);syncCatSel();}
      const rm=MARCAS.find(x=>norm(x)===norm(r.marca||""));if(rm&&!$("#fMarca").value)$("#fMarca").value=rm;syncMarca();
      st.textContent="✨ Rellenado desde la etiqueta. Revísalo antes de guardar.";}
    catch(e){st.textContent=(code?"Código leído. ":"")+errMsg(e);}}
  else st.textContent=code?"Código leído. Pon la clave de Gemini en ⚙ para rellenar también la descripción.":"No he podido leer el código de la etiqueta.";
  const ex=itemOf($("#fCode").value.trim());if(ex)st.textContent=`⚠ Ese código ya existe en el catálogo: ${ex.desc}${ex.obsoleta?" (en Obsoletas)":""}.`;}
$("#addForm").addEventListener("submit",async e=>{e.preventDefault();const err=$("#fErr");err.textContent="";
  let code=$("#fCode").value.replace(/\s/g,"");const desc=$("#fDesc").value.trim(),noCode=$("#fNoCode").checked&&!code;
  let rename=null;if(editCode&&isSinRef(editCode)){if(code&&code!==editCode)rename=code;else code=editCode;}
  if(!noCode&&!/^\d{11}$/.test(code)&&!isSinRef(code)){err.textContent="El código tiene que tener 11 números (o marca «Todavía no tiene código»).";return;}
  const ex=itemOf(code);if(ex&&(!editCode||rename)){err.textContent=`Ese código ya existe: ${ex.desc}${ex.obsoleta?" (está en Obsoletas)":""}.`;return;}
  if(!desc){err.textContent="Escribe una descripción.";return;}
  const s=$("#fSec").value,ci=$("#fCat").value.trim(),known=catsOf(s).find(c=>norm(c)===norm(ci)),cat=known||(ci?ci[0].toUpperCase()+ci.slice(1):"General");
  const ref={code:noCode&&!editCode?"":(rename||code),sinCodigo:noCode&&!editCode,desc,sec:s,cat,medida:$("#fMed").value.trim(),canal:$("#fCanal").value.trim(),modelo:$("#fMod").value.trim(),stock:$("#fStock").value.trim()===""?"":Math.max(0,Math.round(+$("#fStock").value)||0),marca:hasMarca(cat)?$("#fMarca").value:"",also:[...document.querySelectorAll("#fAlso input:checked")].map(x=>x.value).filter(x=>x!==s)};
  const btn=$("#fSave");btn.disabled=true;btn.textContent="Guardando…";
  try{if(editCode){if(rename){await api("renameRef",editCode,rename);refs.find(x=>x.code===editCode).code=rename;fotos.forEach(f=>{if(f.code===editCode)f.code=rename;});}
      await api("saveRef",ref,false);Object.assign(refs.find(x=>x.code===ref.code),ref);code=ref.code;}
    else{const c=await api("saveRef",ref,true);ref.code=code=typeof c==="string"?c:ref.code;refs.push({...ref,obsoleta:false,nueva:true});}
    const sl=slots.filter(x=>x.f).map(x=>({...x})),wasEdit=!!editCode;$("#addSheet").hidden=true;if(!wasEdit){if(sec!==ALL)sec=s;cur=cat;$("#q").value="";}notFound=null;merge();toast(wasEdit?"Cambios guardados.":"Referencia añadida.");
    for(const x of sl)await addPhoto(code,x.k,x.f,true);}
  catch(x){err.textContent=errMsg(x);}finally{btn.disabled=false;btn.textContent="Guardar";}});

/* ---------- documentación en la ficha ---------- */
$("#dTipo").innerHTML=DOC_TIPOS.map(t=>`<option>${t}</option>`).join("");
function renderDocBox(){const box=$("#fDocs");box.hidden=!editCode;if(!editCode)return;const code=editCode,L=docsOf(code);$("#dMsg").textContent="";
  $("#fDocList").innerHTML=L.length?L.map(m=>`<div class="docrow"><span>${docIco(m)}</span>${m.url?`<a class="linkout" href="${esc(m.url)}" target="_blank" rel="noopener">${esc(m.titulo)}</a>`:`<span>${esc(m.titulo)} <small>(falta el PDF ${esc(m.archivo)} en el Drive)</small></span>`}${m.codes.length>1?`<small>· en ${m.codes.length} referencias</small>`:""}<button type="button" class="btn-s" data-dquit="${esc(m.id)}" aria-label="Quitar ${esc(m.titulo)}">✕</button></div>`).join(""):`<p class="note" style="margin:0">Todavía no tiene documentos.</p>`;
  const others=manuales.filter(m=>!m.codes.includes(code));$("#dReuse").hidden=!others.length;
  $("#dReuse").innerHTML=`<option value="">Usar uno que ya existe…</option>`+others.map(m=>`<option value="${esc(m.id)}">${esc(docIco(m)+" "+m.titulo)}</option>`).join("");}
async function docSave(d,ok){const msg=$("#dMsg");msg.textContent="Guardando…";
  try{const r=await api("saveManual",d);if(r){const k=manuales.findIndex(x=>x.id===r.id);if(k>=0)manuales[k]=r;else manuales.push(r);}else manuales=manuales.filter(x=>x.id!==d.id);
    renderDocBox();merge();$("#dMsg").textContent=ok;return true;}catch(e){msg.textContent=errMsg(e);return false;}}
$("#dUrl").addEventListener("input",e=>{if(/youtu\.?be|vimeo/i.test(e.target.value))$("#dTipo").value="Vídeo";});
$("#dAdd").onclick=async()=>{const url=$("#dUrl").value.trim(),tit=$("#dTit").value.trim();
  if(!/^https?:/i.test(url)){$("#dMsg").textContent="Pega el enlace (empieza por https://).";return;}if(!tit){$("#dMsg").textContent="Pon un título, por ejemplo «Manual de programación».";return;}
  if(await docSave({codes:[editCode],tipo:$("#dTipo").value,titulo:tit,url},"Documento añadido.")){$("#dUrl").value="";$("#dTit").value="";}};
$("#dFile").addEventListener("change",async e=>{const f=e.target.files[0];e.target.value="";if(!f||!editCode)return;const msg=$("#dMsg");
  if(f.size>25e6){msg.textContent="El archivo pasa de 25 MB. Súbelo a Drive y pega aquí el enlace.";return;}
  const tit=$("#dTit").value.trim()||f.name.replace(/\.[^.]+$/,"");msg.textContent="Subiendo…";
  try{const r=await api("uploadManual",{codes:[editCode],tipo:$("#dTipo").value,titulo:tit},f.name,f.type||"application/pdf",await blobB64(f));
    if(r)manuales.push(r);$("#dTit").value="";renderDocBox();merge();$("#dMsg").textContent="Documento subido.";}catch(x){msg.textContent=errMsg(x);}});
$("#dReuse").addEventListener("change",e=>{const m=manuales.find(x=>x.id===e.target.value);if(!m)return;docSave({id:m.id,codes:[...m.codes,editCode],tipo:m.tipo,titulo:m.titulo},"Documento añadido.");});
$("#fDocList").addEventListener("click",e=>{const b=e.target.closest("[data-dquit]");if(!b)return;const m=manuales.find(x=>x.id===b.dataset.dquit);if(!m)return;
  docSave({id:m.id,codes:m.codes.filter(c=>c!==editCode),tipo:m.tipo,titulo:m.titulo},"Quitado de esta referencia.");});

/* ---------- buscar con foto ---------- */
function pick(t){const m=String(t||"").replace(/\s/g,"").match(/\d{11}/g)||[];return m.find(c=>items.some(i=>i.code===c))||m[0]||null;}
async function canvasOf(file,max,rot,top){let bmp;try{bmp=await createImageBitmap(file,{imageOrientation:"from-image"});}catch(_){bmp=await createImageBitmap(file);}
  const s=Math.min(1,max/Math.max(bmp.width,bmp.height)),w=Math.round(bmp.width*s),h=Math.round(bmp.height*s);
  const c=document.createElement("canvas");c.width=rot?h:w;c.height=rot?w:h;const g=c.getContext("2d");if(rot){g.translate(h,0);g.rotate(Math.PI/2);}g.drawImage(bmp,0,0,w,h);
  if(!top)return c;const k=document.createElement("canvas");k.width=c.width;k.height=Math.round(c.height*top);k.getContext("2d").drawImage(c,0,0);return k;}
async function readBarcode(file){
  if("BarcodeDetector" in window){try{const r=await new BarcodeDetector().detect(await canvasOf(file,1800,0));for(const x of r){const c=pick(x.rawValue);if(c)return c;}}catch(_){}}
  if(!window.ZXing)return null;const F=ZXing.BarcodeFormat,oneD=[F.CODE_128,F.EAN_13,F.CODE_39,F.CODE_93,F.ITF,F.UPC_A];
  for(const [rot,top] of [[0,0],[0,0.55],[1,0],[1,0.55]]){const c=await canvasOf(file,1600,rot,top);
    for(const fm of [oneD,null]){const h=new Map();h.set(ZXing.DecodeHintType.TRY_HARDER,true);if(fm)h.set(ZXing.DecodeHintType.POSSIBLE_FORMATS,fm);
      const rd=new ZXing.MultiFormatReader();rd.setHints(h);
      try{const code=pick(rd.decode(new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(new ZXing.HTMLCanvasElementLuminanceSource(c)))).getText());if(code)return code;}catch(_){}}}
  return null;}
async function readLabel(file){let code=null;try{code=await readBarcode(file);}catch(_){}
  if(!code&&S.settings.gemini){toast("No se lee el código de barras. Se lo paso a Gemini…",60000);
    try{const r=await api("aiLabel",await blobB64(await shrink(file,2000,0.88)));code=pick(r&&r.codigo);}catch(e){toast(errMsg(e),5000);}}
  return code;}
function showCode(code,file){const it=itemOf(code);$("#q").value=code;recog=null;window.scrollTo({top:0});
  if(it){notFound=null;sec=it.sec;cur=it.obsoleta?OBS:"Todas";chips();render();openFicha(code);toast(`${code} · ${it.desc}${it.obsoleta?" (obsoleta)":""}`,4000);}
  else{notFound={code,file};cur="Todas";chips();render();}}
$("#scanFile").addEventListener("change",async e=>{const f=e.target.files[0];e.target.value="";if(!f)return;const L=$("#scanL");L.classList.add("busy");toast("Leyendo el código de barras…",60000);
  const code=await readLabel(f);L.classList.remove("busy");
  if(code){showCode(code,f);return;}
  if(S.settings.gemini){toast("No veo ningún código: busco la pieza por su aspecto…",4000);recognize(f);}
  else toast("No he podido leer el código. Haz la foto más cerca y con la etiqueta recta.",5500);});

/* ---------- reconocer producto ---------- */
function renderRecog(){const r=recog,img=`<img src="${r.url}" alt="Tu foto">`;
  if(r.loading){$("#list").innerHTML=`<div class="recbox"><div class="top">${img}<p><b>Gemini está mirando la foto…</b><br><span class="sub">Compara con las fotos del catálogo: puede tardar hasta un minuto.</span></p></div></div>`;return;}
  const cands=r.cands.map(c=>({c,i:itemOf(c.codigo)})).filter(x=>x.i);
  $("#list").innerHTML=`<div class="recbox"><div class="top">${img}<p><b>${cands.length?"Posibles coincidencias":"No he encontrado nada parecido en el catálogo"}</b><br>
    <span class="sub">${cands.length?"Comprueba con la etiqueta antes de dar por buena una referencia.":"Prueba con otra foto, con la pieza entera y buena luz."}</span></p></div>
    ${r.queEs?`<p class="sub" style="margin:8px 0 0">Gemini ve: <b>${esc(r.queEs)}</b></p>`:""}
    ${r.aviso?`<p class="aviso">⚠ ${esc(r.aviso)}</p>`:""}
    <div class="acts"><button class="btn-s pick" data-rec-add>Ninguna es · dar de alta con esta foto</button><button class="btn-s" data-rec-close>Cerrar</button></div></div>
    ${cands.length?`<div class="grid">${cands.map(x=>card(x.i,{conf:["alta","media","baja"].includes(x.c.confianza)?x.c.confianza:"baja",motivo:x.c.motivo||""})).join("")}</div>`:""}`;}
$("#recFile").addEventListener("change",e=>{const f=e.target.files[0];e.target.value="";if(f)recognize(f);});
async function recognize(f){
  recog={file:f,url:URL.createObjectURL(f),loading:true};$("#q").value="";notFound=null;window.scrollTo({top:0});render();$("#scanL").classList.add("busy");
  try{const r=await api("aiRecognize",await blobB64(await shrink(f,1600)));if(!recog||recog.file!==f)return;
    recog={file:f,url:recog.url,loading:false,cands:Array.isArray(r&&r.candidatos)?r.candidatos.slice(0,3):[],aviso:(r&&r.aviso)||"",queEs:(r&&r.que_es)||""};render();}
  catch(x){recog=null;render();toast(errMsg(x),5500);}
  finally{$("#scanL").classList.remove("busy");}}

/* ---------- visor de fotos ---------- */
function buildLb(i){lbList=[...i.photos.map(p=>({...p,k:"e",lab:"Etiqueta"})),...i.prod.map(p=>({...p,k:"p",lab:p.from==="web"?"Producto (web del fabricante)":"Producto"}))];}
function openLb(i,k){lbItem=i;buildLb(i);lbI=k==="p"?i.photos.length:0;askDel(false);showLb();$("#lb").hidden=false;$("#lbClose").focus();}
function askDel(on){$("#lbAsk").hidden=!on;$("#lbNav").hidden=on;}
async function fullSrc(p){if(fullCache[p.fileId])return fullCache[p.fileId];const r=await api("getFull",p.fileId);return fullCache[p.fileId]=`data:${r.mime};base64,${r.b64}`;}
function showLb(){const p=lbList[lbI];if(!p)return;const img=$("#lbImg");img.src=fullCache[p.fileId]||p.src;img.alt=lbItem.desc;
  if(!fullCache[p.fileId])fullSrc(p).then(u=>{if(lbList[lbI]===p)img.src=u;}).catch(()=>{});
  img.className=p.rot?"r"+p.rot:"";$("#lbStage").classList.toggle("side",p.rot===90||p.rot===270);
  const same=lbList.filter(x=>x.k===p.k),pos=same.indexOf(p)+1;
  $("#lbCap").innerHTML=`<b>${esc(lbItem.code)}</b>${esc(lbItem.desc)} · ${p.lab}${same.length>1?` ${pos} de ${same.length}`:""}`;
  const T=$("#lbThumbs");T.innerHTML=lbList.map((x,j)=>`${j>0&&x.k!==lbList[j-1].k?'<span class="sep"></span>':""}<button data-j="${j}" aria-current="${j===lbI}" aria-label="${x.lab} ${j+1}"><img src="${x.src}" alt="" class="${x.rot?"r"+x.rot:""}"></button>`).join("");
  T.hidden=lbList.length<2;$("#lbPrev").hidden=$("#lbNext").hidden=lbList.length<2;
  const w=!lbItem.obsoleta;$("#lbAdd").hidden=!w;$("#lbRot").hidden=!w;$("#lbDel").hidden=!w;}
$("#lbThumbs").addEventListener("click",e=>{const b=e.target.closest("button[data-j]");if(!b)return;lbI=+b.dataset.j;askDel(false);showLb();});
const rotT={};
$("#lbRot").onclick=()=>{const p=lbList[lbI],f=fotos.find(x=>x.id===p.id);if(!f)return;f.rot=((f.rot||0)+90)%360;p.rot=f.rot;
  merge();clearTimeout(rotT[p.id]);rotT[p.id]=setTimeout(()=>api("setPhoto",p.id,{rot:f.rot}).then(()=>toast("Giro guardado.")).catch(e=>toast(errMsg(e),5000)),700);};
$("#lbDel").onclick=()=>askDel(true);$("#lbNo").onclick=()=>askDel(false);
$("#lbYes").onclick=async()=>{const p=lbList[lbI];askDel(false);if(!p)return;
  try{await api("setPhoto",p.id,{hidden:true});fotos=fotos.filter(x=>x.id!==p.id);merge();toast("Foto quitada del catálogo.");}catch(e){toast(errMsg(e),5000);}};
$("#lbClose").onclick=()=>{$("#lb").hidden=true};
$("#lbPrev").onclick=()=>{lbI=(lbI-1+lbList.length)%lbList.length;showLb()};
$("#lbNext").onclick=()=>{lbI=(lbI+1)%lbList.length;showLb()};
$("#lbFile").addEventListener("change",async e=>{const fs=[...e.target.files];if(!fs.length)return;const k=lbList[lbI].k,c=lbItem.code;e.target.value="";$("#lb").hidden=true;for(const x of fs)await addPhoto(c,k,x);});
/* selector Cámara / Galería para todos los botones de foto (menos los de documentos) */
let pickInput=null,pickPass=false;
document.addEventListener("click",e=>{if(pickPass)return;const t=e.target;
  const inp=t.matches&&t.matches("input[type=file][data-pick]")?t:(t.closest&&t.closest("label")?t.closest("label").querySelector("input[type=file][data-pick]"):null);
  if(!inp)return;e.preventDefault();e.stopPropagation();pickInput=inp;
  $("#pickGal").textContent=inp.dataset.multi?"🖼 Elegir de la galería (una o varias)":"🖼 Elegir de la galería";$("#pickSheet").hidden=false;},true);
function pickGo(cam){const i=pickInput;$("#pickSheet").hidden=true;pickInput=null;if(!i)return;
  /* En Android, accept="image/*" abre el selector de fotos de Google; sin filtro de tipo abre el explorador de archivos
     con todas las carpetas del móvil (Cámara, WhatsApp, Descargas…). Lo que no sea foto se descarta al elegir. */
  if(cam){i.setAttribute("capture","environment");i.multiple=false;i.accept="image/*";}
  else{i.removeAttribute("capture");i.multiple=!!i.dataset.multi;if(/Android/i.test(navigator.userAgent))i.removeAttribute("accept");else i.accept="image/*";}
  pickPass=true;try{i.click();}finally{pickPass=false;}}
const isImg=f=>/^image\//.test(f.type)||/\.(jpe?g|png|webp|gif|bmp|heic|heif)$/i.test(f.name);
document.addEventListener("change",e=>{const i=e.target;if(!(i.matches&&i.matches("input[type=file][data-pick]"))||!i.files.length)return;
  const ok=[...i.files].filter(isImg);if(ok.length===i.files.length)return;
  if(!ok.length){e.stopImmediatePropagation();i.value="";toast("Eso no es una foto. Elige una imagen (JPG, PNG…).",4000);return;}
  const dt=new DataTransfer();ok.forEach(f=>dt.items.add(f));i.files=dt.files;toast(`Solo se suben las fotos (${ok.length}); el resto se ignora.`,4000);},true);
$("#pickCam").onclick=()=>pickGo(true);$("#pickGal").onclick=()=>pickGo(false);
$("#pickNo").onclick=()=>{$("#pickSheet").hidden=true;pickInput=null;};
$("#pickSheet").addEventListener("click",e=>{if(e.target.id==="pickSheet"){$("#pickSheet").hidden=true;pickInput=null;}});
$("#lb").addEventListener("click",e=>{if(e.target.id==="lb")$("#lb").hidden=true});

/* =========================================================
   Revisión de hojas
   ========================================================= */
const counts=r=>{const L=r.lineas||[];return {f:L.filter(l=>l.tipo==="falta").length,s:L.filter(l=>l.tipo==="sobra").length,p:L.filter(l=>l.tipo==="revisar").length};};
const pills=c=>`${c.f?`<span class="pill f">Faltan ${c.f}</span>`:""}${c.s?`<span class="pill s">Sobran ${c.s}</span>`:""}${c.p?`<span class="pill">Por revisar ${c.p}</span>`:""}`;
function persistRev(){if(!revDraft)return;clearTimeout(revT);revT=setTimeout(()=>{revT=null;writeRev();},800);}
function writeRev(){const r=JSON.parse(JSON.stringify(revDraft));const c=counts(r);
  const s={id:r.id,fecha:r.fecha,hoja:r.hoja,of:r.of,obra:r.obra,estado:r.estado,enviado:r.enviado,f:c.f,s:c.s,p:c.p};
  const j=revs.findIndex(x=>x.id===r.id);if(j>=0)revs[j]=s;else revs.unshift(s);chips();
  return api("saveRev",r).catch(e=>toast("No se pudo guardar la revisión: "+errMsg(e),6000));}
function flushRev(){if(revT){clearTimeout(revT);revT=null;if(revDraft)writeRev();}}
function newRev(){revDraft={id:"h"+Date.now().toString(36),fecha:new Date().toISOString(),hoja:"",of:"",obra:"",obs:"",lineas:[],docs:[],estado:"Pendiente"};
  revId=revDraft.id;sendState="idle";writeRev();renderRev();}
async function openRev(id){revId=id;revDraft=null;$("#rev").innerHTML=`<div class="boot"><div class="spin"></div><p>Abriendo la revisión…</p></div>`;updateFab();
  try{const r=await api("getRev",id);if(revId!==id)return;revDraft=r;sendState="idle";renderRev();window.scrollTo({top:0});}
  catch(e){revId=null;renderRev();toast(errMsg(e),5000);}}
function closeRev(){flushRev();revId=null;revDraft=null;renderRev();}
function renderRev(){updateFab();chips();const R=$("#rev");
  if(revId&&revDraft){renderEditor();return;}
  if(consOn){renderCons();return;}
  R.innerHTML=`<div class="rv-top"><h2 style="margin:0">Revisiones</h2><span><button class="btn-l" data-rv="cons">📊 Consumo</button> <button class="btn-l primary" data-rv="new">＋ Nueva revisión</button></span></div>
    ${revs.length?revs.map(r=>`<button class="rv-item" data-rv-open="${esc(r.id)}"><b>${r.hoja?"Hoja "+esc(r.hoja):"Hoja sin número"}${r.of?" · OF "+esc(r.of):""}${r.obra?" · "+esc(r.obra):""}</b>
      <span class="rv-meta">${fmtF(r.fecha)} ${pills(r)}${r.estado==="Enviada"?`<span class="pill env">Enviado ${fmtF(r.enviado)}</span>`:`<span class="pill">Pendiente</span>`}</span></button>`).join("")
      :`<p class="empty">Todavía no hay revisiones. Pulsa «＋ Nueva revisión» cuando te llegue una hoja.</p>`}`;}
/* ---------- Consumo: suma lo que sale en las hojas revisadas (menos «Sobra») por día, semana, mes o año ---------- */
let consOn=false,consP="mes",consOff=0,consData=null;
const cap=s=>s.charAt(0).toUpperCase()+s.slice(1);
function consRange(p,off){const d=new Date();d.setHours(0,0,0,0);let a,b,lab;
  if(p==="dia"){a=new Date(d);a.setDate(a.getDate()+off);b=new Date(a);b.setDate(b.getDate()+1);lab=cap(a.toLocaleDateString("es-ES",{weekday:"long",day:"numeric",month:"long",year:"numeric"}));}
  else if(p==="semana"){a=new Date(d);a.setDate(a.getDate()-((a.getDay()+6)%7)+7*off);b=new Date(a);b.setDate(b.getDate()+7);const e=new Date(b);e.setDate(e.getDate()-1);lab=`Semana del ${fmtF(a.toISOString())} al ${fmtF(e.toISOString())}`;}
  else if(p==="mes"){a=new Date(d.getFullYear(),d.getMonth()+off,1);b=new Date(a.getFullYear(),a.getMonth()+1,1);lab=cap(a.toLocaleDateString("es-ES",{month:"long",year:"numeric"}));}
  else if(p==="anio"){a=new Date(d.getFullYear()+off,0,1);b=new Date(a.getFullYear()+1,0,1);lab="Año "+a.getFullYear();}
  else{a=new Date(0);b=new Date(8.64e15);lab="Desde el principio";}
  return {a,b,lab};}
function consRows(){const {a,b}=consRange(consP,consOff),m=new Map(),hs=new Set();let first=null;
  for(const x of consData||[]){const t=new Date(x.f);if(!(t>=a&&t<b))continue;hs.add(x.r);if(!first||t<first)first=t;
    const g=m.get(x.c)||{code:x.c,q:0,n:new Set()};g.q=Math.round((g.q+(+x.q||0))*1000)/1000;g.n.add(x.r);m.set(x.c,g);}
  const tom=new Date();tom.setHours(24,0,0,0);const from=consP==="todo"?first:a,to=b<tom?b:tom;
  const dias=from?Math.max(1,Math.round((to-from)/864e5)):0;
  return {rows:[...m.values()].map(g=>({code:g.code,q:g.q,n:g.n.size})).sort((x,y)=>y.q-x.q),hojas:hs.size,dias};}
async function openCons(){consOn=true;consData=null;renderRev();
  try{consData=await api("consumo");}catch(e){toast(errMsg(e),5000);consData=[];}
  if(consOn&&!revId)renderRev();}
function renderCons(){const R=$("#rev"),{lab}=consRange(consP,consOff),P=[["dia","Día"],["semana","Semana"],["mes","Mes"],["anio","Año"],["todo","Todo"]];
  const top=`<div class="rv-top"><h2 style="margin:0">Consumo</h2><button class="btn-l" data-rv="consBack">‹ Revisiones</button></div>
    <p class="note">Sale solo de las hojas que revisas: cuenta cada componente del catálogo con la fecha de la revisión. Lo marcado «Sobra» no cuenta.</p>
    <div class="cons-bar"><span class="seg" role="group" aria-label="Periodo">${P.map(([k,t])=>`<button data-cp="${k}" aria-pressed="${consP===k}">${t}</button>`).join("")}</span>
    ${consP!=="todo"?`<span class="cons-nav"><button class="btn-s" data-co="-1" aria-label="Periodo anterior">‹</button><b>${esc(lab)}</b><button class="btn-s" data-co="1" aria-label="Periodo siguiente"${consOff>=0?" disabled":""}>›</button></span>`:""}</div>`;
  if(!consData){R.innerHTML=top+`<div class="boot"><div class="spin"></div><p>Sumando las revisiones…</p></div>`;return;}
  const {rows,hojas,dias}=consRows();
  R.innerHTML=top+(rows.length?`<div class="box"><p style="margin:0"><b>${rows.length}</b> referencias en <b>${hojas}</b> hoja${hojas>1?"s":""}${dias>1?` · ${dias} días`:""}</p>
    ${rows.map(x=>{const it=itemOf(x.code),src=thumbOf(it);return `<div class="ln">${src?`<img class="th" src="${src}" alt="">`:`<span class="th"></span>`}
      <div class="d"><span><code>${esc(x.code)}</code> ${esc(it?it.desc:"")}</span><span class="note">${it&&it.medida?esc(it.medida)+" · ":""}En ${x.n} hoja${x.n>1?"s":""}${dias>1?` · ${fmtQ(Math.round(x.q/dias*100)/100)} al día`:""}</span></div>
      <b class="cons-q">${fmtQ(x.q)}</b></div>`;}).join("")}
    <div><button class="btn-s" data-rv="consXlsx">⬇ Excel</button> <span id="consMsg" class="note"></span></div></div>`
    :`<p class="empty">No hay consumo en este periodo.</p>`);}
async function consXlsx(){try{await loadLib(XLSX_URL,"XLSX");const {rows,hojas}=consRows(),{lab}=consRange(consP,consOff);
  const aoa=[[`Consumo · ${lab}`],[`${hojas} hojas revisadas`],[],["Código","Descripción","Medida","Cantidad","Nº de hojas"],
    ...rows.map(x=>{const it=itemOf(x.code);return [x.code,it?it.desc:"",it?it.medida:"",x.q,x.n];})];
  const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=[{wch:14},{wch:44},{wch:14},{wch:10},{wch:12}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Consumo");
  await deliver(new Blob([XLSX.write(wb,{type:"array",bookType:"xlsx"})]),`Consumo ${lab.replace(/[^\wÀ-ÿ\-]+/g,"_")}.xlsx`,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","drive",null,$("#consMsg"));}catch(x){toast(errMsg(x),5000);}}
function lineHTML(l,idx){const it=itemOf(l.code),src=thumbOf(it);
  const name=l.code?`<span><code>${esc(l.code)}</code> ${esc(it?it.desc:(l.desc||""))}</span>`:`<input data-desc value="${esc(l.desc||"")}" placeholder="Descripción del componente">`;
  const origen=l.origen==="hoja"?"la hoja":l.origen==="presupuesto"?"el presupuesto":"";
  return `<div class="ln${l.tipo==="revisar"?" rv":""}" data-i="${idx}">
    ${src?`<img class="th" src="${src}" alt="">`:`<span class="th"></span>`}
    <div class="d">${name}
      ${origen&&l.desc&&it?`<span class="note">En ${origen} pone: ${esc(l.desc)}</span>`:""}${origen&&!it?`<span class="note">Leído de ${origen}; no lo encuentro en el catálogo.</span>`:""}
      <div class="ctl"><span class="seg" role="group" aria-label="Estado">
        <button class="f" data-t="falta" aria-pressed="${l.tipo==="falta"}">Falta</button><button class="s" data-t="sobra" aria-pressed="${l.tipo==="sobra"}">Sobra</button><button class="o" data-t="ok" aria-pressed="${l.tipo==="ok"}">OK</button></span>
        <span class="qty"><button data-q="-1" aria-label="Menos">−</button><input inputmode="decimal" value="${esc(fmtQ(l.cant))}" data-qi aria-label="Cantidad"><button data-q="1" aria-label="Más">+</button></span></div>
      <input class="nota" placeholder="Nota (opcional)" value="${esc(l.nota||"")}" data-nota></div>
    <button class="x" data-del aria-label="Quitar línea">✕</button></div>`;}
function renderEditor(){const r=revDraft,c=counts(r),R=$("#rev");
  const names={revisar:"Por revisar (leído de los documentos)",falta:"Faltan",sobra:"Sobran",ok:"Correcto (no sale en el correo)"};
  const groups=["revisar","falta","sobra","ok"].map(t=>({t,ls:r.lineas.map((l,i)=>[l,i]).filter(([l])=>l.tipo===t)})).filter(g=>g.ls.length);
  R.innerHTML=`<div class="rv-top"><button class="btn-s" data-rv="back">‹ Revisiones</button>${r.estado==="Enviada"?`<span class="pill env">Enviado ${fmtF(r.enviado)}</span>`:""}</div>
  <div class="box"><h3>Hoja de trabajo y presupuesto</h3>
    ${[["hoja","Hoja de trabajo"],["presupuesto","Presupuesto"]].map(([k,t])=>`<div class="docrow${docBusy[k]?" busy":""}"><b>${t}</b>
      <label class="btn-l docbtn">📷 Cámara<input type="file" accept="image/*" capture="environment" data-doc="${k}" aria-label="${t}: hacer foto"></label>
      <label class="btn-l docbtn">📎 Archivo<small>PDF o galería</small><input type="file" accept="image/*,application/pdf,.pdf" multiple data-doc="${k}" aria-label="${t}: adjuntar"></label></div>`).join("")}
    ${(r.docs||[]).length?`<div class="docs">${r.docs.map((d,j)=>`<div class="doc"><a href="https:\/\/drive.google.com/file/d/${esc(d.fileId)}/view" target="_blank" rel="noopener"><span class="pdf">${/pdf/.test(d.mime)?"PDF":"📷"}<br><small>${esc((d.name||"").slice(0,20))}</small></span></a>
      <small>${d.tipo==="hoja"?"Hoja":"Presupuesto"}</small><button class="x" data-docdel="${j}" aria-label="Quitar">✕</button></div>`).join("")}</div>`
      :`<p class="note">Haz una foto a cada página o adjunta el PDF. Se guardan en tu Drive, en la carpeta Documentos.</p>`}
    ${(r.docs||[]).length?(S.settings.gemini?`<button class="btn-l primary" data-rv="fill"${filling?" disabled":""}>${filling?"Gemini está leyendo…":"✨ Rellenar con los documentos"}</button>
      <p class="note">Gemini lee el nº de hoja, la OF, la obra y los componentes. Revisa lo que ponga antes de enviar.</p>`
      :`<p class="note" style="color:var(--warn)">Para rellenar solo, pon la clave de Gemini en ⚙ Ajustes.</p>`):""}</div>
  <div class="box">
    <label class="field"><span>Nº de hoja / orden</span><input id="rvHoja" value="${esc(r.hoja)}"></label>
    <label class="field"><span>Nº de orden de fabricación</span><input id="rvOf" value="${esc(r.of||"")}"></label>
    <label class="field"><span>Obra / cliente</span><input id="rvObra" value="${esc(r.obra)}"></label>
    <p class="note">Fecha: ${fmtF(r.fecha)}</p></div>
  <div class="box"><h3>Añadir componentes a mano</h3>
    <div class="addline"><input id="rvQ" type="search" placeholder="Buscar en el catálogo: código o nombre" autocomplete="off">
      <label class="scan" id="rvScanL" title="Añadir con una foto de la etiqueta"><span aria-hidden="true">🏷️</span><input id="rvScan" type="file" accept="image/*" capture="environment" data-pick aria-label="Añadir con una foto de la etiqueta"></label></div>
    <div class="sugg" id="rvSugg"></div>
    <button class="btn-s" data-rv="free" style="align-self:flex-start">＋ Componente que no está en el catálogo</button></div>
  <div class="box"><h3>Componentes</h3><div class="sumrow">${pills(c)}</div>
    ${groups.length?groups.map(g=>`<h2 style="margin:8px 0 0">${names[g.t]}</h2><div>${g.ls.map(([l,i])=>lineHTML(l,i)).join("")}</div>`).join(""):`<p class="note">Aún no hay componentes.</p>`}</div>
  <div class="box"><h3>Observaciones</h3><textarea id="rvObs" placeholder="Comentarios para el correo (opcional)">${esc(r.obs||"")}</textarea></div>
  <div class="box"><h3>Enviar</h3>
    <label class="field"><span>Para</span><input id="rvPara" type="email" value="${esc(S.settings.mailTo)}"></label>
    ${c.p?`<p class="note" style="color:var(--warn)">⚠ Hay ${c.p} componente${c.p>1?"s":""} por revisar. No saldrán en el correo hasta que los marques.</p>`:""}
    <div id="sendBox">${sendBoxHTML()}</div>
    <div class="acts2"><button class="btn-s" data-rv="xlsx">⬇ Excel</button><button class="btn-s" data-rv="copy">Copiar texto</button></div>
    <p class="note" id="rvMsg"></p></div>
  <div class="acts2" id="rvDelBox"><button class="btn-s obs" data-rv="del">Borrar esta revisión</button></div>`;}
function suggest(v){const q=norm(v.trim()),box=$("#rvSugg");if(!q){box.innerHTML="";return;}
  const res=items.filter(i=>!i.obsoleta&&norm([i.code,i.desc,i.modelo,i.medida].join(" ")).includes(q)).slice(0,6);
  box.innerHTML=res.length?res.map(i=>`<button data-add="${i.code}">${thumbOf(i)?`<img src="${thumbOf(i)}" alt="">`:`<span class="th"></span>`}<span><code>${i.code}</code><br>${esc(i.desc)}${i.medida?" · "+esc(i.medida):""}</span></button>`).join("")
    :`<p class="note">Nada en el catálogo con «${esc(v)}». Usa «＋ Componente que no está en el catálogo».</p>`;}
function addLine(code){const ex=revDraft.lineas.find(l=>l.code===code&&l.tipo==="falta");
  if(ex)ex.cant=(+ex.cant||0)+1;else revDraft.lineas.push({code,desc:"",tipo:"falta",cant:1,nota:""});persistRev();renderEditor();
  const it=itemOf(code);toast(`Añadido como «Falta»: ${it?it.desc:code}`);}
async function addDocs(tipo,files,auto){if(!revDraft)return;const rid=revDraft.id;docBusy[tipo]=true;renderEditor();toast("Guardando en tu Drive…",90000);let ok=0;
  try{for(const f of files){const isPdf=f.type==="application/pdf"||/\.pdf$/i.test(f.name||"");
      if(isPdf&&f.size>15e6){toast("Ese PDF pesa demasiado (más de 15 MB).",5000);continue;}
      const blob=isPdf?f:await shrink(f,2000,0.85),name=isPdf?(f.name||"documento.pdf"):`${tipo}_${Date.now()}.jpg`;
      const d=await api("addDoc",tipo,name,isPdf?"application/pdf":"image/jpeg",await blobB64(blob));
      if(!revDraft||revDraft.id!==rid)return;revDraft.docs=[...(revDraft.docs||[]),d];ok++;writeRev();}
    toast(!auto?"Guardada. Haz la siguiente página o pulsa «✨ Rellenar».":ok>1?`${ok} archivos guardados.`:"Archivo guardado.",auto?2600:5000);}
  catch(e){toast(errMsg(e),5000);}
  finally{docBusy[tipo]=false;if(revDraft&&revDraft.id===rid){renderEditor();if(auto&&ok&&S.settings.gemini&&!filling)autoFill();}}}
async function autoFill(){if(!revDraft||filling||!(revDraft.docs||[]).length)return;const rid=revDraft.id;filling=true;renderEditor();toast("Gemini está leyendo los documentos…",400000);
  try{const r=await api("aiFill",revDraft.docs);if(!revDraft||revDraft.id!==rid)return;
    const setIf=(k,v)=>{v=v==null?"":String(v).trim();if(v&&!String(revDraft[k]||"").trim())revDraft[k]=v;};setIf("hoja",r&&r.hoja);setIf("of",r&&r.of);setIf("obra",r&&r.obra);
    revDraft.lineas=revDraft.lineas.filter(l=>!(l.tipo==="revisar"&&(l.origen==="hoja"||l.origen==="presupuesto")));
    /* Solo componentes que están en el catálogo. */
    const L=(Array.isArray(r&&r.lineas)?r.lineas:[]).filter(x=>{const c=pick(x&&x.codigo);return c&&itemOf(c);});
    for(const x of L){const c=pick(x&&x.codigo);revDraft.lineas.push({code:c,desc:String((x&&x.descripcion)||""),tipo:"revisar",
      cant:numQ(x&&x.cantidad)||1,nota:"",origen:x&&x.documento==="presupuesto"?"presupuesto":"hoja"});}
    writeRev();toast(L.length?`He encontrado ${L.length} componentes del catálogo. Márcalos como Falta, Sobra u OK.`:"He rellenado los datos, pero no he encontrado componentes del catálogo.",6000);}
  catch(e){toast(errMsg(e),6000);}
  finally{filling=false;if(revDraft&&revDraft.id===rid)renderEditor();}}
function repRows(t){return revDraft.lineas.filter(l=>l.tipo===t).map(l=>{const it=itemOf(l.code);return {code:l.code||"",desc:it?it.desc:(l.desc||""),medida:it?it.medida:"",cant:+l.cant||1,nota:l.nota||""};});}
const repTitle=()=>`Hoja ${revDraft.hoja||"s/n"}${revDraft.of?" · OF "+revDraft.of:""}${revDraft.obra?" – "+revDraft.obra:""}`;
const repSubject=()=>`${repTitle()} – faltan ${repRows("falta").length}, sobran ${repRows("sobra").length} (${fmtF(revDraft.fecha)})`;
function repText(){const r=revDraft,blk=(t,rows)=>rows.length?`${t}:\n`+rows.map(x=>`- ${fmtQ(x.cant)} x ${x.code?x.code+" ":""}${x.desc}${x.medida?" ("+x.medida+")":""}${x.nota?" – "+x.nota:""}`).join("\n"):`${t}: nada`;
  return `${repTitle()} · ${fmtF(r.fecha)}\n\n${blk("FALTAN",repRows("falta"))}\n\n${blk("SOBRAN",repRows("sobra"))}${r.obs?`\n\nObservaciones: ${r.obs}`:""}`;}
function repHTML(){const r=revDraft,F="font-family:Arial,sans-serif;font-size:13px;",th=`style="text-align:left;padding:6px 10px;border-bottom:2px solid #333;${F}"`,td=`style="padding:6px 10px;border-bottom:1px solid #ddd;${F}"`;
  const tbl=(t,col,rows)=>`<h3 style="font-family:Arial,sans-serif;color:${col};margin:18px 0 6px">${t} (${rows.length})</h3>`+(rows.length?`<table style="border-collapse:collapse"><tr><th ${th}>Cant.</th><th ${th}>Código</th><th ${th}>Descripción</th><th ${th}>Nota</th></tr>${rows.map(x=>`<tr><td ${td}><b>${fmtQ(x.cant)}</b></td><td ${td}>${esc(x.code)}</td><td ${td}>${esc(x.desc)}${x.medida?" · "+esc(x.medida):""}</td><td ${td}>${esc(x.nota)}</td></tr>`).join("")}</table>`:`<p style="${F}">Nada.</p>`);
  return `<div><p style="${F}font-size:14px"><b>${esc(repTitle())}</b> · ${fmtF(r.fecha)}</p>${tbl("Faltan","#B8322A",repRows("falta"))}${tbl("Sobran","#2B64A8",repRows("sobra"))}${r.obs?`<p style="${F}"><b>Observaciones:</b> ${esc(r.obs)}</p>`:""}</div>`;}
function sendBoxHTML(){const r=revDraft,para=esc(($("#rvPara")&&$("#rvPara").value)||S.settings.mailTo),n=repRows("falta").length+repRows("sobra").length;
  const prev=r.estado==="Enviada"?`<p class="note">Enviado el ${fmtF(r.enviado)} a ${esc(r.enviadoA||S.settings.mailTo)}.</p>`:"";
  if(sendState==="sending")return `${prev}<button class="btn-l primary" disabled>Enviando…</button>`;
  if(sendState==="confirm")return `<div class="box" style="border-color:var(--accent)"><p style="margin:0"><b>¿Enviar ahora?</b></p>
    <p class="note">Para: <b>${para}</b><br>Asunto: ${esc(repSubject())}<br>Sale desde tu Gmail con ${n} línea${n!==1?"s":""}.</p>
    ${r.estado==="Enviada"?`<p class="note" style="color:var(--warn)">⚠ Esta hoja ya se envió el ${fmtF(r.enviado)}. Se mandará otra vez.</p>`:""}
    <div class="acts2"><button class="btn-l" data-rv="sendno">Cancelar</button><button class="btn-l primary" data-rv="sendyes">✉️ Sí, enviar</button></div></div>`;
  return `${prev}<button class="btn-l primary" data-rv="send">✉️ Enviar a ${para}</button>`;}
function refreshSend(){const b=$("#sendBox");if(b)b.innerHTML=sendBoxHTML();}
async function doSend(){const to=$("#rvPara").value.trim();sendState="sending";refreshSend();flushRev();
  try{const r=await api("sendRev",revDraft.id,to,repSubject(),repText(),repHTML());Object.assign(revDraft,{estado:"Enviada",enviado:r.enviado,enviadoA:r.enviadoA});
    const j=revs.findIndex(x=>x.id===revDraft.id);if(j>=0)Object.assign(revs[j],{estado:"Enviada",enviado:r.enviado});
    sendState="idle";renderEditor();$("#rvMsg").textContent=`✅ Enviado a ${to}. Lo tienes en «Enviados» de tu Gmail.`;toast("Correo enviado.");}
  catch(e){sendState="idle";refreshSend();$("#rvMsg").textContent=errMsg(e)+" Mira en «Enviados» antes de repetir, por si salió.";}}
$("#rev").addEventListener("click",async e=>{const t=e.target,a=t.closest("[data-rv]"),act=a&&a.dataset.rv;
  if(act==="new"){newRev();return;}
  const op=t.closest("[data-rv-open]");if(op){openRev(op.dataset.rvOpen);return;}
  if(act==="cons"){openCons();return;}
  if(act==="consBack"){consOn=false;renderRev();return;}
  const cp=t.closest("[data-cp]");if(cp){consP=cp.dataset.cp;consOff=0;renderCons();return;}
  const co=t.closest("[data-co]");if(co){consOff=Math.min(0,consOff+(+co.dataset.co));renderCons();return;}
  if(act==="consXlsx"){consXlsx();return;}
  if(!revDraft)return;
  if(act==="back"){closeRev();return;}
  if(act==="fill"){autoFill();return;}
  const dd=t.closest("[data-docdel]");if(dd){const j=+dd.dataset.docdel,d=revDraft.docs[j];revDraft.docs.splice(j,1);writeRev();renderEditor();if(d)api("delDoc",d.fileId).catch(()=>{});return;}
  if(act==="free"){revDraft.lineas.push({code:null,desc:"",tipo:"falta",cant:1,nota:""});persistRev();renderEditor();const ins=$("#rev").querySelectorAll("[data-desc]");ins.length&&ins[ins.length-1].focus();return;}
  if(act==="send"){const f=repRows("falta"),s=repRows("sobra");if(!f.length&&!s.length){toast("No hay nada marcado como Falta o Sobra.",4000);return;}
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test($("#rvPara").value.trim())){toast("Revisa la dirección de correo.",4000);return;}sendState="confirm";refreshSend();return;}
  if(act==="sendno"){sendState="idle";refreshSend();return;}
  if(act==="sendyes"){if(sendState==="confirm")doSend();return;}
  if(act==="xlsx"){try{await loadLib(XLSX_URL,"XLSX");const r=revDraft,H=["Cantidad","Código","Descripción","Medida","Nota"],f=repRows("falta"),s=repRows("sobra");
      const aoa=[["Hoja",r.hoja],["Orden de fabricación",r.of],["Obra / cliente",r.obra],["Fecha",fmtF(r.fecha)],[],["FALTAN"],H,...f.map(x=>[x.cant,x.code,x.desc,x.medida,x.nota]),[],["SOBRAN"],H,...s.map(x=>[x.cant,x.code,x.desc,x.medida,x.nota])];
      const ws=XLSX.utils.aoa_to_sheet(aoa);ws["!cols"]=[{wch:16},{wch:14},{wch:44},{wch:12},{wch:32}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Revisión");
      await deliver(new Blob([XLSX.write(wb,{type:"array",bookType:"xlsx"})]),`Hoja ${String(r.hoja||"sn").replace(/[^\w\-]+/g,"_")} ${hoy()}.xlsx`,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","drive",null,$("#rvMsg"));}catch(x){toast(errMsg(x),5000);}return;}
  if(act==="copy"){const txt=repText();try{await navigator.clipboard.writeText(txt);toast("Texto copiado.");}
    catch(_){$("#rvMsg").innerHTML=`<textarea readonly style="min-height:160px">${esc(txt)}</textarea>`;$("#rvMsg textarea").select();}return;}
  if(act==="del"){$("#rvDelBox").innerHTML=`<span class="q">¿Borrar esta revisión y sus documentos?</span><button class="btn-s" data-rv="delno">No</button><button class="btn-s yes" data-rv="delyes">Sí, borrar</button>`;return;}
  if(act==="delno"){renderEditor();return;}
  if(act==="delyes"){const id=revDraft.id;clearTimeout(revT);revT=null;revId=null;revDraft=null;revs=revs.filter(r=>r.id!==id);renderRev();
    try{await api("deleteRev",id);toast("Revisión borrada.");}catch(x){toast(errMsg(x),5000);}return;}
  const sg=t.closest("[data-add]");if(sg){addLine(sg.dataset.add);return;}
  const ln=t.closest(".ln");if(!ln)return;const l=revDraft.lineas[+ln.dataset.i];if(!l)return;
  const tb=t.closest("[data-t]");if(tb){l.tipo=tb.dataset.t;persistRev();renderEditor();return;}
  const qb=t.closest("[data-q]");if(qb){l.cant=Math.max(1,Math.round(((+l.cant||1)+(+qb.dataset.q))*1000)/1000);ln.querySelector("[data-qi]").value=fmtQ(l.cant);persistRev();return;}
  if(t.closest("[data-del]")){revDraft.lineas.splice(+ln.dataset.i,1);persistRev();renderEditor();}});
$("#rev").addEventListener("input",e=>{if(!revDraft)return;const t=e.target;
  if(t.id==="rvHoja"){revDraft.hoja=t.value;persistRev();return;}if(t.id==="rvOf"){revDraft.of=t.value;persistRev();return;}
  if(t.id==="rvObra"){revDraft.obra=t.value;persistRev();return;}if(t.id==="rvObs"){revDraft.obs=t.value;persistRev();return;}
  if(t.id==="rvQ"){suggest(t.value);return;}
  const ln=t.closest(".ln");if(!ln)return;const l=revDraft.lineas[+ln.dataset.i];if(!l)return;
  if(t.matches("[data-nota]"))l.nota=t.value;else if(t.matches("[data-desc]"))l.desc=t.value;else if(t.matches("[data-qi]")){const n=numQ(t.value);if(n>0)l.cant=n;}
  persistRev();});
$("#rev").addEventListener("change",async e=>{const t=e.target;
  if(t.id==="rvScan"&&t.files[0]){const f=t.files[0];t.value="";toast("Leyendo el código…",60000);const code=await readLabel(f);if(!revDraft)return;
    if(!code){toast("No he podido leer el código.",4500);return;}
    if(!itemOf(code)){revDraft.lineas.push({code,desc:"(no está en el catálogo)",tipo:"falta",cant:1,nota:""});persistRev();renderEditor();toast(`${code} no está en el catálogo; lo añado igualmente.`,4500);return;}
    addLine(code);return;}
  if(t.dataset&&t.dataset.doc&&t.files.length){let fs=[...t.files];const k=t.dataset.doc,cam=t.hasAttribute("capture");t.value="";
    if(cam&&fs[0]&&/^image\//.test(fs[0].type||"image/")){const c=await cropPhoto(fs[0]);if(!c)return;fs=[c];}addDocs(k,fs,!cam);}});
window.addEventListener("pagehide",flushRev);

/* =========================================================
   Ajustes
   ========================================================= */
function renderCfg(){chips();const s=S.settings,o=S.owner;
  $("#cfg").innerHTML=`<div class="cfg">
    <div class="box"><h3>Tu cuenta</h3><p class="note">Estás usando la app como <b>${esc(S.user||"(cuenta de Google)")}</b>${o?" · propietario":""}.</p></div>
    <div class="box"><h3>Correo de las revisiones</h3>
      <label class="field"><span>Enviar las hojas a</span><input id="cMail" type="email" value="${esc(s.mailTo)}"${o?"":" disabled"}></label></div>
    <div class="box"><h3>Gemini (IA)</h3>
      <p class="note">${s.gemini?`✅ Hay una clave guardada. Para leer etiquetas se usa el «flash» (rápido); para reconocer piezas y leer hojas de trabajo, el «pro» (más potente). Si uno no responde, se prueba otro.${s.last?` Último modelo usado: <b>${esc(s.last)}</b>.`:""}`:"Aún no hay clave. Sin ella no funcionan: leer hojas y presupuestos, leer etiquetas difíciles y reconocer productos."}</p>
      ${o?`<ol class="note" style="margin:0;padding-left:18px">
        <li>Abre <a href="https:\/\/aistudio.google.com/apikey" target="_blank" rel="noopener">aistudio.google.com/apikey</a> con tu cuenta de Google.</li>
        <li>Pulsa <b>Create API key</b> (Crear clave de API) y cópiala.</li>
        <li>Pégala aquí debajo y pulsa Guardar. No la escribas en ningún chat.</li></ol>
      <label class="field"><span>Clave de Gemini</span><input id="cKey" type="password" autocomplete="off" placeholder="${s.gemini?"(guardada; escribe otra para cambiarla)":"Pega aquí la clave"}"></label>
      <label class="field"><span>Modelo de reserva (si falla el «pro»)</span><input id="cModel" value="${esc(s.model)}"></label>
      <label class="field" style="display:flex;gap:8px;align-items:flex-start"><input type="checkbox" id="cOnlyPro"${s.onlyPro?" checked":""} style="width:auto;margin-top:4px"><span>Usar solo el «pro». Si no responde, avisa en vez de usar otro modelo</span></label>`:`<p class="note">Solo el propietario puede cambiar la clave.</p>`}</div>
    ${o?`<div class="acts2"><button class="btn-l primary" id="cSave">Guardar</button>${s.gemini?`<button class="btn-l" id="cTest">Probar Gemini</button><button class="btn-s obs" id="cDel">Quitar la clave</button>`:""}</div>`:""}
    <p class="note" id="cMsg"></p>
    <div class="box"><h3>Hojas de referencias electrónicas</h3>
      <p class="note">Se actualizan solas cada vez que guardas una referencia (alta, edición u obsoleta): <b>ELYA</b> → «REFERENCIAS ELÉCTRICAS UMARA.» y <b>PERGOHABITAT</b> → «REFERENCIAS ELECTRÓNICA PERGOHABITAT».</p>
      <div class="acts2"><button class="btn-s" id="cUmara">Actualizar las dos hojas ahora</button>
        <a class="btn-s" href="https://docs.google.com/spreadsheets/d/1wnrZd3f2KfYYpNl9vrYdX4kYDREvIwKs1hGj1rXMqgo/edit" target="_blank" rel="noopener" style="text-decoration:none">Abrir UMARA</a>
        <a class="btn-s" href="https://docs.google.com/spreadsheets/d/1h6ClF1fHd31I9R6iEbsF_i9FMzEZxRBqfgVcy3yIIMA/edit" target="_blank" rel="noopener" style="text-decoration:none">Abrir PERGOHABITAT</a></div>
      <p class="note" id="cUmaraMsg"></p></div>
    <div class="box"><h3>Fotos para ordenar</h3>
      <p class="note">Las fotos de la carpeta «FOTOS PARA ORDENAR» cuyo nombre empiece por el código y el tipo (por ejemplo <b>48009000000_producto_1.jpg</b>) se añaden a su referencia.</p>
      <div class="acts2"><button class="btn-s" id="cImport">Importar fotos ordenadas</button>
        <a class="btn-s" href="https://drive.google.com/drive/folders/1UaQeJOoxsXY7SdoYnfwo1BXPQOTT2hLF" target="_blank" rel="noopener" style="text-decoration:none">Abrir la carpeta</a></div>
      <p class="note" id="cImportMsg"></p></div>
    <div class="box"><h3>Dar acceso a un compañero</h3>
      <ol class="note" style="margin:0;padding-left:18px"><li>En tu Drive, comparte la carpeta <b>APP CATALOGO LLAZA</b> con su correo de Google como <b>Editor</b>.</li>
      <li>Pásale el enlace de esta app. La primera vez Google le pedirá permiso; aparecerá un aviso de «app no verificada»: pulsa <b>Configuración avanzada → Ir a Catálogo Llaza</b>.</li>
      <li>Las revisiones que envíe saldrán desde su propio Gmail.</li></ol></div>
    <div class="box"><h3>En el móvil</h3><p class="note">Abre la app en Chrome (Android) o Safari (iPhone) y usa <b>Añadir a pantalla de inicio</b>. Quedará como un icono más.</p></div>
  </div>`;
  $("#cImport").onclick=async()=>{const m=$("#cImportMsg"),bt=$("#cImport");bt.disabled=true;m.textContent="Importando… puede tardar un poco.";
    try{const r=await api("importSorted");const L=[`✅ ${r.importadas} foto${r.importadas===1?"":"s"} importada${r.importadas===1?"":"s"}.`];
      if(r.pendientes.length)L.push(`⏳ ${r.pendientes.length} aún sin miniatura en Drive; vuelve a pulsar en unos minutos.`);
      if(r.sinReferencia.length)L.push(`⚠ Sin referencia en el catálogo: ${esc(r.sinReferencia.join(", "))}`);
      if(r.sinNombre.length)L.push(`⚠ Sin ordenar (nombre sin código): ${r.sinNombre.length}`);
      m.innerHTML=L.join("<br>");if(r.importadas){const n=await api("boot");refs=n.refs;fotos=n.fotos;merge();}}
    catch(e){m.textContent=errMsg(e);}finally{bt.disabled=false;}};
  $("#cUmara").onclick=async()=>{const m=$("#cUmaraMsg");m.textContent="Actualizando…";
    try{const r=await api("syncSheetsNow");m.innerHTML=r.map(x=>`✅ ${esc(x.hoja)}: ${x.actualizadas} actualizadas, ${x.nuevas} nuevas`).join("<br>");}catch(e){m.textContent=errMsg(e);}};
  if(!o)return;
  $("#cSave").onclick=async()=>{const m=$("#cMsg");m.textContent="Guardando…";
    try{const r=await api("saveSettings",{mailTo:$("#cMail").value,gemini:$("#cKey").value||"",model:$("#cModel").value,onlyPro:$("#cOnlyPro").checked});S.settings=r;renderCfg();$("#cMsg").textContent="✅ Ajustes guardados.";}
    catch(e){m.textContent=errMsg(e);}};
  const tb=$("#cTest");if(tb)tb.onclick=async()=>{const m=$("#cMsg");m.textContent="Probando…";try{const r=await api("testGemini")||{};if(r.model)S.settings.last=r.model;
      const tries=(r.intentos||[]).map(x=>`<li><b>${esc(x.model)}</b>: error ${x.code}. ${esc(x.msg)}</li>`).join("");
      m.innerHTML=(r.ok?`✅ Gemini responde con <b>${esc(r.model)}</b>.${/pro|^gemini-\d+(\.\d+)?-(?!.*(flash|lite))[a-z]+$/.test(r.model)?"":" <b>No es el «pro».</b>"}`:`❌ ${esc(errMsg(new Error(r.error)))}`)+(tries?`<br>Lo que ha contestado Google a los modelos que fallaron:<ul style="margin:4px 0 0 18px">${tries}</ul>`:"");}catch(e){m.textContent=errMsg(e);}};
  const db=$("#cDel");if(db)db.onclick=async()=>{try{const r=await api("saveSettings",{clearGemini:true});S.settings=r;renderCfg();$("#cMsg").textContent="Clave quitada.";}catch(e){$("#cMsg").textContent=errMsg(e);}};}

/* =========================================================
   Exportar
   ========================================================= */
function sortItems(L){const order=[];SECS.forEach(S=>catsOf(S.n).forEach(c=>order.push(S.n+"|"+c)));const idx=i=>{const k=order.indexOf(i.sec+"|"+i.cat);return k<0?999:k;};
  const mi=i=>{const k=[...MARCAS,OTRAS].indexOf(marcaOf(i));return k<0?-1:k;};
  return [...L].sort((a,b)=>idx(a)-idx(b)||mi(a)-mi(b)||a.code.localeCompare(b.code));}
function scopeList(k,obs){const inc=i=>obs||!i.obsoleta;let L;
  if(k==="one")L=items.filter(i=>i.code===expOne);else if(k==="stk")L=items.filter(i=>inS(i,sec)&&hasStock(i)&&inc(i));else if(k==="all")L=items.filter(inc);
  else{const addSel=!!($("#expAddSel")&&$("#expAddSel").checked);L=items.filter(i=>(secsOf(i).some(s=>expCats.has(s+"|"+i.cat)||expCats.has(s+"|"+i.cat+"|"+marcaOf(i)))&&inc(i))||(addSel&&selSet.has(i.code)));}
  if(k!=="one"&&$("#expStkOnly")&&$("#expStkOnly").checked)L=L.filter(hasStock);
  return sortItems(L);}
function scopeName(k){if(k==="one"){const it=itemOf(expOne);return `${expOne}${it?" "+it.desc:""}`;}if(k==="all")return "Completo";if(k==="stk")return `Stock seguridad${sec===ALL?"":" "+sec}`;
  const parts=[...expCats].map(x=>x.split("|").slice(1).join(" · "));if($("#expStkOnly")&&$("#expStkOnly").checked)parts.unshift("Stock seguridad");if($("#expAddSel")&&$("#expAddSel").checked&&selSet.size)parts.push(`${selSet.size} sueltas`);
  return parts.length?(parts.length<=3?parts.join(", "):`${parts.slice(0,3).join(", ")} y más`):"Selección";}
const safeName=s=>String(s).replace(/[\\/:*?"<>|]+/g," ").replace(/\s+/g," ").trim().slice(0,80);
const countCat=(s,c,obs,m)=>items.filter(i=>inS(i,s)&&i.cat===c&&(obs||!i.obsoleta)&&(!m||marcaOf(i)===m)).length;
/* Marcas de una categoría con subcarpetas por marca (para exportar solo una marca). */
const expMarcas=(s,c,obs)=>hasMarca(c)?marcasIn(items.filter(i=>inS(i,s)&&i.cat===c&&(obs||!i.obsoleta))):[];
function renderWhat(val){const obs=$("#expObs").checked,it=itemOf(expOne);
  const tree=SECS.map(S=>{const cs=catsOf(S.n).filter(c=>countCat(S.n,c,obs));if(!cs.length)return "";const allOn=cs.every(c=>expCats.has(S.n+"|"+c));
    return `<div class="ctree"><label class="csec"><input type="checkbox" data-secall="${esc(S.n)}"${allOn?" checked":""}><b>${esc(S.n)}</b><small>todo el apartado</small></label>
      <div class="cgrid">${cs.map(c=>{const k=S.n+"|"+c,on=expCats.has(k),ms=expMarcas(S.n,c,obs);
        return `<label><input type="checkbox" data-cat="${esc(k)}"${on?" checked":""}><span>${esc(c)} <small>(${countCat(S.n,c,obs)})</small></span></label>`+
          (ms.length>1?`<div class="cmarcas">${ms.map(m=>`<label><input type="checkbox" data-marca="${esc(k+"|"+m)}"${on||expCats.has(k+"|"+m)?" checked":""}><span>${esc(m)} <small>(${countCat(S.n,c,obs,m)})</small></span></label>`).join("")}</div>`:"");}).join("")}</div></div>`;}).join("");
  const stkOn=!!($("#expStkOnly")&&$("#expStkOnly").checked);
  $("#expWhat").innerHTML=`${expOne?`<label><input type="radio" name="expWhat" value="one"><span>Solo esta: ${esc(expOne)}${it?" · "+esc(it.desc):""}</span></label>`:""}
    ${items.some(i=>inS(i,sec)&&hasStock(i))?`<label><input type="radio" name="expWhat" value="stk"><span><b>Lista de stock de seguridad</b><small>${sec===ALL?"Todos los apartados":esc(sec)} · ${items.filter(i=>inS(i,sec)&&hasStock(i)&&(obs||!i.obsoleta)).length} referencias. En Excel sale con columnas para el recuento y lo que hay que pedir.</small></span></label>`:""}
    <label><input type="radio" name="expWhat" value="mix"><span><b>Elegir subgrupos</b><small>Marca una o varias categorías, de cualquier apartado</small></span></label>
    <div id="expTree">${tree}<label class="csel"><input type="checkbox" id="expStkOnly"${stkOn?" checked":""}><span>Solo las que tienen stock de seguridad<small>Para sacar la lista de stock de una categoría o marca</small></span></label><label class="csel"><input type="checkbox" id="expAddSel"${selSet.size?(expAddOn?" checked":""):" disabled"}><span>Sumar las referencias sueltas elegidas (${selSet.size})<small>Con «Elegir referencias una a una», aquí abajo</small></span></label></div>
    <label><input type="radio" name="expWhat" value="all"><span>Todo el catálogo (${items.filter(i=>obs||!i.obsoleta).length})</span></label><p class="note" id="expCount"></p>`;
  const r=document.querySelector(`input[name=expWhat][value="${val||(expOne?"one":view==="cat"&&cur===STK?"stk":"mix")}"]`);if(r)r.checked=true;updCount();}
const curWhat=()=>(document.querySelector("input[name=expWhat]:checked")||{}).value||"mix";
function updCount(){const n=scopeList(curWhat(),$("#expObs").checked).length;$("#expCount").textContent=n?`Se exportarán ${n} referencia${n===1?"":"s"}.`:"Marca al menos un subgrupo o una referencia.";$("#expGo").disabled=!n;}
function openExp(one){expOne=one||null;$("#expMsg").textContent="";
  if(!one&&!expCats.size&&view==="cat"&&cur!=="Todas"&&cur!==OBS&&cur!==INC&&cur!==STK&&cur!==DOC&&cur!==PROP)(sec===ALL?SECS.map(x=>x.n):[sec]).forEach(x=>expCats.add(x+"|"+cur+(marca!=="Todas"?"|"+marca:"")));
  renderWhat();let fm=null,sz=null;try{fm=localStorage.getItem("cl_expFmt");sz=localStorage.getItem("cl_expSize");}catch(_){}
  const setR=(n,v)=>{const r=v&&document.querySelector(`input[name=${n}][value="${v}"]`);if(r)r.checked=true;};setR("expFmt",fm||(expOne?"png":"tabla"));setR("expSize",sz);updFmt();
  if(!$("#expMailTo").value)$("#expMailTo").value="";$("#expSheet").hidden=false;}
$("#expWhat").addEventListener("change",e=>{const t=e.target;
  if(t.dataset.cat){const k=t.dataset.cat;[...expCats].forEach(x=>{if(x.startsWith(k+"|"))expCats.delete(x);});t.checked?expCats.add(k):expCats.delete(k);renderWhat("mix");return;}
  if(t.dataset.marca){const km=t.dataset.marca,k=km.split("|").slice(0,2).join("|"),[s,c]=k.split("|");
    if(expCats.has(k)){expCats.delete(k);expMarcas(s,c,$("#expObs").checked).forEach(m=>expCats.add(k+"|"+m));}
    t.checked?expCats.add(km):expCats.delete(km);
    const ms=expMarcas(s,c,$("#expObs").checked);if(ms.length&&ms.every(m=>expCats.has(k+"|"+m))){ms.forEach(m=>expCats.delete(k+"|"+m));expCats.add(k);}
    renderWhat("mix");return;}
  if(t.id==="expStkOnly"){renderWhat(curWhat());return;}
  if(t.dataset.secall){const s=t.dataset.secall;[...expCats].forEach(x=>{if(x.split("|").length>2&&x.startsWith(s+"|"))expCats.delete(x);});catsOf(s).forEach(c=>t.checked?expCats.add(s+"|"+c):expCats.delete(s+"|"+c));renderWhat("mix");return;}
  if(t.id==="expAddSel"){expAddOn=t.checked;renderWhat("mix");return;}updCount();});
function updFmt(){const f=(document.querySelector("input[name=expFmt]:checked")||{}).value;$("#expSizeF").hidden=!(f==="tabla2"||f==="tabla"||f==="stock4");}
$("#expSheet").addEventListener("change",e=>{const t=e.target;if(t.name!=="expFmt"&&t.name!=="expSize")return;try{localStorage.setItem(t.name==="expFmt"?"cl_expFmt":"cl_expSize",t.value);}catch(_){}updFmt();});
$("#expBtn").onclick=()=>openExp();$("#expClose").onclick=()=>{$("#expSheet").hidden=true;};
$("#expObs").addEventListener("change",()=>renderWhat(curWhat()));
$("#expSheet").addEventListener("click",e=>{if(e.target.id==="expSheet")$("#expSheet").hidden=true;});
$("#expMailTo").addEventListener("focus",()=>{const r=document.querySelector('input[name=expDest][value=mail]');if(r)r.checked=true;});
function updSel(){$("#selN").textContent=`${selSet.size} elegida${selSet.size===1?"":"s"}`;$("#selGo").disabled=!selSet.size;}
function setSel(on){selMode=on;$("#selBar").hidden=!on;if(on){updSel();if(view!=="cat"){flushRev();revId=null;revDraft=null;setView("cat");chips();}}updateFab();render();}
$("#expPick").onclick=()=>{$("#expSheet").hidden=true;setSel(true);toast("Toca el círculo de cada referencia que quieras exportar.",4500);};
$("#selAll").onclick=()=>{document.querySelectorAll("#list [data-sel]").forEach(b=>selSet.add(b.dataset.sel));render();updSel();};
$("#selNone").onclick=()=>{selSet.clear();render();updSel();};
$("#selExit").onclick=()=>{selSet.clear();setSel(false);};
$("#selGo").onclick=()=>openExp();

async function deliver(blob,name,mime,dest,mailTo,msg){
  if(dest==="down"){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();
    msg.innerHTML="Descargando… Si no aparece, elige «Guardarlo en mi Drive».";return;}
  msg.textContent="Subiendo…";const data=await blobB64(blob);
  if(dest==="mail"){await api("mailExport",mailTo,`Catálogo Llaza · ${name.replace(/\.[^.]+$/,"")}`,"Te envío el archivo adjunto del catálogo de Llaza.",name,mime,data);msg.textContent=`✅ Enviado a ${mailTo}.`;return;}
  const r=await api("saveExport",name,mime,data);msg.innerHTML=`✅ Guardado en tu Drive (Exportaciones): <a class="linkout" href="${esc(r.url)}" target="_blank" rel="noopener">abrir ${esc(r.name)}</a>`;}
$("#expGo").onclick=async()=>{const what=curWhat(),fmt=(document.querySelector("input[name=expFmt]:checked")||{}).value||"tabla",size=(document.querySelector("input[name=expSize]:checked")||{}).value||"m",dest=(document.querySelector("input[name=expDest]:checked")||{}).value||"drive";
  const allPh=$("#expAllPh").checked,obs=$("#expObs").checked,list=scopeList(what,obs),msg=$("#expMsg"),btn=$("#expGo"),mailTo=$("#expMailTo").value.trim();
  if(!list.length){msg.textContent="No hay referencias que exportar con esa elección.";return;}
  if(dest==="mail"&&!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mailTo)){msg.textContent="Escribe el correo al que quieres enviarlo.";return;}
  const name=safeName(`Catalogo Llaza - ${scopeName(what)} ${hoy()}`);btn.disabled=true;
  try{let out;if(fmt==="xlsx")out=what==="stk"||(what!=="one"&&$("#expStkOnly")&&$("#expStkOnly").checked)?await expStockXlsx(list,name,msg):await expXlsx(list,name,msg);else if(fmt==="png")out=await expPng(list,name,allPh,msg);else out=await expDocx(list,name,scopeName(what),fmt,allPh,msg,size);
    await deliver(out.blob,out.name,out.mime,dest,mailTo,msg);}
  catch(e){msg.textContent=errMsg(e);}finally{btn.disabled=false;}};

async function expXlsx(list,name,msg){msg.textContent="Preparando el Excel…";await loadLib(XLSX_URL,"XLSX");
  const rows=list.map(i=>[secsOf(i).join(", "),i.cat,marcaOf(i),i.code,i.desc,i.medida,i.canal,i.modelo,hasStock(i)?i.stock:"",i.photos.length?"Sí":"No",i.prod.length?"Sí":"No",i.obsoleta?"Sí":"No"]);
  const ws=XLSX.utils.aoa_to_sheet([["Apartado","Categoría","Marca","Código","Descripción","Medida / formato","Canal / color","Modelo / notas","Stock seguridad","Foto etiqueta","Foto producto","Obsoleta"],...rows]);
  ws["!cols"]=[{wch:18},{wch:20},{wch:14},{wch:14},{wch:40},{wch:15},{wch:12},{wch:40},{wch:15},{wch:12},{wch:13},{wch:10}];ws["!autofilter"]={ref:`A1:L${rows.length+1}`};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Catálogo");
  return {blob:new Blob([XLSX.write(wb,{type:"array",bookType:"xlsx"})]),name:name+".xlsx",mime:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};}
/* Lista de stock de seguridad: «Stock actual» se rellena al hacer el recuento y «A pedir» se calcula sola. */
async function expStockXlsx(list,name,msg){msg.textContent="Preparando la lista de stock…";await loadLib(XLSX_URL,"XLSX");
  const H=["Apartado","Categoría","Marca","Código","Descripción","Medida / formato","Stock seguridad","Stock actual","A pedir"];
  const ws=XLSX.utils.aoa_to_sheet([[`Lista de stock de seguridad · ${fmtF(new Date().toISOString())}`],["Rellena «Stock actual» al contar; «A pedir» = stock de seguridad − stock actual (0 si hay suficiente)."],[],H,
    ...list.map(i=>[secsOf(i).join(", "),i.cat,marcaOf(i),i.code,i.desc,i.medida,Number(i.stock),"",""])]);
  list.forEach((_,k)=>{const r=k+5;ws["I"+r]={t:"n",f:`IF(H${r}="","",MAX(0,G${r}-H${r}))`};});
  ws["!cols"]=[{wch:16},{wch:22},{wch:14},{wch:14},{wch:40},{wch:15},{wch:15},{wch:13},{wch:10}];ws["!autofilter"]={ref:`A4:I${list.length+4}`};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Stock seguridad");
  return {blob:new Blob([XLSX.write(wb,{type:"array",bookType:"xlsx"})]),name:name+".xlsx",mime:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};}
async function bmpOf(p,full){const src=full?await fullSrc(p):p.src;return await createImageBitmap(await (await fetch(src)).blob());}
async function imgBuf(p,full,max=620){const bmp=await bmpOf(p,full);
  const rot=p.rot||0,s=Math.min(1,max/Math.max(bmp.width,bmp.height)),w=Math.round(bmp.width*s),h=Math.round(bmp.height*s),side=rot===90||rot===270;
  const c=document.createElement("canvas");c.width=side?h:w;c.height=side?w:h;const g=c.getContext("2d");g.fillStyle="#fff";g.fillRect(0,0,c.width,c.height);
  g.translate(c.width/2,c.height/2);g.rotate(rot*Math.PI/180);g.drawImage(bmp,-w/2,-h/2,w,h);
  const b=await new Promise(res=>c.toBlob(res,"image/jpeg",0.82));return {data:await b.arrayBuffer(),w:c.width,h:c.height};}
const pics=(i,allPh)=>({e:allPh?i.photos:i.photos.slice(0,1),p:allPh?i.prod:i.prod.slice(0,1)});
async function expDocx(list,name,title,fmt,allPh,msg,size){msg.textContent="Preparando el Word…";await loadLib(DOCX_URL,"docx");const D=window.docx;
  const P=(t,{center,...o}={})=>new D.Paragraph({alignment:center?D.AlignmentType.CENTER:undefined,children:[new D.TextRun({text:String(t||""),...o})]});
  const cell=(children,w,shade)=>new D.TableCell({children,width:{size:w,type:D.WidthType.DXA},verticalAlign:D.VerticalAlign.CENTER,margins:{top:60,bottom:60,left:80,right:80},shading:shade?{fill:"D9DEE3",type:D.ShadingType.CLEAR,color:"auto"}:undefined});
  const full=fmt==="fichas";
  async function imgPar(p,maxW,maxH){try{const im=await imgBuf(p,full);let w=maxW,h=Math.round(im.h*w/im.w);if(h>maxH){w=Math.round(w*maxH/h);h=maxH;}
      return new D.Paragraph({alignment:D.AlignmentType.CENTER,spacing:{after:80},children:[new D.ImageRun({data:im.data,transformation:{width:w,height:h}})]});}catch(_){return P("(sin foto)",{italics:true,center:true});}}
  function notesPar(t){const runs=[];let k=0;t=String(t||"");
    for(const m of t.matchAll(URL_RE)){if(m.index>k)runs.push(new D.TextRun(t.slice(k,m.index)));
      runs.push(new D.ExternalHyperlink({link:m[0],children:[new D.TextRun({text:linkLabel(m[0]).replace(/^\S+\s/,""),style:"Hyperlink",color:"0563C1",underline:{}})]}));k=m.index+m[0].length;}
    if(k<t.length)runs.push(new D.TextRun(t.slice(k)));return [new D.Paragraph({children:runs})];}
  let done=0;const step=()=>{msg.textContent=`Preparando el Word… ${++done} de ${list.length}`;};let sections;
  if(fmt==="tabla2"){const W=[5233,5233],Z={s:[300,120],m:[320,185],l:[330,290]}[size]||[320,185];
    const head=()=>new D.TableRow({tableHeader:true,children:["Etiqueta y referencia","Producto"].map((t,j)=>cell([P(t,{bold:true,center:true})],W[j],true))});
    async function picsOf(ps){if(!ps.length)return [P("Falta foto",{italics:true,color:"B03A10",center:true})];const k=ps.length>1?0.6:1,out=[];for(const p of ps)out.push(await imgPar(p,Math.round(Z[0]*k),Math.round(Z[1]*k)));return out;}
    const children=[new D.Paragraph({heading:D.HeadingLevel.TITLE,children:[new D.TextRun(`Catálogo Llaza · ${title}`)]}),P(`${list.length} referencias · ${fmtF(new Date().toISOString())}`)];let first=true;
    for(const S0 of SECS){const si=list.filter(i=>inS(i,S0.n));if(!si.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_1,pageBreakBefore:!first,children:[new D.TextRun(S0.n)]}));first=false;
      for(const c of catsOf(S0.n)){const ci=si.filter(i=>i.cat===c);if(!ci.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_2,children:[new D.TextRun(`${c} (${ci.length})`)]}));
        const rows=[head()];for(const i of ci){step();const ph=pics(i,allPh),extra=[i.medida,i.canal,i.modelo,stockTxt(i)].filter(Boolean).join(" · ");
          const left=[...await picsOf(ph.e),P(isSinRef(i.code)?"Sin código":i.code,{bold:true,center:true,size:26}),P(i.desc,{center:true}),...(extra?[P(extra,{center:true,color:"5A6570",size:18})]:[]),...(i.obsoleta?[P("OBSOLETA",{bold:true,color:"6B4E8C",center:true})]:[])];
          rows.push(new D.TableRow({cantSplit:true,children:[cell(left,W[0]),cell(await picsOf(ph.p),W[1])]}));}
        children.push(new D.Table({rows,columnWidths:W,width:{size:W[0]+W[1],type:D.WidthType.DXA}}));}}
    sections=[{properties:{page:{margin:{top:720,bottom:720,left:720,right:720}}},children}];}
  else if(fmt==="stock4"){const W=[2750,2750,3466,1500],Z={s:[110,120],m:[160,170],l:[200,220]}[size]||[160,170];
    const head=()=>new D.TableRow({tableHeader:true,children:["Foto etiqueta","Foto producto","Descripción","Stock"].map((t,j)=>cell([P(t,{bold:true,center:true})],W[j],true))});
    async function picCell(ps,w){if(!ps.length)return cell([P("Falta",{italics:true,color:"B03A10",center:true})],w);const k=ps.length>1?0.7:1,out=[];for(const p of ps)out.push(await imgPar(p,Math.round(Z[0]*k),Math.round(Z[1]*k)));return cell(out,w);}
    const children=[new D.Paragraph({heading:D.HeadingLevel.TITLE,children:[new D.TextRun(`Catálogo Llaza · ${title}`)]}),P(`${list.length} referencias · ${fmtF(new Date().toISOString())}`)];let first=true;
    for(const S0 of SECS){const si=list.filter(i=>inS(i,S0.n));if(!si.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_1,pageBreakBefore:!first,children:[new D.TextRun(S0.n)]}));first=false;
      for(const c of catsOf(S0.n)){const ci=si.filter(i=>i.cat===c);if(!ci.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_2,children:[new D.TextRun(`${c} (${ci.length})`)]}));
        const rows=[head()];for(const i of ci){step();const ph=pics(i,allPh),extra=[marcaOf(i)&&marcaOf(i)!==OTRAS?marcaOf(i):"",i.medida,i.canal].filter(Boolean).join(" · ");
          rows.push(new D.TableRow({cantSplit:true,children:[await picCell(ph.e,W[0]),await picCell(ph.p,W[1]),
            cell([P(isSinRef(i.code)?"Sin código":i.code,{bold:true}),P(i.desc),...(extra?[P(extra,{color:"5A6570",size:18})]:[]),...(i.obsoleta?[P("OBSOLETA",{bold:true,color:"6B4E8C"})]:[])],W[2]),
            cell([P(hasStock(i)?String(i.stock):"—",{bold:true,center:true,size:32})],W[3])]}));}
        children.push(new D.Table({rows,columnWidths:W,width:{size:W.reduce((a,b)=>a+b,0),type:D.WidthType.DXA}}));}}
    sections=[{properties:{page:{margin:{top:720,bottom:720,left:720,right:720}}},children}];}
  else if(fmt==="tabla"){const W=[2700,2700,1600,2700,3700];
    const head=()=>new D.TableRow({tableHeader:true,children:["Foto etiqueta","Foto producto","Código","Descripción","Notas / especificaciones"].map((t,j)=>cell([P(t,{bold:true,center:true})],W[j],true))});
    async function picCell(ps,w){if(!ps.length)return cell([P("Falta",{italics:true,color:"B03A10",center:true})],w);const out=[];for(const p of ps){const z={s:[100,120],m:[150,180],l:[220,260]}[size]||[150,180],k=ps.length>1?0.75:1;out.push(await imgPar(p,Math.round(z[0]*k),Math.round(z[1]*k)));}return cell(out,w);}
    const children=[new D.Paragraph({heading:D.HeadingLevel.TITLE,children:[new D.TextRun(`Catálogo Llaza · ${title}`)]}),P(`${list.length} referencias · ${fmtF(new Date().toISOString())}`)];let first=true;
    for(const S0 of SECS){const si=list.filter(i=>inS(i,S0.n));if(!si.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_1,pageBreakBefore:!first,children:[new D.TextRun(S0.n)]}));first=false;
      for(const c of catsOf(S0.n)){const ci=si.filter(i=>i.cat===c);if(!ci.length)continue;children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_2,children:[new D.TextRun(`${c} (${ci.length})`)]}));
        const rows=[head()];for(const i of ci){step();const ph=pics(i,allPh);
          rows.push(new D.TableRow({cantSplit:true,children:[await picCell(ph.e,W[0]),await picCell(ph.p,W[1]),cell([P(i.code,{bold:true,center:true})],W[2]),
            cell([P(i.desc),...(i.obsoleta?[P("OBSOLETA",{bold:true,color:"6B4E8C"})]:[])],W[3]),cell(notesPar([i.medida,i.modelo,i.canal,stockTxt(i),...docsOf(i.code).filter(m=>m.url).map(m=>m.titulo+": "+m.url)].filter(Boolean).join(" · ")),W[4])]}));}
        children.push(new D.Table({rows,columnWidths:W,width:{size:W.reduce((a,b)=>a+b,0),type:D.WidthType.DXA}}));}}
    sections=[{properties:{page:{size:{orientation:D.PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children}];}
  else{const children=[];
    for(const [n,i] of list.entries()){step();const ph=pics(i,allPh);
      children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_1,pageBreakBefore:n>0,children:[new D.TextRun(i.code)]}));children.push(P(i.desc,{bold:true,size:30}));
      children.push(P(`${secsOf(i).join(" + ")} · ${i.cat}${i.obsoleta?" · OBSOLETA":""}`,{color:"5A6570"}));
      const det=[["Medida / formato",i.medida],["Canal / color",i.canal],["Modelo / notas",i.modelo],["Stock de seguridad",hasStock(i)?String(i.stock):""]].filter(x=>x[1]);
      if(det.length)children.push(new D.Table({columnWidths:[2600,6400],width:{size:9000,type:D.WidthType.DXA},rows:det.map(([k,v])=>new D.TableRow({children:[cell([P(k,{bold:true})],2600,true),cell([P(v)],6400)]}))}));
      const dl=docsOf(i.code).filter(m=>m.url);if(dl.length){children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_3,children:[new D.TextRun("Documentación")]}));dl.forEach(m=>children.push(...notesPar(`${docIco(m)} ${m.titulo}: ${m.url}`)));}
      for(const [lab,ps] of [["Foto de etiqueta",ph.e],["Foto de producto",ph.p]]){children.push(new D.Paragraph({heading:D.HeadingLevel.HEADING_3,children:[new D.TextRun(lab)]}));
        if(!ps.length){children.push(P("Falta foto",{italics:true,color:"B03A10"}));continue;}for(const p of ps)children.push(await imgPar(p,ps.length>1?260:420,ps.length>1?240:330));}}
    sections=[{properties:{page:{margin:{top:900,bottom:900,left:1000,right:1000}}},children}];}
  const blob=await D.Packer.toBlob(new D.Document({styles:{default:{document:{run:{font:"Arial",size:21}}}},sections}));
  return {blob,name:name+".docx",mime:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"};}
function wrapLines(g,text,maxW){const words=String(text||"").split(/\s+/),lines=[];let line="";for(const w of words){const t=line?line+" "+w:w;if(g.measureText(t).width>maxW&&line){lines.push(line);line=w;}else line=t;}if(line)lines.push(line);return lines;}
function drawFit(g,bmp,rot,x,y,w,h){g.save();g.fillStyle="#FFFFFF";g.fillRect(x,y,w,h);g.beginPath();g.rect(x,y,w,h);g.clip();g.translate(x+w/2,y+h/2);g.rotate((rot||0)*Math.PI/180);
  const side=rot===90||rot===270,bw=side?bmp.height:bmp.width,bh=side?bmp.width:bmp.height,s=Math.min(w/bw,h/bh);g.drawImage(bmp,-bmp.width*s/2,-bmp.height*s/2,bmp.width*s,bmp.height*s);g.restore();}
async function cardPng(i,allPh){try{await Promise.all(["600 60px 'IBM Plex Mono'","600 38px 'IBM Plex Sans'"].map(f=>document.fonts.load(f)));}catch(_){}
  const W=1080,pad=44,ph=pics(i,allPh),cells=[...ph.e.map(p=>["ETIQUETA",p]),...ph.p.map(p=>["PRODUCTO",p])];
  if(!ph.e.length)cells.unshift(["ETIQUETA",null]);if(!ph.p.length)cells.push(["PRODUCTO",null]);
  const cw=(W-pad*2-16)/2,rowsN=Math.ceil(cells.length/2),m=document.createElement("canvas").getContext("2d");
  m.font="600 38px 'IBM Plex Sans', Arial";const dl=wrapLines(m,i.desc,W-pad*2);m.font="500 28px 'IBM Plex Mono', monospace";const ml=i.modelo?wrapLines(m,i.modelo,W-pad*2):[];
  const tags=[i.medida,i.canal,stockTxt(i)].filter(Boolean).join("  ·  "),H=Math.ceil(96+pad+rowsN*(cw+16)+24+70+dl.length*48+(tags?46:0)+ml.length*38+pad+40);
  const c=document.createElement("canvas");c.width=W;c.height=H;const g=c.getContext("2d");g.fillStyle="#FFFFFF";g.fillRect(0,0,W,H);
  g.fillStyle="#C9501C";g.fillRect(0,0,W,96);g.fillStyle="#FFFFFF";g.font="600 30px 'IBM Plex Sans', Arial";g.textBaseline="middle";g.fillText(`LLAZA  ·  ${i.sec.toUpperCase()}  ·  ${i.cat.toUpperCase()}`,pad,48);
  const y0=96+pad;
  for(let k=0;k<cells.length;k++){const [lab,p]=cells[k],x=pad+(k%2)*(cw+16),yy=y0+Math.floor(k/2)*(cw+16);g.fillStyle="#ECEEF0";g.fillRect(x,yy,cw,cw);
    if(p){try{drawFit(g,await bmpOf(p,true),p.rot,x,yy,cw,cw);}catch(_){}}else{g.fillStyle="#8A949E";g.font="500 30px 'IBM Plex Sans', Arial";g.textAlign="center";g.fillText("Sin foto",x+cw/2,yy+cw/2);g.textAlign="left";}
    g.font="600 22px 'IBM Plex Sans', Arial";const tw=g.measureText(lab).width+20;g.fillStyle="rgba(0,0,0,.7)";g.fillRect(x+10,yy+10,tw,34);g.fillStyle="#FFFFFF";g.fillText(lab,x+20,yy+27);}
  let y=y0+rowsN*(cw+16)+24;g.textBaseline="alphabetic";g.fillStyle="#1B2127";g.font="600 60px 'IBM Plex Mono', monospace";g.fillText(i.code,pad,y+56);y+=70;
  g.font="600 38px 'IBM Plex Sans', Arial";for(const l of dl){y+=48;g.fillText(l,pad,y);}
  if(tags){y+=46;g.fillStyle="#5A6570";g.font="500 30px 'IBM Plex Sans', Arial";g.fillText(tags,pad,y);}
  g.fillStyle="#5A6570";g.font="500 28px 'IBM Plex Mono', monospace";for(const l of ml){y+=38;g.fillText(l,pad,y);}
  g.fillStyle="#9AA5B0";g.font="500 22px 'IBM Plex Sans', Arial";g.fillText(`Catálogo Llaza · ${fmtF(new Date().toISOString())}${i.obsoleta?" · OBSOLETA":""}`,pad,H-pad/2);
  return await new Promise(r=>c.toBlob(r,"image/png"));}
async function expPng(list,name,allPh,msg){
  if(list.length===1){msg.textContent="Preparando la imagen…";return {blob:await cardPng(list[0],allPh),name:safeName(`${list[0].code} ${list[0].desc}`)+".png",mime:"image/png"};}
  await loadLib(JSZIP_URL,"JSZip");const zip=new JSZip();let n=0;
  for(const i of list){msg.textContent=`Preparando las imágenes… ${++n} de ${list.length}`;zip.file(safeName(`${i.code} ${i.desc}`)+".png",await cardPng(i,allPh));}
  return {blob:await zip.generateAsync({type:"blob"}),name:name+".zip",mime:"application/zip"};}

document.addEventListener("keydown",e=>{if(e.key==="Escape"){if(fichaCode&&$("#lb").hidden&&$("#addSheet").hidden&&$("#expSheet").hidden){if(history.state&&history.state.ficha)history.back();else closeFicha();}$("#lb").hidden=true;$("#addSheet").hidden=true;$("#expSheet").hidden=true;}
  if($("#lb").hidden)return;if(e.key==="ArrowRight")$("#lbNext").click();else if(e.key==="ArrowLeft")$("#lbPrev").click();});
start();

})();
