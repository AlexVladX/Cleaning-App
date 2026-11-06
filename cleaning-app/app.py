(function(){
"use strict";
const STAGES={
  nou:{label:"Noi",hint:"Încă neanalizată"},
  acum:{label:"Acum",hint:"Lucrez la ea, maximum 2"},
  explorez:{label:"Explorez",hint:"Pași mici, orizont lung"},
  parcare:{label:"Parcare",hint:"O recitesc la o dată"},
  renuntat:{label:"Renunțate",hint:"Nu o mai urmăresc"}
};
const CRIT=[
  ["potential","Potențial de câștig","1 mic · 5 mare"],
  ["viteza","Primii bani","1 târziu · 5 repede"],
  ["cost","Cost de pornire","1 mare · 5 mic"],
  ["timp","Timp cerut","1 mult · 5 puțin"],
  ["durabil","Are sens pe termen lung","1 deloc · 5 sigur"]
];
const DRAFT_KEY="atelier-draft", MAX_ACUM=2, CFG=window.ATELIER_CONFIG||{};
let sb=null, user=null, ideas=[], meta={}, loaded=false, view="azi", filter="acum";
let draft=null, draftId=null, advice="";
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

/* Date locale, ca text AAAA-LL-ZZ */
const iso=d=>d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
const parse=s=>{const[y,m,d]=s.split("-").map(Number);return new Date(y,m-1,d)};
const today=()=>iso(new Date());
const addDays=(s,n)=>{const d=parse(s);d.setDate(d.getDate()+n);return iso(d)};
const diff=(a,b)=>Math.round((parse(a)-parse(b))/864e5);
const nice=s=>parse(s).toLocaleDateString("ro-RO",{day:"numeric",month:"short"});
function dueLabel(s){
  if(!s)return{t:"fără termen",c:""};
  const n=diff(s,today());
  if(n<0)return{t:(n===-1?"întârziat 1 zi":"întârziat "+(-n)+" zile"),c:"late"};
  if(n===0)return{t:"azi",c:"today"};
  if(n===1)return{t:"mâine",c:"today"};
  return{t:nice(s),c:""};
}

/* Recomandarea din note */
function advise(s){
  const v=CRIT.map(c=>s&&s[c[0]]);
  if(v.some(x=>!x))return null;
  const total=v.reduce((a,b)=>a+b,0);
  const quick=(s.viteza+s.cost+s.timp)/3, long=(s.potential+s.durabil)/2;
  if(total<=9)return{stage:"renuntat",total,why:"Notele sunt mici pe aproape toate criteriile."};
  if(quick>=3.5&&s.potential>=3)return{stage:"acum",total,why:"Pornește repede, cu puțini bani și are potențial."};
  if(long>=3.7)return{stage:"explorez",total,why:"Merită pe termen lung. Începe cu pași mici de cercetare."};
  return{stage:"parcare",total,why:"Nu este prioritară acum. Pune o dată la care o recitești."};
}

function attentionItems(){
  const t=today(), out=[];
  ideas.forEach(i=>{
    if(i.stage==="nou"&&diff(t,(i.createdAt||t).slice(0,10))>=3)out.push({i,msg:"Așteaptă analiza"});
    if((i.stage==="acum"||i.stage==="explorez")&&!i.nextStep)out.push({i,msg:"Nu are următorul pas"});
    if((i.stage==="acum"||i.stage==="explorez")&&i.decisionDate&&i.decisionDate<=t)out.push({i,msg:"A venit data deciziei"});
    if(i.stage==="parcare"&&i.reviewDate&&i.reviewDate<=t)out.push({i,msg:"E timpul să o recitești"});
  });
  return out;
}
const steps=()=>ideas.filter(i=>i.nextStep&&(i.stage==="acum"||i.stage==="explorez")).sort((a,b)=>(a.dueDate||"9999").localeCompare(b.dueDate||"9999"));

/* Randare */
function ideaRow(i,opts){
  opts=opts||{};
  const d=dueLabel(i.dueDate), a=advise(i.scores);
  let m="";
  if(opts.step){
    m=`<span class="pill ${d.c}">${esc(d.t)}</span><span>${esc(i.title)}</span>`;
    return `<li><div class="row"><button class="tick" type="button" data-done="${esc(i.id)}" aria-label="Bifează: ${esc(i.nextStep)}"><span></span></button>
    <button class="open" type="button" data-open="${esc(i.id)}"><strong>${esc(i.nextStep)}</strong><span class="meta">${m}</span></button></div></li>`;
  }
  if(opts.msg)m=`<span class="pill today">${esc(opts.msg)}</span>`;
  else if(i.stage==="parcare"&&i.reviewDate)m=`<span>Recitesc pe ${esc(nice(i.reviewDate))}</span>`;
  else if(i.nextStep)m=`<span class="pill ${d.c}">${esc(d.t)}</span><span>${esc(i.nextStep)}</span>`;
  else if(i.stage==="nou")m=`<span>Adăugată pe ${esc(nice((i.createdAt||today()).slice(0,10)))}</span>`;
  if(a&&!opts.msg)m+=`<span>${a.total}/25</span>`;
  return `<li><button class="open pad" type="button" data-open="${esc(i.id)}"><strong>${esc(i.title)}</strong><span class="meta">${m}</span></button></li>`;
}
function renderAzi(){
  const att=attentionItems();
  $("attention").innerHTML=att.length?`<div class="section"><div class="section-head"><h2>Cere o decizie</h2><span class="count">${att.length}</span></div><ul class="list">${att.map(x=>ideaRow(x.i,{msg:x.msg})).join("")}</ul></div>`:"";
  const s=steps();
  let html=`<div class="section"><div class="section-head"><h2>Următorii pași</h2><span class="count">${s.length}</span></div>`;
  if(s.length)html+=`<ul class="list">${s.map(i=>ideaRow(i,{step:true})).join("")}</ul>`;
  else html+=`<p class="empty" style="margin:0">${loaded?"Aici apar pașii cu termen. Adaugă o idee, analizeaz-o și alege „Acum” sau „Explorez”.":"Se încarcă ideile tale…"}</p>`;
  $("todo").innerHTML=html+"</div>";
}
function renderIdei(){
  const n=k=>ideas.filter(i=>i.stage===k).length;
  const active=n("acum")+n("explorez");
  const withStep=steps().length;
  $("pipeline").innerHTML=[["Prind",n("nou")],["Analizez",ideas.filter(i=>advise(i.scores)&&i.stage!=="renuntat").length],["Decid",active+n("parcare")],["Acționez",withStep]]
    .map(p=>`<div><b>${p[1]}</b>${p[0]}</div>`).join("");
  $("chips").innerHTML=Object.keys(STAGES).map(k=>`<button class="chip" type="button" data-filter="${k}" aria-pressed="${filter===k}">${STAGES[k].label} ${n(k)}</button>`).join("");
  $("ideaWarn").innerHTML=(filter==="acum"&&n("acum")>MAX_ACUM)?`<p class="warn" style="margin:0">Ai ${n("acum")} idei la „Acum”. Lasă cel mult ${MAX_ACUM}, ca fiecare să primească timp real.</p>`:"";
  const rows=ideas.filter(i=>i.stage===filter).sort((a,b)=>(advise(b.scores)?.total||0)-(advise(a.scores)?.total||0));
  $("ideaList").innerHTML=rows.length?`<ul class="list">${rows.map(i=>ideaRow(i)).join("")}</ul>`:`<p class="empty" style="margin:0">${loaded?"Nicio idee aici.":"Se încarcă ideile tale…"}</p>`;
}
function renderRevizuire(){
  const t=today(), n=k=>ideas.filter(i=>i.stage===k);
  const late=steps().filter(i=>i.dueDate&&i.dueDate<t).length;
  const nostep=ideas.filter(i=>(i.stage==="acum"||i.stage==="explorez")&&!i.nextStep).length;
  const dec=ideas.filter(i=>(i.stage==="acum"||i.stage==="explorez")&&i.decisionDate&&i.decisionDate<=addDays(t,7)).length;
  const park=n("parcare").filter(i=>i.reviewDate&&i.reviewDate<=addDays(t,7)).length;
  const items=[[n("nou").length,"Analizează ideile noi","Dă note și alege unde merge fiecare."],[late,"Rezolvă pașii întârziați","Fă-i, mută termenul sau schimbă pasul."],[nostep,"Pune un pas fiecărei idei active","O idee fără pas nu avansează."],[dec,"Ia deciziile din următoarele 7 zile","Continui, schimbi direcția sau te oprești."],[park,"Recitește ideile din parcare","Le activezi sau le muți data."]];
  $("reviewList").innerHTML=items.map(x=>`<li><div class="check"><b class="${x[0]?"":"zero"}">${x[0]}</b><span>${x[1]}<small>${x[2]}</small></span></div></li>`).join("");
  $("reviewLast").textContent=meta.lastReview?("Ultima revizuire: "+nice(meta.lastReview)+" · Următoarea: "+nice(addDays(meta.lastReview,7))):"Fă-o o dată pe săptămână, la aceeași oră.";
}
function render(){
  renderAzi();renderIdei();renderRevizuire();
  const c=attentionItems().length+steps().filter(i=>i.dueDate&&i.dueDate<=today()).length;
  $("navBadge").innerHTML=c?`<span class="badge">${c}</span>`:"";
}
function show(v){
  view=v;
  ["azi","idei","revizuire"].forEach(k=>{$("view-"+k).hidden=k!==v});
  document.querySelectorAll("nav button").forEach(b=>{b.dataset.view===v?b.setAttribute("aria-current","page"):b.removeAttribute("aria-current")});
  window.scrollTo(0,0);
}
let toastT;
function toast(m){const t=$("toast");t.textContent=m;t.hidden=false;clearTimeout(toastT);toastT=setTimeout(()=>{t.hidden=true},3200)}

/* Date: Supabase, cu copie locală și coadă pentru ideile noi */
const K=n=>"atelier-"+n+"-"+(user?user.id:"anon");
const lsGet=(k,d)=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(_){return d}};
const lsSet=(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v))}catch(_){}};
const newId=()=>(window.crypto&&crypto.randomUUID)?crypto.randomUUID():"i"+Date.now().toString(36)+Math.random().toString(36).slice(2);
function withTimeout(p){return Promise.race([p,new Promise((_,r)=>setTimeout(()=>r({code:"timeout"}),12000))])}
function failMsg(){return navigator.onLine===false?"Nu ai internet. Încearcă din nou când revii online.":"Nu am putut salva. Încearcă din nou."}
const check=r=>{if(r&&r.error)throw r.error;return r};
async function putIdea(id,data){check(await withTimeout(sb.from("ideas").upsert({user_id:user.id,id,data,updated_at:new Date().toISOString()},{onConflict:"user_id,id"})))}
function setLocal(id,data){const i=ideas.findIndex(x=>x.id===id),row=Object.assign({id},data);if(i<0)ideas.push(row);else ideas[i]=row;lsSet(K("cache"),ideas);render()}
async function flushOutbox(){
  if(!user)return;
  const box=lsGet(K("outbox"),[]);
  for(const it of box){
    try{await putIdea(it.id,it.data);const rest=lsGet(K("outbox"),[]).filter(x=>x.id!==it.id);lsSet(K("outbox"),rest)}catch(_){return false}
  }
  return true;
}
let refreshing=false;
async function refresh(){
  if(!user||refreshing)return; refreshing=true;
  try{
    await flushOutbox();
    const r=check(await withTimeout(sb.from("ideas").select("id,data")));
    const box=lsGet(K("outbox"),[]);
    ideas=r.data.map(x=>Object.assign({id:x.id},x.data));
    box.forEach(b=>{if(!ideas.some(i=>i.id===b.id))ideas.push(Object.assign({id:b.id},b.data))});
    lsSet(K("cache"),ideas);
    const s=check(await withTimeout(sb.from("user_settings").select("last_review,target_date").maybeSingle()));
    meta=s.data?{lastReview:s.data.last_review||"",targetDate:s.data.target_date||""}:{};
    lsSet(K("meta"),meta);
    status("");
  }catch(_){status(navigator.onLine===false?"Lucrezi fără internet. Ideile noi se trimit când revii online.":"Nu pot actualiza ideile acum. Vezi ultima copie salvată pe acest dispozitiv.")}
  loaded=true;refreshing=false;drawTarget();render();
}
async function saveSettings(patch){
  meta=Object.assign({},meta,patch);lsSet(K("meta"),meta);
  check(await withTimeout(sb.from("user_settings").upsert({user_id:user.id,last_review:meta.lastReview||null,target_date:meta.targetDate||null,updated_at:new Date().toISOString()},{onConflict:"user_id"})));
}

/* Captură */
$("captureText").addEventListener("input",e=>{try{localStorage.setItem(DRAFT_KEY,e.target.value)}catch(_){}});
try{const d=localStorage.getItem(DRAFT_KEY);if(d)$("captureText").value=d}catch(_){}
$("captureForm").addEventListener("submit",async e=>{
  e.preventDefault();
  const el=$("captureText"), text=el.value.trim();
  if(!text){el.focus();return}
  const lines=text.split("\n"), title=lines[0].slice(0,90), note=(lines[0].length>90?lines[0].slice(90)+"\n":"")+lines.slice(1).join("\n");
  const id=newId(), data={title,note:note.trim(),stage:"nou",scores:{},nextStep:"",dueDate:"",test:"",decisionDate:"",stopRule:"",reviewDate:"",log:[],createdAt:new Date().toISOString()};
  lsSet(K("outbox"),lsGet(K("outbox"),[]).concat([{id,data}]));
  setLocal(id,data);
  el.value="";try{localStorage.removeItem(DRAFT_KEY)}catch(_){}
  const ok=await flushOutbox();
  toast(ok?"Idee salvată în Noi.":"Idee păstrată pe dispozitiv. O trimit când revii online.");
});

/* Fișa ideii */
function openSheet(id,focusStep){
  const i=ideas.find(x=>x.id===id); if(!i)return;
  draft=JSON.parse(JSON.stringify(i)); draft.scores=draft.scores||{}; draftId=id; advice="";
  drawSheet();
  $("sheet").hidden=false;$("sheetBar").hidden=false;$("sheet").scrollTop=0;document.body.style.overflow="hidden";
  if(focusStep){const f=$("fStep");if(f){f.scrollIntoView({block:"center"});f.focus()}}
}
function closeSheet(){draft=null;draftId=null;$("sheet").hidden=true;$("sheetBar").hidden=true;document.body.style.overflow=""}
function verdictHtml(){
  const a=advise(draft.scores);
  if(!a)return `<div class="verdict"><span>Dă toate cele 5 note ca să primești o recomandare.</span></div>`;
  return `<div class="verdict"><strong>${a.total}/25: recomand ${esc(STAGES[a.stage].label)}</strong><span>${esc(a.why)}</span></div>`;
}
function scoresHtml(){
  return CRIT.map(c=>`<div class="score"><div class="score-top"><span>${c[1]}</span><span>${c[2]}</span></div><div class="seg" role="group" aria-label="${c[1]}">${[1,2,3,4,5].map(n=>`<button type="button" data-crit="${c[0]}" data-val="${n}" aria-pressed="${draft.scores[c[0]]===n}">${n}</button>`).join("")}</div></div>`).join("");
}
function stageFields(){
  const s=draft.stage;
  if(s==="acum"||s==="explorez"){
    const t=today();
    return `<div class="group"><h2>Plan de acțiune</h2>
    <div class="field"><label for="fStep">Următorul pas</label><input type="text" id="fStep" data-f="nextStep" value="${esc(draft.nextStep)}" placeholder="O acțiune clară, de maximum 2 ore"></div>
    <div class="field"><label for="fDue">Termen</label><input type="date" id="fDue" data-f="dueDate" value="${esc(draft.dueDate)}">
    <div class="quick"><button type="button" data-due="${addDays(t,1)}">Mâine</button><button type="button" data-due="${addDays(t,3)}">În 3 zile</button><button type="button" data-due="${addDays(t,7)}">Într-o săptămână</button></div></div>
    <div class="field"><label for="fTest">Testul de validare</label><input type="text" id="fTest" data-f="test" value="${esc(draft.test)}" placeholder="Ce rezultat îmi arată că merită?"></div>
    <div class="field"><label for="fDec">Data deciziei</label><input type="date" id="fDec" data-f="decisionDate" value="${esc(draft.decisionDate)}"></div>
    <div class="field"><label for="fStop">Mă opresc dacă</label><input type="text" id="fStop" data-f="stopRule" value="${esc(draft.stopRule)}" placeholder="Condiția clară de oprire"></div></div>`;
  }
  if(s==="parcare")return `<div class="group"><h2>Parcare</h2><div class="field"><label for="fRev">Recitesc ideea pe</label><input type="date" id="fRev" data-f="reviewDate" value="${esc(draft.reviewDate)}"></div></div>`;
  return "";
}
function drawSheet(){
  const logs=(draft.log||[]).slice().reverse();
  $("sheet").innerHTML=`<div class="sheet-in">
  <div class="field"><label for="fTitle">Ideea</label><input class="title-input" type="text" id="fTitle" data-f="title" value="${esc(draft.title)}"></div>
  <div class="field"><label for="fNote">Note</label><textarea id="fNote" data-f="note" rows="6" placeholder="Pentru cine e, ce problemă rezolvă, cu cine o faci">${esc(draft.note)}</textarea></div>
  <div class="group"><div class="section-head"><h2>Analiză</h2><button class="btn alt" type="button" id="askAi">Analizează cu AI</button></div>
    ${advice?`<div class="verdict"><strong>Propunerea analizei</strong><span>${esc(advice)}</span></div>`:""}
    <button class="btn bare" type="button" id="copyPrompt" style="justify-self:start">Copiază cererea pentru Claude</button>
    <div class="paste" id="pasteBox" hidden><label class="flabel" for="fPaste">Lipește aici răspunsul primit de la Claude</label><textarea id="fPaste" rows="4"></textarea><button class="btn alt" type="button" id="applyPaste">Aplică analiza</button></div>
    <div class="field"><label for="fApproach">Cum aș putea aborda ideea</label><textarea id="fApproach" data-f="approach" rows="3" placeholder="Apasă Analizează cu AI sau scrie propriile idei">${esc(draft.approach)}</textarea></div>
    <div class="group" id="scoreBox">${scoresHtml()}</div><div id="verdict">${verdictHtml()}</div></div>
  <div class="group"><h2>Decizie</h2><div class="stages" role="group" aria-label="Decizie">${Object.keys(STAGES).map(k=>`<button type="button" data-stage="${k}" aria-pressed="${draft.stage===k}">${k==="nou"?"Încă nedecis":k==="renuntat"?"Renunț":STAGES[k].label}<small>${STAGES[k].hint}</small></button>`).join("")}</div></div>
  <div id="stageBox">${stageFields()}</div>
  ${logs.length?`<div class="group"><h2>Pași făcuți</h2><ol class="history">${logs.map(l=>`<li>${esc(l.text)} (${esc(nice(l.date))})</li>`).join("")}</ol></div>`:""}
  <div id="delBox"><button class="btn danger" type="button" id="delAsk">Șterge ideea</button></div>
  </div>`;
}
$("sheet").addEventListener("input",e=>{const f=e.target.dataset.f;if(f&&draft)draft[f]=e.target.value});
$("sheet").addEventListener("click",async e=>{
  const b=e.target.closest("button"); if(!b||!draft)return;
  if(b.dataset.crit){
    const k=b.dataset.crit,v=+b.dataset.val; draft.scores[k]=v;
    $("scoreBox").innerHTML=scoresHtml();$("verdict").innerHTML=verdictHtml();
  }else if(b.dataset.stage){
    draft.stage=b.dataset.stage;
    document.querySelectorAll("[data-stage]").forEach(x=>x.setAttribute("aria-pressed",x.dataset.stage===draft.stage));
    $("stageBox").innerHTML=stageFields();
  }else if(b.dataset.due){
    draft.dueDate=b.dataset.due;$("fDue").value=b.dataset.due;
  }else if(b.id==="askAi"){
    if(!draft.title.trim()){toast("Scrie un titlu pentru idee.");return}
    b.disabled=true;b.textContent="Se analizează…";
    const sheetFor=draftId;
    try{
      const res=await withTimeout(sb.functions.invoke("analyze",{body:{title:draft.title,note:draft.note,today:today()}}));
      if(res.error){const e=new Error("ai");e.status=res.error.context&&res.error.context.status;throw e}
      if(draftId===sheetFor)applyAnalysis(res.data);
    }catch(err){
      toast(err.status===429?"Ai folosit toate analizele de azi. Revino mâine.":err.status===503?"Analiza AI nu este disponibilă acum. Poți da notele singur.":"Analiza nu a mers. Încearcă din nou sau dă notele singur.");
      const x=$("askAi");if(x){x.disabled=false;x.textContent="Analizează cu AI"}
    }
  }else if(b.id==="copyPrompt"){
    const text=analysisPrompt();
    $("pasteBox").hidden=false;
    try{await navigator.clipboard.writeText(text);toast("Copiat. Trimite cererea în Claude și lipește răspunsul aici.")}
    catch(_){$("fPaste").value=text;$("fPaste").select();toast("Copiază manual textul selectat.")}
  }else if(b.id==="applyPaste"){
    const raw=$("fPaste").value;
    try{applyAnalysis(JSON.parse(raw.slice(raw.indexOf("{"),raw.lastIndexOf("}")+1)))}
    catch(_){toast("Nu am găsit analiza în text. Lipește tot răspunsul.")}
  }else if(b.id==="delAsk"){
    $("delBox").innerHTML=`<p style="margin:0 0 8px">Ștergi definitiv ideea?</p><div class="quick"><button type="button" id="delYes">Da, șterge</button><button type="button" id="delNo">Nu</button></div>`;
  }else if(b.id==="delNo"){
    $("delBox").innerHTML=`<button class="btn danger" type="button" id="delAsk">Șterge ideea</button>`;
  }else if(b.id==="delYes"){
    try{check(await withTimeout(sb.from("ideas").delete().eq("id",draftId)));
      lsSet(K("outbox"),lsGet(K("outbox"),[]).filter(x=>x.id!==draftId));
      ideas=ideas.filter(x=>x.id!==draftId);lsSet(K("cache"),ideas);render();closeSheet();toast("Idee ștearsă.")}catch(err){toast(failMsg())}
  }
});
function analysisPrompt(){
  return `Ești un consilier practic pentru cineva care vrea să testeze idei de afaceri cu buget mic. Azi este ${today()}.
Analizează ideea de mai jos și răspunde doar cu JSON valid, în română, cu cheile:
scores (obiect cu potential, viteza, cost, timp, durabil, fiecare număr întreg 1-5; 5 înseamnă mereu favorabil: potențial mare, bani repede, cost mic, timp puțin, are sens pe termen lung),
stage ("acum", "explorez", "parcare" sau "renuntat"),
reason (maximum două fraze, onest, cu principalul risc),
approach (2-3 rânduri cu idei concrete despre cum ar putea fi abordată ideea: unghi, primii clienți, ce să evite),
nextStep (o acțiune concretă de maximum 2 ore),
dueDate (AAAA-LL-ZZ),
test (rezultat măsurabil de validare),
decisionDate (AAAA-LL-ZZ),
stopRule (condiție clară de oprire),
reviewDate (AAAA-LL-ZZ, doar pentru parcare).
Ideea: ${JSON.stringify({titlu:draft.title,note:draft.note})}`;
}
function applyAnalysis(r){
  const ok=n=>Number.isInteger(n)&&n>=1&&n<=5;
  if(r&&r.scores)CRIT.forEach(c=>{if(ok(r.scores[c[0]]))draft.scores[c[0]]=r.scores[c[0]]});
  if(STAGES[r.stage]&&r.stage!=="nou")draft.stage=r.stage;
  ["nextStep","dueDate","test","decisionDate","stopRule","reviewDate"].forEach(k=>{if(typeof r[k]==="string"&&r[k]&&!draft[k])draft[k]=r[k]});
  if(typeof r.approach==="string"&&r.approach.trim())draft.approach=r.approach.trim();
  advice=typeof r.reason==="string"?r.reason:"";
  drawSheet();toast("Verifică propunerea, apoi apasă Salvează.");
}
$("sheetClose").addEventListener("click",closeSheet);
$("sheetSave").addEventListener("click",async()=>{
  if(!draft)return;
  if(!draft.title.trim()){toast("Scrie un titlu pentru idee.");return}
  const {id,...data}=draft; const b=$("sheetSave");b.disabled=true;
  try{await putIdea(draftId,data);
    filter=data.stage;setLocal(draftId,data);closeSheet();toast("Salvat.");
  }catch(err){toast(failMsg())}
  b.disabled=false;
});

/* Clicuri în liste și navigare */
$("app").addEventListener("click",async e=>{
  const b=e.target.closest("button"); if(!b)return;
  if(b.dataset.open)openSheet(b.dataset.open);
  else if(b.dataset.filter){filter=b.dataset.filter;renderIdei()}
  else if(b.dataset.done){
    const i=ideas.find(x=>x.id===b.dataset.done); if(!i)return;
    const log=(i.log||[]).concat([{text:i.nextStep,date:today()}]);
    const {id,...rest}=i, data=Object.assign({},rest,{log,nextStep:"",dueDate:""});
    try{await putIdea(id,data);setLocal(id,data);openSheet(id,true);toast("Bifat. Care e următorul pas?");
    }catch(err){toast(failMsg())}
  }
});
$("nav").addEventListener("click",e=>{const b=e.target.closest("button");if(b)show(b.dataset.view)});
$("reviewDone").addEventListener("click",async()=>{
  try{await saveSettings({lastReview:today()});render();toast("Revizuire notată.")}catch(_){toast(failMsg())}
});
$("targetDate").addEventListener("change",async e=>{
  try{await saveSettings({targetDate:e.target.value});drawTarget();toast("Termen salvat.")}catch(_){toast(failMsg())}
});

/* Export și import */
$("exportBtn").addEventListener("click",()=>{
  const blob=new Blob([JSON.stringify({version:1,exportedAt:new Date().toISOString(),settings:meta,ideas},null,2)],{type:"application/json"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="atelier-de-idei-"+today()+".json";document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),5000);
});
$("importBtn").addEventListener("click",()=>$("importFile").click());
$("importFile").addEventListener("change",async e=>{
  const f=e.target.files[0];e.target.value="";if(!f)return;
  try{
    const o=JSON.parse(await f.text());
    const list=(o.ideas||[]).filter(x=>x&&typeof x.id==="string"&&typeof x.title==="string"&&STAGES[x.stage]).slice(0,1000);
    if(!list.length){toast("Fișierul nu conține idei valide.");return}
    toast("Import "+list.length+" idei…");
    const rows=list.map(({id,...data})=>({user_id:user.id,id,data,updated_at:new Date().toISOString()}));
    for(let k=0;k<rows.length;k+=100)check(await withTimeout(sb.from("ideas").upsert(rows.slice(k,k+100),{onConflict:"user_id,id"})));
    if(o.settings&&(o.settings.lastReview||o.settings.targetDate))await saveSettings({lastReview:o.settings.lastReview||meta.lastReview,targetDate:o.settings.targetDate||meta.targetDate});
    await refresh();toast("Am importat "+list.length+" idei.");
  }catch(_){toast("Importul nu a reușit. Verifică fișierul și internetul.")}
});

/* Cont */
let authMode="login";
function setMode(m){
  authMode=m;
  $("passField").hidden=m==="forgot";
  $("authPass").required=m!=="forgot";
  $("authEmail").closest(".field").hidden=m==="newpass";
  $("authEmail").required=m!=="newpass";
  $("authPass").autocomplete=m==="login"?"current-password":"new-password";
  $("authSubmit").textContent={login:"Intră în cont",signup:"Creează contul",forgot:"Trimite linkul de resetare",newpass:"Salvează parola nouă"}[m];
  $("authMode").textContent=m==="login"?"Nu am cont":"Am deja cont";
  $("authForgot").hidden=m!=="login";
  $("authMsg").textContent="";
}
$("authMode").addEventListener("click",()=>setMode(authMode==="login"?"signup":"login"));
$("authForgot").addEventListener("click",()=>setMode("forgot"));
$("authForm").addEventListener("submit",async e=>{
  e.preventDefault();
  const email=$("authEmail").value.trim(), password=$("authPass").value, back=location.origin+location.pathname, msg=$("authMsg"), btn=$("authSubmit");
  btn.disabled=true;msg.textContent="";
  try{
    if(authMode==="login"){
      const r=await sb.auth.signInWithPassword({email,password});
      if(r.error)msg.textContent="Email sau parolă greșită, ori contul nu e confirmat încă.";
    }else if(authMode==="signup"){
      const r=await sb.auth.signUp({email,password,options:{emailRedirectTo:back}});
      if(r.error)msg.textContent="Nu am putut crea contul. Parola trebuie să aibă cel puțin 8 caractere.";
      else if(!r.data.session)msg.textContent="Verifică emailul și apasă linkul de confirmare.";
    }else if(authMode==="forgot"){
      const r=await sb.auth.resetPasswordForEmail(email,{redirectTo:back});
      msg.textContent=r.error?"Nu am putut trimite emailul. Încearcă mai târziu.":"Dacă adresa are cont, primești un link pe email.";
    }else{
      const r=await sb.auth.updateUser({password});
      if(r.error)msg.textContent="Nu am putut schimba parola. Folosește cel puțin 8 caractere.";
      else{recovering=false;setMode("login");enter(r.data.user)}
    }
  }catch(_){msg.textContent="Nu ajung la server. Verifică internetul."}
  btn.disabled=false;
});
$("logoutBtn").addEventListener("click",async()=>{await sb.auth.signOut();location.reload()});
let recovering=false;
function lock(){user=null;$("auth").hidden=false;$("app").hidden=true;$("nav").hidden=true}
function enter(u){
  if(user&&user.id===u.id)return;
  user=u;$("auth").hidden=true;$("app").hidden=false;$("nav").hidden=false;
  $("accountLine").textContent="Cont: "+(u.email||"");
  ideas=lsGet(K("cache"),[]);meta=lsGet(K("meta"),{});
  lsGet(K("outbox"),[]).forEach(b=>{if(!ideas.some(i=>i.id===b.id))ideas.push(Object.assign({id:b.id},b.data))});
  drawTarget();render();refresh();
}

/* Pornire */
function status(msg){const s=$("status");s.hidden=!msg;s.textContent=msg||""}
function drawTarget(){
  const t=meta.targetDate||"", left=t?diff(t,today()):0;
  $("targetDate").value=t;
  $("countdown").innerHTML=left>0?`<b>${left}</b> zile până la termen`:"";
}
$("today").textContent=new Date().toLocaleDateString("ro-RO",{weekday:"long",day:"numeric",month:"long"});
window.addEventListener("online",refresh);
document.addEventListener("visibilitychange",()=>{if(!document.hidden)refresh()});
if("serviceWorker" in navigator&&location.protocol==="https:")navigator.serviceWorker.register("sw.js").catch(()=>{});
(async function boot(){
  if(!CFG.supabaseUrl||!CFG.supabaseKey){$("auth").hidden=false;$("authForm").hidden=true;$("authMode").hidden=true;$("authForgot").hidden=true;$("auth").querySelector("p").textContent="Aplicația nu este configurată încă. Completează config.js după instrucțiunile din README.";return}
  sb=window.supabase.createClient(CFG.supabaseUrl,CFG.supabaseKey);
  sb.auth.onAuthStateChange((ev,session)=>{
    if(ev==="PASSWORD_RECOVERY"){recovering=true;lock();setMode("newpass");return}
    if(recovering)return;
    if(session&&session.user)setTimeout(()=>enter(session.user),0);
    else if(ev==="SIGNED_OUT"||ev==="INITIAL_SESSION")lock();
  });
})();
})();
