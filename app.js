/* UTSL Forms — เว็บแอป (Microsoft login · เลือกเลขที่ใบสั่งงานจากทะเบียนงาน · 5 แบบฟอร์ม · PDF → SharePoint + อีเมล) */
"use strict";
var C = window.APP_CONFIG || {};
var DEMO = !C.CLIENT_ID;
var FORMS = [
  { id:"asset",   code:"FM-SV-06-04",   title:"ใบรายการทรัพย์สินลูกค้า",            en:"Customer Asset List" },
  { id:"service", code:"FM-SV-02-01",   title:"Service Report",                        en:"Service Report" },
  { id:"record",  code:"FM-SV-03-04",   title:"ใบบันทึกข้อมูลผู้ใช้งานและข้อมูลเครื่อง", en:"User and Equipment Record" },
  { id:"warranty",code:"FM-SV-03-01",   title:"ใบรับประกันสินค้า",                     en:"Warranty Certificate" },
  { id:"qc",      code:"FM-SV-03-02.1", title:"ใบตรวจสอบคุณภาพสินค้า",                 en:"Quality Inspection" }
];
/* ช่องที่เติมให้อัตโนมัติเมื่อเลือกเลขที่ใบสั่งงาน: [ช่องเลขที่ใบสั่งงาน, ช่องชื่อลูกค้า] */
var AUTOFILL = { asset:["docNo","customer"], service:["ref","customer"], record:["jobNo","company"], warranty:["jobNo","customer"], qc:["docNo","customer"] };
var MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var SCOPES = ["User.Read","Files.ReadWrite.All","Sites.ReadWrite.All","Mail.Send"];

/* ================= helpers ================= */
var $ = function(s,r){return (r||document).querySelector(s)};
var $$ = function(s,r){return Array.prototype.slice.call((r||document).querySelectorAll(s))};
var esc = function(s){return String(s==null?"":s).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]})};
var pad = function(n){return String(n).padStart(2,"0")};
var isoOf = function(d){return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())};
var today = function(){return isoOf(new Date())};
var fmtDate = function(iso){ if(!iso) return ""; var p=iso.split("-"); return p.length===3?p[2]+"/"+p[1]+"/"+p[0]:iso };
var emailOk = function(e){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)};
var toast = function(m,ms){ms=ms||3200;var t=$("#toast");t.textContent=m;t.style.display="block";clearTimeout(toast.t);toast.t=setTimeout(function(){t.style.display="none"},ms)};
var busy = function(on,msg){ $("#ov").classList.toggle("on",!!on); if(msg) $("#ovt").textContent=msg };
var bytesToB64 = function(u){var s="";var k=0x8000;for(var i=0;i<u.length;i+=k)s+=String.fromCharCode.apply(null,u.subarray(i,i+k));return btoa(s)};
var dl = function(blob,name){var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},1500)};
var safeName = function(s){return String(s||"").replace(/[\\/:*?"<>|#%\s]+/g,"_").replace(/^_+|_+$/g,"").slice(0,60)};
function go(h){ if(location.hash===h) route(); else location.hash=h; }

/* ================= Microsoft login (MSAL) ================= */
var MSAL=null, ACCOUNT=null;
async function authInit(){
  if(DEMO) return;
  if(typeof msal==="undefined") throw new Error("โหลดระบบเข้าสู่ระบบ Microsoft ไม่สำเร็จ — ตรวจสอบอินเทอร์เน็ต");
  MSAL=new msal.PublicClientApplication({
    auth:{clientId:C.CLIENT_ID,authority:"https://login.microsoftonline.com/"+C.TENANT_ID,redirectUri:location.origin+location.pathname.replace(/index\.html$/,""),navigateToLoginRequestUrl:true},
    cache:{cacheLocation:"localStorage",storeAuthStateInCookie:true}
  });
  var r=await MSAL.handleRedirectPromise();
  ACCOUNT=(r&&r.account)||MSAL.getAllAccounts()[0]||null;
  if(ACCOUNT&&C.DOMAIN&&!String(ACCOUNT.username||"").toLowerCase().endsWith("@"+C.DOMAIN)){
    var bad=ACCOUNT.username; ACCOUNT=null; sessionStorage.setItem("loginErr","บัญชี "+bad+" ไม่ใช่บัญชี @"+C.DOMAIN);
  }
}
function login(){ return MSAL.loginRedirect({scopes:SCOPES,prompt:"select_account",domainHint:C.DOMAIN}); }
function logout(){ return MSAL.logoutRedirect({account:ACCOUNT,postLogoutRedirectUri:location.origin+location.pathname.replace(/index\.html$/,"")}); }
async function token(){
  try{ return (await MSAL.acquireTokenSilent({scopes:SCOPES,account:ACCOUNT})).accessToken; }
  catch(e){ await MSAL.acquireTokenRedirect({scopes:SCOPES,account:ACCOUNT}); throw new Error("กำลังเข้าสู่ระบบใหม่..."); }
}

/* ================= Microsoft Graph ================= */
var GRAPH="https://graph.microsoft.com/v1.0";
async function g(method,path,body,raw){
  var url=/^https:/.test(path)?path:GRAPH+path, h={Authorization:"Bearer "+await token()}, o={method:method,headers:h};
  if(raw){o.body=raw.bytes;h["Content-Type"]=raw.mime}
  else if(body!==undefined){o.body=JSON.stringify(body);h["Content-Type"]="application/json"}
  for(var i=0;i<3;i++){
    var r=await fetch(url,o);
    if((r.status===429||r.status===503||r.status===504)&&i<2){var ra=+r.headers.get("Retry-After")||(2+i*3);await new Promise(function(z){setTimeout(z,Math.min(20,ra)*1000)});continue}
    var t=await r.text(),j=null; try{j=t?JSON.parse(t):null}catch(e){}
    return {status:r.status,json:j};
  }
}
function gok(res,what){
  if(res.status>=200&&res.status<300) return res.json;
  var e=res.json&&res.json.error;
  throw new Error(what+" ไม่สำเร็จ ("+res.status+(e?": "+String(e.message||e.code).slice(0,180):"")+")");
}
var encPath=function(p){return String(p).split("/").filter(String).map(encodeURIComponent).join("/")};
var CACHE={};
async function siteId(){
  if(CACHE.site) return CACHE.site;
  CACHE.site=gok(await g("GET","/sites/"+C.SP_HOST+":"+C.SITE+"?$select=id"),"ค้นหา SharePoint site").id;
  return CACHE.site;
}
async function driveId(library){
  var k="drive:"+library; if(CACHE[k]) return CACHE[k];
  var list=(gok(await g("GET","/sites/"+await siteId()+"/drives?$select=id,name,webUrl"),"อ่าน Document Library")||{}).value||[];
  var want=(C.SITE+"/"+library).toLowerCase();
  var d=list.filter(function(x){try{return decodeURIComponent(x.webUrl||"").toLowerCase().endsWith(want)}catch(e){return false}})[0]||list.filter(function(x){return String(x.name).toLowerCase()===library.toLowerCase()})[0];
  if(!d) throw new Error("ไม่พบ Document Library \""+library+"\" ใน "+C.SITE);
  return (CACHE[k]=d.id);
}
async function itemByPath(drive,path){
  var r=await g("GET","/drives/"+drive+"/root:/"+encPath(path)+"?$select=id,name,webUrl,folder");
  if(r.status===404) return null; return gok(r,"อ่าน "+path);
}
async function folders(drive,id){
  var url="/drives/"+drive+"/items/"+id+"/children?$select=id,name,webUrl,folder&$top=999",out=[];
  while(url){ var j=gok(await g("GET",url),"อ่านรายชื่อโฟลเดอร์"); (j.value||[]).forEach(function(x){if(x.folder)out.push(x)}); url=j["@odata.nextLink"]||""; }
  return out;
}
async function mkdir(drive,parent,name){
  var r=await g("POST","/drives/"+drive+"/items/"+parent+"/children",{name:name,folder:{},"@microsoft.graph.conflictBehavior":"fail"});
  if(r.status===409){var ex=(await folders(drive,parent)).filter(function(x){return x.name.toLowerCase()===name.toLowerCase()})[0]; if(ex) return {item:ex,created:false}}
  return {item:gok(r,"สร้างโฟลเดอร์ "+name),created:true};
}
async function putFile(drive,folderId,name,bytes,mime){
  var base="/drives/"+drive+"/items/"+folderId+":/"+encodeURIComponent(name)+":";
  if(bytes.length<=4*1024*1024) return gok(await g("PUT",base+"/content?@microsoft.graph.conflictBehavior=replace",undefined,{bytes:bytes,mime:mime}),"อัปโหลด "+name);
  var s=gok(await g("POST",base+"/createUploadSession",{item:{"@microsoft.graph.conflictBehavior":"replace"}}),"เริ่มอัปโหลด "+name), CH=320*1024*16, res=null;
  for(var i=0;i<bytes.length;i+=CH){
    var part=bytes.subarray(i,Math.min(bytes.length,i+CH));
    var r=await fetch(s.uploadUrl,{method:"PUT",headers:{"Content-Range":"bytes "+i+"-"+(i+part.length-1)+"/"+bytes.length},body:part});
    if(r.status!==202&&r.status!==200&&r.status!==201) throw new Error("อัปโหลด "+name+" ไม่สำเร็จ ("+r.status+")");
    if(r.status!==202) res=await r.json();
  }
  return res;
}

/* ================= ทะเบียนงาน ================= */
/* เลขที่ใบสั่งงาน: <ประเภท>-<dd>-<MM>-<yy>-<Brand>-<รหัส>-<ชื่อลูกค้า>--<ลำดับ> */
var JOB_RE=/^([A-Za-z]+)-(\d{2})-(\d{2})-(\d{2})-([A-Za-z]*)-([^-]*)-(.+?)--(\d*)$/;
function parseJob(s){
  var t=String(s||"").trim(),m=JOB_RE.exec(t);
  if(!m||!m[7].trim()) return null;
  return {job:t,type:m[1],year:"20"+m[4],brand:m[5].toUpperCase(),code:m[6],customer:m[7]};
}
var DAY_RE=/^(\d{2})-(\d{2})-(\d{2})$/;
/* แบ่งทะเบียนเป็นช่วงวัน: แถวหัวช่วง = คอลัมน์ A มีชื่อวัน + คอลัมน์ B เป็นวันที่ · งานในช่วงคือทุกแถวจนถึงหัวช่วงถัดไป */
function parseRegister(text,firstRow){
  var out={},cur=null;
  text.forEach(function(r,i){
    var a=String(r[0]||"").trim(),b=String(r[1]||"").trim(),m=DAY_RE.exec(b);
    if(a&&m){ cur="20"+m[3]+"-"+m[2]+"-"+m[1]; out[cur]=out[cur]||[]; return; }
    if(!cur) return;
    var job=String(r[12]||"").trim(), brand=String(r[8]||"").trim();
    if(!job||!brand) return;                                     // ข้ามแถวว่าง/ลา (ไม่มีแบรนด์)
    out[cur].push({row:firstRow+i,date:b,type:String(r[7]||"").trim(),brand:brand,customer:String(r[11]||"").trim(),job:job,
      link:String(r[13]||"").trim(),owner:String(r[15]||"").trim(),detail:String(r[16]||"").trim(),time:String(r[2]||"").trim()});
  });
  return out;
}
var REG={};   // แคชต่อเดือน
async function loadJobs(iso){
  var p=iso.split("-"),y=p[0],sheet=MONTHS[+p[1]-1],key=y+"-"+p[1];
  if(DEMO){ var dm={}; dm[iso]=demoJobs(iso); return dm[iso]; }
  if(!REG[key]){
    var drive=await driveId(C.REGISTER.library), path=C.REGISTER.path.replace(/\{YYYY\}/g,y);
    var it=await itemByPath(drive,path); if(!it) throw new Error("ไม่พบไฟล์ทะเบียนงาน "+path);
    var r=await g("GET","/drives/"+drive+"/items/"+it.id+"/workbook/worksheets('"+encodeURIComponent(sheet)+"')/range(address='"+C.REGISTER.range+"')?$select=text");
    if(r.status===404) throw new Error("ไม่พบชีต "+sheet+" ในไฟล์ทะเบียนงาน "+y);
    var rg=gok(r,"อ่านทะเบียนงาน");
    var first=+(/\d+/.exec(C.REGISTER.range)||[1])[0];
    REG[key]={at:Date.now(),days:parseRegister(rg.text||[],first)};
  }
  return REG[key].days[iso]||[];
}
function demoJobs(iso){
  var p=iso.split("-"),d=p[2]+"-"+p[1]+"-"+p[0].slice(2);
  return [["O","XR","S01","Office","ธีรพงษ์",""],["VC","XR","S01","Ford","สิริมา",""],["RJ","XR","04","Ford","ชูเกียรติ","ติดตั้ง SPL QC"],
    ["RJD","XR","02","RMA Automotive","เสฎฐกรณ์","Demo MA-T12"],["O","XR","06,07","Office","วิทิต","แกะเครื่อง Ci64"],["VC","PP","S01","Kimpack","ปฐมาภรณ์",""]]
    .map(function(x,i){return {row:160+i,date:d,type:x[0],brand:x[1],customer:x[3],job:x[0]+"-"+d+"-"+x[1]+"-"+x[2]+"-"+x[3]+"--"+(i>1&&i<5?"1":""),link:"",owner:x[4],detail:x[5],time:"เข้า-บ่าย"}});
}
/* หา/สร้างโฟลเดอร์ <แบรนด์>/<ปี>/<ลูกค้า>/<เลขที่ใบสั่งงาน> */
async function jobFolder(p,b){
  var drive=await driveId(C.JOB_LIBRARY), yp=b.root+"/"+b.year.replace("{YYYY}",p.year), yr=await itemByPath(drive,yp);
  if(!yr) throw new Error("ไม่พบโฟลเดอร์ "+yp+" ใน SharePoint (ต้องสร้างโฟลเดอร์ปีก่อน)");
  var loose=function(s){return String(s||"").toLowerCase().replace(/\s+/g,"")}, cs=await folders(drive,yr.id);
  var cu=cs.filter(function(x){return x.name===p.customer})[0]||cs.filter(function(x){return loose(x.name)===loose(p.customer)})[0],nc=false;
  if(!cu){var m=await mkdir(drive,yr.id,p.customer.trim());cu=m.item;nc=m.created}
  var jf=(await folders(drive,cu.id)).filter(function(x){return x.name===p.job})[0],nj=false;
  if(!jf){var m2=await mkdir(drive,cu.id,p.job);jf=m2.item;nj=m2.created}
  return {drive:drive,id:jf.id,webUrl:jf.webUrl,path:yp+"/"+cu.name+"/"+jf.name,newCust:nc,newJob:nj};
}
async function sendMail(to,subject,html,pdf,name){
  var msg={subject:subject,body:{contentType:"HTML",content:html},toRecipients:[{emailAddress:{address:to}}]};
  if(pdf) msg.attachments=[{"@odata.type":"#microsoft.graph.fileAttachment",name:name,contentType:"application/pdf",contentBytes:bytesToB64(pdf)}];
  gok(await g("POST","/me/sendMail",{message:msg,saveToSentItems:true}),"ส่งอีเมลถึง "+to);
}

/* ================= แม่แบบ PDF ================= */
var TPL={};
async function tpl(code){
  if(TPL[code]) return TPL[code];
  var r=await fetch("templates/"+code+".pdf");
  if(!r.ok) throw new Error("ไม่พบแม่แบบ templates/"+code+".pdf");
  return (TPL[code]=new Uint8Array(await r.arrayBuffer()));
}

/* ================= รูปภาพ ================= */
async function compressImage(file,max,q){
  max=max||1280; q=q||0.72;
  var url=URL.createObjectURL(file);
  try{
    var img=await new Promise(function(res,rej){var i=new Image();i.onload=function(){res(i)};i.onerror=rej;i.src=url});
    var k=Math.min(1,max/Math.max(img.naturalWidth,img.naturalHeight));
    var c=document.createElement("canvas");c.width=Math.round(img.naturalWidth*k);c.height=Math.round(img.naturalHeight*k);
    var x=c.getContext("2d");x.fillStyle="#fff";x.fillRect(0,0,c.width,c.height);x.drawImage(img,0,0,c.width,c.height);
    var blob=await new Promise(function(r){c.toBlob(r,"image/jpeg",q)});
    return {url:URL.createObjectURL(blob),bytes:new Uint8Array(await blob.arrayBuffer()),w:c.width,h:c.height};
  } finally { URL.revokeObjectURL(url); }
}

/* ================= ลายเซ็น ================= */
function sigPad(host){
  host.innerHTML='<canvas class="sig"></canvas><div class="sigrow"><small style="color:var(--mute)">เซ็นด้วยนิ้ว/ปากกา/เมาส์ในกรอบ</small><button type="button" class="ghost sm">ล้าง</button></div>';
  var c=$("canvas",host),x=c.getContext("2d");var drawing=false,dirty=false;
  var fit=function(){var r=c.getBoundingClientRect(),d=window.devicePixelRatio||1;c.width=r.width*d;c.height=r.height*d;x.setTransform(1,0,0,1,0,0);x.scale(d,d);x.lineWidth=2.4;x.lineCap="round";x.lineJoin="round";x.strokeStyle="#0a1a6b"};
  setTimeout(fit,0);
  var pos=function(e){var r=c.getBoundingClientRect();return[e.clientX-r.left,e.clientY-r.top]};
  c.addEventListener("pointerdown",function(e){drawing=true;c.setPointerCapture(e.pointerId);var p=pos(e);x.beginPath();x.moveTo(p[0],p[1]);x.lineTo(p[0]+.1,p[1]+.1);x.stroke();dirty=true});
  c.addEventListener("pointermove",function(e){if(!drawing)return;var p=pos(e);x.lineTo(p[0],p[1]);x.stroke()});
  ["pointerup","pointercancel"].forEach(function(t){c.addEventListener(t,function(){drawing=false})});
  $("button",host).onclick=function(){x.clearRect(0,0,c.width,c.height);dirty=false};
  return { get:async function(){ if(!dirty) return null; var b=await new Promise(function(r){c.toBlob(r,"image/png")}); return new Uint8Array(await b.arrayBuffer()); } };
}

/* ================= ข้อความ -> ภาพ (ภาษาไทยซ้อนวรรณยุกต์ถูกต้อง) ================= */
var PX=4, FONTSTACK='Sarabun,"Noto Sans Thai","Leelawadee UI","Thonburi",Tahoma,sans-serif';
var _mc=document.createElement("canvas").getContext("2d");
function fontStr(size,bold){return (bold?700:400)+" "+(size*PX)+"px "+FONTSTACK}
function measure(t,size,bold){_mc.font=fontStr(size,bold);return _mc.measureText(t).width/PX}
async function loadFonts(){ try{ await Promise.all([document.fonts.load(fontStr(12,false),"กขค"),document.fonts.load(fontStr(12,true),"กขค")]); }catch(e){} }
var seg=(typeof Intl!=="undefined"&&Intl.Segmenter)?new Intl.Segmenter("th",{granularity:"word"}):null;
function toks(s){ return seg?Array.from(seg.segment(s)).map(function(x){return x.segment}):Array.from(s); }
var COMB=new RegExp("[\\u0E31\\u0E34-\\u0E3A\\u0E47-\\u0E4E]");
function wrapAll(text,size,bold,widthOf){
  var out=[]; var paras=String(text||"").replace(/\r/g,"").split("\n");
  for(var pi=0;pi<paras.length;pi++){
    var cur="", W=function(){return widthOf(out.length)}, tl=toks(paras[pi]);
    for(var ti=0;ti<tl.length;ti++){
      var tk=tl[ti];
      if(measure(cur+tk,size,bold)<=W()){cur+=tk;continue}
      if(cur.trim()){out.push(cur.replace(/\s+$/,""));cur=tk.replace(/^\s+/,"")} else cur=tk.replace(/^\s+/,"");
      while(measure(cur,size,bold)>W()){
        var i=1; while(i<cur.length&&measure(cur.slice(0,i+1),size,bold)<=W())i++;
        while(i>1&&COMB.test(cur.charAt(i)||""))i--;
        out.push(cur.slice(0,i)); cur=cur.slice(i);
      }
    }
    out.push(cur.replace(/\s+$/,""));
  }
  while(out.length>1&&out[out.length-1]==="")out.pop();
  return out;
}
async function textPng(text,size,bold,maxW,color,minSx){
  color=color||"#0a1a6b";
  var w=measure(text,size,bold),sx=1;
  if(maxW&&w>maxW){ sx=maxW/w; var mn=minSx||0.62; if(sx<mn) sx=mn }
  var wpx=Math.ceil(w*sx*PX)+6,hpx=Math.ceil(size*1.8*PX),c=document.createElement("canvas");
  c.width=wpx;c.height=hpx;var x=c.getContext("2d");x.font=fontStr(size,bold);x.fillStyle=color;x.textBaseline="alphabetic";
  x.save();x.scale(sx,1);x.fillText(text,2,size*1.3*PX);x.restore();
  var b=await new Promise(function(r){c.toBlob(r,"image/png")});
  return {bytes:new Uint8Array(await b.arrayBuffer()),w:wpx/PX,h:hpx/PX,below:size*0.5};
}

/* ================= PDF: วางข้อมูลลงบนฟอร์มต้นฉบับ ================= */
class Pdf{
  constructor(doc,W,H){this.doc=doc;this.page=doc.getPage(0);this.W=W;this.H=H;this.cache=new Map()}
  static async open(bytes){const d=await PDFLib.PDFDocument.load(bytes);const p=d.getPage(0);const s=p.getSize();return new Pdf(d,s.width,s.height)}
  async text(t,x,base,o){
    o=o||{};
    t=String(t==null?"":t).trim(); if(!t)return; const size=o.size||11;
    const key=[t,size,o.bold,o.maxW,o.minSx].join("|"); let r=this.cache.get(key);
    if(!r){const gl=await textPng(t,size,o.bold,o.maxW,undefined,o.minSx);r={g:gl,img:await this.doc.embedPng(gl.bytes)};this.cache.set(key,r)}
    const w=Math.min(r.g.w,(o.maxW||1e9)+3); let dx=x; if(o.align==="center")dx=x-w/2; else if(o.align==="right")dx=x-w;
    (o.pg||this.page).drawImage(r.img,{x:dx-0.5,y:this.H-base-r.g.below,width:r.g.w,height:r.g.h});
  }
  async fit(img,box,o){
    o=o||{};
    const e=img.jpg?await this.doc.embedJpg(img.bytes):await this.doc.embedPng(img.bytes);
    const k=Math.min(box.w/e.width,box.h/e.height),w=e.width*k,h=e.height*k,pg=o.pg||this.page;
    const px=box.x+(box.w-w)/2,pt=box.top+(box.h-h)/2;
    pg.drawImage(e,{x:px,y:this.H-pt-h,width:w,height:h});
    return {x:px,top:pt,w:w,h:h};
  }
  async lines(text,slots,size){
    const ls=wrapAll(text,size,false,function(i){return slots[Math.min(i,slots.length-1)].w});
    for(let i=0;i<Math.min(ls.length,slots.length);i++) await this.text(ls[i],slots[i].x,slots[i].base,{size:size,maxW:slots[i].w});
    return ls.slice(slots.length);
  }
  check(x,top,s){
    s=s||8;
    const y=this.H-top;const o={thickness:1.6,color:PDFLib.rgb(0.04,0.1,0.42)};
    this.page.drawLine(Object.assign({start:{x:x-s*.38,y:y-s*.1},end:{x:x-s*.1,y:y-s*.38}},o));
    this.page.drawLine(Object.assign({start:{x:x-s*.1,y:y-s*.38},end:{x:x+s*.45,y:y+s*.4}},o));
  }
  newPage(){return this.doc.addPage([this.W,this.H])}
  async textPage(title,sections){
    let pg=this.newPage(),top=50;
    await this.text(title,36,top,{size:14,bold:true,pg:pg,maxW:this.W-72});top+=28;
    for(const s of sections){
      await this.text(s.label,36,top,{size:12,bold:true,pg:pg});top+=18;
      for(const l of wrapAll(s.text,11,false,()=>this.W-72)){
        if(top>this.H-50){pg=this.newPage();top=50}
        await this.text(l,36,top,{size:11,pg:pg});top+=16;
      } top+=10;
    }
  }
  async gridPhotos(photos,fr,colsFn){
    const n=Math.min(photos.length,fr.max); if(!n)return;
    const cols=colsFn(n),rows=Math.ceil(n/cols),gp=2,cw=(fr.w-gp*(cols+1))/cols,ch=(fr.h-gp*(rows+1))/rows;
    for(let i=0;i<n;i++){
      const c=i%cols,r=Math.floor(i/cols),x=fr.x+gp+c*(cw+gp),top=fr.top+gp+r*(ch+gp);
      const q=await this.fit({bytes:photos[i].bytes,jpg:true},{x:x,top:top,w:cw,h:ch});
      if(photos[i].no){this.page.drawRectangle({x:q.x,y:this.H-q.top-9,width:9,height:9,color:PDFLib.rgb(1,1,1),opacity:.9});await this.text(String(photos[i].no),q.x+4.5,q.top+7.4,{size:7,align:"center",bold:true});}
    }
  }
  async photoPages(title,photos){
    if(!photos.length)return; const per=6,cw=(this.W-72-14)/2,ch=(this.H-130-2*14)/3;
    for(let p=0;p<photos.length;p+=per){
      const pg=this.newPage(); await this.text(title,36,50,{size:14,bold:true,pg:pg,maxW:this.W-72});
      for(let i=0;i<Math.min(per,photos.length-p);i++){
        const ph=photos[p+i],col=i%2,row=Math.floor(i/2),x=36+col*(cw+14),top=76+row*(ch+14);
        pg.drawRectangle({x:x,y:this.H-top-ch,width:cw,height:ch,borderColor:PDFLib.rgb(.75,.8,.78),borderWidth:.7});
        await this.fit({bytes:ph.bytes,jpg:true},{x:x+4,top:top+4,w:cw-8,h:ch-30},{pg:pg});
        await this.text(ph.caption||"",x+6,top+ch-9,{size:10,maxW:cw-12,pg:pg});
      }
    }
  }
  async save(){return await this.doc.save()}
}

/* ---------- FM-SV-06-04 ใบรายการทรัพย์สินลูกค้า ---------- */
async function buildAsset(d){
  const P=await Pdf.open(await tpl("FM-SV-06-04"));
  await P.text(d.docNo,364,90.4,{maxW:184});
  await P.text(fmtDate(d.date),364,110.9,{maxW:184});
  await P.text(d.customer,120,152.4,{maxW:426});
  await P.text(d.address,120,174,{maxW:426});
  await P.text(d.contact,120,195.4,{maxW:140});
  await P.text(d.dept,328,195.4,{maxW:218});
  await P.text(fmtDate(d.receiveDate),120,217,{maxW:426});
  const photos=[];
  for(let i=0;i<d.items.length&&i<13;i++){
    const it=d.items[i],top=281.8+21.6*i,base=top+15.2;
    if(!(it.name||it.sn||(i<5&&it.cond)||it.photo))continue;
    await P.text(String(i+1),45,base,{size:11,align:"center"});
    await P.text(it.name,68,base,{maxW:125});
    await P.text(it.sn,200,base,{maxW:119});
    if(i<5) await P.text(it.cond,326,base,{maxW:220});
    if(it.photo) photos.push({bytes:it.photo.bytes,no:i+1,caption:"รายการที่ "+(i+1)+": "+(it.name||"")+(it.cond?" — "+it.cond:"")});
  }
  await P.gridPhotos(photos,{x:327.5,top:390.8,w:222.5,h:171.4,max:9},function(n){return n<=1?1:n<=4?2:3});
  const over=await P.lines(d.note,[{x:70,w:478,base:582.2},{x:30,w:518,base:603.6},{x:30,w:518,base:625.2},{x:30,w:518,base:646.8},{x:30,w:518,base:668.4}],11);
  if(over.length) await P.text("(ต่อหน้าถัดไป)",548,668.4,{size:9,align:"right"});
  for(const s of [{k:"recv",x0:0},{k:"send",x0:325.5}]){
    const sg=d.sig[s.k]; if(sg) await P.fit({bytes:sg},{x:s.x0+42,top:690.5,w:150,h:22});
    const nm=d[s.k+"Name"]; if(nm) await P.text(nm,s.x0+111,737.4,{size:8,align:"center",maxW:170});
    const dt=d[s.k+"Date"]; if(dt){const p=dt.split("-");
      await P.text(p[2],s.x0+77,747.2,{align:"center"});await P.text(p[1],s.x0+111.5,747.2,{align:"center"});await P.text(p[0],s.x0+146,747.2,{align:"center"});}
  }
  if(over.length) await P.textPage("ต่อจากหน้า 1 — หมายเหตุ ("+(d.docNo||d.customer)+")",[{label:"หมายเหตุ (ต่อ)",text:over.join("\n")}]);
  await P.photoPages("ภาพประกอบ — "+(d.customer||"")+(d.docNo?"  ("+d.docNo+")":"")+"  FM-SV-06-04",photos);
  return await P.save();
}
/* ---------- FM-SV-02-01 Service Report ---------- */
async function buildService(d){
  const P=await Pdf.open(await tpl("FM-SV-02-01"));
  const B=8.8;
  await P.text(d.no,439,62.2+B,{maxW:119});
  await P.text(fmtDate(d.date),439,81.6+B,{maxW:119});
  await P.text(d.ref,439,101+B,{maxW:119});
  await P.text(d.customer,86,131.1+B,{maxW:170});
  await P.text(d.address,80,148.7+B,{maxW:176});
  await P.text(d.contact,78,166.1+B,{maxW:180});
  await P.text(d.cdept,96,183.6+B,{maxW:160});
  await P.text([d.tel,d.email].filter(Boolean).join(" / "),96,201+B,{maxW:160});
  const tops=[148.5,166.1,183.5];
  for(let i=0;i<3;i++){ const p=d.products[i]; if(!p)continue;
    await P.text(p.type,361,tops[i]+B,{maxW:67});
    await P.text(p.model,458,tops[i]+B,{maxW:47});
    await P.text(p.sn,520,tops[i]+B,{size:8.5,maxW:52,minSx:0.4});
  }
  await P.text(d.software,345,201+B,{maxW:116});
  await P.text(d.version,498,201+B,{maxW:70});
  const slots=function(first,fx,fw,tops2,x,w){x=x||45;w=w||520;return [{x:fx,w:fw,base:first+B}].concat(tops2.map(function(t){return {x:x,w:w,base:t+B}}))};
  const overC=await P.lines(d.complaint,slots(226.3,136,404,[243.3,258.7,274.3,289.8,305.4]),11);
  const wTops=[343.7,359.2,374.7,390.3,405.8,421.4,436.8,452.3,467.9,483.4,499.0];
  const wSlots=[{x:160,w:404,base:326.8+B}].concat(wTops.map(function(t){return {x:45,w:t>=388?368:520,base:t+B}}));
  const overW=await P.lines(d.work,wSlots,11);
  await P.gridPhotos(d.photos,{x:424,top:389,w:142,h:122,max:6},function(n){return n<=2?1:2});
  const boxTops=[548.5,565.6,582.6,599.7,616.7,633.7,650.8], keys=["install","warranty","charge","free","pickup","return","other"];
  keys.forEach(function(k,i){ if(d.svc[k]) P.check(44.7,boxTops[i]+6.6,7.4) });
  if(d.svc.other) await P.text(d.otherText,80,651.3+B,{maxW:482});
  const rowB=[570,589.6,609.3,629];
  for(let i=0;i<4;i++){ const p=d.parts[i]; if(!p)continue;
    await P.text(p.name,206,rowB[i],{maxW:300}); await P.text(p.qty,542,rowB[i],{align:"center",maxW:48});
  }
  const overR=await P.lines(d.remarks,[{x:90,w:473,base:673+B},{x:42,w:521,base:689.9+B},{x:42,w:521,base:705.4+B}],11);
  if(d.sig.cust) await P.fit({bytes:d.sig.cust},{x:204,top:746,w:162,h:30});
  if(d.sig.eng)  await P.fit({bytes:d.sig.eng}, {x:394,top:746,w:162,h:30});
  await P.text(d.custName,285,791.8,{size:10,align:"center",maxW:150});
  await P.text(d.engName,475.8,791.8,{size:10,align:"center",maxW:150});
  const more=[];
  if(overC.length)more.push({label:"Nature of complaint (ต่อ)",text:overC.join("\n")});
  if(overW.length)more.push({label:"Description of Work done (ต่อ)",text:overW.join("\n")});
  if(overR.length)more.push({label:"Remarks (ต่อ)",text:overR.join("\n")});
  if(more.length) await P.textPage("ต่อจากหน้า 1 — Service Report "+(d.no||d.customer),more);
  await P.photoPages("ภาพประกอบ — Service Report "+(d.no||"")+" "+(d.customer||""),d.photos.map(function(p){return {bytes:p.bytes,caption:p.caption}}));
  return await P.save();
}
/* ---------- FM-SV-03-04 ใบบันทึกข้อมูลผู้ใช้งานและข้อมูลเครื่อง ---------- */
async function buildRecord(d){
  const P=await Pdf.open(await tpl("FM-SV-03-04"));
  await P.text(fmtDate(d.deliverDate),102,114.5,{maxW:136});
  await P.text(d.product,112,139.6,{maxW:126});
  await P.text(d.model,293,139.6,{maxW:248});
  await P.text(d.software,95,164.7,{maxW:143});
  await P.text(d.serial,293,164.7,{maxW:248});
  await P.text(d.aui,76,189.7,{maxW:162});
  await P.text(d.other,293,189.7,{maxW:248});
  await P.text(d.company,102,448.2,{maxW:440});
  await P.text(d.dept,102,473.3,{maxW:185});
  await P.text(d.gmap,347,473.3,{maxW:195});
  await P.text(d.c1Name,111,498.4,{maxW:177});
  await P.text(d.c1Tel,340,498.4,{maxW:202});
  await P.text(d.c1Email,102,523.5,{maxW:186});
  await P.text(d.c1Line,340,523.5,{maxW:202});
  await P.text(d.c2Name,111,548.5,{maxW:177});
  await P.text(d.c2Tel,340,548.5,{maxW:202});
  await P.text(d.c2Email,102,573.6,{maxW:186});
  await P.text(d.c2Line,340,573.6,{maxW:202});
  const oc=await P.lines(d.claim,[{x:388,w:154,base:214.8},{x:54,w:488,base:237.7},{x:54,w:488,base:260.5},{x:54,w:488,base:283.5},{x:54,w:488,base:306.4}],11);
  const og=await P.lines(d.reg,[{x:360,w:182,base:331.6},{x:54,w:488,base:354.4},{x:54,w:488,base:377.3},{x:54,w:488,base:400.2},{x:54,w:488,base:423.1}],11);
  const ok2=await P.lines(d.contract,[{x:187,w:355,base:600.6},{x:54,w:488,base:623.5},{x:54,w:488,base:646.5}],11);
  const on=await P.lines(d.note,[{x:105,w:437,base:669.4},{x:54,w:488,base:692.3},{x:54,w:488,base:715.2}],11);
  const more=[];
  if(oc.length)more.push({label:"ข้อมูลที่ต้องใช้กรณีเคลมสินค้ากับผู้ผลิต (ต่อ)",text:oc.join("\n")});
  if(og.length)more.push({label:"การลงทะเบียนกับผู้ผลิต (ต่อ)",text:og.join("\n")});
  if(ok2.length)more.push({label:"สัญญาการให้บริการหลังการขาย (ต่อ)",text:ok2.join("\n")});
  if(on.length)more.push({label:"บันทึกอื่นๆ (ต่อ)",text:on.join("\n")});
  if(more.length) await P.textPage("ต่อจากหน้า 1 — ใบบันทึกข้อมูลผู้ใช้งานและข้อมูลเครื่อง ("+(d.company||"")+")",more);
  return await P.save();
}
/* ---------- FM-SV-03-01 ใบรับประกันสินค้า ---------- */
async function buildWarranty(d){
  const P=await Pdf.open(await tpl("FM-SV-03-01"));
  await P.lines(d.no,[{x:170,w:88,base:133.6},{x:170,w:88,base:148.6},{x:170,w:88,base:163.6}],11);
  await P.text(d.period,170,199.5,{maxW:88});
  await P.text(fmtDate(d.issued),170,243.4,{maxW:88});
  await P.text(fmtDate(d.expire),170,287.4,{maxW:88});
  await P.text(d.customer,361,133.6,{maxW:200});
  await P.text(d.product,361,177.6,{maxW:200});
  await P.text(d.brand,361,221.5,{maxW:200});
  await P.text(d.model,361,265.3,{maxW:200});
  await P.text(d.serial,361,309.2,{maxW:200});
  const cb=[[442.6,464.6,486.6,508.5],[552.4,574.4,596.4,618.3],[662.2,684.2,706.2,728.1]];
  for(let k=0;k<3;k++){ const c=d.contacts[k]||{};
    await P.text(c.name,118,cb[k][0],{maxW:138});
    await P.text(c.office,118,cb[k][1],{maxW:138});
    await P.text(c.mobile,118,cb[k][2],{maxW:138});
    await P.text(c.email,118,cb[k][3],{maxW:138});
  }
  return await P.save();
}
/* ---------- FM-SV-03-02.1 ใบตรวจสอบคุณภาพสินค้า ---------- */
async function buildQc(d){
  const P=await Pdf.open(await tpl("FM-SV-03-02.1"));
  await P.text(fmtDate(d.date),123,80.5,{maxW:175});
  await P.text(d.docNo,377,80.5,{maxW:182});
  await P.text(d.product,123,102.2,{maxW:175});
  await P.text(d.brand,377,102.2,{maxW:182});
  await P.text(d.model,123,123.9,{maxW:175});
  await P.text(d.serial,377,123.9,{maxW:182});
  if(d.qtyPl)  P.check(99.2,168.3,7.2);
  if(d.qtyPo)  P.check(99.2,190.2,7.2);
  if(d.qtyBad) P.check(99.2,211.5,7.2);
  await P.text(d.plRef,308,173.8,{maxW:252});
  await P.text(d.poRef,267,195.0,{maxW:292});
  const oM=await P.lines(d.missing,[{x:170,w:390,base:216.3},{x:66,w:493,base:236.6},{x:66,w:493,base:257.9},{x:66,w:493,base:279.1}],11);
  if(d.boot==="ok")  P.check(98.1,314.1,7.2);
  if(d.boot==="bad") P.check(98.1,334.5,7.2);
  const oB=await P.lines(d.bootNote,[{x:267,w:292,base:343.0},{x:66,w:493,base:363.1},{x:66,w:493,base:384.1},{x:66,w:493,base:405.1}],11);
  const pb=[472.5,493.8,514.9,536.2,557.4,578.7,599.9], cy=[465.4,487.5,507.9,530.0,550.4,572.4,593.7];
  for(let i=0;i<7;i++){ const p=d.params[i]; if(!p)continue;
    await P.text(p.name,123,pb[i],{maxW:243});
    if(p.res==="ผ่าน") P.check(410.05,cy[i],7.2); else if(p.res==="ไม่ผ่าน") P.check(506.15,cy[i],7.2);
  }
  const oN=await P.lines(d.note,[{x:90,w:470,base:620.9},{x:66,w:493,base:639.9},{x:66,w:493,base:659.9},{x:66,w:493,base:679.8}],11);
  if(d.final==="ผ่าน")   P.check(255.2,694.3,7.2);
  if(d.final==="ไม่ผ่าน") P.check(345.8,694.3,7.2);
  if(d.sig) await P.fit({bytes:d.sig},{x:266,top:705,w:100,h:17.5});
  if(d.inspName) await P.text(d.inspName,415,719.8,{size:10,align:"center",maxW:100});
  const more=[];
  if(oM.length)more.push({label:"1. จำนวนสินค้า — ไม่ครบ ขาด (ต่อ)",text:oM.join("\n")});
  if(oB.length)more.push({label:"2. คุณภาพสินค้า — ขั้นตอนการดำเนินการต่อไป (ต่อ)",text:oB.join("\n")});
  if(oN.length)more.push({label:"บันทึก (ต่อ)",text:oN.join("\n")});
  if(more.length) await P.textPage("ต่อจากหน้า 1 — ใบตรวจสอบคุณภาพสินค้า ("+(d.docNo||d.product||"")+")",more);
  return await P.save();
}

/* ================= หน้าจอ ================= */
var view=$("#view");
var PICK=null;   // งานที่เลือก {job, customer, brand, ...}
try{ PICK=JSON.parse(sessionStorage.getItem("pick")||"null"); }catch(e){}
function setPick(j){ PICK=j; try{ j?sessionStorage.setItem("pick",JSON.stringify(j)):sessionStorage.removeItem("pick"); }catch(e){} }
function crumb(){
  var p=Array.prototype.slice.call(arguments);
  return '<div class="crumb">'+p.map(function(x){return Array.isArray(x)?'<a onclick="go(\''+x[1]+'\')">'+esc(x[0])+'</a>':esc(x)}).join(" › ")+'</div>';
}
function chrome(){
  $("#demo").style.display=DEMO?"":"none";
  $("#me").textContent=DEMO?"โหมดทดลอง":(ACCOUNT?ACCOUNT.username:"");
  $("#btnOut").style.display=(!DEMO&&ACCOUNT)?"":"none";
}
$("#btnOut").onclick=function(){ setPick(null); logout(); };

function vLogin(){
  var err=sessionStorage.getItem("loginErr")||""; sessionStorage.removeItem("loginErr");
  view.innerHTML='<div class="login"><div class="card"><div style="font-size:42px">📝</div><h1>UTSL Forms</h1>'+
    '<p class="sub">เข้าสู่ระบบด้วยบัญชี Microsoft ของบริษัท (<b>@'+esc(C.DOMAIN)+'</b>)</p>'+
    '<button id="lgo" style="width:100%">เข้าสู่ระบบด้วย Microsoft</button>'+
    (err?'<p class="err" style="margin-top:10px">'+esc(err)+' — กรุณาเลือกบัญชีของบริษัท</p>':'')+'</div></div>';
  $("#lgo").onclick=function(){ busy(true,"กำลังไปหน้าเข้าสู่ระบบ Microsoft..."); login().catch(function(e){busy(false);toast(e.message||String(e),6000)}); };
}

/* ---------- ขั้นที่ 1: เลือกเลขที่ใบสั่งงาน ---------- */
var JOBDATE=null;
function vJobs(){
  if(!JOBDATE) JOBDATE=today();
  view.innerHTML='<h1>เลือกเลขที่ใบสั่งงาน</h1><p class="sub">ขั้นที่ 1 จาก 2 · เลือกงานจากทะเบียนงาน แล้วเลือกแบบฟอร์ม — ระบบเติมเลขที่ใบสั่งงานและชื่อลูกค้าลงฟอร์มให้</p>'+
    '<div class="card"><div class="row"><button class="ghost sm" id="dPrev" aria-label="วันก่อนหน้า">‹</button>'+
    '<input id="jd" type="date" style="max-width:190px"><button class="ghost sm" id="dNext" aria-label="วันถัดไป">›</button><button class="ghost sm" id="dToday">วันนี้</button><button class="ghost sm" id="dRe">⟳</button></div>'+
    '<input id="jq" type="search" placeholder="ค้นหา: เลขที่ใบสั่งงาน / ชื่อลูกค้า / ผู้รับผิดชอบ" style="margin-top:10px" autocomplete="off"></div>'+
    '<div id="jl" class="jobs"><p class="sub">กำลังโหลดทะเบียนงาน...</p></div>'+
    '<div class="card" style="margin-top:16px"><h2>ไม่พบเลขในรายการ?</h2><div class="row"><input id="jm" placeholder="พิมพ์เลขที่ใบสั่งงาน เช่น RJ-11-11-26-XR-04-Ford--1" style="flex:1;min-width:220px"><button id="jmGo" class="ghost">ใช้เลขนี้</button></div>'+
    '<p class="sub" style="margin:8px 0 0">หรือ <a href="javascript:void 0" id="jSkip" style="color:var(--teal)">กรอกฟอร์มโดยไม่ระบุเลขที่ใบสั่งงาน</a> (จะไม่บันทึกลง SharePoint)</p></div>';
  var list=[];
  var draw=function(){
    var q=String($("#jq").value||"").trim().toLowerCase();
    var show=list.filter(function(j){return !q||[j.job,j.customer,j.owner,j.detail].join(" ").toLowerCase().indexOf(q)>=0});
    $("#jl").innerHTML=show.map(function(j,i){
      return '<div class="job" tabindex="0" data-i="'+list.indexOf(j)+'"><div><div class="no">'+esc(j.job)+'</div><div class="cu">'+esc(j.customer||"-")+'</div></div>'+
        '<div><span class="tag">'+esc(j.brand)+'</span>'+(j.link?'<span class="tag done">มีโฟลเดอร์แล้ว</span>':'')+'</div>'+
        '<small>'+[j.type,j.time,j.owner,j.detail].filter(Boolean).map(esc).join(" · ")+' · แถว '+j.row+'</small></div>';
    }).join("")||'<div class="card"><p class="sub" style="margin:0">'+(list.length?"ไม่พบงานที่ค้นหา":"ไม่มีงานในวันที่ "+esc(fmtDate(JOBDATE))+" (หรือยังไม่ได้ลงทะเบียนงาน)")+'</p></div>';
  };
  var load=async function(force){
    $("#jd").value=JOBDATE;
    $("#jl").innerHTML='<p class="sub">กำลังโหลดทะเบียนงาน '+esc(fmtDate(JOBDATE))+'...</p>';
    if(force) REG={};
    try{ list=await loadJobs(JOBDATE); draw(); }
    catch(e){ $("#jl").innerHTML='<div class="card"><p class="err" style="margin:0">'+esc(e.message||String(e))+'</p></div>'; }
  };
  var shift=function(n){var p=JOBDATE.split("-");JOBDATE=isoOf(new Date(+p[0],+p[1]-1,+p[2]+n));load()};
  $("#jd").onchange=function(){ if(this.value){JOBDATE=this.value;load()} };
  $("#dPrev").onclick=function(){shift(-1)}; $("#dNext").onclick=function(){shift(1)};
  $("#dToday").onclick=function(){JOBDATE=today();load()}; $("#dRe").onclick=function(){load(true)};
  $("#jq").oninput=draw;
  var pick=function(el){ var j=el&&list[+el.dataset.i]; if(!j) return; setPick(j); go("#/forms"); };
  $("#jl").addEventListener("click",function(e){pick(e.target.closest(".job"))});
  $("#jl").addEventListener("keydown",function(e){if(e.key==="Enter")pick(e.target.closest(".job"))});
  $("#jmGo").onclick=function(){
    var t=$("#jm").value.trim(); if(!t) return toast("กรุณาพิมพ์เลขที่ใบสั่งงาน");
    var p=parseJob(t); if(!p&&!confirm("เลขนี้ไม่ตรงรูปแบบทะเบียนงาน ระบบจะอ่านชื่อลูกค้า/แบรนด์ไม่ได้ และจะไม่บันทึกลง SharePoint\nใช้เลขนี้ต่อหรือไม่?")) return;
    setPick({job:t,customer:p?p.customer:"",brand:p?p.brand:"",manual:true}); go("#/forms");
  };
  $("#jSkip").onclick=function(){ setPick({job:"",customer:"",none:true}); go("#/forms"); };
  load();
}
function pickedBar(){
  if(!PICK) return "";
  return '<div class="picked">'+(PICK.none?'<span>ไม่ได้ระบุเลขที่ใบสั่งงาน</span>':'<span>ใบสั่งงาน <b>'+esc(PICK.job)+'</b>'+(PICK.customer?' · '+esc(PICK.customer):'')+'</span>')+
    '<span class="sp"></span><button class="ghost sm" onclick="go(\'#/home\')">เปลี่ยนเลข</button></div>';
}
/* ---------- ขั้นที่ 2: เลือกแบบฟอร์ม ---------- */
function vForms(){
  if(!PICK) return go("#/home");
  view.innerHTML=crumb(["เลือกเลขที่ใบสั่งงาน","#/home"],"เลือกแบบฟอร์ม")+'<h1>เลือกแบบฟอร์ม</h1><p class="sub">ขั้นที่ 2 จาก 2</p>'+pickedBar()+
    '<div class="grid">'+FORMS.map(function(f){return '<div class="tile" onclick="go(\'#/form/'+f.id+'\')"><span class="ic">📄</span><b>'+esc(f.title)+'</b><small>'+f.code+' Rev.00 · '+esc(f.en)+'</small></div>'}).join("")+'</div>';
}

/* ---------- ส่วนประกอบฟอร์ม ---------- */
function fld(id,label,o){
  o=o||{};
  return '<label class="f '+(o.wide?"wide":"")+'"><span>'+label+(o.req?' <b>*</b>':'')+'</span>'+
   (o.type==="textarea"?'<textarea id="'+id+'" rows="'+(o.rows||3)+'" placeholder="'+(o.ph||"")+'"></textarea>':
    '<input id="'+id+'" type="'+(o.type||"text")+'" '+(o.im?'inputmode="'+o.im+'"':'')+' placeholder="'+(o.ph||"")+'" autocomplete="off">')+'</label>';
}
function val(id){return (($("#"+id)||{}).value||"").trim()}
function stickyBar(){return '<div class="sticky"><div class="in"><button class="ghost" id="bClear">ล้างฟอร์ม</button><button id="bSubmit">บันทึก · สร้าง PDF · ส่งอีเมล</button></div></div>'}
function formHead(F){ return crumb(["เลือกเลขที่ใบสั่งงาน","#/home"],["เลือกแบบฟอร์ม","#/forms"],F.code)+'<h1>'+esc(F.title)+'</h1><p class="sub">'+F.code+' Rev.00</p>'+pickedBar(); }
/* เติมเลขที่ใบสั่งงาน + ชื่อลูกค้าจากงานที่เลือก */
function autofill(fid){
  var m=AUTOFILL[fid]; if(!m||!PICK) return;
  [[m[0],PICK.job],[m[1],PICK.customer]].forEach(function(x){ var el=$("#"+x[0]); if(el&&x[1]){ el.value=x[1]; el.classList.add("auto"); el.addEventListener("input",function(){el.classList.remove("auto")},{once:true}); } });
}
function ctx(fid){ if(!PICK) { go("#/home"); return null; } return FORMS.find(function(x){return x.id===fid}); }

function vAsset(){
  var F=ctx("asset"); if(!F) return;
  view.innerHTML=formHead(F)+
  '<div class="card"><div class="cols">'+fld("docNo","เลขใบสั่งงาน")+fld("date","วันที่",{type:"date"})+
   fld("customer","ชื่อลูกค้า/บริษัท",{req:1,wide:1})+fld("address","ที่อยู่/เบอร์โทรศัพท์",{wide:1})+
   fld("contact","ชื่อผู้ติดต่อ")+fld("dept","แผนก")+fld("receiveDate","วันที่รับของ",{type:"date"})+fld("custEmail","อีเมลลูกค้า (ส่งสำเนา PDF)",{type:"email",im:"email",req:1})+'</div></div>'+
  '<div class="card"><h2>รายละเอียดของทรัพย์สิน</h2><div id="items"></div><div class="row" style="margin-top:10px"><button class="ghost sm" id="addRow">+ เพิ่มแถว</button><small style="color:var(--mute)">สูงสุด 13 แถว · รูปจะวางในกรอบ "สภาพ" (แถว 6–13) บนฟอร์ม และแนบรูปเต็มต่อท้าย PDF</small></div></div>'+
  '<div class="card"><h2>หมายเหตุ</h2>'+fld("note","",{type:"textarea",rows:4,wide:1})+'</div>'+
  '<div class="card cols"><div><h2>ผู้รับสินค้า / ลูกค้า</h2><div id="sgR"></div>'+fld("recvName","ชื่อ (ตัวเขียน)")+fld("recvDate","วันที่",{type:"date"})+'</div>'+
   '<div><h2>ผู้ส่งสินค้า / ลูกค้ารับสินค้าคืน</h2><div id="sgS"></div>'+fld("sendName","ชื่อ (ตัวเขียน)")+fld("sendDate","วันที่",{type:"date"})+'</div></div>'+stickyBar();
  $("#date").value=today(); autofill("asset");
  var photos={},n=0;
  var addItem=function(){ if(n>=13)return toast("ครบ 13 แถวตามฟอร์มแล้ว"); var i=n++;
    var row=document.createElement("div");row.className="item";
    row.innerHTML='<div class="n">'+(i+1)+'</div><input id="i_n'+i+'" placeholder="รายการ"><input id="i_s'+i+'" placeholder="SN">'+(i<5?'<input class="c" id="i_c'+i+'" placeholder="สภาพ">':'<span class="c" style="color:var(--mute);font-size:13px">สภาพ: ใช้เป็นพื้นที่แสดงรูป</span>')+'<div class="ph"><label class="fbtn">📷 รูป<input type="file" accept="image/*" hidden id="i_f'+i+'"></label><span id="i_p'+i+'"></span></div>';
    $("#items").appendChild(row);
    $("#i_f"+i).onchange=async function(e){var f=e.target.files[0];if(!f)return;try{photos[i]=await compressImage(f);$("#i_p"+i).innerHTML='<img src="'+photos[i].url+'">'}catch(x){toast("อ่านรูปไม่สำเร็จ")}};
    if(n===13)$("#addRow").disabled=true; };
  for(var k=0;k<5;k++)addItem(); $("#addRow").onclick=addItem;
  var sgR=sigPad($("#sgR")),sgS=sigPad($("#sgS"));
  $("#bClear").onclick=function(){if(confirm("ล้างข้อมูลที่กรอกทั้งหมด?"))vAsset()};
  $("#bSubmit").onclick=function(){submitForm(F,async function(){
    var items=[];for(var i=0;i<n;i++)items.push({name:val("i_n"+i),sn:val("i_s"+i),cond:val("i_c"+i),photo:photos[i]||null});
    var d={docNo:val("docNo"),date:val("date"),customer:val("customer"),address:val("address"),contact:val("contact"),dept:val("dept"),receiveDate:val("receiveDate"),custEmail:val("custEmail"),items:items,note:val("note"),
      recvName:val("recvName"),recvDate:val("recvDate"),sendName:val("sendName"),sendDate:val("sendDate"),sig:{recv:await sgR.get(),send:await sgS.get()}};
    if(!d.customer)throw new Error("กรุณากรอก ชื่อลูกค้า/บริษัท");
    if(!d.custEmail)throw new Error("กรุณากรอก อีเมลลูกค้า");
    if(!items.some(function(i){return i.name||i.sn||i.cond}))throw new Error("กรุณากรอกรายการทรัพย์สินอย่างน้อย 1 แถว");
    return {d:d,email:d.custEmail,jobNo:d.docNo,key:d.docNo||d.customer,build:buildAsset};
  })};
}
function vService(){
  var F=ctx("service"); if(!F) return;
  var svcOpts=[["install","Installation"],["warranty","Warranty"],["charge","Chargeable"],["free","Free of Charge"],["pickup","รับสินค้า (Pickup from customer)"],["return","คืนสินค้า (Return to customer)"],["other","Other"]];
  view.innerHTML=formHead(F)+
  '<div class="card"><div class="cols">'+fld("no","No.")+fld("date","Date",{type:"date"})+fld("ref","เอกสารอ้างอิง / ใบสั่งงาน",{wide:1})+'</div></div>'+
  '<div class="card"><h2>ลูกค้า</h2><div class="cols">'+fld("customer","Customer",{req:1,wide:1})+fld("address","Address",{wide:1})+fld("contact","Contact")+fld("cdept","Department")+fld("tel","Tel",{type:"tel",im:"tel"})+fld("email","E-mail (ส่งสำเนา PDF ให้ลูกค้า)",{type:"email",im:"email",req:1})+'</div></div>'+
  '<div class="card"><h2>Detail</h2>'+[0,1,2].map(function(i){return '<div class="cols" style="grid-template-columns:repeat(3,1fr);margin-bottom:8px">'+fld("pt"+i,"Product type "+(i+1))+fld("pm"+i,"Model")+fld("ps"+i,"SN")+'</div>'}).join("")+
   '<div class="cols">'+fld("software","Software")+fld("version","Version")+'</div></div>'+
  '<div class="card"><h2>Nature of complaint</h2>'+fld("complaint","",{type:"textarea",rows:5,wide:1})+'<h2 style="margin-top:14px">Description of Work done</h2>'+fld("work","",{type:"textarea",rows:8,wide:1})+'</div>'+
  '<div class="card"><h2>Service Detail</h2><div class="chk">'+svcOpts.map(function(o){return '<label><input type="checkbox" id="sv_'+o[0]+'">'+o[1]+'</label>'}).join("")+'</div>'+fld("otherText","Other (ระบุ)",{wide:1})+'</div>'+
  '<div class="card"><h2>Part used</h2>'+[0,1,2,3].map(function(i){return '<div class="cols" style="grid-template-columns:3fr 1fr;margin-bottom:8px">'+fld("pn"+i,(i+1)+".")+fld("pq"+i,"Quantity",{im:"numeric"})+'</div>'}).join("")+'</div>'+
  '<div class="card"><h2>Remarks</h2>'+fld("remarks","",{type:"textarea",rows:3,wide:1})+'</div>'+
  '<div class="card"><h2>รูปภาพประกอบ <small style="font-weight:400;color:var(--mute)">(6 รูปแรกวางในกรอบด้านขวาของ Description of Work done)</small></h2><label class="fbtn lg">📷 ถ่ายรูป / เลือกรูป<input type="file" accept="image/*" multiple hidden id="gf"></label><small style="color:var(--mute);margin-left:8px">รูปทั้งหมดแนบเป็นหน้าต่อท้าย PDF ด้วย</small><div class="gal" id="gal"></div></div>'+
  '<div class="card cols"><div><h2>Customer\'s Signature</h2><div id="sgC"></div>'+fld("custName","ชื่อลูกค้า (ในวงเล็บ)")+'</div><div><h2>Engineer\'s Signature</h2><div id="sgE"></div>'+fld("engName","ชื่อวิศวกร (ในวงเล็บ)")+'</div></div>'+stickyBar();
  $("#date").value=today(); autofill("service");
  var photos=[]; var drawGal=function(){$("#gal").innerHTML=photos.map(function(p,i){return '<figure><img src="'+p.url+'"><button data-i="'+i+'" type="button">×</button><input data-c="'+i+'" placeholder="คำอธิบายรูป" value="'+esc(p.caption||"")+'"></figure>'}).join("");
    $$("#gal button").forEach(function(b){b.onclick=function(){photos.splice(+b.dataset.i,1);drawGal()}});$$("#gal input").forEach(function(x){x.oninput=function(){photos[+x.dataset.c].caption=x.value}})};
  $("#gf").onchange=async function(e){var fs=Array.prototype.slice.call(e.target.files);for(var q=0;q<fs.length;q++){try{var p=await compressImage(fs[q]);p.caption="";photos.push(p)}catch(x){toast("อ่านรูปไม่สำเร็จ")}}e.target.value="";drawGal()};
  var sgC=sigPad($("#sgC")),sgE=sigPad($("#sgE"));
  $("#bClear").onclick=function(){if(confirm("ล้างข้อมูลที่กรอกทั้งหมด?"))vService()};
  $("#bSubmit").onclick=function(){submitForm(F,async function(){
    var svc={};svcOpts.forEach(function(o){svc[o[0]]=$("#sv_"+o[0]).checked});
    var d={no:val("no"),date:val("date"),ref:val("ref"),customer:val("customer"),address:val("address"),contact:val("contact"),cdept:val("cdept"),tel:val("tel"),email:val("email"),
      products:[0,1,2].map(function(i){return {type:val("pt"+i),model:val("pm"+i),sn:val("ps"+i)}}),software:val("software"),version:val("version"),complaint:val("complaint"),work:val("work"),
      svc:svc,otherText:val("otherText"),parts:[0,1,2,3].map(function(i){return {name:val("pn"+i),qty:val("pq"+i)}}),remarks:val("remarks"),
      photos:photos,sig:{cust:await sgC.get(),eng:await sgE.get()},custName:val("custName"),engName:val("engName")};
    if(!d.customer)throw new Error("กรุณากรอก Customer");
    if(!d.email)throw new Error("กรุณากรอก E-mail");
    return {d:d,email:d.email,jobNo:d.ref,key:d.ref||d.no||d.customer,build:buildService};
  })};
}
function vRecord(){
  var F=ctx("record"); if(!F) return;
  view.innerHTML=formHead(F)+
  '<div class="card"><div class="cols">'+fld("jobNo","เลขที่ใบสั่งงาน (ใช้จัดเก็บ · ไม่พิมพ์ลง PDF)",{wide:1})+fld("deliverDate","วันที่ส่งมอบ",{type:"date"})+fld("product","ชื่อสินค้า/ยี่ห้อ")+fld("model","Model")+fld("software","Software")+fld("serial","Serial")+fld("aui","AUI")+fld("other","อื่นๆ",{wide:1})+'</div></div>'+
  '<div class="card"><h2>ข้อมูลที่ต้องใช้กรณีเคลมสินค้ากับผู้ผลิต</h2><small style="color:var(--mute)">เช่น วันรับสินค้าเข้าคลัง / Invoice No. supplier</small>'+fld("claim","",{type:"textarea",rows:4,wide:1})+
   '<h2 style="margin-top:14px">การลงทะเบียนกับผู้ผลิต</h2><small style="color:var(--mute)">เช่น Registration / E-mail / Account / User / Password</small>'+fld("reg","",{type:"textarea",rows:4,wide:1})+'</div>'+
  '<div class="card"><h2>บริษัทและผู้ติดต่อ</h2><div class="cols">'+fld("company","บริษัท",{req:1,wide:1})+fld("dept","แผนก")+fld("gmap","Google Map")+
   fld("c1Name","ชื่อผู้ติดต่อ (1)")+fld("c1Tel","โทรศัพท์ (1)",{type:"tel",im:"tel"})+fld("c1Email","E-mail (1)",{type:"email",im:"email"})+fld("c1Line","Line (1)")+
   fld("c2Name","ชื่อผู้ติดต่อ (2)")+fld("c2Tel","โทรศัพท์ (2)",{type:"tel",im:"tel"})+fld("c2Email","E-mail (2)",{type:"email",im:"email"})+fld("c2Line","Line (2)")+'</div></div>'+
  '<div class="card"><h2>สัญญาการให้บริการหลังการขาย</h2>'+fld("contract","",{type:"textarea",rows:3,wide:1})+'<h2 style="margin-top:14px">บันทึกอื่นๆ</h2>'+fld("note","",{type:"textarea",rows:3,wide:1})+'</div>'+stickyBar();
  $("#deliverDate").value=today(); autofill("record");
  $("#bClear").onclick=function(){if(confirm("ล้างข้อมูลที่กรอกทั้งหมด?"))vRecord()};
  $("#bSubmit").onclick=function(){submitForm(F,async function(){
    var d={};["jobNo","deliverDate","product","model","software","serial","aui","other","claim","reg","company","dept","gmap","c1Name","c1Tel","c1Email","c1Line","c2Name","c2Tel","c2Email","c2Line","contract","note"].forEach(function(k){d[k]=val(k)});
    if(!d.company)throw new Error("กรุณากรอก บริษัท");
    ["c1Email","c2Email"].forEach(function(k){if(d[k]&&!emailOk(d[k]))throw new Error("รูปแบบ E-mail ผู้ติดต่อไม่ถูกต้อง")});
    return {d:d,email:"",jobNo:d.jobNo,key:d.jobNo||d.company,build:buildRecord};
  })};
}
function vWarranty(){
  var F=ctx("warranty"); if(!F) return;
  var contact=function(k){return '<div class="card"><h2>Technical Service — ผู้ติดต่อ '+k+'</h2><div class="cols">'+fld("cn"+k,"ชื่อ")+fld("co"+k,"Office",{type:"tel",im:"tel"})+fld("cm"+k,"Mobile",{type:"tel",im:"tel"})+fld("ce"+k,"Email",{type:"email",im:"email"})+'</div></div>'};
  view.innerHTML=formHead(F)+
  '<div class="card"><h2>การรับประกัน (Warranty)</h2><div class="cols">'+fld("jobNo","เลขที่ใบสั่งงาน (ใช้จัดเก็บ · ไม่พิมพ์ลง PDF)",{wide:1})+fld("no","เลขที่ใบรับประกัน")+fld("period","ระยะเวลารับประกัน (Warranty period)")+fld("issued","วันที่รับประกันสินค้า (Issued Date)",{type:"date"})+fld("expire","วันสิ้นสุดการรับประกัน (Expire Date)",{type:"date"})+'</div></div>'+
  '<div class="card"><h2>รายละเอียดสินค้า (Detail)</h2><div class="cols">'+fld("customer","ชื่อลูกค้า (Customer Name)",{req:1,wide:1})+fld("product","สินค้า (Product)")+fld("brand","ยี่ห้อ (Brand)")+fld("model","รุ่น (Model)")+fld("serial","หมายเลขเครื่อง (Serial No.)")+fld("custEmail","อีเมลลูกค้า (ส่งสำเนา PDF — ไม่บังคับ)",{type:"email",im:"email",wide:1})+'</div></div>'+
  contact(1)+contact(2)+contact(3)+stickyBar();
  $("#issued").value=today(); autofill("warranty");
  $("#bClear").onclick=function(){if(confirm("ล้างข้อมูลที่กรอกทั้งหมด?"))vWarranty()};
  $("#bSubmit").onclick=function(){submitForm(F,async function(){
    var d={jobNo:val("jobNo"),no:val("no"),period:val("period"),issued:val("issued"),expire:val("expire"),customer:val("customer"),product:val("product"),brand:val("brand"),model:val("model"),serial:val("serial"),custEmail:val("custEmail"),
      contacts:[1,2,3].map(function(k){return {name:val("cn"+k),office:val("co"+k),mobile:val("cm"+k),email:val("ce"+k)}})};
    if(!d.customer)throw new Error("กรุณากรอก ชื่อลูกค้า");
    d.contacts.forEach(function(c){if(c.email&&!emailOk(c.email))throw new Error("รูปแบบ Email ผู้ติดต่อไม่ถูกต้อง")});
    return {d:d,email:d.custEmail,jobNo:d.jobNo,key:d.no||d.jobNo||d.customer,build:buildWarranty};
  })};
}
function vQc(){
  var F=ctx("qc"); if(!F) return;
  var cb=function(id,label){return '<div class="chk" style="grid-template-columns:1fr"><label><input type="checkbox" id="'+id+'">'+label+'</label></div>'};
  var sel=function(id,label,opts){return '<label class="f"><span>'+label+'</span><select id="'+id+'"><option value="">— เลือก —</option>'+opts.map(function(o){return '<option>'+o+'</option>'}).join("")+'</select></label>'};
  var rows="";for(var i=0;i<7;i++)rows+='<div class="qrow"><div class="n">'+(i+1)+'</div><input id="qp'+i+'" maxlength="90" placeholder="พารามิเตอร์ / ฟังก์ชั่น"><select id="qr'+i+'"><option value="">— ผลทดสอบ —</option><option>ผ่าน</option><option>ไม่ผ่าน</option></select></div>';
  view.innerHTML=formHead(F)+
  '<div class="card"><div class="cols">'+fld("date","วันที่",{type:"date"})+fld("docNo","เลขที่ใบสั่งงาน")+fld("product","Product",{req:1})+fld("brand","Brand")+fld("model","Model")+fld("serial","Serial")+
   fld("customer","ชื่อลูกค้า/บริษัท (ไม่พิมพ์ลง PDF)",{wide:1})+'</div></div>'+
  '<div class="card"><h2>1. จำนวนสินค้า</h2>'+
   '<div class="blk">'+cb("q1","ครบตาม Packing List (ตาม Installation Manual /WI..)")+fld("plRef","ระบุ Installation Manual / WI ..",{wide:1})+'</div>'+
   '<div class="blk">'+cb("q2","ครบตาม PO และ Quotation No.")+fld("poRef","Quotation No.",{wide:1})+'</div>'+
   '<div class="blk">'+cb("q3","ไม่ครบ ขาด")+fld("missing","รายการที่ขาด",{type:"textarea",rows:3,wide:1})+'</div></div>'+
  '<div class="card"><h2>2. คุณภาพสินค้า</h2>'+cb("b1","เปิดใช้งานได้")+cb("b2","เปิดใช้งานไม่ได้ ขั้นตอนการดำเนินการต่อไป")+fld("bootNote","ขั้นตอนการดำเนินการต่อไป (กรณีเปิดใช้งานไม่ได้)",{type:"textarea",rows:3,wide:1})+'</div>'+
  '<div class="card"><h2>3. ทดสอบพารามิเตอร์-ฟังก์ชั่นการใช้งาน ตาม Operation Manual จากผู้ผลิต</h2>'+rows+'</div>'+
  '<div class="card"><h2>บันทึก</h2>'+fld("note","",{type:"textarea",rows:4,wide:1})+'</div>'+
  '<div class="card"><h2>สรุปผล</h2><div class="cols">'+sel("final","สามารถใช้งานได้ปกติ พร้อมส่ง",["ผ่าน","ไม่ผ่าน"])+fld("inspName","ชื่อผู้ตรวจสอบ (ตัวเขียน)")+'</div>'+
   '<h2 style="margin-top:14px">ลายเซ็นผู้ตรวจสอบ</h2><div id="sgQ"></div></div>'+stickyBar();
  $("#date").value=today(); autofill("qc");
  $("#b1").onchange=function(){if(this.checked)$("#b2").checked=false}; $("#b2").onchange=function(){if(this.checked)$("#b1").checked=false};
  var sg=sigPad($("#sgQ"));
  $("#bClear").onclick=function(){if(confirm("ล้างข้อมูลที่กรอกทั้งหมด?"))vQc()};
  $("#bSubmit").onclick=function(){submitForm(F,async function(){
    var params=[];for(var i=0;i<7;i++)params.push({name:val("qp"+i),res:val("qr"+i)});
    var d={date:val("date"),docNo:val("docNo"),customer:val("customer"),product:val("product"),brand:val("brand"),model:val("model"),serial:val("serial"),
      qtyPl:$("#q1").checked,plRef:val("plRef"),qtyPo:$("#q2").checked,poRef:val("poRef"),qtyBad:$("#q3").checked,missing:val("missing"),
      boot:$("#b1").checked?"ok":$("#b2").checked?"bad":"",bootNote:val("bootNote"),params:params,note:val("note"),final:val("final"),inspName:val("inspName"),sig:await sg.get()};
    if(!d.product)throw new Error("กรุณากรอก Product");
    return {d:d,email:"",jobNo:d.docNo,key:d.docNo||d.product,build:buildQc};
  })};
}

/* ---------- บันทึก: PDF → SharePoint → อีเมล ---------- */
async function submitForm(F,collect){
  try{
    busy(true,"กำลังตรวจสอบข้อมูล...");
    var c=await collect();
    if(c.email&&!emailOk(c.email))throw new Error("รูปแบบอีเมลลูกค้าไม่ถูกต้อง");
    busy(true,"กำลังสร้าง PDF จากฟอร์มต้นฉบับ...");
    await loadFonts();
    var pdf=await c.build(c.d);
    var now=new Date(),stamp=now.getFullYear()+pad(now.getMonth()+1)+pad(now.getDate())+"-"+pad(now.getHours())+pad(now.getMinutes())+pad(now.getSeconds());
    var name=F.code+"_"+(safeName(c.key)||"doc")+"_"+stamp+".pdf";
    busy(false);
    vDone({F:F,pdf:pdf,name:name,c:c});
  }catch(e){busy(false);console.error(e);toast(e.message||"เกิดข้อผิดพลาด",6000)}
}
function vDone(o){
  var F=o.F,c=o.c,st={sp:"run",spMsg:"",spUrl:"",mail:"wait",mailMsg:""};
  view.innerHTML='<div class="card" style="text-align:center;padding:28px"><div style="font-size:46px">✅</div><h1>สร้าง PDF เรียบร้อย</h1>'+
    '<p class="sub" style="word-break:break-all">'+esc(o.name)+'</p><div id="dS" style="text-align:left;max-width:640px;margin:0 auto"></div>'+
    '<div class="row" style="justify-content:center;margin-top:16px"><button id="dP">⬇ ดาวน์โหลด PDF</button></div>'+
    '<div class="row" style="justify-content:center;margin-top:10px"><button class="ghost" id="bSame">กรอกฟอร์มอื่น (ใบสั่งงานเดิม)</button><button class="ghost" onclick="go(\'#/home\')">เลือกเลขที่ใบสั่งงานใหม่</button></div></div>';
  $("#dP").onclick=function(){dl(new Blob([o.pdf],{type:"application/pdf"}),o.name)};
  $("#bSame").onclick=function(){go("#/forms")};
  var paint=function(){
    var el=$("#dS"); if(!el) return;
    var sp=st.sp==="run"?"⏳ กำลังบันทึก PDF ลง SharePoint...":st.sp==="ok"?'✔ บันทึกลง SharePoint: <b>'+esc(st.spMsg)+'</b><br><a href="'+esc(st.spUrl)+'" target="_blank" rel="noopener" style="color:var(--teal)">เปิดโฟลเดอร์ใบสั่งงาน</a>':
      st.sp==="skip"?'<span style="color:var(--mute)">– '+esc(st.spMsg)+'</span>':'<span class="err">⚠ บันทึกลง SharePoint ไม่สำเร็จ: '+esc(st.spMsg)+'</span> <button class="ghost sm" id="rS">ลองใหม่</button>';
    var ml=st.mail==="wait"?"⏳ รอส่งอีเมล...":st.mail==="run"?"⏳ กำลังส่งอีเมล...":st.mail==="ok"?"✔ "+st.mailMsg:st.mail==="skip"?'<span style="color:var(--mute)">– '+esc(st.mailMsg)+'</span>':'<span class="err">⚠ ส่งอีเมลไม่สำเร็จ: '+esc(st.mailMsg)+'</span> <button class="ghost sm" id="rM">ลองใหม่</button>';
    el.innerHTML='<p>'+sp+'</p><p>'+ml+'</p>';
    var a=$("#rS"),b=$("#rM"); if(a)a.onclick=doSp; if(b)b.onclick=doMail;
    window.onbeforeunload=(st.sp==="run"||st.mail==="run"||st.mail==="wait")?function(){return "ยังบันทึก/ส่งอีเมลไม่เสร็จ"}:null;
  };
  var doSp=async function(){
    st.sp="run";paint();
    try{
      if(DEMO){st.sp="skip";st.spMsg="โหมดทดลอง — ไม่ได้บันทึกลง SharePoint";return}
      var job=String(c.jobNo||"").trim();
      if(!job){st.sp="skip";st.spMsg="ไม่ได้ระบุเลขที่ใบสั่งงาน จึงไม่ได้บันทึกลง SharePoint";return}
      var p=parseJob(job);
      if(!p){st.sp="skip";st.spMsg="เลขที่ใบสั่งงาน \""+job+"\" ไม่ตรงรูปแบบทะเบียนงาน — ไม่ได้บันทึกลง SharePoint";return}
      var b=(C.BRANDS||{})[p.brand];
      if(!b){st.sp="skip";st.spMsg="แบรนด์ "+(p.brand||"(ว่าง)")+" ยังไม่ได้ตั้งค่าโฟลเดอร์ใน SharePoint — ดาวน์โหลด PDF ไว้ก่อน";return}
      var fd=await jobFolder(p,b), up=await putFile(fd.drive,fd.id,o.name,o.pdf,"application/pdf");
      st.sp="ok";st.spUrl=fd.webUrl;st.spMsg=fd.path+(fd.newCust?" (สร้างโฟลเดอร์ลูกค้าใหม่)":"")+(fd.newJob?" (สร้างโฟลเดอร์ใบสั่งงานใหม่)":"");
      o.fileUrl=up&&up.webUrl;
    }catch(e){st.sp="err";st.spMsg=e.message||String(e)}
    finally{paint()}
  };
  var doMail=async function(){
    st.mail="run";paint();
    try{
      if(DEMO){st.mail="skip";st.mailMsg="โหมดทดลอง — ไม่ได้ส่งอีเมล";return}
      var big=o.pdf.length>3*1024*1024, link=st.sp==="ok"?(o.fileUrl||st.spUrl):"";
      if(big&&!link) throw new Error("ไฟล์ PDF ใหญ่เกิน 3 MB แนบอีเมลไม่ได้ และยังไม่ได้บันทึกลง SharePoint — ดาวน์โหลด PDF แล้วส่งเอง");
      var att=big?null:o.pdf, sent=[], subj=F.code+" — "+(c.jobNo||c.key||"")+" — "+o.name;
      var me=ACCOUNT.username, extra=link?'<p>ไฟล์ใน SharePoint: <a href="'+esc(link)+'">'+esc(link)+'</a></p>':"";
      await sendMail(me,subj,'<p>สำเนาเอกสาร <b>'+esc(F.code)+' '+esc(F.title)+'</b>'+(c.jobNo?' ใบสั่งงาน <b>'+esc(c.jobNo)+'</b>':'')+(att?' แนบมาพร้อมอีเมลนี้':'')+'</p>'+extra,att,o.name); sent.push(me);
      var cu=String(c.email||"").trim();
      if(cu&&cu.toLowerCase()!==me.toLowerCase()){
        if(big) throw new Error("ส่งให้ตัวเองแล้ว แต่ PDF ใหญ่เกิน 3 MB จึงส่งให้ลูกค้าไม่ได้ — ดาวน์โหลดแล้วส่งเอง");
        await sendMail(cu,F.code+" — "+o.name,'<p>เรียนลูกค้า</p><p>บริษัท ยูเนียน ทีเอสแอล จำกัด / ยูเนียน เคมเทค จำกัด ขอส่งสำเนาเอกสาร <b>'+esc(F.code)+'</b> ตามไฟล์แนบ</p>',att,o.name); sent.push(cu);
      }
      st.mail="ok";st.mailMsg="ส่ง PDF ทางอีเมลถึง "+sent.join(" และ ")+" แล้ว";
    }catch(e){st.mail="err";st.mailMsg=e.message||String(e)}
    finally{paint()}
  };
  paint(); scrollTo(0,0);
  doSp().then(doMail);   // อัปโหลดก่อน เพื่อแนบลิงก์ในอีเมลได้กรณีไฟล์ใหญ่
}

/* ---------- router ---------- */
function route(){
  chrome(); window.onbeforeunload=null;
  var h=(location.hash||"#/").replace(/^#\/?/,"").split("/");
  if(!DEMO&&!ACCOUNT) return vLogin();
  scrollTo(0,0);
  switch(h[0]){
    case "forms": return vForms();
    case "form": return ({asset:vAsset,service:vService,record:vRecord,warranty:vWarranty,qc:vQc}[h[1]]||vForms)();
    default: return vJobs();
  }
}
addEventListener("hashchange",route);
(async function(){
  try{ busy(true,"กำลังเริ่มระบบ..."); await authInit(); }
  catch(e){ toast(e.message||String(e),8000); }
  finally{ busy(false); }
  route();
})();
