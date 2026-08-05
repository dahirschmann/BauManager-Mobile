(() => {
"use strict";
const config=window.BAUMANAGER_CONFIG||window.BAUAUFMASS_CONFIG||{};
const GRAPH="https://graph.microsoft.com/v1.0";
const INDEX_PATH=`${config.baseFolder||"BauAufmass/Mobile"}/project_index.json`;
const state={account:null,user:null,projects:[],selectedProject:null,selectedPosition:null,category:null,files:[],previousView:"dashboardView"};
const $=id=>document.getElementById(id);
const views=["setupView","loginView","dashboardView","positionView","positionDetailView","accountView"];
let msalInstance=null;

function showView(id){views.forEach(v=>$(v).classList.toggle("hidden",v!==id));window.scrollTo({top:0,behavior:"smooth"})}
function setStatus(el,msg,error=false,success=false){el.textContent=msg||"";el.classList.toggle("error",error);el.classList.toggle("success",success)}
function isConfigured(){return config.clientId&&!config.clientId.includes("HIER_")}
function initials(name){return String(name||"BM").split(/\s+/).slice(0,2).map(x=>x[0]||"").join("").toUpperCase()}

async function initialize(){
 if(!isConfigured()){showView("setupView");return}
 msalInstance=new msal.PublicClientApplication({auth:{clientId:config.clientId,authority:config.authority,redirectUri:config.redirectUri},cache:{cacheLocation:"localStorage",storeAuthStateInCookie:true}});
 await msalInstance.initialize();
 const redirect=await msalInstance.handleRedirectPromise();
 if(redirect?.account)msalInstance.setActiveAccount(redirect.account);
 state.account=msalInstance.getActiveAccount()||msalInstance.getAllAccounts()[0]||null;
 if(state.account){msalInstance.setActiveAccount(state.account);await enterApp()}else showView("loginView");
}
async function login(){try{setStatus($("loginStatus"),"Microsoft-Anmeldung wird geöffnet …");await msalInstance.loginRedirect({scopes:config.graphScopes})}catch(e){setStatus($("loginStatus"),`Anmeldung fehlgeschlagen: ${e.message}`,true)}}
function logout(){msalInstance.logoutRedirect({account:state.account,postLogoutRedirectUri:config.redirectUri})}
async function getToken(){const req={scopes:config.graphScopes,account:state.account};try{return(await msalInstance.acquireTokenSilent(req)).accessToken}catch{await msalInstance.acquireTokenRedirect(req);throw new Error("Anmeldung wird erneuert.")}}
async function graphFetch(url,options={}){const token=await getToken();const response=await fetch(url,{...options,headers:{Authorization:`Bearer ${token}`,...(options.headers||{})}});if(!response.ok){const body=await response.text();throw new Error(`${response.status}: ${body||response.statusText}`)}return response}
function graphPath(path){const safe=path.split("/").filter(Boolean).map(encodeURIComponent).join("/");return `${GRAPH}/me/drive/root:/${safe}`}

async function enterApp(){
 $("accountButton").classList.remove("hidden");
 try{state.user=await(await graphFetch(`${GRAPH}/me`)).json()}catch{state.user={displayName:state.account?.name||"Benutzer",mail:state.account?.username||""}}
 $("accountInitials").textContent=initials(state.user.displayName);
 $("welcomeName").textContent=`Guten Tag, ${(state.user.displayName||"").split(" ")[0]||""}`;
 $("accountName").textContent=state.user.displayName||"Benutzerkonto";
 $("accountEmail").textContent=state.user.mail||state.user.userPrincipalName||state.account?.username||"";
 await loadProjects();
}
async function loadProjects(){
 showView("dashboardView");$("projectList").innerHTML='<div class="panel">Projektdaten werden geladen …</div>';
 try{const response=await graphFetch(`${graphPath(INDEX_PATH)}:/content`);const data=await response.json();state.projects=data.projects||[];renderProjects(state.projects)}
 catch(e){$("projectList").innerHTML=`<div class="panel"><h2>Keine Projektdaten gefunden</h2><p>Starte auf dem Büro-PC <strong>Mobile_Sync_Start.bat</strong>. Die Projektübersicht wird anschließend nach OneDrive synchronisiert.</p><p class="status error">${escapeHtml(e.message)}</p></div>`}
}
function renderProjects(projects){
 const list=$("projectList");list.innerHTML="";
 if(!projects.length){list.innerHTML='<div class="panel">Keine Projekte vorhanden.</div>';return}
 const template=$("projectCardTemplate");
 projects.forEach(project=>{const node=template.content.cloneNode(true);const btn=node.querySelector(".project-card");node.querySelector(".project-number").textContent=project.project_number;node.querySelector(".project-name").textContent=project.name;node.querySelector(".project-meta").textContent=`${project.positions?.length||0} Positionen · ${project.client_name||"Auftraggeber nicht angegeben"}`;btn.addEventListener("click",()=>selectProject(project));list.appendChild(node)})
}
function selectProject(project){
 state.selectedProject=project;$("projectNumber").textContent=project.project_number;$("projectName").textContent=project.name;$("projectParties").textContent=`AG: ${project.client_name||"—"} · AN: ${project.contractor_name||"—"}`;$("positionSearch").value="";renderPositions(project.positions||[]);showView("positionView")
}
function titleKey(position){return position.title||position.parent_title||"Weitere Positionen"}
function renderPositions(positions){
 const list=$("positionList");list.innerHTML="";
 if(!positions.length){list.innerHTML='<div class="panel">Keine Positionen gefunden.</div>';return}
 const groups=new Map();
 positions.forEach(p=>{const key=titleKey(p);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(p)});
 const template=$("positionCardTemplate");
 groups.forEach((items,title)=>{
  const group=document.createElement("section");group.className="title-group";
  const titleButton=document.createElement("button");titleButton.className="title-button";titleButton.innerHTML=`<span>${escapeHtml(title)}</span><span>⌄</span>`;
  const container=document.createElement("div");container.className="title-positions";
  items.forEach(position=>{const node=template.content.cloneNode(true);const btn=node.querySelector(".position-card");node.querySelector(".position-number").textContent=position.ordinal||"—";node.querySelector(".position-text").textContent=position.short_text||"Ohne Kurztext";node.querySelector(".position-meta").textContent=`${position.unit||""} · LV-Menge ${formatNumber(position.quantity)} · ${position.measurement_sheets?.length||0} Aufmaßblätter`;btn.addEventListener("click",()=>selectPosition(position));container.appendChild(node)});
  titleButton.addEventListener("click",()=>{container.classList.toggle("hidden");titleButton.lastElementChild.textContent=container.classList.contains("hidden")?"›":"⌄"});
  group.append(titleButton,container);list.appendChild(group)
 })
}
function selectPosition(position){
 state.selectedPosition=position;$("selectedProject").textContent=`${state.selectedProject.project_number} – ${state.selectedProject.name}`;$("selectedPosition").textContent=`Position ${position.ordinal||""}`;$("selectedShortText").textContent=position.short_text||"";$("selectedUnit").textContent=`Einheit: ${position.unit||"—"}`;$("selectedQuantity").textContent=`LV-Menge: ${formatNumber(position.quantity)}`;closeUploadPanel();showView("positionDetailView")
}
function openUploadPanel(category){
 state.category=category;state.files=[];renderFileQueue();
 const names={aufmass:"Aufmaß hinzufügen",fotos:"Fotos hinzufügen",lieferscheine:"Lieferscheine hinzufügen",wiegescheine:"Wiegescheine hinzufügen",sonstige:"Sonstige Dokumente hinzufügen"};
 $("uploadTitle").textContent=names[category];$("sheetFields").classList.toggle("hidden",category!=="aufmass");if(category==="aufmass")renderSheetNumbers();$("uploadPanel").classList.remove("hidden");$("uploadPanel").scrollIntoView({behavior:"smooth",block:"start"})
}
function closeUploadPanel(){$("uploadPanel").classList.add("hidden");state.files=[];renderFileQueue();setStatus($("uploadProgress"),"")}
function currentExternalId(){const no=Number($("sheetNumber").value||0);return `${state.selectedProject.project_number}-${compactOrdinal(state.selectedPosition.ordinal)}-${String(no).padStart(3,"0")}`}
function renderSheetNumbers(){
 const select=$("sheetNumber");select.innerHTML="";const sheets=state.selectedPosition.measurement_sheets||[];
 sheets.forEach(sheet=>{const option=document.createElement("option");option.value=String(sheet.sheet_no);option.textContent=`Blatt ${String(sheet.sheet_no).padStart(3,"0")} – ${sheet.status||"gedruckt"}`;select.appendChild(option)});
 const next=Math.max(0,...sheets.map(s=>Number(s.sheet_no)||0))+1;const option=document.createElement("option");option.value=String(next);option.textContent=`Neues Blatt ${String(next).padStart(3,"0")}`;select.appendChild(option);updateIdPreview()
}
function updateIdPreview(){$("sheetIdPreview").textContent=`ID: ${currentExternalId()}`}
function addFiles(files){state.files.push(...Array.from(files||[]));renderFileQueue()}
function renderFileQueue(){
 const queue=$("fileQueue");queue.innerHTML="";
 state.files.forEach((file,index)=>{const item=document.createElement("div");item.className="file-item";item.innerHTML=`<span>${escapeHtml(file.name)} · ${humanSize(file.size)}</span><button aria-label="Entfernen">×</button>`;item.querySelector("button").onclick=()=>{state.files.splice(index,1);renderFileQueue()};queue.appendChild(item)});
 $("uploadButton").disabled=!state.files.length
}
async function uploadSelectedFiles(){
 if(!state.files.length)return;const button=$("uploadButton");button.disabled=true;const total=state.files.length;let completed=0;const failures=[];
 for(const file of [...state.files]){try{setStatus($("uploadProgress"),`Datei ${completed+1} von ${total} wird hochgeladen …`);await uploadOne(file);completed++}catch(e){failures.push(`${file.name}: ${e.message}`)}}
 if(failures.length){setStatus($("uploadProgress"),`${completed} von ${total} Datei(en) hochgeladen.\nFehler:\n${failures.join("\n")}`,true)}
 else{setStatus($("uploadProgress"),`${completed} Datei(en) erfolgreich übertragen. Sie werden im Büro automatisch übernommen.`,false,true);state.files=[];renderFileQueue()}
 button.disabled=!state.files.length
}
async function uploadOne(file){
 const uploadId=createUploadId(),base=`${config.baseFolder}/uploads/${uploadId}`;await ensureFolderPath(base);
 const project=state.selectedProject,position=state.selectedPosition,category=state.category;let blob=file,name=sanitizeFilename(file.name),sheetNo=null,externalId=null;
 if(category==="aufmass"){sheetNo=Number($("sheetNumber").value);externalId=currentExternalId();if(file.type.startsWith("image/")){blob=await imageToPdf(file);name=`${externalId}.pdf`}else if(file.type==="application/pdf"||file.name.toLowerCase().endsWith(".pdf"))name=`${externalId}.pdf`;else throw new Error("Für Aufmaße sind Bilder oder PDF-Dateien zulässig.")}
 const metadata={version:2,upload_id:uploadId,uploaded_at:new Date().toISOString(),project_id:project.id,project_number:project.project_number,project_name:project.name,position_id:position.id,position_ordinal:position.ordinal,position_short_text:position.short_text,category,original_filename:file.name,stored_filename:name,sheet_no:sheetNo,external_id:externalId,client_source:"BauManager Mobile v1.0"};
 await uploadFile(`${base}/${name}`,blob);await uploadFile(`${base}/metadata.json`,new Blob([JSON.stringify(metadata,null,2)],{type:"application/json"}))
}
async function ensureFolderPath(path){
 const parts=path.split("/").filter(Boolean);let parentId=null;
 for(const part of parts){let child;
  try{const url=parentId?`${GRAPH}/me/drive/items/${parentId}:/${encodeURIComponent(part)}`:`${GRAPH}/me/drive/root:/${encodeURIComponent(part)}`;child=await(await graphFetch(url)).json()}
  catch{const url=parentId?`${GRAPH}/me/drive/items/${parentId}/children`:`${GRAPH}/me/drive/root/children`;child=await(await graphFetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:part,folder:{},["@microsoft.graph.conflictBehavior"]:"fail"})})).json()}
  parentId=child.id
 }return parentId
}
async function uploadFile(path,blob){if(blob.size>4*1024*1024)return uploadLargeFile(path,blob);return(await(await graphFetch(`${graphPath(path)}:/content`,{method:"PUT",headers:{"Content-Type":blob.type||"application/octet-stream"},body:blob})).json())}
async function uploadLargeFile(path,blob){
 const name=path.split("/").pop(),folder=path.split("/").slice(0,-1).join("/"),parentId=await ensureFolderPath(folder);
 const session=await(await graphFetch(`${GRAPH}/me/drive/items/${parentId}:/${encodeURIComponent(name)}:/createUploadSession`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({item:{["@microsoft.graph.conflictBehavior"]:"rename",name}})})).json();
 const chunkSize=5*1024*1024;let start=0,result=null;while(start<blob.size){const end=Math.min(start+chunkSize,blob.size),chunk=blob.slice(start,end);const response=await fetch(session.uploadUrl,{method:"PUT",headers:{"Content-Length":String(chunk.size),"Content-Range":`bytes ${start}-${end-1}/${blob.size}`},body:chunk});if(!response.ok)throw new Error(`Upload fehlgeschlagen: ${response.status}`);result=await response.json();start=end}return result
}
async function imageToPdf(file){
 const bytes=await file.arrayBuffer(),pdf=await PDFLib.PDFDocument.create();let image;
 if(file.type==="image/png")image=await pdf.embedPng(bytes);else image=await pdf.embedJpg(bytes);
 const pageSize=[595.28,841.89],page=pdf.addPage(pageSize),margin=24,scale=Math.min((pageSize[0]-margin*2)/image.width,(pageSize[1]-margin*2)/image.height),width=image.width*scale,height=image.height*scale;
 page.drawImage(image,{x:(pageSize[0]-width)/2,y:(pageSize[1]-height)/2,width,height});return new Blob([await pdf.save()],{type:"application/pdf"})
}
function compactOrdinal(v){return String(v||"").replace(/[^a-zA-Z0-9]/g,"")}
function sanitizeFilename(v){return String(v||"Datei").replace(/[<>:"/\\|?*\x00-\x1F]/g,"_").replace(/\s+/g,"_").slice(0,160)}
function createUploadId(){const r=crypto.getRandomValues(new Uint32Array(2));return `${Date.now()}-${r[0].toString(16)}${r[1].toString(16)}`}
function formatNumber(v){const n=Number(v);return Number.isFinite(n)?n.toLocaleString("de-DE",{maximumFractionDigits:3}):"—"}
function humanSize(b){if(b<1024)return`${b} B`;if(b<1048576)return`${(b/1024).toFixed(1)} KB`;return`${(b/1048576).toFixed(1)} MB`}
function escapeHtml(v){const d=document.createElement("div");d.textContent=String(v||"");return d.innerHTML}

$("loginButton").onclick=login;$("logoutButton").onclick=logout;$("refreshButton").onclick=loadProjects;
$("homeButton").onclick=()=>state.account&&showView("dashboardView");
$("accountButton").onclick=()=>{state.previousView=views.find(v=>!$(v).classList.contains("hidden"))||"dashboardView";showView("accountView")};
$("backFromAccount").onclick=()=>showView(state.previousView);
$("backToProjects").onclick=()=>showView("dashboardView");$("backToPositions").onclick=()=>showView("positionView");$("closeUploadPanel").onclick=closeUploadPanel;
$("cameraInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("libraryInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("uploadButton").onclick=uploadSelectedFiles;$("sheetNumber").onchange=updateIdPreview;
document.querySelectorAll(".action-card").forEach(b=>b.onclick=()=>openUploadPanel(b.dataset.category));
$("projectSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase();renderProjects(state.projects.filter(p=>`${p.project_number||""} ${p.name||""}`.toLowerCase().includes(q)))};
$("positionSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase(),positions=state.selectedProject?.positions||[];renderPositions(positions.filter(p=>`${p.ordinal||""} ${p.short_text||""} ${titleKey(p)}`.toLowerCase().includes(q)))};
if("serviceWorker"in navigator)navigator.serviceWorker.register("sw.js").catch(()=>{});
initialize().catch(e=>{showView("loginView");setStatus($("loginStatus"),e.message,true)});
})();
