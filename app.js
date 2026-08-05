(() => {
"use strict";
const config=window.BAUMANAGER_CONFIG||window.BAUAUFMASS_CONFIG||{};
const GRAPH="https://graph.microsoft.com/v1.0";
const INDEX_FILE_NAME="project_index.json";
const CONFIGURED_BASE_FOLDER=config.baseFolder||"";
const PRIMARY_CLOUD_PATH=["BauManager","BauManagerCloud"];
const CLOUD_DIAGNOSTICS=[];
const state={account:null,user:null,projects:[],selectedProject:null,selectedPosition:null,category:null,files:[],previousView:"dashboardView",cloudBaseFolder:"",projectIndexItemId:""};
const $=id=>document.getElementById(id);
const views=["setupView","loginView","dashboardView","positionView","positionDetailView","digitalMeasurementView","accountView"];
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

function addCloudDiagnostic(step,ok,detail=""){
 CLOUD_DIAGNOSTICS.push({
  step,
  ok:Boolean(ok),
  detail:String(detail||"")
 });
}
function resetCloudDiagnostics(){
 CLOUD_DIAGNOSTICS.length=0;
}
function cloudDiagnosticText(){
 return CLOUD_DIAGNOSTICS.map(entry=>
  `${entry.ok?"✓":"✗"} ${entry.step}${entry.detail?`\n   ${entry.detail}`:""}`
 ).join("\n");
}
async function graphJsonDiagnostic(step,url,options={}){
 try{
  const response=await graphFetch(url,options);
  const data=await response.json();
  addCloudDiagnostic(step,true,url);
  return data;
 }catch(error){
  addCloudDiagnostic(step,false,`${url}\n${error.message}`);
  throw error;
 }
}
async function fetchIndexByAbsolutePath(){
 const relative="BauManager/BauManagerCloud/project_index.json";
 const url=`${graphPath(relative)}:/content`;
 try{
  const response=await graphFetch(url);
  const data=await response.json();
  state.cloudBaseFolder="BauManager/BauManagerCloud";
  addCloudDiagnostic(
   "Direkter Dateizugriff",
   true,
   relative
  );
  return data;
 }catch(error){
  addCloudDiagnostic(
   "Direkter Dateizugriff",
   false,
   `${relative}\n${error.message}`
  );
  return null;
 }
}

function accountCacheKey(){
 const accountId=state.account?.homeAccountId||state.account?.username||"default";
 return `baumanager-cloud-location:${accountId}`;
}
function normaliseDriveFolder(path){
 let value=String(path||"");
 value=value.replace(/^\/drive\/root:\/?/i,"");
 value=value.replace(/^\/drives\/[^/]+\/root:\/?/i,"");
 value=value.replace(/^\/+/,"").replace(/\/+$/,"");
 return value;
}
function folderFromParentReference(parentReference){
 return normaliseDriveFolder(parentReference?.path||"");
}
function saveCloudLocation(item){
 const folder=folderFromParentReference(item.parentReference);
 state.cloudBaseFolder=folder;
 state.projectIndexItemId=item.id||"";
 localStorage.setItem(accountCacheKey(),JSON.stringify({
  itemId:state.projectIndexItemId,
  folder:state.cloudBaseFolder,
  savedAt:new Date().toISOString()
 }));
}
function loadCachedCloudLocation(){
 try{
  const value=JSON.parse(localStorage.getItem(accountCacheKey())||"{}");
  state.projectIndexItemId=value.itemId||"";
  state.cloudBaseFolder=value.folder||"";
 }catch{
  state.projectIndexItemId="";
  state.cloudBaseFolder="";
 }
}
async function fetchIndexByItemId(itemId){
 if(!itemId)return null;
 try{
  const metadata=await(await graphFetch(`${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,parentReference,lastModifiedDateTime`)).json();
  if(String(metadata.name||"").toLowerCase()!==INDEX_FILE_NAME)return null;
  const data=await(await graphFetch(`${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/content`)).json();
  saveCloudLocation(metadata);
  return data;
 }catch{
  return null;
 }
}
async function listFolderChildren(parentId){
 let url=parentId
  ? `${GRAPH}/me/drive/items/${encodeURIComponent(parentId)}/children?$select=id,name,file,folder,parentReference,lastModifiedDateTime,size&$top=200`
  : `${GRAPH}/me/drive/root/children?$select=id,name,file,folder,parentReference,lastModifiedDateTime,size&$top=200`;
 const items=[];
 while(url){
  const response=await graphFetch(url);
  const page=await response.json();
  items.push(...(page.value||[]));
  url=page["@odata.nextLink"]||"";
 }
 addCloudDiagnostic(
  parentId?"Ordnerinhalt gelesen":"OneDrive-Stammordner gelesen",
  true,
  items.map(item=>item.name).join(", ")||"(leer)"
 );
 return items;
}

async function findChildByName(parentId,name,wantFolder=null){
 const items=await listFolderChildren(parentId);
 const target=String(name||"").toLowerCase();
 const found=items.find(item=>{
  if(String(item.name||"").toLowerCase()!==target)return false;
  if(wantFolder===true)return Boolean(item.folder);
  if(wantFolder===false)return Boolean(item.file);
  return true;
 })||null;
 addCloudDiagnostic(
  `Element „${name}“`,
  Boolean(found),
  found?`ID: ${found.id}`:`Vorhanden: ${items.map(item=>item.name).join(", ")||"(keine)"}`
 );
 return found;
}

async function fetchIndexFromFolderParts(parts){
 let parentId=null;
 for(const part of parts){
  const folder=await findChildByName(parentId,part,true);
  if(!folder)return null;
  parentId=folder.id;
 }

 const indexItem=await findChildByName(parentId,INDEX_FILE_NAME,false);
 if(!indexItem)return null;

 const data=await(
  await graphFetch(
   `${GRAPH}/me/drive/items/${encodeURIComponent(indexItem.id)}/content`
  )
 ).json();

 saveCloudLocation(indexItem);
 return data;
}

async function fetchIndexByConfiguredPath(){
 if(!CONFIGURED_BASE_FOLDER)return null;
 const parts=normaliseDriveFolder(CONFIGURED_BASE_FOLDER)
  .split("/")
  .filter(Boolean);
 if(!parts.length)return null;

 try{
  return await fetchIndexFromFolderParts(parts);
 }catch{
  return null;
 }
}

async function fetchIndexByPrimaryPath(){
 try{
  return await fetchIndexFromFolderParts(PRIMARY_CLOUD_PATH);
 }catch{
  return null;
 }
}

async function searchProjectIndexItems(){
 let url=`${GRAPH}/me/drive/root/search(q='${encodeURIComponent(INDEX_FILE_NAME)}')?$select=id,name,file,folder,parentReference,lastModifiedDateTime,size&$top=200`;
 const matches=[];
 while(url){
  const response=await graphFetch(url);
  const page=await response.json();
  for(const item of page.value||[]){
   if(String(item.name||"").toLowerCase()===INDEX_FILE_NAME && item.file){
    matches.push(item);
   }
  }
  url=page["@odata.nextLink"]||"";
 }
 return matches;
}

function scoreIndexItem(item){
 const folder=folderFromParentReference(item.parentReference).toLowerCase();
 let score=0;
 if(folder.endsWith("/baumanager/bauermanagercloud"))score+=500;
 if(folder.endsWith("/baumanager/baumanagercloud"))score+=1000;
 if(folder.includes("baumanagercloud"))score+=400;
 if(CONFIGURED_BASE_FOLDER && folder===normaliseDriveFolder(CONFIGURED_BASE_FOLDER).toLowerCase())score+=2000;
 score+=new Date(item.lastModifiedDateTime||0).getTime()/1e13;
 return score;
}
async function findProjectIndex(){
 resetCloudDiagnostics();

 try{
  const drive=await graphJsonDiagnostic(
   "OneDrive erreichbar",
   `${GRAPH}/me/drive?$select=id,driveType,owner,quota`
  );
  addCloudDiagnostic(
   "OneDrive-Typ",
   true,
   `${drive.driveType||"unbekannt"} · ${drive.owner?.user?.displayName||""}`
  );
 }catch(error){
  throw new Error(
   "Das OneDrive des angemeldeten Kontos ist nicht erreichbar.\n\n"+
   cloudDiagnosticText()
  );
 }

 // Most reliable method: direct path content endpoint.
 const absolute=await fetchIndexByAbsolutePath();
 if(absolute)return absolute;

 loadCachedCloudLocation();
 const cached=await fetchIndexByItemId(state.projectIndexItemId);
 if(cached){
  addCloudDiagnostic("Gespeicherte Datei-ID",true,state.projectIndexItemId);
  return cached;
 }
 addCloudDiagnostic("Gespeicherte Datei-ID",false,state.projectIndexItemId||"(keine)");

 const primary=await fetchIndexByPrimaryPath();
 if(primary){
  addCloudDiagnostic("Direkte Ordnernavigation",true,PRIMARY_CLOUD_PATH.join("/"));
  return primary;
 }
 addCloudDiagnostic("Direkte Ordnernavigation",false,PRIMARY_CLOUD_PATH.join("/"));

 const configured=await fetchIndexByConfiguredPath();
 if(configured){
  addCloudDiagnostic("Konfigurierter Pfad",true,CONFIGURED_BASE_FOLDER);
  return configured;
 }
 addCloudDiagnostic("Konfigurierter Pfad",false,CONFIGURED_BASE_FOLDER||"(leer)");

 try{
  const matches=await searchProjectIndexItems();
  addCloudDiagnostic(
   "OneDrive-Suche",
   matches.length>0,
   `${matches.length} Treffer`
  );
  if(matches.length){
   matches.sort((a,b)=>scoreIndexItem(b)-scoreIndexItem(a));
   const selected=matches[0];
   const data=await(
    await graphFetch(
     `${GRAPH}/me/drive/items/${encodeURIComponent(selected.id)}/content`
    )
   ).json();
   saveCloudLocation(selected);
   return data;
  }
 }catch(error){
  addCloudDiagnostic("OneDrive-Suche",false,error.message);
 }

 throw new Error(
  "project_index.json konnte nicht geladen werden.\n\n"+
  cloudDiagnosticText()
 );
}

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
 showView("dashboardView");
 $("projectList").innerHTML='<div class="panel">BauManagerCloud wird direkt in OneDrive geöffnet …</div>';
 try{
  const data=await findProjectIndex();
  state.projects=data.projects||[];
  renderProjects(state.projects);
 }catch(e){
  $("projectList").innerHTML=`<div class="panel"><h2>Keine Projektdaten gefunden</h2><p>Die App hat zuerst den Ordner <strong>BauManager/BauManagerCloud</strong> direkt geöffnet und anschließend die Ausweichsuche ausgeführt.</p><p>Starte auf dem Büro-PC <strong>Mobile_Daten_einmal_synchronisieren.bat</strong> und warte auf den grünen OneDrive-Haken.</p><p class="status error">${escapeHtml(e.message)}</p></div>`;
 }
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
 if(!state.cloudBaseFolder)await findProjectIndex();
 const uploadId=createUploadId(),base=`${state.cloudBaseFolder}/uploads/${uploadId}`;await ensureFolderPath(base);
 const project=state.selectedProject,position=state.selectedPosition,category=state.category;let blob=file,name=sanitizeFilename(file.name),sheetNo=null,externalId=null;
 if(category==="aufmass"){sheetNo=Number($("sheetNumber").value);externalId=currentExternalId();if(file.type.startsWith("image/")){blob=await imageToPdf(file);name=`${externalId}.pdf`}else if(file.type==="application/pdf"||file.name.toLowerCase().endsWith(".pdf"))name=`${externalId}.pdf`;else throw new Error("Für Aufmaße sind Bilder oder PDF-Dateien zulässig.")}
 const metadata={version:2,upload_id:uploadId,uploaded_at:new Date().toISOString(),project_id:project.id,project_number:project.project_number,project_name:project.name,position_id:position.id,position_ordinal:position.ordinal,position_short_text:position.short_text,category,original_filename:file.name,stored_filename:name,sheet_no:sheetNo,external_id:externalId,client_source:"BauManager Mobile v2.0"};
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

// Digitales Aufmaß – Grundgerüst
const digitalState={strokes:[],redo:[],drawing:false,current:null,tool:"pen",lineWidth:2.4};
function digitalDraftKey(){return `digital-measurement:${currentExternalId()}`;}
function resizeDrawingCanvas(){const c=$("drawingCanvas"); if(!c)return; const r=c.getBoundingClientRect(),dpr=window.devicePixelRatio||1; const old=document.createElement("canvas"); old.width=c.width;old.height=c.height;old.getContext("2d").drawImage(c,0,0); c.width=Math.max(1,Math.round(r.width*dpr));c.height=Math.max(1,Math.round(r.height*dpr)); const x=c.getContext("2d");x.scale(dpr,dpr); redrawDigitalCanvas();}
function canvasPoint(e){const r=$("drawingCanvas").getBoundingClientRect();return {x:e.clientX-r.left,y:e.clientY-r.top,p:e.pressure||0.5};}
function redrawDigitalCanvas(){const c=$("drawingCanvas"),x=c.getContext("2d"),dpr=window.devicePixelRatio||1;x.setTransform(1,0,0,1,0,0);x.clearRect(0,0,c.width,c.height);x.scale(dpr,dpr);for(const s of digitalState.strokes){x.beginPath();x.lineCap="round";x.lineJoin="round";x.strokeStyle=s.tool==="eraser"?"#fff":"#17212b";x.lineWidth=s.tool==="eraser"?18:s.width; s.points.forEach((p,i)=>i?x.lineTo(p.x,p.y):x.moveTo(p.x,p.y));x.stroke();}}
function bindDrawingCanvas(){const c=$("drawingCanvas");c.style.touchAction="none";c.onpointerdown=e=>{c.setPointerCapture(e.pointerId);digitalState.drawing=true;digitalState.current={tool:digitalState.tool,width:digitalState.lineWidth,points:[canvasPoint(e)]};digitalState.redo=[];};c.onpointermove=e=>{if(!digitalState.drawing)return;digitalState.current.points.push(canvasPoint(e));digitalState.strokes.push(digitalState.current);redrawDigitalCanvas();digitalState.strokes.pop();};c.onpointerup=e=>{if(!digitalState.drawing)return;digitalState.current.points.push(canvasPoint(e));digitalState.strokes.push(digitalState.current);digitalState.current=null;digitalState.drawing=false;redrawDigitalCanvas();autoSaveDigitalDraft();};}
function renderDigitalSheetNumbers(){const s=$("digitalSheetNumber");s.innerHTML="";const sheets=state.selectedPosition.measurement_sheets||[];const next=Math.max(0,...sheets.map(v=>Number(v.sheet_no)||0))+1;for(const sh of sheets){const o=document.createElement("option");o.value=sh.sheet_no;o.textContent=`Blatt ${String(sh.sheet_no).padStart(3,"0")} – ${sh.status||"gedruckt"}`;s.appendChild(o);}const n=document.createElement("option");n.value=next;n.textContent=`Neues Blatt ${String(next).padStart(3,"0")}`;s.appendChild(n);}
function digitalExternalId(){const n=Number($("digitalSheetNumber").value||0);return `${state.selectedProject.project_number}-${compactOrdinal(state.selectedPosition.ordinal)}-${String(n).padStart(3,"0")}`;}
function openDigitalMeasurement(){renderDigitalSheetNumbers();$("digitalDate").value=new Date().toISOString().slice(0,10);$("digitalProject").textContent=`${state.selectedProject.project_number} – ${state.selectedProject.name}`;$("digitalPosition").textContent=`${state.selectedPosition.ordinal} – ${state.selectedPosition.short_text||""}`;$("digitalUnit").textContent=state.selectedPosition.unit||"";updateDigitalId();showView("digitalMeasurementView");setTimeout(()=>{resizeDrawingCanvas();bindDrawingCanvas();loadDigitalDraft();},50);}
function updateDigitalId(){$("digitalId").textContent=digitalExternalId();loadDigitalDraft();}
function draftObject(){return {version:1,id:digitalExternalId(),date:$("digitalDate").value,project:state.selectedProject,position:state.selectedPosition,strokes:digitalState.strokes,updated_at:new Date().toISOString(),status:"Entwurf"};}
function autoSaveDigitalDraft(){localStorage.setItem(digitalDraftKey(),JSON.stringify(draftObject()));$("digitalStatus").textContent="Entwurf lokal gespeichert";}
function loadDigitalDraft(){try{const raw=localStorage.getItem(digitalDraftKey());digitalState.strokes=raw?JSON.parse(raw).strokes||[]:[];}catch{digitalState.strokes=[];}digitalState.redo=[];redrawDigitalCanvas();}
async function digitalPreviewBlob(){const c=$("drawingCanvas");return await new Promise(r=>c.toBlob(r,"image/png"));}
async function digitalPdfBlob(){const pdf=await PDFLib.PDFDocument.create(),page=pdf.addPage([595.28,841.89]);const f=await pdf.embedFont(PDFLib.StandardFonts.Helvetica),fb=await pdf.embedFont(PDFLib.StandardFonts.HelveticaBold);const id=digitalExternalId();page.drawText("Aufmaßblatt",{x:230,y:805,size:18,font:fb});page.drawText(`ID: ${id}`,{x:430,y:808,size:8,font:fb});page.drawText(`Projekt: ${state.selectedProject.name}`,{x:35,y:775,size:9,font:fb});page.drawText(`Position: ${state.selectedPosition.ordinal} – ${state.selectedPosition.short_text||""}`.slice(0,95),{x:35,y:755,size:8,font:f});page.drawText(`Datum: ${$("digitalDate").value}`,{x:430,y:775,size:8,font:f});page.drawRectangle({x:35,y:95,width:525,height:640,borderWidth:1,borderColor:PDFLib.rgb(.35,.4,.45)});const png=await digitalPreviewBlob();const img=await pdf.embedPng(await png.arrayBuffer());page.drawImage(img,{x:38,y:98,width:519,height:634});page.drawText(`BauManager Digitales Aufmaß | ${id}`,{x:35,y:55,size:7,font:f});return new Blob([await pdf.save()],{type:"application/pdf"});}
async function saveDigitalToCloud(){const btn=$("digitalCloudSave");btn.disabled=true;setStatus($("digitalStatus"),"Digitales Aufmaß wird übertragen …");try{if(!state.cloudBaseFolder)await findProjectIndex();const id=digitalExternalId(),uploadId=createUploadId(),base=`${state.cloudBaseFolder}/uploads/${uploadId}`;await ensureFolderPath(base);const pdf=await digitalPdfBlob(),preview=await digitalPreviewBlob(),draft=draftObject();const jsonBlob=new Blob([JSON.stringify(draft,null,2)],{type:"application/json"});await uploadFile(`${base}/${id}.pdf`,pdf);await uploadFile(`${base}/${id}.json`,jsonBlob);await uploadFile(`${base}/${id}.preview.png`,preview);const metadata={version:3,upload_id:uploadId,uploaded_at:new Date().toISOString(),project_id:state.selectedProject.id,project_number:state.selectedProject.project_number,project_name:state.selectedProject.name,position_id:state.selectedPosition.id,position_ordinal:state.selectedPosition.ordinal,position_short_text:state.selectedPosition.short_text,category:"aufmass",original_filename:`${id}.json`,stored_filename:`${id}.pdf`,sheet_no:Number($("digitalSheetNumber").value),external_id:id,digital_draft_filename:`${id}.json`,preview_filename:`${id}.preview.png`,client_source:"BauManager Mobile v2.0"};await uploadFile(`${base}/metadata.json`,new Blob([JSON.stringify(metadata,null,2)],{type:"application/json"}));autoSaveDigitalDraft();setStatus($("digitalStatus"),"Aufmaßblatt erfolgreich übertragen.",false,true);}catch(e){setStatus($("digitalStatus"),e.message,true);}finally{btn.disabled=false;}}

$("loginButton").onclick=login;$("logoutButton").onclick=logout;$("refreshButton").onclick=loadProjects;
$("homeButton").onclick=()=>state.account&&showView("dashboardView");
$("accountButton").onclick=()=>{state.previousView=views.find(v=>!$(v).classList.contains("hidden"))||"dashboardView";showView("accountView")};
$("backFromAccount").onclick=()=>showView(state.previousView);
$("backToProjects").onclick=()=>showView("dashboardView");$("backToPositions").onclick=()=>showView("positionView");$("backFromDigital").onclick=()=>showView("positionDetailView");$("digitalMeasurementButton").onclick=openDigitalMeasurement;$("digitalSheetNumber").onchange=updateDigitalId;$("digitalDate").onchange=autoSaveDigitalDraft;$("digitalPen").onclick=()=>{digitalState.tool="pen"};$("digitalEraser").onclick=()=>{digitalState.tool="eraser"};$("digitalUndo").onclick=()=>{const s=digitalState.strokes.pop();if(s)digitalState.redo.push(s);redrawDigitalCanvas();autoSaveDigitalDraft()};$("digitalRedo").onclick=()=>{const s=digitalState.redo.pop();if(s)digitalState.strokes.push(s);redrawDigitalCanvas();autoSaveDigitalDraft()};$("digitalClear").onclick=()=>{if(confirm("Skizze wirklich löschen?")){digitalState.strokes=[];digitalState.redo=[];redrawDigitalCanvas();autoSaveDigitalDraft()}};$("digitalLocalSave").onclick=autoSaveDigitalDraft;$("digitalCloudSave").onclick=saveDigitalToCloud;$("closeUploadPanel").onclick=closeUploadPanel;
$("cameraInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("libraryInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("uploadButton").onclick=uploadSelectedFiles;$("sheetNumber").onchange=updateIdPreview;
document.querySelectorAll(".action-card").forEach(b=>b.onclick=()=>openUploadPanel(b.dataset.category));
$("projectSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase();renderProjects(state.projects.filter(p=>`${p.project_number||""} ${p.name||""}`.toLowerCase().includes(q)))};
$("positionSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase(),positions=state.selectedProject?.positions||[];renderPositions(positions.filter(p=>`${p.ordinal||""} ${p.short_text||""} ${titleKey(p)}`.toLowerCase().includes(q)))};
if("serviceWorker"in navigator)navigator.serviceWorker.register("sw.js").catch(()=>{});
initialize().catch(e=>{showView("loginView");setStatus($("loginStatus"),e.message,true)});
})();
