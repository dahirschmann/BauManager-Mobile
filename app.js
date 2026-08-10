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
 const metadata={version:2,upload_id:uploadId,uploaded_at:new Date().toISOString(),project_id:project.id,project_number:project.project_number,project_name:project.name,position_id:position.id,position_ordinal:position.ordinal,position_short_text:position.short_text,category,original_filename:file.name,stored_filename:name,sheet_no:sheetNo,external_id:externalId,client_source:"BauManager Mobile v2.9.1"};
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
const digitalState={strokes:[],redo:[],drawing:false,current:null,tool:"pen",texts:[],calculations:[],formulaPoint:null,lineWidth:2.4,background:"#ffffff",strokeColor:"#17212b",bound:false,zoom:1,panX:0,panY:0,panMode:false,pointers:new Map(),pinchStart:null};
function digitalDraftKey(){return `digital-measurement:${digitalExternalId()}`;}

function clamp(value,min,max){return Math.max(min,Math.min(max,value));}

function constrainDigitalPan(){
 const viewport=$("digitalViewport"),sheet=$("digitalSheetViewport"); if(!viewport||!sheet)return;
 const sheetWidth=sheet.offsetWidth||900,sheetHeight=sheet.offsetHeight||1120;
 const sw=sheetWidth*digitalState.zoom,sh=sheetHeight*digitalState.zoom,vw=viewport.clientWidth,vh=viewport.clientHeight,margin=28;
 if(sw<=vw-margin*2)digitalState.panX=(vw-sw)/2;else digitalState.panX=clamp(digitalState.panX,vw-sw-margin,margin);
 if(sh<=vh-margin*2)digitalState.panY=(vh-sh)/2;else digitalState.panY=clamp(digitalState.panY,vh-sh-margin,margin);
}
function updateZoomUi(){
 constrainDigitalPan();
 if($("digitalZoomValue"))$("digitalZoomValue").textContent=`${Math.round(digitalState.zoom*100)} %`;
 const sheet=$("digitalSheetViewport");
 if(sheet){
  sheet.style.setProperty("--sheet-scale",digitalState.zoom);
  sheet.style.setProperty("--sheet-pan-x",`${digitalState.panX}px`);
  sheet.style.setProperty("--sheet-pan-y",`${digitalState.panY}px`);
 }
}
function setDigitalZoom(next,anchorX=null,anchorY=null){
 const viewport=$("digitalViewport");
 if(!viewport)return;
 const old=digitalState.zoom;
 const value=clamp(Number(next)||1,0.45,4);
 if(anchorX!==null&&anchorY!==null&&old!==value){
  const rect=viewport.getBoundingClientRect();
  const lx=anchorX-rect.left, ly=anchorY-rect.top;
  const wx=(lx-digitalState.panX)/old, wy=(ly-digitalState.panY)/old;
  digitalState.panX=lx-wx*value;
  digitalState.panY=ly-wy*value;
 }
 digitalState.zoom=value;
 updateZoomUi();
}
function resetDigitalView(){
 const viewport=$("digitalViewport"),sheet=$("digitalSheetViewport");if(!viewport||!sheet)return;
 const desktop=window.matchMedia("(min-width: 901px)").matches;
 const sheetWidth=sheet.offsetWidth||900,sheetHeight=sheet.offsetHeight||1120,pad=desktop?36:16;
 const fit=Math.min((viewport.clientWidth-pad*2)/sheetWidth,(viewport.clientHeight-pad*2)/sheetHeight);
 digitalState.zoom=clamp(fit,0.30,desktop?0.95:0.90);
 digitalState.panX=(viewport.clientWidth-sheetWidth*digitalState.zoom)/2;
 digitalState.panY=(viewport.clientHeight-sheetHeight*digitalState.zoom)/2;
 updateZoomUi();
}
function setDigitalPanMode(active){
 digitalState.panMode=Boolean(active);
 if($("digitalPan"))$("digitalPan").classList.toggle("active-tool",digitalState.panMode);
 if($("digitalViewport"))$("digitalViewport").classList.toggle("pan-mode",digitalState.panMode);
 if(digitalState.panMode){
  digitalState.tool="pan";
  ["digitalPen","digitalEraser","digitalText","digitalFormula"].forEach(id=>$(id)?.classList.remove("active-tool"));
 }
}
function screenToCanvasPoint(event){
 const canvas=$("drawingCanvas"),r=canvas.getBoundingClientRect();
 return {x:(event.clientX-r.left)*((canvas.clientWidth||900)/r.width),y:(event.clientY-r.top)*((canvas.clientHeight||720)/r.height),p:event.pointerType==="pen"&&event.pressure>0?event.pressure:0.5};
}
function beginPinch(){
 if(digitalState.pointers.size!==2)return;const p=[...digitalState.pointers.values()],cx=(p[0].x+p[1].x)/2,cy=(p[0].y+p[1].y)/2;
 digitalState.pinchStart={distance:Math.hypot(p[1].x-p[0].x,p[1].y-p[0].y),zoom:digitalState.zoom,centerX:cx,centerY:cy,lastCenterX:cx,lastCenterY:cy};
}
function updatePinch(){
 if(digitalState.pointers.size!==2||!digitalState.pinchStart)return;const p=[...digitalState.pointers.values()],d=Math.max(1,Math.hypot(p[1].x-p[0].x,p[1].y-p[0].y)),cx=(p[0].x+p[1].x)/2,cy=(p[0].y+p[1].y)/2;
 digitalState.panX+=cx-digitalState.pinchStart.lastCenterX;digitalState.panY+=cy-digitalState.pinchStart.lastCenterY;digitalState.pinchStart.lastCenterX=cx;digitalState.pinchStart.lastCenterY=cy;
 setDigitalZoom(digitalState.pinchStart.zoom*(d/digitalState.pinchStart.distance),cx,cy);updateZoomUi();
}


function openFormulaModalAt(x,y){
 digitalState.formulaPoint={x,y};
 $("formulaModal").classList.remove("hidden");
 $("formulaType").value="flaeche";
 $("formulaName").value="";
 $("formulaA").value="";
 $("formulaB").value="";
 $("formulaC").value="";
 $("formulaExpression").value="";
 updateFormulaFields();
 setTimeout(()=>$("formulaName").focus(),50);
}
function closeFormulaModal(){
 $("formulaModal").classList.add("hidden");
 digitalState.formulaPoint=null;
}
function updateFormulaFields(){
 const type=$("formulaType").value;
 $("formulaFieldB").classList.toggle("hidden",type==="strecke"||type==="frei");
 $("formulaFieldC").classList.toggle("hidden",type!=="volumen");
 $("formulaExpressionField").classList.toggle("hidden",type!=="frei");
 $("formulaFieldA").classList.toggle("hidden",type==="frei");
 const labels={
  strecke:["Länge / Wert","",""],
  flaeche:["Länge","Breite / Höhe",""],
  volumen:["Länge","Breite","Höhe"]
 };
 const values=labels[type]||labels.flaeche;
 if($("formulaALabel"))$("formulaALabel").textContent=values[0];
 if($("formulaBLabel"))$("formulaBLabel").textContent=values[1];
 if($("formulaCLabel"))$("formulaCLabel").textContent=values[2];
}
function parseGermanNumber(value){
 const text=String(value??"").trim().replace(/\s/g,"").replace(/\./g,"").replace(",",".");
 const number=Number(text);
 return Number.isFinite(number)?number:0;
}
function formulaTextFromDialog(){
 const type=$("formulaType").value;
 const name=$("formulaName").value.trim() || (
  type==="strecke"?"Strecke":type==="flaeche"?"Fläche":type==="volumen"?"Volumen":"Berechnung"
 );
 const a=parseGermanNumber($("formulaA").value);
 const b=parseGermanNumber($("formulaB").value);
 const c=parseGermanNumber($("formulaC").value);
 let result=0, formula="";
 if(type==="strecke"){
  result=a; formula=`${formatCalcNumber(a)}`;
 }else if(type==="flaeche"){
  result=a*b; formula=`${formatCalcNumber(a)} × ${formatCalcNumber(b)}`;
 }else if(type==="volumen"){
  result=a*b*c; formula=`${formatCalcNumber(a)} × ${formatCalcNumber(b)} × ${formatCalcNumber(c)}`;
 }else{
  formula=$("formulaExpression").value.trim();
  result=parseGermanNumber($("formulaA").value || prompt("Ergebnis:", "0"));
 }
 return {type,name,result,formula,text:`${name}: ${formula} = ${formatCalcNumber(result)} ${state.selectedPosition?.unit||""}`};
}
function insertFormulaFromModal(){
 if(!digitalState.formulaPoint)return;
 const item=formulaTextFromDialog();
 digitalState.texts.push({
  x:digitalState.formulaPoint.x,
  y:digitalState.formulaPoint.y,
  text:item.text,
  size:Number($("digitalTextSize")?.value||16),
  color:digitalState.strokeColor||"#17212b",
  kind:"formula",
  formula:item.formula,
  result:item.result
 });
 redrawDigitalCanvas();
 autoSaveDigitalDraft();
 closeFormulaModal();
}

function activateDigitalTool(tool){
 setDigitalPanMode(false);
 setDigitalTool(tool);
 if($("digitalText"))$("digitalText").classList.toggle("active-tool",tool==="text");
}
function addTextAt(x,y){
 const text=prompt("Text eingeben:");
 if(text===null||!String(text).trim())return;
 digitalState.texts.push({
  x,y,text:String(text),
  size:Number($("digitalTextSize")?.value||16),
  color:digitalState.strokeColor||"#17212b"
 });
 redrawDigitalCanvas();
 autoSaveDigitalDraft();
}
function formatCalcNumber(value){
 const n=Number(value);
 return Number.isFinite(n)?n.toLocaleString("de-DE",{minimumFractionDigits:0,maximumFractionDigits:3}):"0";
}
function calculationResult(calc){
 if(calc.type==="strecke")return Number(calc.a||0);
 if(calc.type==="flaeche")return Number(calc.a||0)*Number(calc.b||0);
 if(calc.type==="volumen")return Number(calc.a||0)*Number(calc.b||0)*Number(calc.c||0);
 if(calc.type==="frei")return Number(calc.result||0);
 return 0;
}
function calculationLabel(calc){
 const r=formatCalcNumber(calculationResult(calc));
 if(calc.type==="strecke")return `${calc.name||"Strecke"}: ${formatCalcNumber(calc.a)} = ${r}`;
 if(calc.type==="flaeche")return `${calc.name||"Fläche"}: ${formatCalcNumber(calc.a)} × ${formatCalcNumber(calc.b)} = ${r}`;
 if(calc.type==="volumen")return `${calc.name||"Volumen"}: ${formatCalcNumber(calc.a)} × ${formatCalcNumber(calc.b)} × ${formatCalcNumber(calc.c)} = ${r}`;
 return `${calc.name||"Freie Formel"}: ${calc.formula||""} = ${r}`;
}
function renderCalculationList(){
 const box=$("calculationList");
 if(!box)return;
 box.innerHTML="";
 for(const calc of digitalState.calculations||[]){
  const row=document.createElement("div");
  row.className="calc-row";
  const checked=calc.include!==false?"checked":"";
  row.innerHTML=`<label><input type="checkbox" ${checked}> <span>${escapeHtml(calculationLabel(calc))} ${escapeHtml(state.selectedPosition?.unit||"")}</span></label><button>Ins Blatt</button>`;
  row.querySelector("input").onchange=e=>{calc.include=e.target.checked;updateSheetTotal();autoSaveDigitalDraft();};
  row.querySelector("button").onclick=()=>insertCalculationText(calc);
  box.appendChild(row);
 }
 updateSheetTotal();
}
function updateSheetTotal(){
 const total=(digitalState.calculations||[]).filter(c=>c.include!==false).reduce((sum,c)=>sum+calculationResult(c),0);
 if($("calculationTotal"))$("calculationTotal").textContent=`Blattsumme: ${formatCalcNumber(total)} ${state.selectedPosition?.unit||""}`;
 if($("digitalCalculatedSum"))$("digitalCalculatedSum").textContent=formatCalcNumber(total);
}
function insertCalculationText(calc){
 const x=45, y=110+(digitalState.texts.length*24)%430;
 digitalState.texts.push({x,y,text:calculationLabel(calc),size:16,color:digitalState.strokeColor||"#17212b"});
 redrawDigitalCanvas();autoSaveDigitalDraft();
}
function openCalculationDialog(type){
 const name=prompt("Bezeichnung:",type==="strecke"?"Strecke":type==="flaeche"?"Fläche":type==="volumen"?"Volumen":"Berechnung");
 if(name===null)return;
 let calc={id:Date.now(),type,name,include:true};
 if(type==="strecke"){
  calc.a=prompt("Länge / Wert:","0")?.replace(",",".")||0;
 }else if(type==="flaeche"){
  calc.a=prompt("Länge:","0")?.replace(",",".")||0;
  calc.b=prompt("Breite / Höhe:","0")?.replace(",",".")||0;
 }else if(type==="volumen"){
  calc.a=prompt("Länge:","0")?.replace(",",".")||0;
  calc.b=prompt("Breite:","0")?.replace(",",".")||0;
  calc.c=prompt("Höhe:","0")?.replace(",",".")||0;
 }else{
  calc.formula=prompt("Formel / Beschreibung:","")||"";
  calc.result=prompt("Ergebnis:","0")?.replace(",",".")||0;
 }
 digitalState.calculations.push(calc);
 renderCalculationList();
 autoSaveDigitalDraft();
}
function setDigitalTool(tool){
 digitalState.tool=tool;
 $("digitalPen").classList.toggle("active-tool",tool==="pen");
 $("digitalEraser").classList.toggle("active-tool",tool==="eraser");
}
function updateDigitalWidth(){
 const value=Number($("digitalLineWidth")?.value||2.4);
 digitalState.lineWidth=Math.max(1,Math.min(10,value));
 if($("digitalLineWidthValue"))$("digitalLineWidthValue").textContent=`${digitalState.lineWidth.toFixed(1)} mm`;
}
function resizeDrawingCanvas(){
 const canvas=$("drawingCanvas");
 if(!canvas)return;
 const logicalWidth=Math.max(1,canvas.clientWidth||900);
 const logicalHeight=Math.max(1,canvas.clientHeight||720);
 const dpr=Math.max(1,window.devicePixelRatio||1);
 const width=Math.round(logicalWidth*dpr),height=Math.round(logicalHeight*dpr);
 if(canvas.width!==width||canvas.height!==height){canvas.width=width;canvas.height=height;}
 redrawDigitalCanvas();
}
function canvasPoint(event){return screenToCanvasPoint(event);}
function drawStroke(context,stroke){
 if(!stroke?.points?.length)return;
 context.save();
 context.beginPath();
 context.lineCap="round";
 context.lineJoin="round";
 context.globalCompositeOperation=stroke.tool==="eraser"?"destination-out":"source-over";
 context.strokeStyle=stroke.color||"#17212b";
 context.lineWidth=stroke.tool==="eraser"
  ? Math.max(18,stroke.width*6)
  : Math.max(1,stroke.width*(0.65+(stroke.points[0]?.p||0.5)*0.7));
 stroke.points.forEach((point,index)=>{
  if(index===0)context.moveTo(point.x,point.y);
  else context.lineTo(point.x,point.y);
 });
 context.stroke();
 context.restore();
}
function redrawDigitalCanvas(){
 const canvas=$("drawingCanvas"); if(!canvas)return;
 const context=canvas.getContext("2d");
 const dpr=Math.max(1,window.devicePixelRatio||1);
 const logicalWidth=Math.max(1,canvas.clientWidth||900);
 const logicalHeight=Math.max(1,canvas.clientHeight||720);
 context.setTransform(1,0,0,1,0,0);
 context.clearRect(0,0,canvas.width,canvas.height);
 context.fillStyle=digitalState.background;
 context.fillRect(0,0,canvas.width,canvas.height);
 context.setTransform(dpr,0,0,dpr,0,0);
 context.save();context.beginPath();context.rect(0,0,logicalWidth,logicalHeight);context.clip();
 for(const stroke of digitalState.strokes)drawStroke(context,stroke);
 for(const item of digitalState.texts||[]){
  context.save();
  context.fillStyle=item.color||"#17212b";
  context.font=`${item.size||16}px Arial`;
  context.textBaseline="top";
  String(item.text||"").split("\n").forEach((line,index)=>context.fillText(line,item.x,item.y+index*(item.size||16)*1.25));
  context.restore();
 }
 if(digitalState.current)drawStroke(context,digitalState.current);
 context.restore();
}
function finishDigitalStroke(event){
 if(!digitalState.drawing)return;
 if(event)digitalState.current.points.push(canvasPoint(event));
 if(digitalState.current?.points?.length){
  digitalState.strokes.push(digitalState.current);
 }
 digitalState.current=null;
 digitalState.drawing=false;
 redrawDigitalCanvas();
 autoSaveDigitalDraft();
}
function bindDrawingCanvas(){
 const canvas=$("drawingCanvas"), viewport=$("digitalViewport");
 if(!canvas||!viewport||digitalState.bound)return;
 digitalState.bound=true;
 canvas.style.touchAction="none";
 viewport.style.touchAction="none";
 const updatePointer=e=>digitalState.pointers.set(e.pointerId,{x:e.clientX,y:e.clientY});
 viewport.addEventListener("pointerdown",e=>{
  updatePointer(e);
  if(digitalState.pointers.size===2){digitalState.drawing=false;digitalState.current=null;beginPinch();return;}
  if(digitalState.panMode){
   viewport.setPointerCapture(e.pointerId);
   viewport.dataset.panPointer=String(e.pointerId);
   viewport.dataset.panStartX=String(e.clientX); viewport.dataset.panStartY=String(e.clientY);
   viewport.dataset.panBaseX=String(digitalState.panX); viewport.dataset.panBaseY=String(digitalState.panY);
   return;
  }
  if(e.target!==canvas)return;
  if(digitalState.tool==="pan")return;
  if(digitalState.tool==="text"){
   e.preventDefault();
   const point=canvasPoint(e);
   addTextAt(point.x,point.y);
   return;
  }
  if(digitalState.tool==="formula"){
   e.preventDefault();
   const point=canvasPoint(e);
   openFormulaModalAt(point.x,point.y);
   return;
  }
  e.preventDefault();
  canvas.setPointerCapture(e.pointerId);
  digitalState.drawing=true; digitalState.redo=[];
  digitalState.current={tool:digitalState.tool,width:digitalState.lineWidth,color:digitalState.strokeColor,points:[canvasPoint(e)]};
  redrawDigitalCanvas();
 });
 viewport.addEventListener("pointermove",e=>{
  if(!digitalState.pointers.has(e.pointerId))return;
  updatePointer(e);
  if(digitalState.pointers.size===2){e.preventDefault();updatePinch();return;}
  if(viewport.dataset.panPointer===String(e.pointerId)){
   digitalState.panX=Number(viewport.dataset.panBaseX||0)+(e.clientX-Number(viewport.dataset.panStartX||0));
   digitalState.panY=Number(viewport.dataset.panBaseY||0)+(e.clientY-Number(viewport.dataset.panStartY||0));
   updateZoomUi(); return;
  }
  if(!digitalState.drawing||!digitalState.current)return;
  e.preventDefault();
  const events=e.getCoalescedEvents?e.getCoalescedEvents():[e];
  for(const item of events)digitalState.current.points.push(canvasPoint(item));
  redrawDigitalCanvas();
 });
 const finish=e=>{
  digitalState.pointers.delete(e.pointerId);
  if(digitalState.pointers.size<2)digitalState.pinchStart=null;
  if(viewport.dataset.panPointer===String(e.pointerId)){delete viewport.dataset.panPointer;return;}
  if(digitalState.drawing)finishDigitalStroke(e);
 };
 viewport.addEventListener("pointerup",finish);
 viewport.addEventListener("pointercancel",finish);
 viewport.addEventListener("wheel",e=>{
  e.preventDefault();
  if(e.ctrlKey||e.metaKey) setDigitalZoom(digitalState.zoom*(e.deltaY<0?1.12:0.88),e.clientX,e.clientY);
  else {digitalState.panX-=e.deltaX;digitalState.panY-=e.deltaY;updateZoomUi();}
 },{passive:false});
}
function renderDigitalSheetNumbers(){
 const select=$("digitalSheetNumber");
 select.innerHTML="";
 const sheets=state.selectedPosition.measurement_sheets||[];
 const used=new Set(sheets.map(value=>Number(value.sheet_no)||0));
 const prefix=`digital-measurement:${state.selectedProject.project_number}-${compactOrdinal(state.selectedPosition.ordinal)}-`;
 for(let index=0;index<localStorage.length;index++){
  const key=localStorage.key(index)||"";
  if(key.startsWith(prefix)){
   const match=key.match(/-(\d{3})$/);
   if(match)used.add(Number(match[1]));
  }
 }
 for(const sheet of sheets){
  const option=document.createElement("option");
  option.value=sheet.sheet_no;
  option.textContent=`Blatt ${String(sheet.sheet_no).padStart(3,"0")} – ${sheet.status||"gedruckt"}`;
  select.appendChild(option);
 }
 let next=1;
 while(used.has(next))next++;
 const option=document.createElement("option");
 option.value=next;
 option.textContent=`Neues digitales Blatt ${String(next).padStart(3,"0")}`;
 option.selected=true;
 select.appendChild(option);
}
function digitalExternalId(){const n=Number($("digitalSheetNumber").value||0);return `${state.selectedProject.project_number}-${compactOrdinal(state.selectedPosition.ordinal)}-${String(n).padStart(3,"0")}`;}

function toggleMobileEditorPanel(name){
 const editor=$("digitalMeasurementView"); if(!editor)return;
 const current=editor.dataset.mobilePanel||"";
 editor.dataset.mobilePanel=current===name?"":name;
 document.querySelectorAll(".mobile-editor-dock button[data-panel]").forEach(b=>b.classList.toggle("active",b.dataset.panel===editor.dataset.mobilePanel));
}
function closeMobileEditorPanels(){
 const editor=$("digitalMeasurementView"); if(editor)editor.dataset.mobilePanel="";
 document.querySelectorAll(".mobile-editor-dock button[data-panel]").forEach(b=>b.classList.remove("active"));
}

function openDigitalMeasurement(){
 closeUploadPanel();
 closeMobileEditorPanels();
 renderDigitalSheetNumbers();
 $("digitalDate").value=new Date().toISOString().slice(0,10);

 const project=state.selectedProject||{};
 const position=state.selectedPosition||{};

 $("digitalProject").textContent=project.name||"";
 $("digitalPositionNumber").textContent=position.ordinal||"";
 $("digitalPositionText").textContent=position.short_text||"";
 $("digitalUnit").textContent=position.unit||"";
 $("digitalQuantity").textContent=formatNumber(position.quantity);
 $("digitalClientName").textContent=project.client_name||"";
 $("digitalClientAddress").textContent=project.client_address||"";
 $("digitalContractorName").textContent=project.contractor_name||"";
 $("digitalContractorAddress").textContent=project.contractor_address||"";

 setDigitalTool("pan");
 setDigitalPanMode(true);
 updateDigitalWidth();
 showView("digitalMeasurementView");
 updateDigitalId();
 requestAnimationFrame(()=>requestAnimationFrame(()=>{
  resizeDrawingCanvas();
  bindDrawingCanvas();
  resetDigitalView();
  loadDigitalDraft(); renderCalculationList();
 }));
}
function updateDigitalId(){
 const id=digitalExternalId();
 $("digitalId").textContent=id;
 if($("digitalPageId"))$("digitalPageId").textContent=`ID: ${id}`;
 if($("digitalTopUnit")){
  const unit=state.selectedPosition?.unit||"";
  $("digitalTopUnit").textContent=unit;
  $("digitalBottomUnit1").textContent=unit;
  $("digitalBottomUnit2").textContent=unit;
 }
 if($("digitalFooterId"))$("digitalFooterId").textContent=`BauAufmaß | ID: ${id}`;
 const dateValue=$("digitalDate")?.value||"";
 if($("digitalDateDisplay")){
  const parts=dateValue.split("-");
  $("digitalDateDisplay").textContent=parts.length===3?`${parts[2]}.${parts[1]}.${parts[0]}`:dateValue;
 }
 if($("digitalPageNumber")){
  $("digitalPageNumber").textContent=String(Number($("digitalSheetNumber").value||0)).padStart(3,"0");
 }
 if($("digitalPreviousSum")){
  const current=Number($("digitalSheetNumber").value||0);
  $("digitalPreviousSum").textContent=current>1
   ? `Übertrag Summe Blatt ${String(current-1).padStart(3,"0")}`
   : "Übertrag Summe";
 }
 if($("digitalSheetSum")){
  $("digitalSheetSum").textContent=`Summe Blatt ${String(Number($("digitalSheetNumber").value||0)).padStart(3,"0")}`;
 }
 loadDigitalDraft();
}
function draftObject(){return {version:1,id:digitalExternalId(),date:$("digitalDate").value,project:state.selectedProject,position:state.selectedPosition,strokes:digitalState.strokes,texts:digitalState.texts,calculations:digitalState.calculations,updated_at:new Date().toISOString(),status:"Entwurf"};}
function autoSaveDigitalDraft(){localStorage.setItem(digitalDraftKey(),JSON.stringify(draftObject()));$("digitalStatus").textContent="Entwurf lokal gespeichert";}
function loadDigitalDraft(){
 try{
  const raw=localStorage.getItem(digitalDraftKey());
  const draft=raw?JSON.parse(raw):null;
  digitalState.strokes=draft?.strokes||[]; digitalState.texts=draft?.texts||[]; digitalState.calculations=draft?.calculations||[]; renderCalculationList();
  if(draft?.date)$("digitalDate").value=draft.date;
  $("digitalStatus").textContent=draft
   ? `Lokaler Entwurf geladen · ${new Date(draft.updated_at||Date.now()).toLocaleString("de-DE")}`
   : "Neues digitales Aufmaßblatt";
 }catch{
  digitalState.strokes=[];
  $("digitalStatus").textContent="Neues digitales Aufmaßblatt";
 }
 digitalState.redo=[];
 redrawDigitalCanvas();
}
async function digitalPreviewBlob(){
 const canvas=$("drawingCanvas");
 return await new Promise((resolve,reject)=>{
  canvas.toBlob(blob=>blob?resolve(blob):reject(new Error("Vorschaubild konnte nicht erstellt werden.")),"image/png",0.96);
 });
}
async function digitalTransparentPreviewBlob(){
 const source=$("drawingCanvas");
 const logicalWidth=Math.max(1,source.clientWidth||900);
 const logicalHeight=Math.max(1,source.clientHeight||720);
 const exportScale=2;
 const canvas=document.createElement("canvas");
 canvas.width=Math.round(logicalWidth*exportScale);
 canvas.height=Math.round(logicalHeight*exportScale);
 const context=canvas.getContext("2d");
 context.setTransform(exportScale,0,0,exportScale,0,0);
 context.save();context.beginPath();context.rect(0,0,logicalWidth,logicalHeight);context.clip();
 for(const stroke of digitalState.strokes)drawStroke(context,stroke);
 for(const item of digitalState.texts||[]){
  context.save();
  context.fillStyle=item.color||"#17212b";
  context.font=`${item.size||16}px Arial`;
  context.textBaseline="top";
  String(item.text||"").split("\n").forEach((line,index)=>context.fillText(line,item.x,item.y+index*(item.size||16)*1.25));
  context.restore();
 }
 context.restore();
 return await new Promise((resolve,reject)=>canvas.toBlob(blob=>blob?resolve(blob):reject(new Error("Zeichnung konnte nicht erstellt werden.")),"image/png"));
}
function pdfMm(value){return value*72/25.4;}
function pdfFormatQuantity(value){
 const number=Number(value);
 if(!Number.isFinite(number))return String(value??"");
 return number.toLocaleString("de-DE",{minimumFractionDigits:0,maximumFractionDigits:3});
}
function pdfTextWidth(font,text,size){return font.widthOfTextAtSize(String(text||""),size);}
function pdfDrawRight(page,font,text,right,y,size,options={}){
 page.drawText(String(text||""),{x:right-pdfTextWidth(font,text,size),y,size,font,...options});
}
function pdfDrawCentered(page,font,text,center,y,size,options={}){
 page.drawText(String(text||""),{x:center-pdfTextWidth(font,text,size)/2,y,size,font,...options});
}
function pdfWrapLines(font,text,size,maxWidth,maxLines=3){
 const words=String(text||"").replace(/\r/g,"").split(/\s+/).filter(Boolean);
 const lines=[];
 let current="";
 for(const word of words){
  const candidate=current?`${current} ${word}`:word;
  if(pdfTextWidth(font,candidate,size)<=maxWidth){
   current=candidate;
  }else{
   if(current)lines.push(current);
   current=word;
   if(lines.length>=maxLines)break;
  }
 }
 if(current&&lines.length<maxLines)lines.push(current);
 if(lines.length===maxLines){
  const used=lines.join(" ").split(/\s+/).length;
  if(used<words.length){
   let last=lines[lines.length-1];
   while(last&&pdfTextWidth(font,`${last}...`,size)>maxWidth)last=last.slice(0,-1);
   lines[lines.length-1]=`${last.trim()}...`;
  }
 }
 return lines;
}
function pdfDrawWrapped(page,font,text,x,y,maxWidth,size,lineHeight,maxLines=3,options={}){
 const lines=pdfWrapLines(font,text,size,maxWidth,maxLines);
 lines.forEach((line,index)=>page.drawText(line,{x,y:y-index*lineHeight,size,font,...options}));
 return y-lines.length*lineHeight;
}
function pdfDrawRect(page,x,y,width,height,borderWidth=.7,fillColor=null){
 page.drawRectangle({
  x,y,width,height,
  borderWidth,
  borderColor:PDFLib.rgb(0,0,0),
  ...(fillColor?{color:fillColor}:{})
 });
}
function pdfDrawSumBox(page,font,bold,x,y,width,height,rows){
 pdfDrawRect(page,x,y,width,height,.65,PDFLib.rgb(1,1,1));
 const rowHeight=height/rows.length;
 rows.forEach((row,index)=>{
  const rowBottom=y+height-(index+1)*rowHeight;
  if(index>0)page.drawLine({
   start:{x,y:rowBottom+rowHeight},
   end:{x:x+width,y:rowBottom+rowHeight},
   thickness:.65
  });
  page.drawText(row.label,{x:x+pdfMm(2.5),y:rowBottom+rowHeight-pdfMm(4.2),size:7.5,font:bold});
  const writeY=rowBottom+pdfMm(3.1);
  page.drawLine({
   start:{x:x+pdfMm(3),y:writeY},
   end:{x:x+width-pdfMm(12),y:writeY},
   thickness:.45
  });
  pdfDrawRight(page,font,row.unit||"",x+width-pdfMm(2.5),writeY-pdfMm(1),7.5);
 });
}
function digitalProjectAddress(value){
 return String(value||"").replace(/\r/g,"").split("\n").map(line=>line.trim()).filter(Boolean).join("\n");
}
async function digitalPdfBlob(){
 const pdf=await PDFLib.PDFDocument.create();
 const page=pdf.addPage([595.28,841.89]);
 const font=await pdf.embedFont(PDFLib.StandardFonts.Helvetica);
 const bold=await pdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
 const pageWidth=page.getWidth();
 const pageHeight=page.getHeight();
 const margin=pdfMm(12);
 const contentWidth=pageWidth-2*margin;
 const id=digitalExternalId();
 const sheetNo=Number($("digitalSheetNumber").value||0);
 const project=state.selectedProject||{};
 const position=state.selectedPosition||{};
 const unit=position.unit||"";
 const quantity=pdfFormatQuantity(position.quantity);
 const gridColor=PDFLib.rgb(0.72,0.745,0.776);

 pdf.setTitle(`Aufmaßblatt ${id}`);
 pdf.setAuthor("BauAufmaß");

 pdfDrawCentered(page,bold,"Aufmaßblatt",pageWidth/2,pageHeight-pdfMm(15)-5,20);
 pdfDrawRight(page,bold,`ID: ${id}`,pageWidth-margin,pageHeight-pdfMm(15)-2,8);

 const topY=pageHeight-pdfMm(22);
 const headerH=pdfMm(25);
 const partyW=pdfMm(77);
 const numberW=contentWidth-2*partyW;
 pdfDrawRect(page,margin,topY-headerH,partyW,headerH,.75);
 pdfDrawRect(page,margin+partyW,topY-headerH,partyW,headerH,.75);
 pdfDrawRect(page,margin+2*partyW,topY-headerH,numberW,headerH,.75);

 page.drawText("Auftraggeber:",{x:margin+pdfMm(3),y:topY-pdfMm(5)-2,size:8.5,font:bold});
 page.drawText("Auftragnehmer:",{x:margin+partyW+pdfMm(3),y:topY-pdfMm(5)-2,size:8.5,font:bold});
 page.drawText("Blatt-Nr.:",{x:margin+2*partyW+pdfMm(3),y:topY-pdfMm(5)-2,size:8.5,font:bold});

 pdfDrawWrapped(page,bold,project.client_name||"",margin+pdfMm(3),topY-pdfMm(10.5)-2,partyW-pdfMm(6),8,pdfMm(3.7),2);
 pdfDrawWrapped(page,font,digitalProjectAddress(project.client_address),margin+pdfMm(3),topY-pdfMm(18)-2,partyW-pdfMm(6),7.5,pdfMm(3.5),2);
 pdfDrawWrapped(page,bold,project.contractor_name||"",margin+partyW+pdfMm(3),topY-pdfMm(10.5)-2,partyW-pdfMm(6),8,pdfMm(3.7),2);
 pdfDrawWrapped(page,font,digitalProjectAddress(project.contractor_address),margin+partyW+pdfMm(3),topY-pdfMm(18)-2,partyW-pdfMm(6),7.5,pdfMm(3.5),2);
 pdfDrawCentered(page,font,String(sheetNo).padStart(3,"0"),margin+2*partyW+numberW/2,topY-pdfMm(13)-3,10);

 let y=topY-headerH;
 const projectH=pdfMm(17);
 pdfDrawRect(page,margin,y-projectH,contentWidth,projectH,.75);
 page.drawText("Bezeichnung der Bauleistung / Projekt:",{x:margin+pdfMm(3),y:y-pdfMm(4.7)-2,size:8,font:bold});
 pdfDrawWrapped(page,bold,project.name||"",margin+pdfMm(3),y-pdfMm(10.5)-2,contentWidth-pdfMm(6),8.5,pdfMm(3.8),2);

 y-=projectH;
 const serviceH=pdfMm(24);
 const rightW=pdfMm(42);
 const leftW=contentWidth-rightW;
 pdfDrawRect(page,margin,y-serviceH,leftW,serviceH,.75);
 pdfDrawRect(page,margin+leftW,y-serviceH,rightW,serviceH,.75);
 page.drawText("Position / Kurzbeschreibung:",{x:margin+pdfMm(3),y:y-pdfMm(4.8)-2,size:8,font:bold});
 page.drawText(position.ordinal||"",{x:margin+pdfMm(3),y:y-pdfMm(11)-2,size:8.5,font:bold});
 pdfDrawWrapped(page,bold,position.short_text||"",margin+pdfMm(24),y-pdfMm(11)-2,leftW-pdfMm(28),8.5,pdfMm(3.8),3);

 const rightX=margin+leftW;
 page.drawText("Einheit:",{x:rightX+pdfMm(3),y:y-pdfMm(5)-2,size:8,font:bold});
 pdfDrawRight(page,font,unit,rightX+rightW-pdfMm(3),y-pdfMm(5)-2,9);
 page.drawLine({start:{x:rightX,y:y-pdfMm(10)},end:{x:rightX+rightW,y:y-pdfMm(10)},thickness:.75});
 page.drawText("LV-Menge:",{x:rightX+pdfMm(3),y:y-pdfMm(15)-2,size:8,font:bold});
 pdfDrawRight(page,font,quantity,rightX+rightW-pdfMm(3),y-pdfMm(21)-2,9);

 y-=serviceH;
 const labelH=pdfMm(7);
 pdfDrawRect(page,margin,y-labelH,contentWidth,labelH,.75);
 page.drawText("Aufmaß / Skizze / Berechnung",{x:margin+pdfMm(3),y:y-pdfMm(4.7)-2,size:8.5,font:bold});

 const gridTop=y-labelH;
 const gridBottom=pdfMm(27);
 const gridHeight=gridTop-gridBottom;
 pdfDrawRect(page,margin,gridBottom,contentWidth,gridHeight,.7);

 for(let gx=margin+pdfMm(5);gx<margin+contentWidth;gx+=pdfMm(5)){
  page.drawLine({start:{x:gx,y:gridBottom},end:{x:gx,y:gridTop},thickness:.32,color:gridColor});
 }
 for(let gy=gridBottom+pdfMm(5);gy<gridTop;gy+=pdfMm(5)){
  page.drawLine({start:{x:margin,y:gy},end:{x:margin+contentWidth,y:gy},thickness:.32,color:gridColor});
 }

 const transparent=await digitalTransparentPreviewBlob();
 const drawing=await pdf.embedPng(await transparent.arrayBuffer());
 const canvas=$("drawingCanvas");
 const canvasRect=canvas.getBoundingClientRect();
 const sourceRatio=Math.max(.01,canvasRect.width/canvasRect.height);
 const targetRatio=contentWidth/gridHeight;
 let drawWidth=contentWidth;
 let drawHeight=gridHeight;
 let drawX=margin;
 let drawY=gridBottom;
 if(sourceRatio>targetRatio){
  drawHeight=contentWidth/sourceRatio;
  drawY=gridBottom+(gridHeight-drawHeight)/2;
 }else{
  drawWidth=gridHeight*sourceRatio;
  drawX=margin+(contentWidth-drawWidth)/2;
 }
 page.drawImage(drawing,{x:drawX,y:drawY,width:drawWidth,height:drawHeight});

 const sumW=pdfMm(56);
 const upperH=pdfMm(14);
 const lowerH=pdfMm(27);
 const inset=pdfMm(3);
 const previousLabel=sheetNo>1?`Übertrag Summe Blatt ${String(sheetNo-1).padStart(3,"0")}`:"Übertrag Summe";
 pdfDrawSumBox(page,font,bold,pageWidth-margin-sumW-inset,gridTop-upperH-inset,sumW,upperH,[
  {label:previousLabel,unit}
 ]);
 pdfDrawSumBox(page,font,bold,pageWidth-margin-sumW-inset,gridBottom+inset,sumW,lowerH,[
  {label:`Summe Blatt ${String(sheetNo).padStart(3,"0")}`,unit},
  {label:"Gesamtsumme",unit}
 ]);

 const signatureY=pdfMm(17);
 const lineY=pdfMm(12);
 const leftStart=margin+pdfMm(4);
 const leftEnd=margin+pdfMm(62);
 const dateStart=pageWidth/2-pdfMm(23);
 const dateEnd=pageWidth/2+pdfMm(23);
 const rightStart=pageWidth-margin-pdfMm(62);
 const rightEnd=pageWidth-margin-pdfMm(4);

 pdfDrawCentered(page,bold,"Aufgestellt:",pageWidth/2,signatureY+pdfMm(5)-2,7.5);
 page.drawText("für den Auftragnehmer:",{x:leftStart,y:signatureY-2,size:7.5,font});
 pdfDrawCentered(page,font,"Datum:",(dateStart+dateEnd)/2,signatureY-2,7.5);
 pdfDrawRight(page,font,"für den Auftraggeber:",rightEnd,signatureY-2,7.5);
 page.drawLine({start:{x:leftStart,y:lineY},end:{x:leftEnd,y:lineY},thickness:.55});
 page.drawLine({start:{x:dateStart,y:lineY},end:{x:dateEnd,y:lineY},thickness:.55});
 page.drawLine({start:{x:rightStart,y:lineY},end:{x:rightEnd,y:lineY},thickness:.55});

 const dateValue=$("digitalDate").value||"";
 if(dateValue){
  const [year,month,day]=dateValue.split("-");
  const formatted=day&&month&&year?`${day}.${month}.${year}`:dateValue;
  pdfDrawCentered(page,font,formatted,(dateStart+dateEnd)/2,lineY+3,8);
 }

 page.drawText(`BauAufmaß | ID: ${id}`,{
  x:margin,y:pdfMm(6)-1,size:6.8,font,color:PDFLib.rgb(.47,.47,.47)
 });

 return new Blob([await pdf.save()],{type:"application/pdf"});
}
async function saveDigitalToCloud(){const btn=$("digitalCloudSave");btn.disabled=true;setStatus($("digitalStatus"),"Digitales Aufmaß wird übertragen …");try{if(!state.cloudBaseFolder)await findProjectIndex();const id=digitalExternalId(),uploadId=createUploadId(),base=`${state.cloudBaseFolder}/uploads/${uploadId}`;await ensureFolderPath(base);const pdf=await digitalPdfBlob(),preview=await digitalPreviewBlob(),draft=draftObject();const jsonBlob=new Blob([JSON.stringify(draft,null,2)],{type:"application/json"});await uploadFile(`${base}/${id}.pdf`,pdf);await uploadFile(`${base}/${id}.json`,jsonBlob);await uploadFile(`${base}/${id}.preview.png`,preview);const metadata={version:3,upload_id:uploadId,uploaded_at:new Date().toISOString(),project_id:state.selectedProject.id,project_number:state.selectedProject.project_number,project_name:state.selectedProject.name,position_id:state.selectedPosition.id,position_ordinal:state.selectedPosition.ordinal,position_short_text:state.selectedPosition.short_text,category:"aufmass",original_filename:`${id}.json`,stored_filename:`${id}.pdf`,sheet_no:Number($("digitalSheetNumber").value),external_id:id,digital_draft_filename:`${id}.json`,preview_filename:`${id}.preview.png`,client_source:"BauManager Mobile v2.9.1"};await uploadFile(`${base}/metadata.json`,new Blob([JSON.stringify(metadata,null,2)],{type:"application/json"}));autoSaveDigitalDraft();setStatus($("digitalStatus"),`Aufmaßblatt ${id} erfolgreich übertragen.`,false,true);const current=state.selectedPosition.measurement_sheets||[];if(!current.some(item=>Number(item.sheet_no)===Number($("digitalSheetNumber").value))){current.push({sheet_no:Number($("digitalSheetNumber").value),status:"übertragen"});state.selectedPosition.measurement_sheets=current;}}catch(e){setStatus($("digitalStatus"),e.message,true);}finally{btn.disabled=false;}}

$("loginButton").onclick=login;$("logoutButton").onclick=logout;$("refreshButton").onclick=loadProjects;
$("homeButton").onclick=()=>state.account&&showView("dashboardView");
$("accountButton").onclick=()=>{state.previousView=views.find(v=>!$(v).classList.contains("hidden"))||"dashboardView";showView("accountView")};
$("backFromAccount").onclick=()=>showView(state.previousView);
$("backToProjects").onclick=()=>showView("dashboardView");$("backToPositions").onclick=()=>showView("positionView");$("backFromDigital").onclick=()=>showView("positionDetailView");$("digitalMeasurementButton").onclick=openDigitalMeasurement;$("digitalSheetNumber").onchange=updateDigitalId;$("digitalDate").onchange=()=>{
 autoSaveDigitalDraft();
 const value=$("digitalDate").value||"";
 const parts=value.split("-");
 $("digitalDateDisplay").textContent=parts.length===3?`${parts[2]}.${parts[1]}.${parts[0]}`:value;
};$("digitalPen").onclick=()=>activateDigitalTool("pen");$("digitalEraser").onclick=()=>activateDigitalTool("eraser");$("digitalUndo").onclick=()=>{const s=digitalState.strokes.pop();if(s)digitalState.redo.push(s);redrawDigitalCanvas();autoSaveDigitalDraft()};$("digitalRedo").onclick=()=>{const s=digitalState.redo.pop();if(s)digitalState.strokes.push(s);redrawDigitalCanvas();autoSaveDigitalDraft()};$("digitalClear").onclick=()=>{if(confirm("Skizze wirklich löschen?")){digitalState.strokes=[];digitalState.redo=[];redrawDigitalCanvas();autoSaveDigitalDraft()}};$("digitalLineWidth").oninput=updateDigitalWidth;
$("digitalColor").onchange=()=>digitalState.strokeColor=$("digitalColor").value;
$("digitalText").onclick=()=>activateDigitalTool("text");
$("digitalFormula").onclick=()=>activateDigitalTool("formula");
$("formulaType").onchange=updateFormulaFields;
$("formulaCancel").onclick=closeFormulaModal;
$("formulaInsert").onclick=insertFormulaFromModal;
$("formulaModalBackdrop").onclick=closeFormulaModal;

$("digitalPan").onclick=()=>setDigitalPanMode(!digitalState.panMode);
$("digitalZoomIn").onclick=()=>setDigitalZoom(digitalState.zoom*1.25);
$("digitalZoomOut").onclick=()=>setDigitalZoom(digitalState.zoom/1.25);
$("digitalZoomReset").onclick=resetDigitalView;$("digitalLocalSave").onclick=autoSaveDigitalDraft;$("digitalCloudSave").onclick=saveDigitalToCloud;

document.querySelectorAll(".mobile-editor-dock button[data-panel]").forEach(button=>button.onclick=()=>toggleMobileEditorPanel(button.dataset.panel));
$("mobileDraftSave").onclick=()=>{$("digitalLocalSave").click();closeMobileEditorPanels();};
$("mobilePdfUpload").onclick=()=>{$("digitalCloudSave").click();};
$("mobileLineWidth").oninput=()=>{$("digitalLineWidth").value=$("mobileLineWidth").value;updateDigitalWidth();};
$("mobileColor").onchange=()=>{$("digitalColor").value=$("mobileColor").value;digitalState.strokeColor=$("mobileColor").value;};
$("closeUploadPanel").onclick=closeUploadPanel;
$("cameraInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("libraryInput").onchange=e=>{addFiles(e.target.files);e.target.value=""};$("uploadButton").onclick=uploadSelectedFiles;$("sheetNumber").onchange=updateIdPreview;
document.querySelectorAll(".action-card[data-category]").forEach(b=>b.onclick=()=>openUploadPanel(b.dataset.category));
$("projectSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase();renderProjects(state.projects.filter(p=>`${p.project_number||""} ${p.name||""}`.toLowerCase().includes(q)))};
$("positionSearch").oninput=e=>{const q=e.target.value.trim().toLowerCase(),positions=state.selectedProject?.positions||[];renderPositions(positions.filter(p=>`${p.ordinal||""} ${p.short_text||""} ${titleKey(p)}`.toLowerCase().includes(q)))};

window.addEventListener("resize",()=>{
 const view=$("digitalMeasurementView");
 if(view&&!view.classList.contains("hidden")){
  clearTimeout(window.__bauManagerResizeTimer);
  window.__bauManagerResizeTimer=setTimeout(()=>{resizeDrawingCanvas();resetDigitalView();},140);
 }
});

if("serviceWorker"in navigator)navigator.serviceWorker.register("sw.js").catch(()=>{});
initialize().catch(e=>{showView("loginView");setStatus($("loginStatus"),e.message,true)});
})();
