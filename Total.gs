const TOTAL_SHEET="Total";
const TOTAL_HISTORY_SHEET="Total_History";
const TOTAL_HEADERS=["№","Овог","Нэр","Албан тушаал","FitFood нийт","ServiceFood нийт","Hool нийт","Өглөөний цай нийт","Нийт хоол","Нийт үнэ","Ажилтан зөвшөөрөл","Зөвшөөрсөн огноо","Тайлбар"];
// Өмнөх хугацааны дүнг шинэ хугацааны эхний хэдэн өдөр зөвшөөрөх/татгалзах боломжтой
const APPROVAL_GRACE_DAYS=5;
// "Одоогийн хугацаа", "Өмнөх хугацаа", Total sheet дээр нийт тоог гараар засахад зөрүүг Залруулга sheet-д хадгалж нийт дүнд нэмнэ
const ADJ_SHEET="Залруулга";
const VIEW_SHEETS={"Одоогийн хугацаа":0,"Өмнөх хугацаа":-1};
const ADJ_FIELDS={5:"fit",6:"service",7:"hool",8:"breakfast"};

function createTotalSheet(){
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(TOTAL_SHEET);if(!sh)sh=ss.insertSheet(TOTAL_SHEET);
  const period=getCurrentPeriod_(),approvals=Object.assign(readHistoryApprovalMap_(period.key),readApprovalMap_(sh,period)),out=buildTotalRows_(period,approvals);
  writeTotalSheet_(sh,out);
  syncTotalHistory_(false,{period:period,rows:out});
  writePeriodViews_();
  return true;
}

function buildTotalRows_(period,approvals){
  const ss=getSpreadsheet_(),names=[FIT_SHEET,SERVICE_SHEET,HOOL_SHEET,BREAKFAST_SHEET],people={};
  names.forEach(name=>{const sh=ss.getSheetByName(name);if(!sh)return;const d=sh.getDataRange().getValues();if(d.length<2)return;const tc=d[0].indexOf("Нийт");if(tc<0)return;
    for(let r=1;r<d.length;r++){if(String(d[r][0]).trim()==="Нийт"||(!d[r][0]&&!d[r][2]))continue;const k=String(d[r][0]).trim()||String(d[r][1])+String(d[r][2]);if(!people[k])people[k]={no:d[r][0],ovog:d[r][1],ner:d[r][2],job:d[r][3],fit:0,service:0,hool:0,breakfast:0};let sum=0;
      for(let c=4;c<tc;c++){if(!isDateInPeriod_(d[0][c],period))continue;const n=Number(d[r][c]);if(Number.isFinite(n))sum+=n;}
      if(name===FIT_SHEET)people[k].fit+=sum;if(name===SERVICE_SHEET)people[k].service+=sum;if(name===HOOL_SHEET)people[k].hool+=sum;if(name===BREAKFAST_SHEET)people[k].breakfast+=sum;
    }
  });
  // Гараар хийсэн залруулга (Залруулга sheet); хоолны sheet-д мөргүй ажилтныг Employees-ээс нэмнэ
  const adj=readAdjustments_(period.key);
  if(Object.keys(adj).length){const emps={};getAdminEmployees_().forEach(x=>emps[x.no]=x);
    Object.keys(adj).forEach(no=>{if(!people[no]){const x=emps[no]||{};people[no]={no:no,ovog:x.ovog||"",ner:x.ner||"",job:x.job||"",fit:0,service:0,hool:0,breakfast:0};}["fit","service","hool","breakfast"].forEach(f=>people[no][f]+=adj[no][f]);});}
  const out=[TOTAL_HEADERS.slice()];
  Object.keys(people).sort((a,b)=>String(people[a].no).localeCompare(String(people[b].no),"mn",{numeric:true})).forEach(k=>{const p=people[k],count=p.fit+p.hool+p.breakfast,price=(p.fit+p.hool)*15000+p.service+p.breakfast*12000,s=approvals[String(p.no).trim()]||{};out.push([p.no,p.ovog,p.ner,p.job,p.fit,p.service,p.hool,p.breakfast,count,price,s.status||"",s.date||"",s.reason||""]);});
  return out;
}

function writeTotalSheet_(sh,out){
  sh.clear();sh.getRange(1,1,out.length,out[0].length).setValues(out);styleHeader_(sh,out[0].length);
  if(out.length>1){sh.getRange(2,1,out.length-1,4).setNumberFormat("@");sh.getRange(2,5,out.length-1,5).setNumberFormat("0");sh.getRange(2,10,out.length-1,1).setNumberFormat("#,##0₮");sh.getRange(2,11,out.length-1,1).setNumberFormat("@");sh.getRange(2,12,out.length-1,1).setNumberFormat("yyyy-MM-dd HH:mm");}
  sh.getRange(1,1,out.length,out[0].length).setVerticalAlignment("middle");sh.getRange(1,1,1,out[0].length).setWrap(true);sh.setColumnWidth(1,65);sh.setColumnWidths(2,2,120);sh.setColumnWidth(4,220);sh.setColumnWidths(5,5,115);sh.setColumnWidth(10,150);sh.setColumnWidths(11,2,165);
}

function readApprovalMap_(sh,period){
  const map={};if(!sh||sh.getLastRow()<2)return map;const d=sh.getDataRange().getValues(),h=d[0]||[],statusCol=h.indexOf("Ажилтан зөвшөөрөл"),dateCol=h.indexOf("Зөвшөөрсөн огноо"),reasonCol=h.indexOf("Тайлбар");
  for(let i=1;i<d.length;i++){const no=String(d[i][0]||"").trim(),status=statusCol>=0?String(d[i][statusCol]||"").trim():"",date=dateCol>=0?d[i][dateCol]:"";if(no&&status&&(!period||approvalInPeriod_(date,period.key)))map[no]={status:status,date:date,reason:reasonCol>=0?String(d[i][reasonCol]||""):""};}
  return map;
}
function readHistoryApprovalMap_(periodKey){
  const sh=getSpreadsheet_().getSheetByName(TOTAL_HISTORY_SHEET),map={};if(!sh||sh.getLastRow()<2)return map;const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++)if(histKey_(d[i][0])===periodKey){const no=String(d[i][3]||"").trim(),status=String(d[i][13]||"").trim();if(no&&status)map[no]={status:status,date:d[i][14]||"",reason:String(d[i][15]||"")};}
  return map;
}
function syncTotalHistory_(forcePrevious,currentData){
  const ss=getSpreadsheet_();let sh=ss.getSheetByName(TOTAL_HISTORY_SHEET);if(!sh)sh=ss.insertSheet(TOTAL_HISTORY_SHEET);
  const headers=["Хугацааны түлхүүр","Эхлэх огноо","Дуусах огноо"].concat(TOTAL_HEADERS),existing=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],current=currentData&&currentData.period?currentData.period:getCurrentPeriod_(),previous=getPeriodByOffset_(-1),periods=[current,previous];
  // Өмнөх хугацааг ч үргэлж дахин тооцно: өмнөх сарын баганыг гараар засахад Total_History / "Өмнөх хугацаа" шинэчлэгдэнэ
  const rebuildKeys={};periods.forEach(p=>rebuildKeys[p.key]=true);
  const kept=existing.filter(r=>!rebuildKeys[histKey_(r[0])]);
  periods.forEach(period=>{const approvals=readHistoryApprovalMap_(period.key);if(period.offset===0)Object.assign(approvals,readApprovalMap_(ss.getSheetByName(TOTAL_SHEET),period));const rows=currentData&&currentData.period&&currentData.period.key===period.key?currentData.rows:buildTotalRows_(period,approvals);for(let i=1;i<rows.length;i++)kept.push([period.key,period.start,period.end].concat(rows[i]));});
  dedupeHistoryRows_(kept);
  kept.sort((a,b)=>String(b[0]).localeCompare(String(a[0]))||String(a[3]).localeCompare(String(b[3]),"mn",{numeric:true}));
  sh.clear();sh.getRange(1,1,1,headers.length).setValues([headers]);if(kept.length){sh.getRange(2,1,kept.length,1).setNumberFormat("@");sh.getRange(2,1,kept.length,headers.length).setValues(kept);}styleHeader_(sh,headers.length);sh.setFrozenRows(1);sh.getRange(2,2,Math.max(kept.length,1),2).setNumberFormat("yyyy-MM-dd");if(kept.length){sh.getRange(2,8,kept.length,5).setNumberFormat("0");sh.getRange(2,13,kept.length,1).setNumberFormat("#,##0₮");sh.getRange(2,15,kept.length,1).setNumberFormat("yyyy-MM-dd HH:mm");}sh.autoResizeColumns(1,headers.length);
  return{success:true,message:"Total_History шинэчлэгдлээ",periods:periods.map(p=>p.label)};
}

function ensureTotalHistory_(){const sh=getSpreadsheet_().getSheetByName(TOTAL_HISTORY_SHEET);if(!sh||sh.getLastRow()<2)syncTotalHistory_(true);}
function restoreTotalHistory(){createTotalSheet();return syncTotalHistory_(true);}

function getMyTotal(user){
  ensureTotalHistory_();const period=getPeriodByOffset_(user.periodOffset),approvals=Object.assign(readHistoryApprovalMap_(period.key),period.offset===0?readApprovalMap_(getSpreadsheet_().getSheetByName(TOTAL_SHEET),period):{}),out=buildTotalRows_(period,approvals),no=String(user.no||"").trim();
  for(let i=1;i<out.length;i++)if(String(out[i][0]).trim()===no)return{success:true,data:{no:out[i][0],ovog:out[i][1],ner:out[i][2],job:out[i][3],fit:Number(out[i][4])||0,service:Number(out[i][5])||0,hool:Number(out[i][6])||0,breakfast:Number(out[i][7])||0,total:Number(out[i][8])||0,price:Number(out[i][9])||0,status:String(out[i][10]||""),approvedDate:out[i][11]||"",reason:String(out[i][12]||""),period:period.label,periodOffset:period.offset,canApprove:canDecide_(period),deadline:period.offset===0?"":decideDeadlineLabel_(period)}};
  return{success:false,message:period.label+" хугацаанд таны нийт дүн олдсонгүй"};
}

function userApproveTotal(user){
  const period=getPeriodByOffset_(user.periodOffset),reject=String(user.decision||"")==="reject",reason=String(user.reason||"").trim().slice(0,300);
  if(!canDecide_(period))return{success:false,message:period.label+" хугацааны дүнд хариулах хугацаа дууссан"};
  if(reject&&!reason)return{success:false,message:"Зөвшөөрөөгүй шалтгаанаа бичнэ үү"};
  const status=reject?"Зөвшөөрөөгүй":"Зөвшөөрсөн";
  // Бүх sheet-ийг дахин бүтээхгүй: зөвхөн тухайн хүний мөрийг Total (одоогийн хугацаа) болон Total_History дээр шинэчилнэ
  const lock=LockService.getDocumentLock();if(lock)lock.waitLock(20000);
  try{
    const ss=getSpreadsheet_(),no=String(user.no||"").trim(),now=new Date();let row="";
    if(period.offset===0){let sh=ss.getSheetByName(TOTAL_SHEET);row=findTotalUserRow_(sh,no);
      if(!row){createTotalSheet();sh=ss.getSheetByName(TOTAL_SHEET);row=findTotalUserRow_(sh,no);}
      if(!row)return{success:false,message:"Таны мөр Total sheet дээр олдсонгүй"};
      sh.getRange(row,11,1,3).setValues([[status,now,reject?reason:""]]);}
    const hs=ss.getSheetByName(TOTAL_HISTORY_SHEET);let saved=false;
    if(hs&&hs.getLastRow()>=2){const v=hs.getRange(2,1,hs.getLastRow()-1,4).getValues();for(let i=0;i<v.length;i++)if(histKey_(v[i][0])===period.key&&String(v[i][3]).trim()===no){hs.getRange(i+2,14,1,3).setValues([[status,now,reject?reason:""]]);saved=true;}}
    if(!saved){if(period.offset!==0)return{success:false,message:period.label+" хугацаанд таны мөр олдсонгүй"};PropertiesService.getScriptProperties().setProperty("REPORTS_DIRTY","1");}
    writeLog_(reject?"TOTAL_REJECTED":"TOTAL_APPROVED",user,(reject?"Нийт дүнг зөвшөөрөөгүй: "+reason:"Нийт дүнгээ зөвшөөрсөн")+" • "+period.label,period.offset===0?TOTAL_SHEET:TOTAL_HISTORY_SHEET,row,11);
    return{success:true,message:reject?"Таны \"Зөвшөөрөөгүй\" хариу болон шалтгаан хадгалагдлаа":"Та нийт дүнгээ зөвшөөрлөө"};
  }finally{if(lock)lock.releaseLock();}
}
function findTotalUserRow_(sh,no){if(!sh||sh.getLastRow()<2)return null;const v=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();for(let i=0;i<v.length;i++)if(String(v[i][0]).trim()===no)return i+2;return null;}

// Total_History-ийн хугацааны түлхүүр Sheets дээр огноо болж хадгалагдсан ч "yyyy-MM-dd" текст болгож харьцуулна
function histKey_(v){return v instanceof Date&&!isNaN(v.getTime())?Utilities.formatDate(v,APP_TZ,"yyyy-MM-dd"):String(v||"").trim();}
// Огноо аль хугацаанд (26-аас дараа сарын 25) багтахыг түлхүүрээр буцаана
function periodKeyForDate_(d){const p=Utilities.formatDate(d,APP_TZ,"yyyy-MM-dd").split("-").map(Number);let y=p[0],m=p[1];if(p[2]<26){m-=1;if(m<1){m=12;y-=1;}}return y+"-"+String(m).padStart(2,"0")+"-26";}
// Сар солигдоход өмнөх сарын зөвшөөрөл шинэ сард шилжихгүйн тулд зөвшөөрлийн огноо тухайн хугацаанд багтах ёстой
function approvalInPeriod_(date,periodKey){const d=date instanceof Date?date:new Date(String(date||"").replace(" ","T"));return !isNaN(d.getTime())&&periodKeyForDate_(d)===periodKey;}
// Нэг хугацаанд нэг ажилтан нэг л мөртэй байна; зөвшөөрөлтэй мөрийг үлдээнэ
function dedupeHistoryRows_(rows){const best={};rows.forEach((r,i)=>{r[0]=histKey_(r[0]);const k=r[0]+"|"+String(r[3]).trim();if(!(k in best)||(!String(rows[best[k]][13]||"").trim()&&String(r[13]||"").trim()))best[k]=i;});const keep=new Set(Object.values(best));for(let i=rows.length-1;i>=0;i--)if(!keep.has(i))rows.splice(i,1);return rows;}
// Одоогийн хугацаа үргэлж нээлттэй; өмнөх хугацаа нь дуусснаас хойш APPROVAL_GRACE_DAYS өдөр нээлттэй
function decideCutoff_(period){return new Date(period.end.getFullYear(),period.end.getMonth(),period.end.getDate()+1+APPROVAL_GRACE_DAYS);}
function canDecide_(period){if(period.offset===0)return true;if(period.offset!==-1)return false;const now=new Date(Utilities.formatDate(new Date(),APP_TZ,"yyyy-MM-dd'T'HH:mm:ss"));return now<decideCutoff_(period);}
function decideDeadlineLabel_(period){const c=decideCutoff_(period);return Utilities.formatDate(new Date(c.getFullYear(),c.getMonth(),c.getDate()-1),APP_TZ,"yyyy.MM.dd");}

function readAdjustments_(key){
  const sh=getSpreadsheet_().getSheetByName(ADJ_SHEET),map={};if(!sh||sh.getLastRow()<2)return map;
  sh.getRange(2,1,sh.getLastRow()-1,4).getValues().forEach(r=>{const no=String(r[1]).trim(),f=String(r[2]).trim(),n=Number(r[3]);if(histKey_(r[0])!==key||!no||!Number.isFinite(n)||["fit","service","hool","breakfast"].indexOf(f)<0)return;(map[no]=map[no]||{fit:0,service:0,hool:0,breakfast:0})[f]+=n;});
  return map;
}
// Total_History-ээс одоогийн / өмнөх хугацааг утгаар нь бичнэ (томьёо биш тул гараар засаж болно)
function writePeriodViews_(){
  const ss=getSpreadsheet_(),hs=ss.getSheetByName(TOTAL_HISTORY_SHEET),n=TOTAL_HEADERS.length,d=hs&&hs.getLastRow()>=2?hs.getRange(2,1,hs.getLastRow()-1,3+n).getValues():[];
  Object.keys(VIEW_SHEETS).forEach(name=>{
    const p=getPeriodByOffset_(VIEW_SHEETS[name]),rows=d.filter(r=>histKey_(r[0])===p.key).map(r=>r.slice(3)).sort((a,b)=>String(a[0]).localeCompare(String(b[0]),"mn",{numeric:true}));
    let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);sh.clear();sh.getRange(3,11,Math.max(sh.getMaxRows()-2,1),1).clearDataValidations();
    sh.getRange(1,1,1,n).merge();sh.getRange(1,1).setValue(name+": "+p.label+"  •  E–H, K, M баганыг засаж болно").setFontSize(13).setFontWeight("bold").setFontColor("#1d4ed8").setBackground("#eff6ff").setHorizontalAlignment("center");
    sh.getRange(2,1,1,n).setValues([TOTAL_HEADERS]).setFontWeight("bold").setBackground("#17233c").setFontColor("white").setHorizontalAlignment("center").setWrap(true);sh.setFrozenRows(2);
    if(!rows.length){sh.getRange(3,1).setValue("Энэ хугацааны мэдээлэл алга");return;}
    sh.getRange(3,1,rows.length,n).setValues(rows);sh.getRange(3,5,rows.length,5).setNumberFormat("0");sh.getRange(3,10,rows.length,1).setNumberFormat("#,##0₮");sh.getRange(3,12,rows.length,1).setNumberFormat("yyyy-MM-dd HH:mm");
    sh.getRange(3,11,rows.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Зөвшөөрсөн","Зөвшөөрөөгүй"],true).setAllowInvalid(true).build());
  });
}
// onEdit-ээс: Total / хугацааны sheet дээрх нэг нүдний засвар. E–H тоо → Залруулга, K төлөв / M тайлбар → Total_History (ба Total)
function handleTotalEdit_(e,sh,name){
  const hr=name===TOTAL_SHEET?1:2,r=e.range.getRow(),c=e.range.getColumn();
  if(r<=hr||e.range.getNumRows()>1||e.range.getNumColumns()>1)return false;
  const no=String(sh.getRange(r,1).getDisplayValue()).trim();if(!no)return false;
  const ss=getSpreadsheet_(),period=getPeriodByOffset_(name===TOTAL_SHEET?0:VIEW_SHEETS[name]),who=(e.user&&e.user.getEmail&&e.user.getEmail())||"Гараар";
  if(ADJ_FIELDS[c]){
    const oldN=Number(e.oldValue||0),newN=Number(e.value===undefined?0:e.value);if(!Number.isFinite(oldN)||!Number.isFinite(newN)||newN===oldN)return false;
    let adj=ss.getSheetByName(ADJ_SHEET);if(!adj){adj=ss.insertSheet(ADJ_SHEET);adj.getRange(1,1,1,7).setValues([["Хугацааны түлхүүр","№","Төрөл","Залруулга","Хэн","Огноо","Тайлбар"]]);styleHeader_(adj,7);adj.getRange("A:A").setNumberFormat("@");}
    adj.appendRow([period.key,no,ADJ_FIELDS[c],newN-oldN,who,new Date(),name+" "+getColumnLetter(c)+r+": "+oldN+" → "+newN]);
    const v=sh.getRange(r,5,1,4).getValues()[0].map(x=>Number(x)||0);sh.getRange(r,9,1,2).setValues([[v[0]+v[2]+v[3],(v[0]+v[2])*15000+v[1]+v[3]*12000]]);
    writeLog_("MANUAL_TOTAL_EDIT",{no:no,ner:who},period.label+" • "+TOTAL_HEADERS[c-1]+" • "+oldN+" → "+newN,name,r,c);
    return true;
  }
  if(c===11||c===13){
    const status=String(sh.getRange(r,11).getDisplayValue()).trim(),reason=String(sh.getRange(r,13).getDisplayValue()).trim(),vals=[[status,status?new Date():"",reason]];
    const hs=ss.getSheetByName(TOTAL_HISTORY_SHEET);if(hs&&hs.getLastRow()>=2){const v=hs.getRange(2,1,hs.getLastRow()-1,4).getValues();for(let i=0;i<v.length;i++)if(histKey_(v[i][0])===period.key&&String(v[i][3]).trim()===no)hs.getRange(i+2,14,1,3).setValues(vals);}
    if(period.offset===0){const ts=ss.getSheetByName(TOTAL_SHEET),tr=findTotalUserRow_(ts,no);if(tr)ts.getRange(tr,11,1,3).setValues(vals);}
    if(name!==TOTAL_SHEET)sh.getRange(r,12).setValue(vals[0][1]);
    writeLog_("MANUAL_TOTAL_STATUS",{no:no,ner:who},period.label+" • "+(status||"Хариулаагүй")+(reason?" • "+reason:""),name,r,c);
    return true;
  }
  return false;
}

// Нэг удаа ажиллуулна: хугацааны sheet-үүдийг үүсгэж баганын өргөнийг тохируулна (агуулгыг createTotalSheet бүр шинэчилнэ)
function createPeriodViewSheets(){
  writePeriodViews_();const ss=getSpreadsheet_();
  Object.keys(VIEW_SHEETS).forEach(name=>{const sh=ss.getSheetByName(name);sh.setColumnWidth(1,65);sh.setColumnWidths(2,2,120);sh.setColumnWidth(4,220);sh.setColumnWidths(5,5,115);sh.setColumnWidth(10,150);sh.setColumnWidths(11,2,165);sh.setColumnWidth(13,260);});
  return true;
}

// Admin: тухайн хугацааны бүх ажилтны нийт дүн болон зөвшөөрлийн төлөв (Хариулаагүй = хоосон)
function adminListTotals(p){
  ensureTotalHistory_();const period=getPeriodByOffset_(p.periodOffset),approvals=Object.assign(readHistoryApprovalMap_(period.key),period.offset===0?readApprovalMap_(getSpreadsheet_().getSheetByName(TOTAL_SHEET),period):{}),out=buildTotalRows_(period,approvals);
  const fmt=d=>d instanceof Date&&!isNaN(d.getTime())?Utilities.formatDate(d,APP_TZ,"yyyy-MM-dd HH:mm"):String(d||"");
  const rows=out.slice(1).map(r=>({no:String(r[0]),ovog:String(r[1]||""),ner:String(r[2]||""),job:String(r[3]||""),total:Number(r[8])||0,price:Number(r[9])||0,status:String(r[10]||""),date:fmt(r[11]),reason:String(r[12]||"")}));
  return{success:true,period:period.label,periodOffset:period.offset,rows:rows};
}

// Admin: ажилтны нийт дүнгийн хариуг өөрчлөх. status "" бол хариуг устгаж "Хариулаагүй" болгоно (ажилтан дахин хариулж болно).
// Total (одоогийн хугацаа) болон Total_History-г хамт шинэчилж, Logs-д ADMIN_TOTAL_STATUS бичнэ.
function adminSetTotalStatus(p){
  const period=getPeriodByOffset_(p.periodOffset),no=String(p.no||"").trim(),status=String(p.status||""),reason=status==="Зөвшөөрөөгүй"?String(p.reason||"").trim().slice(0,300):"";
  if(!no)return{success:false,message:"Ажилтан сонгоно уу"};
  if(["","Зөвшөөрсөн","Зөвшөөрөөгүй"].indexOf(status)<0)return{success:false,message:"Төлөв буруу байна"};
  if(status==="Зөвшөөрөөгүй"&&!reason)return{success:false,message:"Зөвшөөрөөгүй шалтгаан бичнэ үү"};
  const lock=LockService.getDocumentLock();if(lock)lock.waitLock(20000);
  try{
    const ss=getSpreadsheet_(),vals=[[status,status?new Date():"",reason]];let old="",found=false;
    if(period.offset===0){const sh=ss.getSheetByName(TOTAL_SHEET),row=findTotalUserRow_(sh,no);if(row){old=String(sh.getRange(row,11).getValue()||"");sh.getRange(row,11,1,3).setValues(vals);found=true;}}
    const hs=ss.getSheetByName(TOTAL_HISTORY_SHEET);
    if(hs&&hs.getLastRow()>=2){const v=hs.getRange(2,1,hs.getLastRow()-1,14).getValues();for(let i=0;i<v.length;i++)if(histKey_(v[i][0])===period.key&&String(v[i][3]).trim()===no){if(!old)old=String(v[i][13]||"");hs.getRange(i+2,14,1,3).setValues(vals);found=true;}}
    if(!found)return{success:false,message:period.label+" хугацаанд ажилтны мөр олдсонгүй"};
    writeLog_("ADMIN_TOTAL_STATUS",{no:no,ner:"Admin"},period.label+" • "+(old||"Хариулаагүй")+" → "+(status||"Хариулаагүй")+(reason?" • "+reason:""),period.offset===0?TOTAL_SHEET:TOTAL_HISTORY_SHEET,"",11);
    return{success:true,message:status?'Төлөв "'+status+'" боллоо':'Хариуг устгалаа ("Хариулаагүй")'};
  }finally{if(lock)lock.releaseLock();}
}

