const TOTAL_SHEET="Total";
const TOTAL_HISTORY_SHEET="Total_History";
const TOTAL_HEADERS=["№","Овог","Нэр","Албан тушаал","FitFood нийт","ServiceFood нийт","Hool нийт","Өглөөний цай нийт","Нийт хоол","Нийт үнэ","Ажилтан зөвшөөрөл","Зөвшөөрсөн огноо","Тайлбар"];
// Өмнөх хугацааны дүнг шинэ хугацааны эхний хэдэн өдөр зөвшөөрөх/татгалзах боломжтой
const APPROVAL_GRACE_DAYS=5;

function createTotalSheet(){
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(TOTAL_SHEET);if(!sh)sh=ss.insertSheet(TOTAL_SHEET);
  const period=getCurrentPeriod_(),approvals=Object.assign(readHistoryApprovalMap_(period.key),readApprovalMap_(sh,period)),out=buildTotalRows_(period,approvals);
  writeTotalSheet_(sh,out);
  syncTotalHistory_(false,{period:period,rows:out});
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
  const headers=["Хугацааны түлхүүр","Эхлэх огноо","Дуусах огноо"].concat(TOTAL_HEADERS),existing=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],current=currentData&&currentData.period?currentData.period:getCurrentPeriod_(),previous=getPeriodByOffset_(-1),hasPrevious=existing.some(r=>histKey_(r[0])===previous.key),periods=[current];
  if(forcePrevious||!hasPrevious)periods.push(previous);
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

// Нэг удаа ажиллуулна: Total_History-ээс одоогийн болон өмнөх хугацааг тусдаа sheet-д харуулна.
// Томьёогоор ажилладаг тул шинэ бүртгэл, зөвшөөрөл, сар солигдолт автоматаар тусна.
function createPeriodViewSheets(){
  const ss=getSpreadsheet_(),n=TOTAL_HEADERS.length,start='IF(DAY(TODAY())>=26,DATE(YEAR(TODAY()),MONTH(TODAY()),26),DATE(YEAR(TODAY()),MONTH(TODAY())-1,26))';
  [["Одоогийн хугацаа",0],["Өмнөх хугацаа",-1]].forEach(([name,offset])=>{
    let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);sh.clear();
    const s="EDATE("+start+","+offset+")";
    sh.getRange(1,1,1,n).merge();sh.getRange(1,1).setFormula('="'+name+': "&TEXT('+s+',"yyyy.mm.dd")&" – "&TEXT(EDATE('+s+',1)-1,"yyyy.mm.dd")').setFontSize(13).setFontWeight("bold").setFontColor("#1d4ed8").setBackground("#eff6ff").setHorizontalAlignment("center");
    sh.getRange(2,1,1,n).setValues([TOTAL_HEADERS]).setFontWeight("bold").setBackground("#17233c").setFontColor("white").setHorizontalAlignment("center").setWrap(true);
    sh.getRange(3,1).setFormula('=IFERROR(SORT(FILTER('+TOTAL_HISTORY_SHEET+'!D2:P,TEXT('+TOTAL_HISTORY_SHEET+'!A2:A,"yyyy-mm-dd")=TEXT('+s+',"yyyy-mm-dd")),1,TRUE),"Энэ хугацааны мэдээлэл алга")');
    sh.getRange("J3:J").setNumberFormat("#,##0₮");sh.getRange("L3:L").setNumberFormat("yyyy-MM-dd HH:mm");sh.setFrozenRows(2);
    sh.setColumnWidth(1,65);sh.setColumnWidths(2,2,120);sh.setColumnWidth(4,220);sh.setColumnWidths(5,5,115);sh.setColumnWidth(10,150);sh.setColumnWidths(11,2,165);sh.setColumnWidth(13,260);
  });
  Logger.log("Spreadsheet timezone: "+ss.getSpreadsheetTimeZone());
  return true;
}
