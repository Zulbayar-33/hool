const TOTAL_SHEET="Total";
const TOTAL_HISTORY_SHEET="Total_History";
const TOTAL_HEADERS=["№","Овог","Нэр","Албан тушаал","FitFood нийт","ServiceFood нийт","Hool нийт","Өглөөний цай нийт","Нийт хоол","Нийт үнэ","Ажилтан зөвшөөрөл","Зөвшөөрсөн огноо"];

function createTotalSheet(){
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(TOTAL_SHEET);if(!sh)sh=ss.insertSheet(TOTAL_SHEET);
  const approvals=readApprovalMap_(sh),period=getCurrentPeriod_(),out=buildTotalRows_(period,approvals);
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
  Object.keys(people).sort((a,b)=>String(people[a].no).localeCompare(String(people[b].no),"mn",{numeric:true})).forEach(k=>{const p=people[k],count=p.fit+p.hool+p.breakfast,price=(p.fit+p.hool)*15000+p.service+p.breakfast*12000,s=approvals[String(p.no).trim()]||{};out.push([p.no,p.ovog,p.ner,p.job,p.fit,p.service,p.hool,p.breakfast,count,price,s.status||"",s.date||""]);});
  return out;
}

function writeTotalSheet_(sh,out){
  sh.clear();sh.getRange(1,1,out.length,out[0].length).setValues(out);styleHeader_(sh,out[0].length);
  if(out.length>1){sh.getRange(2,1,out.length-1,4).setNumberFormat("@");sh.getRange(2,5,out.length-1,5).setNumberFormat("0");sh.getRange(2,10,out.length-1,1).setNumberFormat("#,##0₮");sh.getRange(2,11,out.length-1,1).setNumberFormat("@");sh.getRange(2,12,out.length-1,1).setNumberFormat("yyyy-MM-dd HH:mm");}
  sh.getRange(1,1,out.length,out[0].length).setVerticalAlignment("middle");sh.getRange(1,1,1,out[0].length).setWrap(true);sh.setColumnWidth(1,65);sh.setColumnWidths(2,2,120);sh.setColumnWidth(4,220);sh.setColumnWidths(5,5,115);sh.setColumnWidth(10,150);sh.setColumnWidths(11,2,165);
}

function readApprovalMap_(sh){
  const map={};if(!sh||sh.getLastRow()<2)return map;const d=sh.getDataRange().getValues(),h=d[0]||[],statusCol=h.indexOf("Ажилтан зөвшөөрөл"),dateCol=h.indexOf("Зөвшөөрсөн огноо");
  for(let i=1;i<d.length;i++){const no=String(d[i][0]||"").trim(),status=statusCol>=0?String(d[i][statusCol]||"").trim():"";if(no&&status)map[no]={status:status,date:dateCol>=0?d[i][dateCol]:""};}
  return map;
}

function readHistoryApprovalMap_(periodKey){
  const sh=getSpreadsheet_().getSheetByName(TOTAL_HISTORY_SHEET),map={};if(!sh||sh.getLastRow()<2)return map;const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++)if(String(d[i][0])===periodKey){const no=String(d[i][3]||"").trim(),status=String(d[i][13]||"").trim();if(no&&status)map[no]={status:status,date:d[i][14]||""};}
  return map;
}

function syncTotalHistory_(forcePrevious,currentData){
  const ss=getSpreadsheet_();let sh=ss.getSheetByName(TOTAL_HISTORY_SHEET);if(!sh)sh=ss.insertSheet(TOTAL_HISTORY_SHEET);
  const headers=["Хугацааны түлхүүр","Эхлэх огноо","Дуусах огноо"].concat(TOTAL_HEADERS),existing=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[],current=currentData&&currentData.period?currentData.period:getCurrentPeriod_(),previous=getPeriodByOffset_(-1),hasPrevious=existing.some(r=>String(r[0])===previous.key),periods=[current];
  if(forcePrevious||!hasPrevious)periods.push(previous);
  const rebuildKeys={};periods.forEach(p=>rebuildKeys[p.key]=true);
  const kept=existing.filter(r=>!rebuildKeys[String(r[0])]);
  periods.forEach(period=>{const approvals=readHistoryApprovalMap_(period.key);if(period.offset===0)Object.assign(approvals,readApprovalMap_(ss.getSheetByName(TOTAL_SHEET)));const rows=currentData&&currentData.period&&currentData.period.key===period.key?currentData.rows:buildTotalRows_(period,approvals);for(let i=1;i<rows.length;i++)kept.push([period.key,period.start,period.end].concat(rows[i]));});
  kept.sort((a,b)=>String(b[0]).localeCompare(String(a[0]))||String(a[3]).localeCompare(String(b[3]),"mn",{numeric:true}));
  sh.clear();sh.getRange(1,1,1,headers.length).setValues([headers]);if(kept.length)sh.getRange(2,1,kept.length,headers.length).setValues(kept);styleHeader_(sh,headers.length);sh.setFrozenRows(1);sh.getRange(2,2,Math.max(kept.length,1),2).setNumberFormat("yyyy-MM-dd");if(kept.length){sh.getRange(2,8,kept.length,5).setNumberFormat("0");sh.getRange(2,13,kept.length,1).setNumberFormat("#,##0₮");sh.getRange(2,15,kept.length,1).setNumberFormat("yyyy-MM-dd HH:mm");}sh.autoResizeColumns(1,headers.length);
  return{success:true,message:"Total_History шинэчлэгдлээ",periods:periods.map(p=>p.label)};
}

function ensureTotalHistory_(){const sh=getSpreadsheet_().getSheetByName(TOTAL_HISTORY_SHEET);if(!sh||sh.getLastRow()<2)syncTotalHistory_(true);}
function restoreTotalHistory(){createTotalSheet();return syncTotalHistory_(true);}

function getMyTotal(user){
  ensureTotalHistory_();const period=getPeriodByOffset_(user.periodOffset),approvals=readHistoryApprovalMap_(period.key),out=buildTotalRows_(period,approvals),no=String(user.no||"").trim();
  for(let i=1;i<out.length;i++)if(String(out[i][0]).trim()===no)return{success:true,data:{no:out[i][0],ovog:out[i][1],ner:out[i][2],job:out[i][3],fit:Number(out[i][4])||0,service:Number(out[i][5])||0,hool:Number(out[i][6])||0,breakfast:Number(out[i][7])||0,total:Number(out[i][8])||0,price:Number(out[i][9])||0,status:String(out[i][10]||""),approvedDate:out[i][11]||"",period:period.label,periodOffset:period.offset,canApprove:period.offset===0}};
  return{success:false,message:period.label+" хугацаанд таны нийт дүн олдсонгүй"};
}

function userApproveTotal(user){
  const period=getPeriodByOffset_(user.periodOffset);if(period.offset!==0)return{success:false,message:"Өмнөх хугацааны дүнг дахин зөвшөөрөх боломжгүй"};
  createTotalSheet();const sh=getSpreadsheet_().getSheetByName(TOTAL_SHEET),d=sh.getDataRange().getDisplayValues(),no=String(user.no||"").trim();
  for(let i=1;i<d.length;i++)if(String(d[i][0]).trim()===no){sh.getRange(i+1,11).setValue("Зөвшөөрсөн");sh.getRange(i+1,12).setValue(new Date());syncTotalHistory_(false);writeLog_("TOTAL_APPROVED",user,"Нийт дүнгээ зөвшөөрсөн • "+period.label,TOTAL_SHEET,i+1,11);return{success:true,message:"Та нийт дүнгээ зөвшөөрлөө"};}
  return{success:false,message:"Таны мөр Total sheet дээр олдсонгүй"};
}
