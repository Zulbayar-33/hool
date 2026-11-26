const GRAPHIC_SHEET="Graphic";

function createGraphicSheet(){
  const ss=getSpreadsheet_(),period=getCurrentPeriod_();let sh=ss.getSheetByName(GRAPHIC_SHEET);if(!sh)sh=ss.insertSheet(GRAPHIC_SHEET);
  const summary={FitFood:{count:0,money:0},ServiceFood:{count:0,money:0},Hool:{count:0,money:0},OglooniiTsai:{count:0,money:0}},daily={};
  const configs=[{sheet:FIT_SHEET,key:"FitFood",price:15000},{sheet:SERVICE_SHEET,key:"ServiceFood",service:true},{sheet:HOOL_SHEET,key:"Hool",price:15000},{sheet:BREAKFAST_SHEET,key:"OglooniiTsai",price:12000}];
  configs.forEach(cfg=>{const src=ss.getSheetByName(cfg.sheet);if(!src)return;const data=src.getDataRange().getValues();if(data.length<2)return;const totalCol=data[0].indexOf("Нийт");if(totalCol<0)return;
    for(let c=4;c<totalCol;c++){if(!isDateInCurrentPeriod_(data[0][c]))continue;const dateKey=graphicDateKey_(data[0][c],period);if(!daily[dateKey])daily[dateKey]={FitFood:0,ServiceFood:0,Hool:0,OglooniiTsai:0};
      for(let r=1;r<data.length;r++){if(String(data[r][0]).trim()==="Нийт")continue;const n=Number(data[r][c]);if(!Number.isFinite(n)||n<=0)continue;if(cfg.service){summary[cfg.key].count+=1;summary[cfg.key].money+=n;daily[dateKey][cfg.key]+=1;}else{summary[cfg.key].count+=n;summary[cfg.key].money+=n*cfg.price;daily[dateKey][cfg.key]+=n;}}
    }
  });
  const totalCount=summary.FitFood.count+summary.ServiceFood.count+summary.Hool.count+summary.OglooniiTsai.count,totalFood=summary.FitFood.count+summary.Hool.count+summary.OglooniiTsai.count,totalMoney=Object.keys(summary).reduce((s,k)=>s+summary[k].money,0);
  sh.getCharts().forEach(c=>sh.removeChart(c));sh.clear();
  sh.getRange("A1:H1").merge().setValue("АКУМА • ХООЛНЫ НЭГДСЭН ГРАФИК").setFontSize(18).setFontWeight("bold").setFontColor("white").setBackground("#17233c").setHorizontalAlignment("center");
  sh.getRange("A2:H2").merge().setValue("Тооцооны хугацаа: "+period.label).setFontWeight("bold").setFontColor("#1d4ed8").setBackground("#eff6ff").setHorizontalAlignment("center");
  sh.getRange("A4:D4").setValues([["Хоолны төрөл","Бүртгэлийн тоо","Нэгж үнэ","Мөнгөн дүн"]]);
  const summaryRows=[["FitFood",summary.FitFood.count,15000,summary.FitFood.money],["ServiceFood",summary.ServiceFood.count,"Үнийн дүнгээр",summary.ServiceFood.money],["Хоол",summary.Hool.count,15000,summary.Hool.money],["Өглөөний цай",summary.OglooniiTsai.count,12000,summary.OglooniiTsai.money],["Нийт",totalCount,"Нийт хоол: "+totalFood,totalMoney]];
  sh.getRange(5,1,summaryRows.length,4).setValues(summaryRows);sh.getRange("C5:C8").setNumberFormat("#,##0₮");sh.getRange("D5:D9").setNumberFormat("#,##0₮");
  const dates=Object.keys(daily).sort(),dailyRows=dates.map(d=>[d,daily[d].FitFood,daily[d].ServiceFood,daily[d].Hool,daily[d].OglooniiTsai,daily[d].FitFood+daily[d].ServiceFood+daily[d].Hool+daily[d].OglooniiTsai]);
  sh.getRange("A12:F12").setValues([["Огноо","FitFood","ServiceFood","Хоол","Өглөөний цай","Нийт бүртгэл"]]);if(dailyRows.length)sh.getRange(13,1,dailyRows.length,6).setValues(dailyRows);
  [sh.getRange("A4:D4"),sh.getRange("A12:F12")].forEach(r=>r.setFontWeight("bold").setFontColor("white").setBackground("#2563eb").setHorizontalAlignment("center"));sh.getRange("A9:D9").setFontWeight("bold").setBackground("#dbeafe");sh.setFrozenRows(2);sh.setColumnWidth(1,150);sh.setColumnWidths(2,5,125);sh.setColumnWidth(7,30);sh.setColumnWidth(8,30);
  const pie=sh.newChart().setChartType(Charts.ChartType.PIE).addRange(sh.getRange("A4:B8")).setPosition(4,6,0,0).setOption("title","Хоолны төрлүүдийн бүртгэлийн хувь").setOption("pieHole",0.42).setOption("legend",{position:"right"}).build();
  const money=sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sh.getRange("A4:A8")).addRange(sh.getRange("D4:D8")).setPosition(20,6,0,0).setOption("title","Хоолны төрлүүдийн мөнгөн дүн").setOption("legend",{position:"none"}).setOption("colors",["#2563eb"]).build();sh.insertChart(pie);sh.insertChart(money);
  if(dailyRows.length){const trend=sh.newChart().setChartType(Charts.ChartType.LINE).addRange(sh.getRange(12,1,dailyRows.length+1,6)).setPosition(20,1,0,0).setOption("title","Өдөр тутмын бүртгэлийн хөдөлгөөн").setOption("curveType","function").setOption("pointSize",4).build();sh.insertChart(trend);}
  return{success:true,period:period.label,totalFood:totalFood,totalMoney:totalMoney};
}

function graphicDateKey_(value,period){const text=String(value||"").trim();if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(text))return text;const p=normalizeDate(text).split("/").map(Number);let year=period.start.getFullYear();if(p[0]<period.start.getMonth()+1)year=period.end.getFullYear();return year+"-"+String(p[0]).padStart(2,"0")+"-"+String(p[1]).padStart(2,"0");}

// ===== total_finance_akuma: 21→20 санхүүгийн нэгтгэл (тусгай ажилтнууд). Total sheet-тэй хамт 5 минут тутам шинэчлэгдэнэ =====
const FINANCE_SHEET="total_finance_akuma";
const FINANCE_PEOPLE=["575","608","637"];
const FINANCE_START_DAY=21;
function getFinancePeriod_(offset){
  const now=new Date(Utilities.formatDate(new Date(),APP_TZ,"yyyy-MM-dd'T'HH:mm:ss"));
  const base=now.getDate()>=FINANCE_START_DAY?new Date(now.getFullYear(),now.getMonth(),FINANCE_START_DAY):new Date(now.getFullYear(),now.getMonth()-1,FINANCE_START_DAY);
  const start=new Date(base.getFullYear(),base.getMonth()+offset,FINANCE_START_DAY),end=new Date(start.getFullYear(),start.getMonth()+1,FINANCE_START_DAY-1);
  return{start:start,end:end,offset:offset,key:"fin-"+Utilities.formatDate(start,APP_TZ,"yyyy-MM-dd"),label:Utilities.formatDate(start,APP_TZ,"yyyy.MM.dd")+" – "+Utilities.formatDate(end,APP_TZ,"yyyy.MM.dd")};
}
const TEMKA_SHEET="temka",TEMKA_PEOPLE=["990"];
// total_finance_akuma болон temka (гадны хүн) хоёр sheet-ийг хамт шинэчилнэ
function createFinanceSheet(){buildSeparateSheet_(FINANCE_SHEET,FINANCE_PEOPLE,"AKUMA санхүүгийн нэгтгэл",()=>getFinancePeriod_(-1));buildSeparateSheet_(TEMKA_SHEET,TEMKA_PEOPLE,"temka нэгтгэл (26–25)",()=>getPeriodByOffset_(0));return true;}
function buildSeparateSheet_(SHEET_NAME,PEOPLE,TITLE,PERIOD_FN){
  const ss=getSpreadsheet_();let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);
  // Сүүлд дууссан 21→20 хугацаа. 21-ний өдөр болоход автоматаар дараагийн сар руу шилжинэ.
  const p=PERIOD_FN();
  // 575, 608, 637 нь Employees-ийн "Код"; хоолны sheet-үүд "№"-аар холбогддог тул Код -> № хөрвүүлнэ
  const emps={},esh=ss.getSheetByName(EMP_SHEET);
  if(esh&&esh.getLastRow()>1)esh.getRange(2,1,esh.getLastRow()-1,6).getDisplayValues().forEach(x=>{const c=String(x[1]).trim();if(c)emps[c]={no:String(x[0]).trim(),ovog:x[3],ner:x[4],job:x[5]};});
  const rows=buildTotalRowsAll_(p,{}),byNo={};for(let i=1;i<rows.length;i++)byNo[String(rows[i][0]).trim()]=rows[i];
  // Зөвшөөрөл/тайлбар: Total sheet-ээс шууд уншина
  const appr=getFinApprovals_(p.key);
  const head=["Код","№","Овог","Нэр","Албан тушаал","FitFood","ServiceFood (₮)","Hool","Өглөөний цай","Нийт хоол","Нийт үнэ","Ажилтан зөвшөөрөл","Зөвшөөрсөн огноо","Тайлбар"],W=head.length;
  const title=[TITLE+"  •  "+p.label+"  •  Шинэчилсэн: "+Utilities.formatDate(new Date(),APP_TZ,"yyyy-MM-dd HH:mm")];while(title.length<W)title.push("");
  const out=[title,head],sum=new Array(6).fill(0);
  PEOPLE.forEach(code=>{
    const e=emps[code]||{no:"",ovog:"",ner:"Employees дээр олдсонгүй",job:""},r=byNo[e.no],a=appr[e.no]||{};
    const v=r?[r[4],r[5],r[6],r[7],r[8],r[9]]:[0,0,0,0,0,0];
    out.push([code,e.no,e.ovog,e.ner,e.job].concat(v,[a.status||"",a.date?new Date(a.date):"",a.reason||""]));
    for(let k=0;k<6;k++)sum[k]+=Number(v[k])||0;
  });
  out.push(["","","","","НИЙТ"].concat(sum,["","",""]));
  sh.clear();sh.getRange(1,1,1,W).breakApart();
  sh.getRange(1,1,out.length,W).setValues(out);
  styleHeader_(sh,W);
  sh.getRange(2,1,1,W).setFontWeight("bold").setBackground("#17233c").setFontColor("white").setHorizontalAlignment("center").setWrap(true);
  sh.getRange(3,1,out.length-2,5).setNumberFormat("@");
  sh.getRange(3,6,out.length-2,5).setNumberFormat("0");
  sh.getRange(3,7,out.length-2,1).setNumberFormat("#,##0₮");
  sh.getRange(3,11,out.length-2,1).setNumberFormat("#,##0₮");
  sh.getRange(3,12,out.length-2,1).setNumberFormat("@");
  sh.getRange(3,13,out.length-2,1).setNumberFormat("yyyy-MM-dd HH:mm");
  sh.getRange(out.length,1,1,W).setFontWeight("bold").setBackground("#e8f0fe");
  sh.getRange(1,1,1,W).merge().setFontWeight("bold").setFontSize(12).setBackground("#e8f0fe").setFontColor("#17233c").setHorizontalAlignment("left").setVerticalAlignment("middle");
  sh.setRowHeight(1,30);
  sh.setColumnWidths(1,2,60);sh.setColumnWidths(3,2,120);sh.setColumnWidth(5,200);sh.setColumnWidths(6,6,115);sh.setColumnWidths(12,2,150);sh.setColumnWidth(14,160);
  sh.setFrozenRows(2);
  return true;
}
// Нэг удаа ажиллуулна: total_finance_akuma-г 5 минут тутам Total sheet-тэй тулган шинэчилнэ
function installFinanceTrigger(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==="createFinanceSheet").forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("createFinanceSheet").timeBased().everyMinutes(5).create();
}

// ===== Санхүүгийн ажилтнууд (575, 608, 637) Total-оос бүрэн тусгаарлагдсан =====
// Тэдний нийт дүн, зөвшөөрөл нь 21→20 хугацаагаар total_finance_akuma дээр; зөвшөөрлийг Script Properties (FIN_APPR) дээр хадгална.
function financeNos_(){
  const esh=getSpreadsheet_().getSheetByName(EMP_SHEET);if(!esh||esh.getLastRow()<2)return[];
  const out=[];esh.getRange(2,1,esh.getLastRow()-1,2).getDisplayValues().forEach(x=>{if(FINANCE_PEOPLE.concat(TEMKA_PEOPLE).indexOf(String(x[1]).trim())>=0)out.push(String(x[0]).trim());});
  return out;
}
function isFinanceNo_(no){return financeNos_().indexOf(String(no||"").trim())>=0;}
function getFinApprovals_(key){let o={};try{o=JSON.parse(PropertiesService.getScriptProperties().getProperty("FIN_APPR")||"{}");}catch(e){o={};}return o[key]||{};}
function setFinApproval_(key,no,rec){
  const props=PropertiesService.getScriptProperties();let o={};try{o=JSON.parse(props.getProperty("FIN_APPR")||"{}");}catch(e){o={};}
  if(!o[key])o[key]={};if(rec&&rec.status)o[key][no]=rec;else delete o[key][no];
  const keys=Object.keys(o).sort();while(keys.length>12)delete o[keys.shift()];
  props.setProperty("FIN_APPR",JSON.stringify(o));
}
function temkaNos_(){
  const esh=getSpreadsheet_().getSheetByName(EMP_SHEET);if(!esh||esh.getLastRow()<2)return[];
  const out=[];esh.getRange(2,1,esh.getLastRow()-1,2).getDisplayValues().forEach(x=>{if(TEMKA_PEOPLE.indexOf(String(x[1]).trim())>=0)out.push(String(x[0]).trim());});
  return out;
}
// finance: 21→20; temka (гадны хүн): Total-тэй адил 26→25
function finPeriodFor_(no,offset){const o=Number(offset)===-1?-1:0;return temkaNos_().indexOf(String(no||"").trim())>=0?getPeriodByOffset_(o):getFinancePeriod_(o);}
function financeMyTotal_(user){
  const no=String(user.no||"").trim(),p=finPeriodFor_(no,user.periodOffset),out=buildTotalRowsAll_(p,{}),ap=getFinApprovals_(p.key)[no]||{};
  for(let i=1;i<out.length;i++)if(String(out[i][0]).trim()===no)return{success:true,data:{no:out[i][0],ovog:out[i][1],ner:out[i][2],job:out[i][3],fit:Number(out[i][4])||0,service:Number(out[i][5])||0,hool:Number(out[i][6])||0,breakfast:Number(out[i][7])||0,total:Number(out[i][8])||0,price:Number(out[i][9])||0,status:String(ap.status||""),approvedDate:ap.date||"",reason:String(ap.reason||""),period:p.label,periodOffset:p.offset,canApprove:canDecide_(p),deadline:p.offset===0?"":decideDeadlineLabel_(p)}};
  return{success:false,message:p.label+" хугацаанд таны нийт дүн олдсонгүй"};
}
function financeApproveTotal_(user){
  const p=finPeriodFor_(user.no,user.periodOffset),reject=String(user.decision||"")===  "reject",reason=String(user.reason||"").trim().slice(0,300);
  if(!canDecide_(p))return{success:false,message:p.label+" хугацааны дүнд хариулах хугацаа дууссан"};
  if(reject&&!reason)return{success:false,message:"Зөвшөөрөөгүй шалтгаанаа бичнэ үү"};
  const status=reject?"Зөвшөөрөөгүй":"Зөвшөөрсөн",no=String(user.no||"").trim();
  const lock=LockService.getDocumentLock();if(lock)lock.waitLock(20000);
  try{
    setFinApproval_(p.key,no,{status:status,date:new Date().toISOString(),reason:reject?reason:""});
    writeLog_(reject?"TOTAL_REJECTED":"TOTAL_APPROVED",user,(reject?"Нийт дүнг зөвшөөрөөгүй: "+reason:"Нийт дүнгээ зөвшөөрсөн")+" • "+p.label,FINANCE_SHEET,"",12);
    try{createFinanceSheet();}catch(e){}
    return{success:true,message:reject?"Таны \"Зөвшөөрөөгүй\" хариу болон шалтгаан хадгалагдлаа":"Та нийт дүнгээ зөвшөөрлөө"};
  }finally{if(lock)lock.releaseLock();}
}
function financeSetStatus_(p){
  const no=String(p.no||"").trim(),per=finPeriodFor_(no,p.periodOffset),status=String(p.status||""),reason=status==="Зөвшөөрөөгүй"?String(p.reason||"").trim().slice(0,300):"";
  if(["","Зөвшөөрсөн","Зөвшөөрөөгүй"].indexOf(status)<0)return{success:false,message:"Төлөв буруу байна"};
  if(status==="Зөвшөөрөөгүй"&&!reason)return{success:false,message:"Зөвшөөрөөгүй шалтгаан бичнэ үү"};
  const lock=LockService.getDocumentLock();if(lock)lock.waitLock(20000);
  try{
    const old=(getFinApprovals_(per.key)[no]||{}).status||"";
    setFinApproval_(per.key,no,status?{status:status,date:new Date().toISOString(),reason:reason}:null);
    writeLog_("ADMIN_TOTAL_STATUS",{no:no,ner:"Admin"},per.label+" • "+(old||"Хариулаагүй")+" → "+(status||"Хариулаагүй")+(reason?" • "+reason:""),FINANCE_SHEET,"",12);
    try{createFinanceSheet();}catch(e){}
    return{success:true,message:status?'Төлөв "'+status+'" боллоо':'Хариуг устгалаа ("Хариулаагүй")'};
  }finally{if(lock)lock.releaseLock();}
}
function financeAdminRows_(offset){
  const o=Number(offset)===-1?-1:0,pf=getFinancePeriod_(o),pt=getPeriodByOffset_(o),allF=buildTotalRowsAll_(pf,{}),allT=buildTotalRowsAll_(pt,{}),fn=financeNos_(),tn=temkaNos_(),rows=[];
  const fmt=d=>{if(!d)return"";const x=new Date(d);return isNaN(x.getTime())?String(d):Utilities.formatDate(x,APP_TZ,"yyyy-MM-dd HH:mm");};
  [{per:pf,all:allF,tag:"[21–20] ",pick:no=>tn.indexOf(no)<0},{per:pt,all:allT,tag:"[temka 26–25] ",pick:no=>tn.indexOf(no)>=0}].forEach(g=>{
    const ap=getFinApprovals_(g.per.key);
    for(let i=1;i<g.all.length;i++){const r=g.all[i],no=String(r[0]).trim();if(fn.indexOf(no)<0||!g.pick(no))continue;const x=ap[no]||{};
      rows.push({no:no,ovog:String(r[1]||""),ner:String(r[2]||""),job:g.tag+String(r[3]||""),total:Number(r[8])||0,price:Number(r[9])||0,status:String(x.status||""),date:fmt(x.date),reason:String(x.reason||"")});}
  });
  return rows;
}
// Нэг удаа ажиллуулна: одоо Total дээр байгаа зөвшөөрлийг санхүүгийн хугацаанд шилжүүлээд Total, Total_History-г тэднийгүйгээр дахин бүтээнэ
function applyFinanceSeparation(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty("FIN_MIGRATED")!=="1"){
    const map=readApprovalMap_(getSpreadsheet_().getSheetByName(TOTAL_SHEET),null),key=getFinancePeriod_(-1).key;
    financeNos_().forEach(no=>{const r=map[no];if(r&&r.status)setFinApproval_(key,no,{status:r.status,date:r.date instanceof Date?r.date.toISOString():String(r.date||""),reason:r.reason||""});});
    props.setProperty("FIN_MIGRATED","1");
  }
  if(props.getProperty("TEMKA_MIGRATED")!=="1"){
    const map2=readApprovalMap_(getSpreadsheet_().getSheetByName(TOTAL_SHEET),null),key2=getFinancePeriod_(-1).key,esh=getSpreadsheet_().getSheetByName(EMP_SHEET);
    if(esh&&esh.getLastRow()>1)esh.getRange(2,1,esh.getLastRow()-1,2).getDisplayValues().forEach(x=>{if(TEMKA_PEOPLE.indexOf(String(x[1]).trim())<0)return;const no=String(x[0]).trim(),r=map2[no];if(r&&r.status)setFinApproval_(key2,no,{status:r.status,date:r.date instanceof Date?r.date.toISOString():String(r.date||""),reason:r.reason||""});});
    props.setProperty("TEMKA_MIGRATED","1");
  }
  if(props.getProperty("TEMKA_FIX")!=="1"){
    let o={};try{o=JSON.parse(props.getProperty("FIN_APPR")||"{}");}catch(e){o={};}
    const oldKey=getFinancePeriod_(-1).key,newKey=getPeriodByOffset_(0).key;
    temkaNos_().forEach(no=>{if(o[oldKey]&&o[oldKey][no]){setFinApproval_(newKey,no,o[oldKey][no]);setFinApproval_(oldKey,no,null);}});
    props.setProperty("TEMKA_FIX","1");
  }
  createTotalSheet();createFinanceSheet();
}
