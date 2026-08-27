const EMP_SHEET="Employees",FIT_SHEET="FitFood",SERVICE_SHEET="ServiceFood",HOOL_SHEET="Hool",LOG_SHEET="Logs";
const APP_TZ="Asia/Ulaanbaatar";

function doGet(e){
  if(e&&e.parameter&&e.parameter.api==="1")return handleApi_(e);
  return ContentService.createTextOutput("AKUMA Food API ажиллаж байна");
}

function handleApi_(e){
  const callback=String(e.parameter.callback||"callback").replace(/[^a-zA-Z0-9_$]/g,"");
  let result;
  try{
    const p=JSON.parse(e.parameter.payload||"{}");
    switch(String(e.parameter.action||"")){
      case "login":result=loginUser(p.code,p.password);break;
      case "saveFood":result=saveFood(p);break;
      case "getMyTotal":result=getMyTotal(p);break;
      case "approveTotal":result=userApproveTotal(p);break;
      case "adminDashboard":requireAdmin_(p.password);result=getAdminDashboard();break;
      case "adminRemove":requireAdmin_(p.password);result=adminRemove(p);break;
      case "adminAction":requireAdmin_(p.password);result=adminAction(p);break;
      default:throw new Error("API action буруу байна");
    }
  }catch(err){result={success:false,message:err.message};}
  return ContentService.createTextOutput(callback+"("+JSON.stringify(result)+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setupSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  setupOneSheet_(ss,EMP_SHEET,false);[FIT_SHEET,SERVICE_SHEET,HOOL_SHEET].forEach(n=>setupOneSheet_(ss,n,true));
  setupLogSheet_(ss);createTotalSheet();
}

function requireAdmin_(password){const expected=PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");if(!expected)throw new Error("Apps Script Properties дээр ADMIN_PASSWORD тохируулна уу");if(String(password||"")!==expected)throw new Error("Admin нууц үг буруу байна");}

function setupOneSheet_(ss,name,withDates){
  let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){let h=name===EMP_SHEET?["№","Код","Нууц үг","Овог","Нэр","Албан тушаал"]:["№","Овог","Нэр","Ажил үүрэг","Нийт"];
    sh.getRange(1,1,1,h.length).setNumberFormat("@").setValues([h]);styleHeader_(sh,h.length);}
  if(withDates&&!findTotalRow(sh)){const r=Math.max(sh.getLastRow()+1,2);sh.getRange(r,1).setValue("Нийт");}
}

function setupLogSheet_(ss){let sh=ss.getSheetByName(LOG_SHEET);if(!sh)sh=ss.insertSheet(LOG_SHEET);if(sh.getLastRow()===0){const h=["Огноо цаг","Үйлдэл","Ажилтны №","Хэрэглэгч","Дэлгэрэнгүй","Sheet","Мөр","Багана"];sh.getRange(1,1,1,h.length).setValues([h]);styleHeader_(sh,h.length);}}
function styleHeader_(sh,n){sh.getRange(1,1,1,n).setFontWeight("bold").setBackground("#17233c").setFontColor("white").setHorizontalAlignment("center");sh.setFrozenRows(1);}

function loginUser(code,password){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(EMP_SHEET);if(!sh){setupSheets();sh=ss.getSheetByName(EMP_SHEET);}
  const data=sh.getDataRange().getDisplayValues();
  for(let i=1;i<data.length;i++)if(String(data[i][1]).trim()===String(code||"").trim()&&String(data[i][2]).trim()===String(password||"").trim()){
    const user={no:data[i][0],code:data[i][1],ovog:data[i][3],ner:data[i][4],job:data[i][5]};writeLog_("LOGIN",user,"Амжилттай нэвтэрсэн");return{success:true,user};}
  writeLog_("LOGIN_FAILED",{no:String(code||""),ner:""},"Код эсвэл нууц үг буруу");return{success:false,message:"Код эсвэл нууц үг буруу байна"};
}

function saveFood(data){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const type=String(data.type||"").trim();if(!["fit","service","hool","ahool"].includes(type))throw new Error("Хоолны төрөл буруу байна");
    const sheetName=type==="service"?SERVICE_SHEET:(type==="hool"||type==="ahool")?HOOL_SHEET:FIT_SHEET;
    const no=String(data.no||"").trim(),ner=String(data.ner||"").trim(),date=String(data.date||"").trim(),amount=Number(data.amount||0);
    if(!no||!ner||!date)throw new Error("Мэдээлэл дутуу байна");
    if(type==="service"&&(!Number.isFinite(amount)||amount<=0))throw new Error("ServiceFood үнэ оруулна уу");
    const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(sheetName);if(!sh){sh=ss.insertSheet(sheetName);setupOneSheet_(ss,sheetName,true);}
    const col=ensureDateColumn(sh,date);let totalRow=findTotalRow(sh);if(!totalRow){totalRow=Math.max(sh.getLastRow()+1,2);sh.getRange(totalRow,1).setValue("Нийт");}
    let row=findEmployeeRow(sh,no,totalRow);if(!row){sh.insertRowBefore(totalRow);row=totalRow;totalRow++;}
    const value=type==="service"?amount:(isAfterCutoff_()?"Pending":1);
    sh.getRange(row,1,1,4).setValues([[no,String(data.ovog||""),ner,String(data.job||"")]]);sh.getRange(row,col).setValue(value);
    updateRowTotal(sh,row);updateTotalRow(sh);createTotalSheet();writeLog_(value==="Pending"?"PENDING_CREATED":"FOOD_SAVED",data,(value==="Pending"?"Admin зөвшөөрөл хүлээж байна":value+" • "+date),sheetName,row,col);
    return{success:true,pending:value==="Pending",message:value==="Pending"?"11:30 өнгөрсөн тул Admin зөвшөөрөл хүлээж байна":sheetName+" амжилттай хадгалагдлаа"};
  }catch(err){return{success:false,message:err.message};}finally{lock.releaseLock();}
}

function isAfterCutoff_(){const hm=Utilities.formatDate(new Date(),APP_TZ,"HH:mm").split(":").map(Number);return hm[0]>11||(hm[0]===11&&hm[1]>30);}

function getAdminDashboard(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),rows=[],stats={fit:0,service:0,hool:0,totalFood:0,totalMoney:0,all:0,pending:0,period:getCurrentPeriod_().label};
  [FIT_SHEET,SERVICE_SHEET,HOOL_SHEET].forEach(name=>{const sh=ss.getSheetByName(name);if(!sh)return;const data=sh.getDataRange().getDisplayValues();if(data.length<2)return;
    for(let r=1;r<data.length;r++){if(String(data[r][0]).trim()==="Нийт")continue;for(let c=4;c<data[0].length;c++){const v=String(data[r][c]||"").trim();if(!v||data[0][c]==="Нийт"||!isDateInCurrentPeriod_(data[0][c]))continue;rows.push({sheet:name,row:r+1,col:c+1,no:data[r][0],ovog:data[r][1],ner:data[r][2],job:data[r][3],date:data[0][c],value:v,pending:v==="Pending"});stats.all++;if(v==="Pending")stats.pending++;const n=Number(v)||0;if(name===FIT_SHEET)stats.fit+=n;if(name===SERVICE_SHEET)stats.service+=n;if(name===HOOL_SHEET)stats.hool+=n;}}}
  );stats.totalFood=stats.fit+stats.hool;stats.totalMoney=(stats.totalFood*15000)+stats.service;rows.sort((a,b)=>b.col-a.col||b.row-a.row);return{success:true,rows,stats,logs:getRecentLogs_(80)};
}

function adminRemove(p){const allowed=[FIT_SHEET,SERVICE_SHEET,HOOL_SHEET];if(!allowed.includes(String(p.sheet)))return{success:false,message:"Sheet буруу байна"};const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.sheet);if(!sh)return{success:false,message:"Sheet олдсонгүй"};const row=Number(p.row),col=Number(p.col);if(row<2||col<5||row>sh.getLastRow()||col>sh.getLastColumn())return{success:false,message:"Мөр/багана буруу байна"};const old=sh.getRange(row,col).getDisplayValue();sh.getRange(row,col).clearContent();updateRowTotal(sh,row);updateTotalRow(sh);createTotalSheet();writeLog_("ADMIN_REMOVE",{no:sh.getRange(row,1).getDisplayValue(),ner:"Admin"},p.sheet+" • "+old,p.sheet,row,col);return{success:true,message:"Устгалаа"};}

function adminAction(p){const allowed=[FIT_SHEET,HOOL_SHEET];if(!allowed.includes(String(p.sheet)))return{success:false,message:"Зөвхөн FitFood/Хоол Pending зөвшөөрнө"};const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.sheet),row=Number(p.row),col=Number(p.col);if(!sh||String(sh.getRange(row,col).getDisplayValue()).trim()!=="Pending")return{success:false,message:"Pending бүртгэл олдсонгүй"};const value=p.action==="approve"?1:p.action==="reject"?"Rejected":null;if(value===null)return{success:false,message:"Үйлдэл буруу"};sh.getRange(row,col).setValue(value);updateRowTotal(sh,row);updateTotalRow(sh);createTotalSheet();writeLog_(p.action==="approve"?"ADMIN_APPROVED":"ADMIN_REJECTED",{no:sh.getRange(row,1).getDisplayValue(),ner:"Admin"},p.sheet,p.sheet,row,col);return{success:true,message:p.action==="approve"?"Зөвшөөрлөө":"Татгалзлаа"};}

function getCurrentPeriod_(){const now=new Date(Utilities.formatDate(new Date(),APP_TZ,"yyyy-MM-dd'T'HH:mm:ss"));let start,end;if(now.getDate()>=26){start=new Date(now.getFullYear(),now.getMonth(),26);end=new Date(now.getFullYear(),now.getMonth()+1,25);}else{start=new Date(now.getFullYear(),now.getMonth()-1,26);end=new Date(now.getFullYear(),now.getMonth(),25);}return{start:start,end:end,label:Utilities.formatDate(start,APP_TZ,"yyyy.MM.dd")+" – "+Utilities.formatDate(end,APP_TZ,"yyyy.MM.dd")};}
function isDateInCurrentPeriod_(value){const p=getCurrentPeriod_(),text=String(value||"").trim();let d;if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)){const a=text.split("-").map(Number);d=new Date(a[0],a[1]-1,a[2]);}else{const a=normalizeDate(text).split("/").map(Number);if(a.length!==2)return false;let year=p.start.getFullYear();if(a[0]<p.start.getMonth()+1)year=p.end.getFullYear();d=new Date(year,a[0]-1,a[1]);}return d>=p.start&&d<=p.end;}

function writeLog_(action,user,detail,sheet,row,col){const ss=SpreadsheetApp.getActiveSpreadsheet();setupLogSheet_(ss);ss.getSheetByName(LOG_SHEET).appendRow([new Date(),action,String(user.no||""),[user.ovog,user.ner].filter(Boolean).join(" "),String(detail||""),sheet||"",row||"",col||""]);}
function getRecentLogs_(limit){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);if(!sh||sh.getLastRow()<2)return[];const start=Math.max(2,sh.getLastRow()-limit+1),v=sh.getRange(start,1,sh.getLastRow()-start+1,8).getDisplayValues();return v.reverse().map(x=>({time:x[0],action:x[1],user:x[3]||x[2],detail:x[4]}));}

function ensureDateColumn(sh,date){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],target=String(date);for(let i=0;i<h.length;i++)if(String(h[i])===target)return i+1;let total=h.indexOf("Нийт")+1;if(!total){total=sh.getLastColumn()+1;sh.getRange(1,total).setValue("Нийт");}sh.insertColumnBefore(total);sh.getRange(1,total).setNumberFormat("@").setValue(target);return total;}
function normalizeDate(v){if(v instanceof Date)return(v.getMonth()+1)+"/"+v.getDate();const t=String(v||"").trim();if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(t)){const p=t.split("-");return Number(p[1])+"/"+Number(p[2]);}if(t.includes("/")){const p=t.split("/");return Number(p[p.length-2])+"/"+Number(p[p.length-1]);}return t;}
function findEmployeeRow(sh,no,totalRow){if(totalRow<=2)return null;const v=sh.getRange(2,1,totalRow-2,1).getDisplayValues();for(let i=0;i<v.length;i++)if(String(v[i][0]).trim()===String(no).trim())return i+2;return null;}
function findTotalRow(sh){if(sh.getLastRow()<1)return null;const v=sh.getRange(1,1,sh.getLastRow(),1).getDisplayValues();for(let i=0;i<v.length;i++)if(String(v[i][0]).trim()==="Нийт")return i+1;return null;}
function updateRowTotal(sh,row){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],tc=h.indexOf("Нийт")+1;if(!tc)return;sh.getRange(row,tc).setFormula(`=SUM(E${row}:${getColumnLetter(tc-1)}${row})`).setNumberFormat("0");}
function updateTotalRow(sh){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],tc=h.indexOf("Нийт")+1;if(!tc)return;let tr=findTotalRow(sh);if(!tr){tr=Math.max(sh.getLastRow()+1,2);sh.getRange(tr,1).setValue("Нийт");}sh.getRange(tr,1).setValue("Нийт");sh.getRange(tr,2,1,3).clearContent();for(let c=5;c<=tc;c++){if(tr<=2)sh.getRange(tr,c).setValue(0);else{const l=getColumnLetter(c);sh.getRange(tr,c).setFormula(`=SUM(${l}2:${l}${tr-1})`).setNumberFormat("0");}}sh.getRange(tr,1,1,tc).setFontWeight("bold").setBackground("#fff2cc");}
function getColumnLetter(c){let s="";while(c>0){const t=(c-1)%26;s=String.fromCharCode(t+65)+s;c=(c-t-1)/26;}return s;}
