let adminRows=[];
let adminPassword=sessionStorage.getItem("akumaAdmin")||"";
const $=id=>document.getElementById(id);
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

async function load(){
  $("refreshBtn").disabled=true;
  $("tableBox").innerHTML='<div class="notice info">Уншиж байна...</div>';
  try{
    const r=await Api.call("adminDashboard",{password:adminPassword,periodOffset:Number($("adminPeriod").value||0)});
    if(!r.success)throw new Error(r.message);
    $("adminLogin").classList.add("hidden");$("dashboard").classList.remove("hidden");
    adminRows=r.rows||[];populateEmployees(r.employees||[]);
    $("fitStat").textContent=r.stats.fit;
    $("serviceStat").textContent=Number(r.stats.service||0).toLocaleString()+"₮";
    $("hoolStat").textContent=r.stats.hool;
    $("breakfastStat").textContent=r.stats.breakfast;
    $("totalFoodStat").textContent=r.stats.totalFood;
    $("totalMoneyStat").textContent=Number(r.stats.totalMoney||0).toLocaleString()+"₮";
    $("pendingStat").textContent=r.stats.pending;
    $("periodLabel").textContent="Тооцооны хугацаа: "+r.stats.period;
    render();loadApprovals();
    $("logs").innerHTML=(r.logs||[]).map(x=>`<div class="log"><strong>${esc(x.action)}</strong><p>${esc(x.user)} • ${esc(x.detail)}</p><small class="muted">${esc(x.time)}</small></div>`).join("")||'<p class="muted">Лог байхгүй</p>';
  }catch(e){
    sessionStorage.removeItem("akumaAdmin");$("dashboard").classList.add("hidden");$("adminLogin").classList.remove("hidden");$("loginMsg").innerHTML=`<div class="notice error">${esc(e.message)}</div>`;
  }finally{$("refreshBtn").disabled=false;}
}

function populateEmployees(items){const el=$("manualEmployee"),selected=el.value;el.innerHTML=`<option value="">Ажилтан сонгох</option>`+items.map(x=>`<option value="${esc(x.no)}">${esc(x.label)}</option>`).join("");el.value=selected;}
function render(){const f=$("sheetFilter").value,q=$("search").value.trim().toLowerCase();const rows=adminRows.filter(x=>(f==="all"||x.sheet===f)&&(!q||`${x.no} ${x.ovog} ${x.ner} ${x.job}`.toLowerCase().includes(q)));$("tableBox").innerHTML=`<table><thead><tr><th>Төрөл</th><th>№</th><th>Нэр</th><th>Албан тушаал</th><th>Огноо</th><th>Утга</th><th>Үйлдэл</th></tr></thead><tbody>${rows.map(x=>`<tr><td><span class="badge blue">${esc(x.sheet)}</span></td><td>${esc(x.no)}</td><td>${esc(x.ovog)} ${esc(x.ner)}</td><td>${esc(x.job)}</td><td>${esc(x.date)}</td><td><b>${esc(x.value)}</b></td><td>${x.pending?`<button class="btn btn-primary" onclick="pendingAction('${esc(x.sheet)}',${x.row},${x.col},'approve')">Зөвшөөрөх</button> <button class="btn btn-soft" onclick="pendingAction('${esc(x.sheet)}',${x.row},${x.col},'reject')">Татгалзах</button> `:``}<button class="btn btn-danger" onclick="removeEntry('${esc(x.sheet)}',${x.row},${x.col})">Устгах</button></td></tr>`).join("")||'<tr><td colspan="7">Мэдээлэл олдсонгүй</td></tr>'}</tbody></table>`;}
async function pendingAction(sheet,row,col,action){const r=await Api.call("adminAction",{password:adminPassword,sheet,row,col,action});if(!r.success){alert(r.message);return;}load();}
async function removeEntry(sheet,row,col){if(!confirm("Энэ бүртгэлийг устгах уу?"))return;const r=await Api.call("adminRemove",{password:adminPassword,sheet,row,col});if(!r.success){alert(r.message);return;}load();}
// ---------- Нийт дүнгийн зөвшөөрөл ----------
let approvalRows=[];
async function loadApprovals(){$("approvalBox").innerHTML='<div class="notice info">Уншиж байна...</div>';try{const r=await Api.call("adminTotals",{password:adminPassword,periodOffset:Number($("adminPeriod").value||0)});if(!r.success)throw new Error(r.message);approvalRows=r.rows||[];renderApprovals();}catch(e){approvalRows=[];$("approvalAlert").innerHTML="";$("approvalBox").innerHTML=`<div class="notice error">${esc(e.message)}</div>`;}}
function renderApprovals(){
  const c={ok:0,no:0,none:0};approvalRows.forEach(x=>{if(x.status==="Зөвшөөрсөн")c.ok++;else if(x.status==="Зөвшөөрөөгүй")c.no++;else c.none++;});
  $("approvalCounts").textContent=`${c.ok} зөвшөөрсөн · ${c.no} зөвшөөрөөгүй · ${c.none} хариулаагүй`;
  const rej=approvalRows.filter(x=>x.status==="Зөвшөөрөөгүй");
  $("approvalAlert").innerHTML=rej.length?`<div class="notice error approval-alert">⚠️ <b>${rej.length} хүн нийт дүнгээ зөвшөөрөөгүй байна:</b> ${rej.map(x=>`<a href="#" onclick="focusApproval('${esc(x.no)}');return false">${esc(x.ovog)} ${esc(x.ner)}</a>${x.reason?` (${esc(x.reason)})`:""}`).join(", ")}</div>`:"";
  const f=$("approvalFilter").value,q=$("approvalSearch").value.trim().toLowerCase();
  const rows=approvalRows.filter(x=>(f==="all"||(f==="none"?!x.status:x.status===f))&&(!q||`${x.no} ${x.ovog} ${x.ner} ${x.job}`.toLowerCase().includes(q)));
  const badge=s=>s==="Зөвшөөрсөн"?'<span class="badge green">Зөвшөөрсөн</span>':s==="Зөвшөөрөөгүй"?'<span class="badge red">Зөвшөөрөөгүй</span>':'<span class="badge blue">Хариулаагүй</span>';
  $("approvalBox").innerHTML=`<table><thead><tr><th>№</th><th>Нэр</th><th>Албан тушаал</th><th>Нийт үнэ</th><th>Төлөв</th><th>Огноо</th><th>Шалтгаан</th><th>Үйлдэл</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.no)}</td><td>${esc(x.ovog)} ${esc(x.ner)}</td><td>${esc(x.job)}</td><td><b>${Number(x.price||0).toLocaleString()}₮</b></td><td>${badge(x.status)}</td><td>${esc(x.date)}</td><td>${esc(x.reason)}</td><td><div class="approval-actions">${x.status!=="Зөвшөөрсөн"?`<button class="btn btn-primary" onclick="setApproval('${esc(x.no)}','Зөвшөөрсөн')">Зөвшөөрсөн</button>`:""}<button class="btn btn-soft" onclick="setApproval('${esc(x.no)}','Зөвшөөрөөгүй')">${x.status==="Зөвшөөрөөгүй"?"Шалтгаан засах":"Зөвшөөрөөгүй"}</button>${x.status?`<button class="btn btn-danger" onclick="setApproval('${esc(x.no)}','')">Устгах</button>`:""}</div></td></tr>`).join("")||'<tr><td colspan="8">Мэдээлэл олдсонгүй</td></tr>'}</tbody></table>`;
}
function focusApproval(no){$("approvalFilter").value="all";$("approvalSearch").value=no;renderApprovals();$("approvalBox").scrollIntoView({behavior:"smooth",block:"center"});}
async function setApproval(no,status){
  const x=approvalRows.find(r=>String(r.no)===String(no));if(!x)return;const name=`${x.ovog} ${x.ner}`;let reason="";
  if(status===""){if(!confirm(`${name}-ийн хариуг устгаж "Хариулаагүй" болгох уу?\nАжилтан дахин хариулах боломжтой болно.`))return;}
  else if(status==="Зөвшөөрөөгүй"){const v=prompt(`${name} — Зөвшөөрөөгүй шалтгаан:`,x.reason||"");if(v===null)return;reason=v.trim();if(!reason){alert("Шалтгаан бичнэ үү");return;}}
  else if(!confirm(`${name}-ийн нийт дүнг "Зөвшөөрсөн" болгох уу?`))return;
  const r=await Api.call("adminTotalStatus",{password:adminPassword,no,periodOffset:Number($("adminPeriod").value||0),status,reason});alert(r.message);if(r.success)loadApprovals();
}
$("approvalFilter").onchange=renderApprovals;$("approvalSearch").oninput=renderApprovals;

function adminToday(){const parts=new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Ulaanbaatar",year:"numeric",month:"2-digit",day:"2-digit"}).formatToParts(new Date()),x={};parts.forEach(p=>x[p.type]=p.value);return x.year+"-"+x.month+"-"+x.day;}

$("manualDate").value=adminToday();
$("manualType").onchange=()=>{const service=$("manualType").value==="service";$("manualAmountBox").classList.toggle("hidden",!service);$("manualQuantityBox").classList.toggle("hidden",service);};
$("manualAddBtn").onclick=async()=>{const btn=$("manualAddBtn");btn.disabled=true;$("manualMsg").innerHTML="";try{const r=await Api.call("adminAdd",{password:adminPassword,employeeNo:$("manualEmployee").value,date:$("manualDate").value,type:$("manualType").value,quantity:Number($("manualQuantity").value||1),amount:Number($("manualAmount").value||0)});if(!r.success)throw new Error(r.message);$("manualMsg").innerHTML=`<div class="notice ok">${esc(r.message)}</div>`;$("manualAmount").value="";await load();}catch(e){$("manualMsg").innerHTML=`<div class="notice error">${esc(e.message)}</div>`;}finally{btn.disabled=false;}};
$("adminLoginBtn").onclick=()=>{adminPassword=$("adminPassword").value;sessionStorage.setItem("akumaAdmin",adminPassword);load();};
$("refreshBtn").onclick=load;$("adminPeriod").onchange=load;$("sheetFilter").onchange=render;$("search").oninput=render;if(adminPassword)load();
