const Api=(()=>{let seq=0;
const TIMEOUT_MSG="Сервер удаан хариулж байна. Дахин дарж шалгана уу — давхар бүртгэгдэхгүй";
function url(action,payload,cb){const base=window.APP_CONFIG.API_URL;const q=new URLSearchParams({api:"1",action,callback:cb,payload:JSON.stringify(payload),_:Date.now()});return base+(base.includes("?")?"&":"?")+q}
// 1) fetch (cookie-гүй): олон Google account-тай browser дээр /u/1 redirect-ээс сэргийлнэ
async function viaFetch(action,payload){const ctrl=new AbortController(),t=setTimeout(()=>ctrl.abort(),60000);try{const r=await fetch(url(action,payload,"cb"),{credentials:"omit",signal:ctrl.signal});const txt=await r.text();const m=txt.match(/^[^(]*\(([\s\S]*)\);?\s*$/);if(!m)throw new Error("bad response");return JSON.parse(m[1])}catch(e){if(e.name==="AbortError"){const x=new Error(TIMEOUT_MSG);x.timeout=true;throw x}throw e}finally{clearTimeout(t)}}
// 2) script tag (JSONP) — fallback
function viaScript(action,payload){return new Promise((resolve,reject)=>{const cb="__akumaCb"+(++seq);const script=document.createElement("script");const timer=setTimeout(()=>finish(new Error(TIMEOUT_MSG)),60000);function finish(error,data){clearTimeout(timer);delete window[cb];script.remove();error?reject(error):resolve(data)}window[cb]=data=>finish(null,data);script.onerror=()=>finish(new Error("Сервертэй холбогдож чадсангүй"));script.src=url(action,payload,cb);document.body.appendChild(script)})}
async function call(action,payload={}){const base=window.APP_CONFIG.API_URL;if(!base||base.includes("PASTE_"))throw new Error("config.js дотор API_URL тохируулна уу");try{return await viaFetch(action,payload)}catch(e){if(e.timeout)throw e;return viaScript(action,payload)}}
return{call}})();
