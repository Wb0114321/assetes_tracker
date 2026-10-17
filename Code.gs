const CONFIG={ASSETS:"Assets",DEPARTMENTS:"Departments",USERS:"Users"};
const HEADERS={
  Assets:["S.no","Item Name","Serial Number","Specification","Qty","Department","Location","Other Remarks"],
  Departments:["Dept_ID","DeptName"],
  Users:["User_id","User_name","Password"]
};
function doGet(){return HtmlService.createHtmlOutput("Asset API is running. Use the GitHub Pages frontend.");}
function doPost(e){
  let requestId="", response;
  try{
    const p=(e&&e.parameter)||{};
    requestId=String(p.requestId||"");
    if(!/^[a-zA-Z0-9_-]{8,100}$/.test(requestId))throw new Error("Invalid request ID.");
    const action=String(p.action||"");
    let payload={};try{payload=p.payload?JSON.parse(p.payload):{};}catch(_){throw new Error("Invalid request payload.");}
    response=dispatch_(action,payload);
  }catch(err){response={success:false,error:String(err&&err.message?err.message:err||"Unknown error").slice(0,500)};}
  const encoded=Utilities.base64EncodeWebSafe(JSON.stringify(response),Utilities.Charset.UTF_8);
  const html='<!doctype html><html><head><meta charset="utf-8"></head><body><script>parent.postMessage({channel:"asset-management-api",requestId:'+JSON.stringify(requestId)+',encoded:'+JSON.stringify(encoded)+'},"*");</script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function dispatch_(action,p){switch(action){
  case "login":return apiLogin_(p.userId,p.password);
  case "validateSession":return apiValidateSession_(p.token);
  case "logout":return apiLogout_(p.token);
  case "dashboard":return apiDashboard_(p.token);
  case "addDepartment":return apiAddDepartment_(p.token,p.name);
  case "addAsset":return apiAddAsset_(p.token,p.form);
  default:throw new Error("Unknown API action.");
}}
function setupAssetApp(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(HEADERS).forEach(([name,headers])=>{
    let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
    const current=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
    if(current.every(v=>!String(v).trim()))sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setBackground("#17365D").setFontColor("#FFFFFF").setFontWeight("bold");
  });
  const users=ss.getSheetByName(CONFIG.USERS);
  if(users.getLastRow()<2)users.appendRow(["admin","Administrator","DemoAdmin@123"]);
  return "Setup complete. Existing rows preserved.";
}
function getSheet_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error("Sheet not found: "+name);return sh;}
function rows_(name){const sh=getSheet_(name),last=sh.getLastRow();if(last<2)return [];return sh.getRange(2,1,last-1,HEADERS[name].length).getValues().filter(row=>row.some(v=>String(v==null?"":v).trim()!==""));}
function sessionKey_(token){return "ASSET_SESSION_"+token;}
function apiLogin_(userId,password){
  userId=String(userId||"").trim();password=String(password||"");
  if(!userId||!password)throw new Error("Enter User ID and password.");
  const user=rows_(CONFIG.USERS).find(r=>String(r[0]).trim()===userId&&String(r[2])===password);
  if(!user)throw new Error("Invalid User ID or Password.");
  const token=Utilities.getUuid()+Utilities.getUuid(),name=String(user[1]||user[0]);
  PropertiesService.getScriptProperties().setProperty(sessionKey_(token),JSON.stringify({id:String(user[0]),name,created:new Date().toISOString()}));
  return {success:true,token,name};
}
function requireLogin_(token){
  token=String(token||"");if(!/^[a-f0-9-]{60,80}$/i.test(token))throw new Error("Session invalid. Please log in again.");
  const value=PropertiesService.getScriptProperties().getProperty(sessionKey_(token));if(!value)throw new Error("Session invalid. Please log in again.");return JSON.parse(value);
}
function apiValidateSession_(token){const s=requireLogin_(token);return {success:true,valid:true,name:s.name};}
function apiLogout_(token){token=String(token||"");if(/^[a-f0-9-]{60,80}$/i.test(token))PropertiesService.getScriptProperties().deleteProperty(sessionKey_(token));return {success:true};}
function apiDashboard_(token){
  requireLogin_(token);const assets=rows_(CONFIG.ASSETS),departments=rows_(CONFIG.DEPARTMENTS);
  return {success:true,totalAssets:assets.length,totalQty:assets.reduce((s,r)=>s+(Number(r[4])||0),0),totalDepartments:departments.length,
    departments:departments.map(r=>({id:String(r[0]||""),name:String(r[1]||"")})),
    assets:assets.slice(-1000).reverse().map(r=>({sno:r[0],item:String(r[1]||""),serial:String(r[2]||""),specification:String(r[3]||""),qty:r[4],department:String(r[5]||""),location:String(r[6]||""),remarks:String(r[7]||"")}))};
}
function apiAddDepartment_(token,name){
  requireLogin_(token);name=String(name||"").trim();if(!name)throw new Error("Department name is required.");if(name.length>100)throw new Error("Department name is too long.");
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_(CONFIG.DEPARTMENTS),headers=sh.getRange("A1:B1").getDisplayValues()[0];
    if(headers[0].trim()!=="Dept_ID"||headers[1].trim()!=="DeptName")throw new Error("Departments headers must be Dept_ID and DeptName.");
    const last=sh.getLastRow(),data=last>=2?sh.getRange(2,1,last-1,2).getDisplayValues():[],filled=data.filter(r=>String(r[1]||"").trim()!=="");
    const dup=filled.find(r=>String(r[1]).trim().toLowerCase()===name.toLowerCase());if(dup)throw new Error('Department "'+name+'" already exists (ID: '+(dup[0]||"not assigned")+").");
    let max=0;filled.forEach(r=>{const m=String(r[0]).trim().match(/^D(\d+)$/i);if(m)max=Math.max(max,Number(m[1]));});
    const id="D"+String(max+1).padStart(3,"0");let target=2;
    while(target<=sh.getLastRow()){const r=sh.getRange(target,1,1,2).getDisplayValues()[0];if(!r[0].trim()&&!r[1].trim())break;target++;}
    sh.getRange(target,1,1,2).setValues([[id,name]]);SpreadsheetApp.flush();
    const saved=sh.getRange(target,1,1,2).getDisplayValues()[0];if(saved[0]!==id||saved[1]!==name)throw new Error("Department save verification failed.");
    return {success:true,id,name,departments:rows_(CONFIG.DEPARTMENTS).map(r=>({id:String(r[0]||""),name:String(r[1]||"")}))};
  }finally{lock.releaseLock();}
}
function apiAddAsset_(token,form){
  requireLogin_(token);if(!form)throw new Error("Asset details are missing.");
  const item=String(form.item||"").trim(),serial=String(form.serial||"").trim(),spec=String(form.specification||"").trim(),department=String(form.department||"").trim(),location=String(form.location||"").trim(),remarks=String(form.remarks||"").trim(),qty=Number(form.qty);
  if(!item)throw new Error("Item Name is required.");if(!serial)throw new Error("Serial Number is required.");if(!Number.isSafeInteger(qty)||qty<=0)throw new Error("Quantity must be a positive whole number.");if(!department)throw new Error("Select a department first.");
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_(CONFIG.ASSETS),departments=rows_(CONFIG.DEPARTMENTS),assets=rows_(CONFIG.ASSETS);
    if(!departments.some(r=>String(r[1]).trim()===department))throw new Error("Department not found. Refresh the page.");
    if(assets.some(r=>String(r[2]).trim().toLowerCase()===serial.toLowerCase()))throw new Error("Serial number already exists: "+serial);
    const sno=assets.reduce((max,r)=>{const n=Number(r[0]);return Number.isFinite(n)?Math.max(max,n):max;},0)+1;
    sh.getRange(Math.max(sh.getLastRow()+1,2),1,1,8).setValues([[sno,item,serial,spec,qty,department,location,remarks]]);SpreadsheetApp.flush();
    return {success:true,sno,message:"Asset saved successfully."};
  }finally{lock.releaseLock();}
}
