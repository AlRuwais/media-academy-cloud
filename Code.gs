const APP = { VERSION: "1.0.0", SESSION_SECONDS: 21600 };
const TABLES = {
Institutions:["institution_id","institution_name","status","created_at"],
Users:["user_id","institution_id","full_name","email","phone","role","group_id","password_hash","password_salt","status","created_at","updated_at","last_login_at"],
Groups:["group_id","institution_id","group_name","description","created_at"],
Subjects:["subject_id","institution_id","subject_name","description","created_at"],
Sessions:["session_id","institution_id","group_id","subject_id","title","session_date","start_time","end_time","location","notes","created_at"],
Attendance:["attendance_id","institution_id","session_id","student_id","status","notes","created_at"],
Exams:["exam_id","institution_id","group_id","subject_id","title","exam_date","total_marks","created_at"],
Results:["result_id","institution_id","exam_id","student_id","marks","notes","created_at"],
Content:["content_id","institution_id","group_id","subject_id","title","content_type","url","description","created_at"],
Announcements:["announcement_id","institution_id","title","body","created_by","created_at"],
Messages:["message_id","institution_id","sender_id","recipient_id","body","created_at","read_at"],
Complaints:["complaint_id","institution_id","user_id","subject","body","status","created_at","updated_at"],
Audit_Logs:["log_id","institution_id","actor_user_id","action","entity_type","entity_id","details","created_at"],
Settings:["setting_key","setting_value","updated_at"]
};
function doGet(){return json_({ok:true,app:"Media Academy Cloud",version:APP.VERSION,message:"API is running"});}
function doPost(e){try{const b=JSON.parse((e&&e.postData&&e.postData.contents)||"{}"),a=String(b.action||""),d=b.data||{};
if(a==="health")return json_({ok:true,app:"Media Academy Cloud",version:APP.VERSION});
if(a==="initialize")return json_(initializeDatabase_());
if(a==="login")return json_(login_(d));
if(a==="logout")return json_(logout_(b.token));
const u=requireSession_(b.token);
if(a==="me")return json_({ok:true,user:publicUser_(u)});
if(a==="dashboard")return json_(dashboard_(u));
if(a==="list")return json_(listRows_(u,d.table));
if(a==="create")return json_(createRow_(u,d.table,d.row||{}));
if(a==="update")return json_(updateRow_(u,d.table,d.id,d.row||{}));
if(a==="delete")return json_(deleteRow_(u,d.table,d.id));
if(a==="changePassword")return json_(changePassword_(u,d));
throw new Error("Unknown action");
}catch(err){return json_({ok:false,error:String(err.message||err)});}}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function db_(){const id=PropertiesService.getScriptProperties().getProperty("DATABASE_ID");if(!id)throw new Error("Set DATABASE_ID in Script Properties first.");return SpreadsheetApp.openById(id);}
function initializeDatabase_(){const ss=db_(),created=[];Object.keys(TABLES).forEach(n=>{let sh=ss.getSheetByName(n);if(!sh){sh=ss.insertSheet(n);created.push(n);}const h=TABLES[n],last=sh.getLastColumn();if(!last||sh.getLastRow()===0)sh.getRange(1,1,1,h.length).setValues([h]);else{const cur=sh.getRange(1,1,1,last).getValues()[0].map(String),missing=h.filter(x=>!cur.includes(x));if(missing.length)throw new Error("Sheet "+n+" is missing columns: "+missing.join(", "));}sh.setFrozenRows(1);});return{ok:true,message:"Database initialized",createdSheets:created,sheets:Object.keys(TABLES)};}
function initializeDatabase(){return initializeDatabase_();}
function rows_(t){const sh=db_().getSheetByName(t);if(!sh)throw new Error("Unknown table: "+t);const v=sh.getDataRange().getValues();if(v.length<2)return[];const h=v[0].map(String);return v.slice(1).map((r,i)=>{const o={__row:i+2};h.forEach((k,j)=>o[k]=r[j] instanceof Date?r[j].toISOString():r[j]);return o;}).filter(o=>o[h[0]]!=="");}
function publicUser_(u){const o={};["user_id","institution_id","full_name","email","phone","role","group_id","status"].forEach(k=>o[k]=u[k]||"");return o;}
function hash_(v,s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s+String(v),Utilities.Charset.UTF_8).map(b=>(b+256).toString(16).slice(-2)).join("");}
function login_(d){const e=String(d.email||"").trim().toLowerCase(),p=String(d.password||""),u=rows_("Users").find(x=>String(x.email).toLowerCase()===e&&x.status==="active");if(!u||hash_(p,u.password_salt)!==u.password_hash)throw new Error("Email or password is incorrect.");const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put("session:"+token,u.user_id,APP.SESSION_SECONDS);db_().getSheetByName("Users").getRange(u.__row,13).setValue(new Date().toISOString());audit_(u.institution_id,u.user_id,"login","User",u.user_id,"Successful login");return{ok:true,token,user:publicUser_(u)};}
function sessionUser_(t){if(!t)return null;const id=CacheService.getScriptCache().get("session:"+t);return id?rows_("Users").find(x=>x.user_id===id&&x.status==="active")||null:null;}
function requireSession_(t){const u=sessionUser_(t);if(!u)throw new Error("Session expired. Please log in again.");return u;}
function logout_(t){const u=sessionUser_(t);if(t)CacheService.getScriptCache().remove("session:"+t);if(u)audit_(u.institution_id,u.user_id,"logout","User",u.user_id,"Logout");return{ok:true};}
function canRead_(u,t){if(u.role==="owner")return true;if(u.role==="institution_admin")return t!=="Settings";if(u.role==="instructor")return["Groups","Subjects","Sessions","Attendance","Exams","Results","Content","Announcements","Users"].includes(t);return["Sessions","Content","Announcements","Results","Attendance","Groups","Subjects"].includes(t);}
function canManage_(u,t){if(u.role==="owner")return true;if(u.role==="institution_admin")return t!=="Institutions"&&t!=="Settings"&&t!=="Audit_Logs";if(u.role==="instructor")return["Groups","Subjects","Sessions","Attendance","Exams","Results","Content","Announcements"].includes(t);return false;}
function listRows_(u,t){if(!TABLES[t]||!canRead_(u,t))throw new Error("You do not have permission to view this section.");let a=rows_(t);if(u.role!=="owner"&&TABLES[t].includes("institution_id"))a=a.filter(x=>String(x.institution_id)===String(u.institution_id));if(u.role==="student"){if(t==="Users")a=a.filter(x=>x.user_id===u.user_id);if(t==="Results"||t==="Attendance")a=a.filter(x=>x.student_id===u.user_id);}return{ok:true,rows:a.map(x=>{delete x.password_hash;delete x.password_salt;delete x.__row;return x;})};}
function createRow_(u,t,r){if(!TABLES[t]||!canManage_(u,t))throw new Error("You do not have permission to create records here.");const h=TABLES[t],o={};h.forEach(k=>o[k]=r[k]===undefined?"":r[k]);o[h[0]]=Utilities.getUuid();if(h.includes("institution_id"))o.institution_id=u.role==="owner"?(r.institution_id||"main"):u.institution_id;
if(t==="Users"){const p=String(r.password||"");if(p.length<8)throw new Error("Temporary password must be at least 8 characters.");o.password_salt=Utilities.getUuid();o.password_hash=hash_(p,o.password_salt);o.status=o.status||"active";o.created_at=new Date().toISOString();o.updated_at=o.created_at;}else if(h.includes("created_at"))o.created_at=new Date().toISOString();
db_().getSheetByName(t).appendRow(h.map(k=>o[k]));audit_(o.institution_id||u.institution_id,u.user_id,"create",t,o[h[0]],"Record created");return{ok:true,id:o[h[0]]};}
function updateRow_(u,t,id,r){if(!TABLES[t]||!canManage_(u,t))throw new Error("You do not have permission to edit this section.");const h=TABLES[t],hit=rows_(t).find(x=>String(x[h[0]])===String(id));if(!hit||(u.role!=="owner"&&hit.institution_id!==u.institution_id))throw new Error("Record not found.");const sh=db_().getSheetByName(t);h.forEach((k,i)=>{if(k!==h[0]&&k!=="institution_id"&&k!=="password_hash"&&k!=="password_salt"&&r[k]!==undefined)sh.getRange(hit.__row,i+1).setValue(r[k]);});if(t==="Users"&&r.password){const s=Utilities.getUuid();sh.getRange(hit.__row,9).setValue(s);sh.getRange(hit.__row,8).setValue(hash_(r.password,s));}if(h.includes("updated_at"))sh.getRange(hit.__row,h.indexOf("updated_at")+1).setValue(new Date().toISOString());audit_(hit.institution_id||u.institution_id,u.user_id,"update",t,id,"Record updated");return{ok:true};}
function deleteRow_(u,t,id){if(!TABLES[t]||!canManage_(u,t))throw new Error("You do not have permission to delete records here.");const h=TABLES[t],hit=rows_(t).find(x=>String(x[h[0]])===String(id));if(!hit||(u.role!=="owner"&&hit.institution_id!==u.institution_id))throw new Error("Record not found.");db_().getSheetByName(t).deleteRow(hit.__row);audit_(hit.institution_id||u.institution_id,u.user_id,"delete",t,id,"Record deleted");return{ok:true};}
function dashboard_(u){const c={};["Users","Groups","Subjects","Sessions","Exams","Complaints"].forEach(t=>{if(u.role==="owner"||canRead_(u,t)){let a=rows_(t);if(u.role!=="owner"&&TABLES[t].includes("institution_id"))a=a.filter(x=>x.institution_id===u.institution_id);c[t]=a.length;}});return{ok:true,counts:c,user:publicUser_(u)};}
function changePassword_(u,d){if(hash_(String(d.oldPassword||""),u.password_salt)!==u.password_hash)throw new Error("Current password is incorrect.");const p=String(d.newPassword||"");if(p.length<10)throw new Error("New password must be at least 10 characters.");const s=Utilities.getUuid(),sh=db_().getSheetByName("Users");sh.getRange(u.__row,8).setValue(hash_(p,s));sh.getRange(u.__row,9).setValue(s);audit_(u.institution_id,u.user_id,"change_password","User",u.user_id,"Password changed");return{ok:true};}
function audit_(i,a,act,t,id,d){try{const sh=db_().getSheetByName("Audit_Logs");if(sh)sh.appendRow([Utilities.getUuid(),i,a,act,t,id,d,new Date().toISOString()]);}catch(e){}}
