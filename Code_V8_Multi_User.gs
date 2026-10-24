/** RAJ GROUP V8 — five isolated users, one alphabetically ordered Master sheet.
 * Deploy via Manage Deployments > Edit > New Version.
 * IMPORTANT: keep legacy CloudState/Entries as backup; this script does not delete them.
 */
const USERS=['Darshil','Dipak','Shailesh','User 4','User 5']; // alphabetical order
const USER_STATES='UserStateV8';
const HEADERS=['User','Bill Number','Bill Remark','Sr. No','Code','Name','HSN','Unit','Qty','Rate ₹','Discount %','Discount Amount ₹','Scheme Discount %','Scheme Rs ₹','CD %','Limit ₹','T-Fit ₹','Status','Approvals','Item Remark','Manual Changes','Updated At'];
const COL={hsn:7,unit:8,rate:10,disc:11,discAmt:12,schDisc:13,schRs:14,cd:15};
const REQUIRED=['hsn','unit','rate','disc'];
function bookV8_(){const p=PropertiesService.getScriptProperties().getProperty('SHEET_ID');const b=p?SpreadsheetApp.openById(p):SpreadsheetApp.getActiveSpreadsheet();if(!b)throw Error('Bind this Apps Script to the Google Sheet or configure SHEET_ID');return b;}
function tabV8_(name){const b=bookV8_();return b.getSheetByName(name)||b.insertSheet(name);}
function userV8_(name){name=String(name||'');if(!USERS.includes(name))throw Error('Unknown user: '+name);return name;}
function chunksV8_(val){return JSON.stringify(val).match(/[\s\S]{1,30000}/g)||[''];}
function readV8_(user){const t=tabV8_(USER_STATES);const vals=t.getDataRange().getValues();const subset=vals.filter(r=>String(r[0])===user);if(!subset.length)return {data:{bills:[],other:[]},token:''};subset.sort((a,b)=>Number(a[1])-Number(b[1]));return {data:JSON.parse(subset.map(r=>r[3]).join('')),token:String(subset[0][2]||'')};}
function writeUserV8_(user,data,token){const t=tabV8_(USER_STATES);const old=t.getDataRange().getValues().filter(r=>String(r[0])!==user&&r.some(c=>c!==''));const newRows=chunksV8_(data).map((piece,i)=>[user,i,token,piece]);const rows=old.concat(newRows);t.clearContents();if(rows.length)t.getRange(1,1,rows.length,4).setValues(rows);if(!t.isSheetHidden())t.hideSheet();}
function nV8_(v){const n=Number.parseFloat(String(v??'').replace(/[,₹%\s]/g,''));return Number.isFinite(n)?n:0;}
function calcV8_(r){let x=Math.max(0,nV8_(r.rate));x=Math.max(0,x*(1-nV8_(r.disc)/100));x=Math.max(0,x-nV8_(r.discAmt));x=Math.max(0,x*(1-nV8_(r.schDisc)/100));x=Math.max(0,x-nV8_(r.schRs));x=Math.max(0,x*(1-nV8_(r.cd)/100));const a=Math.round((x+Number.EPSILON)*100)/100;return [a,Math.max(0,Math.round((a-.1+Number.EPSILON)*100)/100)];}
function colorV8_(r){let count=Number(r.approvals)||0;if(!r.status)return count===0?'#f49aa8':count===1?'#ffc0c5':'#ffdee2';return count===1?'#b8eccb':count===2?'#ffcf91':count===3?'#ffe36f':'#e5a52c';}
function rebuildV8_(){const master=tabV8_('Master');let rows=[],colors=[],notes=[],stamp=new Date();
 USERS.forEach(user=>{const data=readV8_(user).data;(data.bills||[]).forEach(b=>(b.rows||[]).forEach(r=>{
 const [lim,fit]=calcV8_(r),orig=r.original||{},modified=Object.keys(COL).filter(k=>String(r[k]??'')!==String(orig[k]??''));
 rows.push([user,b.number||'',b.billRemark||'',r.sr??'',r.code??'',r.name??'',r.hsn??'',r.unit??'',r.qty??'',r.rate??'',r.disc??'',r.discAmt??'',r.schDisc??'',r.schRs??'',r.cd??'',lim,fit,r.status?'YES':'NO',Number(r.approvals)||0,r.itemRemark||'',modified.join(', '),stamp]);
 const c=Array(HEADERS.length).fill('#ffffff'),note=Array(HEADERS.length).fill('');
 Object.keys(COL).forEach(k=>{const ix=COL[k]-1,old=String(orig[k]??''),cur=String(r[k]??'');if(REQUIRED.includes(k)&&!cur.trim()){c[ix]='#ffd1d1';note[ix]='Required field blank';}else if(old!==cur){c[ix]='#ffe1e1';note[ix]='Changed from: '+(old||'(blank)');}});
 c[17]=r.status?'#daf4e6':'#ffe5e8';c[18]=colorV8_(r);colors.push(c);notes.push(note);
 }));(data.other||[]).forEach(r=>{rows.push([user,'','', '',r.part||'', '', '', '', '', '', '', '', '', '', '', '', '',r.status?'YES':'NO',Number(r.approvals)||0,r.remark||'','',stamp]);const c=Array(HEADERS.length).fill('#ffffff');c[17]=r.status?'#daf4e6':'#ffe5e8';c[18]=colorV8_(r);colors.push(c);notes.push(Array(HEADERS.length).fill(''));});});
 const last=Math.max(master.getLastRow(),1),width=HEADERS.length;if(last>1){const old=master.getRange(2,1,last-1,width);old.clearContent();old.setBackground('#ffffff');old.clearNote();}
 master.getRange(1,1,1,width).setValues([HEADERS]).setBackground('#165bb7').setFontColor('#ffffff').setFontWeight('bold');if(rows.length){const range=master.getRange(2,1,rows.length,width);range.setValues(rows);range.setBackgrounds(colors);range.setNotes(notes);}master.setFrozenRows(1);
 // Five human-readable individual tabs, derived from the same authoritative snapshots.
 const userCol=HEADERS.length;
 USERS.forEach(user=>{let t=tabV8_(user);let indices=rows.map((r,i)=>r[0]===user?i:-1).filter(i=>i!==-1),values=indices.map(i=>rows[i]);const last=Math.max(t.getLastRow(),1);if(last>1){const old=t.getRange(2,1,last-1,userCol);old.clearContent();old.setBackground('#ffffff');old.clearNote();}t.getRange(1,1,1,userCol).setValues([HEADERS]).setBackground('#165bb7').setFontColor('#ffffff').setFontWeight('bold');if(values.length){const r=t.getRange(2,1,values.length,userCol);r.setValues(values);r.setBackgrounds(indices.map(i=>colors[i]));r.setNotes(indices.map(i=>notes[i]));}t.setFrozenRows(1);});
}
function responseV8_(obj,cb){let result=JSON.stringify(obj);if(cb){if(!/^[a-zA-Z_$][\w$]*$/.test(cb))throw Error('Bad callback');result=cb+'('+result+');';}return ContentService.createTextOutput(result).setMimeType(cb?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);}
function doGet(e){const cb=String(e?.parameter?.callback||'');try{const user=userV8_(e?.parameter?.user);return responseV8_({ok:true,...readV8_(user)},cb);}catch(err){return responseV8_({ok:false,error:String(err.message||err)},cb);}}
function doPost(e){const lock=LockService.getScriptLock();try{const body=JSON.parse(e.postData.contents);if(body.action!=='saveUser')throw Error('Unsupported action');const user=userV8_(body.user);if(!body.token||!body.data||!Array.isArray(body.data.bills))throw Error('Invalid payload');lock.waitLock(30000);const current=readV8_(user);if(String(body.baseToken||'')!==String(current.token||''))throw Error('Conflict: this user was edited elsewhere. Reload.');writeUserV8_(user,{bills:body.data.bills,other:Array.isArray(body.data.other)?body.data.other:[]},String(body.token));rebuildV8_();SpreadsheetApp.flush();return responseV8_({ok:true,token:body.token});}catch(err){return responseV8_({ok:false,error:String(err.message||err)});}finally{if(lock.hasLock())lock.releaseLock();}}

/** Run once manually from Apps Script editor to create Master and all five user tabs. */
function setupV8(){const lock=LockService.getScriptLock();lock.waitLock(30000);try{rebuildV8_();SpreadsheetApp.flush();}finally{lock.releaseLock();}}
