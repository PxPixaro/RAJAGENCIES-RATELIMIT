/** RAJ GROUP V16 — Complete V12-based multi-user backend with AKKI changed-cell display fix.
 * Replace old Code.gs, run setupV10 once, then redeploy EXISTING web app as New Version.
 * Employee deletion removes only their active sheet; Master records remain until manually deleted in Google Sheet.
 */
const USERS=['Darshil','Dipak','Shailesh','User 4','User 5'];
const USER_STATES='UserStateV8',SEEN_TAB='_MasterSeenV10';
const HEADERS=['User','Bill Number','Bill Remark','Sr. No','Code','Name','HSN','Unit','Qty','Rate ₹','Discount %','Discount Amount ₹','Scheme Discount %','Scheme Rs ₹','CD %','Limit ₹','T-Fit ₹','Status','Approvals','Item Remark','Manual Changes','Updated At','Bill Type','Row ID','Bill ID','Bill Date'];
const EDIT_COL={hsn:7,unit:8,rate:10,disc:11,discAmt:12,schDisc:13,schRs:14,cd:15},REQUIRED=['hsn','unit','rate','disc'];
function book_(){const p=PropertiesService.getScriptProperties().getProperty('SHEET_ID');const b=p?SpreadsheetApp.openById(p):SpreadsheetApp.getActiveSpreadsheet();if(!b)throw Error('Bind to spreadsheet or set SHEET_ID');return b;}
function tab_(name){const b=book_();return b.getSheetByName(name)||b.insertSheet(name);}
function user_(u){u=String(u||'');if(!USERS.includes(u))throw Error('Unknown employee: '+u);return u;}
function split_(v){return JSON.stringify(v).match(/[\s\S]{1,30000}/g)||[''];}
function read_(user){let t=tab_(USER_STATES),rows=t.getDataRange().getValues().filter(r=>String(r[0])===user).sort((a,b)=>Number(a[1])-Number(b[1]));return rows.length?{data:JSON.parse(rows.map(r=>r[3]).join('')),token:String(rows[0][2]||'')}:{data:{bills:[],other:[]},token:''};}
function write_(user,data,token){let t=tab_(USER_STATES),old=t.getDataRange().getValues().filter(r=>String(r[0])!==user&&r.some(x=>x!=='')),rows=old.concat(split_(data).map((chunk,i)=>[user,i,token,chunk]));t.clearContents();if(rows.length)t.getRange(1,1,rows.length,4).setValues(rows);if(!t.isSheetHidden())t.hideSheet();}
function num_(v){let n=parseFloat(String(v??'').replace(/[,₹%\s]/g,''));return Number.isFinite(n)?n:0;}
function calc_(r){let a=Math.max(0,num_(r.rate));a=Math.max(0,a*(1-num_(r.disc)/100));a=Math.max(0,a-num_(r.discAmt));a=Math.max(0,a*(1-num_(r.schDisc)/100));a=Math.max(0,a-num_(r.schRs));a=Math.max(0,a*(1-num_(r.cd)/100));a=Math.round((a+Number.EPSILON)*100)/100;return [a,Math.max(0,Math.round((a-.1+Number.EPSILON)*100)/100)];}
function approvalColor_(r){const n=Number(r.approvals)||0;return !r.status?(n===0?'#f49aa8':n===1?'#ffc0c5':'#ffdee2'):(n===1?'#b8eccb':n===2?'#ffcf91':n===3?'#ffe36f':'#e5a52c');}
function rowData_(user,b,r){const [lim,fit]=calc_(r),orig=r.original||{},modified=Object.keys(EDIT_COL).filter(k=>String(r[k]??'')!==String(orig[k]??''));const stamp=b.createdAt||new Date().toISOString();return [user,b.number||'',b.billRemark||'',r.sr??'',r.code??'',r.name??'',r.hsn??'',r.unit??'',r.qty??'',r.rate??'',r.disc??'',r.discAmt??'',r.schDisc??'',r.schRs??'',r.cd??'',lim,fit,r.status?'YES':'NO',Number(r.approvals)||0,r.itemRemark||'',modified.join(', '),stamp,b.billType||'',r.uid||'',b.id||'',String(stamp).slice(0,10)];}
function style_(r){const colors=Array(HEADERS.length).fill('#fff'),notes=Array(HEADERS.length).fill(''),o=r.original||{};Object.keys(EDIT_COL).forEach(k=>{let ix=EDIT_COL[k]-1,old=String(o[k]??''),v=String(r[k]??'');if(REQUIRED.includes(k)&&!v.trim()){colors[ix]='#ffd1d1';notes[ix]='Required field blank';}else if(old!==v){colors[ix]='#ffe1e1';notes[ix]='Changed from: '+(old||'(blank)');}});colors[17]=r.status?'#daf4e6':'#ffe5e8';colors[18]=approvalColor_(r);return {colors,notes};}
function headers_(t){t.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setBackground('#165bb7').setFontColor('#fff').setFontWeight('bold');t.setFrozenRows(1);}
function flat_(user,data){let out=[];(data.bills||[]).forEach(b=>(b.rows||[]).forEach(r=>out.push({user,b,r})));return out;}
function unique_(user,b,r,i){if(!b.id)b.id='legacy-bill-'+user+'-'+i;if(!r.uid)r.uid='legacy-row-'+user+'-'+b.id+'-'+i+'-'+String(r.code||'');if(!b.createdAt)b.createdAt=new Date().toISOString();}
function norm_(user,data){(data.bills||[]).forEach((b,bi)=>(b.rows||[]).forEach((r,ri)=>unique_(user,b,r,bi+'-'+ri)));return data;}
function seen_(){const t=tab_(SEEN_TAB);const values=t.getLastRow()?t.getRange(1,1,t.getLastRow(),1).getValues().flat().filter(Boolean).map(String):[];if(!t.isSheetHidden())t.hideSheet();return new Set(values);}
function markSeen_(ids){if(!ids.length)return;const t=tab_(SEEN_TAB);t.getRange(t.getLastRow()+1,1,ids.length,1).setValues(ids.map(x=>[x]));if(!t.isSheetHidden())t.hideSheet();}
function refreshUsers_(){USERS.forEach(user=>{const data=read_(user).data,t=tab_(user),flat=flat_(user,data),last=t.getLastRow();if(last>1){const g=t.getRange(2,1,last-1,HEADERS.length);g.clearContent();g.setBackground('#fff');g.clearNote();}headers_(t);if(flat.length){const rr=t.getRange(2,1,flat.length,HEADERS.length);rr.setValues(flat.map(x=>rowData_(user,x.b,x.r)));rr.setBackgrounds(flat.map(x=>style_(x.r).colors));rr.setNotes(flat.map(x=>style_(x.r).notes));}});}
// Faster saves: update just the employee being edited, not five tabs on every request.
function refreshOneUser_(user){
 const data=read_(user).data,t=tab_(user),flat=flat_(user,data),last=t.getLastRow();
 if(last>1){const old=t.getRange(2,1,last-1,HEADERS.length);old.clearContent();old.setBackground('#fff');old.clearNote();}
 headers_(t);
 if(flat.length){const r=t.getRange(2,1,flat.length,HEADERS.length);r.setValues(flat.map(x=>rowData_(user,x.b,x.r)));r.setBackgrounds(flat.map(x=>style_(x.r).colors));r.setNotes(flat.map(x=>style_(x.r).notes));}
}
// 48h expiry applies ONLY to active employee snapshots. Master history is never purged.
function expireEmployeeBillsV11(){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const cutoff=Date.now()-48*60*60*1000;
  USERS.forEach(user=>{
   const state=read_(user),data=state.data;
   const originalBills=(data.bills||[]).slice();
   const before=originalBills.length;
   data.bills=(data.bills||[]).filter(b=>{
    const time=Date.parse(b.createdAt||'');
    return !Number.isFinite(time)||time>cutoff;
   });
   if(data.bills.length!==before){
    // Archive any not-yet-indexed data before clearing active seat.
    syncMaster_(user,{bills:originalBills});
    write_(user,data,'expiry-'+Date.now()+'-'+Math.random().toString(36).slice(2));
    refreshOneUser_(user);
   }
  });
  SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
}
function installV11ExpiryTrigger(){
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='expireEmployeeBillsV11').forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('expireEmployeeBillsV11').timeBased().everyHours(1).create();
}
function setupV11(){setupV10();installV11ExpiryTrigger();}
function setupV12(){setupV11();tab_(EVENTS_V12);}
// Historical Master is incrementally updated. A deleted employee bill remains.
// A row manually removed from Master never reappears because its ID remains in _MasterSeenV10.
function syncMaster_(user,data){const t=tab_('Master');headers_(t);const seen=seen_(),last=t.getLastRow();const existing=last>1?t.getRange(2,24,last-1,1).getValues().flat():[],positions=new Map(existing.map((x,i)=>[String(x),i+2]).filter(x=>x[0]));const inserted=[];
 flat_(user,data).forEach(({b,r},ix)=>{unique_(user,b,r,ix);const uid=String(r.uid),entry=rowData_(user,b,r),format=style_(r);if(positions.has(uid)){let at=positions.get(uid),range=t.getRange(at,1,1,HEADERS.length);range.setValues([entry]);range.setBackgrounds([format.colors]);range.setNotes([format.notes]);}else if(!seen.has(uid)){const at=t.getLastRow()+1,range=t.getRange(at,1,1,HEADERS.length);range.setValues([entry]);range.setBackgrounds([format.colors]);range.setNotes([format.notes]);seen.add(uid);positions.set(uid,at);inserted.push(uid);}});
 markSeen_(inserted); // Not rewriting all Master rows preserves historical items and manual deletions.
}
// V16: reconstruct the original pasted value from saved Master cell notes.
// Older records that only have column U keep the changed-field flag, so their
// AKKI view is highlighted even if the previous value was not recorded.
function masterState_(){
 const t=tab_('Master'),n=t.getLastRow(),bills=[],groups=new Map();
 // V20: Master tab remains permanent history, while AKKI displays ACTIVE bill IDs only.
 // Employee delete (or 48h expiry) removes bill from employee snapshot, not Master rows.
 const activeBills=new Set();
 USERS.forEach(user=>{
   const state=read_(user).data;
   (state.bills||[]).forEach(b=>activeBills.add(user+'|'+String(b.id)));
 });
 if(n>1){
  const grid=t.getRange(2,1,n-1,HEADERS.length);
  const data=grid.getValues(),notes=grid.getNotes();
  const fields=Object.keys(EDIT_COL);
  data.forEach((a,i)=>{
   const owner=String(a[0]||''),uid=String(a[23]||'').trim();
   if(!USERS.includes(owner)||!uid)return;
   const bid=String(a[24]||('historic-'+i)),key=owner+'|'+bid;
   if(!activeBills.has(key))return; // Keep archived rows in Master Google Sheet, hide from AKKI dashboard.
   const createdAt=a[25] instanceof Date?Utilities.formatDate(a[25],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(a[25]||a[21]||'');
   let b=groups.get(key);
   if(!b){b={id:bid,owner,number:String(a[1]||''),billRemark:String(a[2]||''),billType:String(a[22]||''),createdAt,rows:[]};groups.set(key,b);bills.push(b);}
   const r={uid,owner,sr:a[3],code:a[4],name:a[5],hsn:a[6],unit:a[7],qty:a[8],rate:a[9],disc:a[10],discAmt:a[11],schDisc:a[12],schRs:a[13],cd:a[14],status:String(a[17]).toUpperCase()==='YES',approvals:Number(a[18])||0,itemRemark:String(a[19]||'')};
   r.original={};
   const changed=new Set(String(a[20]||'').split(',').map(x=>x.trim()).filter(Boolean));
   fields.forEach(k=>{
    let v=r[k],note=String(notes[i][EDIT_COL[k]-1]||'');
    const match=note.match(/^Changed from:\s*([\s\S]*)$/);
    if(match){
     const old=match[1]==='(blank)'?'':match[1];
     r.original[k]=old;
     // The Master original value is not required for calculations, only for changed-cell highlighting.
    }else if(changed.has(k)){
     // Old notes unavailable. Use a sentinel unequal to any current value.
     r.original[k]='__RAJ_GROUP_ORIGINAL_CHANGED__';
    }else r.original[k]=v;
   });
   b.rows.push(r);
  });
 }
 bills.sort((a,b)=>String(a.createdAt).localeCompare(String(b.createdAt))||a.owner.localeCompare(b.owner));
 const signature=bills.map(b=>[b.id,b.number,b.billRemark,b.billType,b.rows.map(r=>[r.uid,r.status,r.approvals,r.rate,r.disc,r.hsn,r.unit,r.itemRemark,r.original])]);
 const token=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,JSON.stringify(signature))).slice(0,24);
 return {data:{bills,other:[]},token};
}
// V16: AKKI edits update just the affected row rather than rebuilding the entire employee tab.
function patchAkki_(body){
 const user=user_(body.owner),uid=String(body.uid||''),changes=body.changes||{};
 const allowed=['status','approvals','itemRemark','hsn','unit','rate','disc','discAmt','schDisc','schRs','cd'];
 if(!uid||!Object.keys(changes).length||Object.keys(changes).some(k=>!allowed.includes(k)))throw Error('Invalid patch');
 const master=tab_('Master'),last=master.getLastRow();
 if(last<2)throw Error('Master is empty');
 const ids=master.getRange(2,24,last-1,1).getValues().flat(),ix=ids.findIndex(v=>String(v)===uid);
 if(ix<0)throw Error('Item not found in Master');
 const at=ix+2;
 if(String(master.getRange(at,1).getValue())!==user)throw Error('Owner mismatch');
 const state=read_(user),rows=flat_(user,state.data),hit=rows.findIndex(x=>String(x.r.uid)===uid);
 if(hit>=0){
  const item=rows[hit];Object.assign(item.r,changes);
  // Only the active employee's snapshot and the edited record are rewritten.
  write_(user,state.data,'akki-'+Date.now()+'-'+Utilities.getUuid().slice(0,8));
  const vals=rowData_(user,item.b,item.r),format=style_(item.r);
  const target=master.getRange(at,1,1,HEADERS.length);
  target.setValues([vals]);target.setBackgrounds([format.colors]);target.setNotes([format.notes]);
  const employeeTab=tab_(user),employeeAt=hit+2;
  const employeeRange=employeeTab.getRange(employeeAt,1,1,HEADERS.length);
  employeeRange.setValues([vals]);employeeRange.setBackgrounds([format.colors]);employeeRange.setNotes([format.notes]);
 }else{
  // Historic bills remain editable in Master only, without recreating deleted employee bills.
  const map={status:18,approvals:19,itemRemark:20,hsn:7,unit:8,rate:10,disc:11,discAmt:12,schDisc:13,schRs:14,cd:15};
  Object.entries(changes).forEach(([k,v])=>master.getRange(at,map[k]).setValue(k==='status'?(v?'YES':'NO'):v));
  const row=master.getRange(at,1,1,HEADERS.length),a=row.getValues()[0],notes=row.getNotes()[0];
  const changed=new Set(String(a[20]||'').split(',').map(x=>x.trim()).filter(Boolean));
  for(const k of Object.keys(EDIT_COL)){
   if(!(k in changes))continue;
   const col=EDIT_COL[k]-1,oldNote=String(notes[col]||'');
   if(!oldNote.startsWith('Changed from:')){
    changed.add(k);
    notes[col]='Changed from: (historical original unavailable)';
   }
  }
  a[20]=Array.from(changed).join(', ');
  const r={rate:a[9],disc:a[10],discAmt:a[11],schDisc:a[12],schRs:a[13],cd:a[14],status:String(a[17]).toUpperCase()==='YES',approvals:Number(a[18])||0};
  const [lim,fit]=calc_(r);a[15]=lim;a[16]=fit;
  row.setValues([a]);row.setNotes([notes]);
  const bg=row.getBackgrounds()[0];
  bg[17]=r.status?'#daf4e6':'#ffe5e8';bg[18]=approvalColor_(r);
  Object.keys(EDIT_COL).forEach(k=>{if(changed.has(k))bg[EDIT_COL[k]-1]='#ffe1e1';});
  row.setBackgrounds([bg]);
 }
}
// V12: Cloud notification event feed. Polling works while the website is open.
const EVENTS_V12='_NotificationsV12';
function recordEventsV12_(events){
 if(!events.length)return;
 const t=tab_(EVENTS_V12),rows=events.map(ev=>[new Date(),Utilities.getUuid(),ev.target,ev.kind,ev.owner,ev.bill,ev.message]);
 t.getRange(t.getLastRow()+1,1,rows.length,7).setValues(rows);
 if(!t.isSheetHidden())t.hideSheet();
}
function getEventsV12_(target,since){
 if(target!=='AKKI')user_(target);
 const t=tab_(EVENTS_V12),n=t.getLastRow();
 const cursor=Math.max(0,Number(since)||0);
 if(n<1)return {cursor:0,events:[]};
 // For first connection, send recent events; for subsequent polling use exact row cursor.
 const first=cursor>0?Math.max(cursor+1,n-249):Math.max(1,n-99);
 if(first>n)return {cursor:n,events:[]};
 const rows=t.getRange(first,1,n-first+1,7).getValues(),events=[];
 rows.forEach((row,i)=>{
  if(String(row[2])!==target)return;
  const stamp=row[0] instanceof Date?row[0].getTime():Date.parse(row[0]);
  events.push({cursor:first+i,time:stamp||0,id:String(row[1]),kind:String(row[3]),owner:String(row[4]),bill:String(row[5]),message:String(row[6])});
 });
 return {cursor:n,events:events.slice(-40)};
}
function eventDetailsV22_(oldRow,newRow){
 const labels={hsn:'HSN',unit:'Unit',rate:'Rate',disc:'Discount %',discAmt:'Discount Amt',schDisc:'Scheme Discount %',schRs:'Scheme Rs',cd:'CD %',itemRemark:'Item Remark',status:'Status',approvals:'Approval'};
 const fields=['hsn','unit','rate','disc','discAmt','schDisc','schRs','cd','itemRemark','status','approvals'];
 const result=[];
 fields.forEach(k=>{
  const left=String(oldRow?.[k]??''),right=String(newRow?.[k]??'');
  if(oldRow&&left===right)return;
  if(!oldRow && !['rate','disc','itemRemark'].includes(k))return;
  if(!oldRow && right==='')return;
  const before=oldRow?(left||'blank'):'new';
  const after=(right||'blank');
  result.push(labels[k]+': '+before+' → '+after);
 });
 return result;
}
function diffEventsV12_(user,before,after){
 const old=new Map((before.bills||[]).map(b=>[String(b.id),b]));const ev=[];
 (after.bills||[]).forEach(b=>{
  const prev=old.get(String(b.id)),bill=String(b.number||('Bill '+b.id));
  if(!prev){
   if((b.rows||[]).length||b.number){
    const codes=(b.rows||[]).slice(0,4).map(r=>String(r.code||r.name||r.sr||'Item')).join(', ');
    ev.push({target:'AKKI',kind:'new',owner:user,bill,message:user+' ने नया Bill '+bill+' जोड़ा ('+(b.rows||[]).length+' items)'+(codes?' · '+codes:'')});
   }return;
  }
  const oldRows=new Map((prev.rows||[]).map(r=>[String(r.uid),r]));
  const updates=[];
  (b.rows||[]).forEach(r=>{
   const p=oldRows.get(String(r.uid));const changes=eventDetailsV22_(p,r);
   if(changes.length){
    const code=String(r.code||r.name||r.sr||'Item');
    updates.push(code+' — '+changes.join('; '));
   }
  });
  if(String(b.number||'')!==String(prev.number||''))updates.push('Bill Number: '+String(prev.number||'')+' → '+String(b.number||''));
  if(String(b.billRemark||'')!==String(prev.billRemark||''))updates.push('Bill Remark: '+String(prev.billRemark||'')+' → '+String(b.billRemark||''));
  if(String(b.billType||'')!==String(prev.billType||''))updates.push('Bill Type: '+String(prev.billType||'')+' → '+String(b.billType||''));
  if(updates.length){
   // Keep notification cell concise, but show the actual changed parts and values.
   const detailed=updates.slice(0,8).join(' | ')+(updates.length>8?' | +'+(updates.length-8)+' more':'');
   ev.push({target:'AKKI',kind:'changed',owner:user,bill,message:user+' · '+bill+' · '+detailed.slice(0,1800)});
  }
 });
 return ev;
}
function sendAdminMessageV22_(body){
 if(String(body.sender)!=='AKKI')throw Error('Only AKKI can send messages');
 const message=String(body.message||'').trim();
 if(!message||message.length>1000)throw Error('Message must be 1–1000 characters');
 const target=String(body.target||'');
 const targets=target==='ALL'?USERS.slice():[user_(target)];
 recordEventsV12_(targets.map(u=>({target:u,kind:'admin-message',owner:'AKKI',bill:'',message:'AKKI: '+message})));
 return {ok:true,count:targets.length};
}
function response_(obj,callback){let s=JSON.stringify(obj);if(callback){if(!/^[a-zA-Z_$][\w$]*$/.test(callback))throw Error('Bad callback');s=callback+'('+s+');';}return ContentService.createTextOutput(s).setMimeType(callback?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);}
function doGet(e){const p=e?.parameter||{},cb=String(p.callback||'');try{if(p.action==='notifications'){return response_({ok:true,...getEventsV12_(String(p.user||''),Number(p.since)||0)},cb);}if(p.user==='AKKI')return response_({ok:true,...masterState_()},cb);let u=user_(p.user);return response_({ok:true,...read_(u)},cb);}catch(err){return response_({ok:false,error:String(err.message||err)},cb);}}
function doPost(e){const lock=LockService.getScriptLock();try{const body=JSON.parse(e.postData.contents);lock.waitLock(30000);if(body.action==='adminMessage'){const result=sendAdminMessageV22_(body);SpreadsheetApp.flush();return response_(result);} if(body.action==='deleteMasterBill'){deleteMasterBillV19_(body);SpreadsheetApp.flush();return response_({ok:true});}if(body.action==='patchAkki'){patchAkki_(body);const owner=user_(body.owner);const master=tab_('Master');const uids=master.getLastRow()>1?master.getRange(2,24,master.getLastRow()-1,1).getValues().flat():[];const at=uids.findIndex(x=>String(x)===String(body.uid));const bill=at>=0?String(master.getRange(at+2,2).getValue()):'';recordEventsV12_([{target:owner,kind:'approval',owner:'AKKI',bill,message:'AKKI ने Bill '+(bill||'')+' का '+(body.changes?.status?'Approval ✓':'Status / Rate')+' अपडेट किया'}]);SpreadsheetApp.flush();return response_({ok:true});}if(body.action!=='saveUser')throw Error('Unknown action');const u=user_(body.user);if(!body.token||!body.data||!Array.isArray(body.data.bills))throw Error('Invalid payload');const current=read_(u);if(String(body.baseToken||'')!==String(current.token||''))throw Error('Conflict; reload user');const data=norm_(u,{bills:body.data.bills,other:Array.isArray(body.data.other)?body.data.other:[]});write_(u,data,String(body.token));recordEventsV12_(diffEventsV12_(u,current.data,data));syncMaster_(u,data);refreshOneUser_(u);SpreadsheetApp.flush();return response_({ok:true,token:body.token});}catch(err){return response_({ok:false,error:String(err.message||err)});}finally{if(lock.hasLock())lock.releaseLock();}}
function setupV10(){const lock=LockService.getScriptLock();lock.waitLock(30000);try{let master=tab_('Master');headers_(master);const seen=seen_();const old=master.getLastRow()>1?master.getRange(2,1,master.getLastRow()-1,HEADERS.length).getValues():[];// Link rows created by V8/V9 (before Row ID existed) to active snapshots when possible.
 const reserved=new Set();USERS.forEach(user=>{const current=read_(user).data;norm_(user,current);flat_(user,current).forEach(({b,r})=>{for(let i=0;i<old.length;i++){const row=old[i];if(reserved.has(i)||String(row[23]||'')||String(row[0])!==user||String(row[1]||'')!==String(b.number||'')||String(row[4]||'')!==String(r.code||'')||String(row[3]||'')!==String(r.sr||''))continue;const at=i+2;master.getRange(at,24,1,3).setValues([[r.uid,b.id,b.createdAt.slice(0,10)]]);row[23]=r.uid;reserved.add(i);break;}});});
 const previous=old.map(x=>String(x[23]||'')).filter(Boolean);markSeen_(previous.filter(id=>!seen.has(id)));USERS.forEach(user=>{const state=read_(user);const data=norm_(user,state.data);write_(user,data,state.token||'v10-init-'+Date.now());syncMaster_(user,data);});refreshUsers_();SpreadsheetApp.flush();}finally{lock.releaseLock();}}
  
/** V19: Delete AKKI master bill and active employee copy, without restoring it on later saves. */
function deleteMasterBillV19_(body){
 const user=user_(body.owner),bid=String(body.billId||'');
 if(!bid)throw Error('Bill ID missing');
 const master=tab_('Master'),last=master.getLastRow();
 const rows=last>1?master.getRange(2,1,last-1,HEADERS.length).getValues():[];
 const indices=[],deletedIds=[];
 rows.forEach((row,i)=>{if(String(row[0])===user&&String(row[24])===bid){indices.push(i+2);if(row[23])deletedIds.push(String(row[23]));}});
 if(!indices.length)throw Error('Master bill not found');
 // Record IDs BEFORE removing rows, preventing reappearance during a delayed employee save.
 const seen=seen_();markSeen_(deletedIds.filter(uid=>!seen.has(uid)));
 indices.sort((a,b)=>b-a).forEach(n=>master.deleteRow(n));
 // Remove current employee copy, if bill is still active (Master history could be older).
 const state=read_(user),data=state.data;
 const oldCount=(data.bills||[]).length;
 data.bills=(data.bills||[]).filter(b=>String(b.id)!==bid);
 if(data.bills.length!==oldCount){
  write_(user,data,'master-delete-'+Date.now()+'-'+Utilities.getUuid());
  refreshOneUser_(user);
 }
}
