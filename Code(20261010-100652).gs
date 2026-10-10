/**
 * RAJ GROUP Rate Limit — Google Sheet Sync (V5)
 * Replace ALL previous Code.gs code with this file, save, then deploy a NEW VERSION
 * of the existing Web App deployment (same /exec URL).
 * Entries is a display/reporting sheet. Edit records through the HTML app.
 */
const ENTRIES_SHEET = 'Entries';
const STATE_SHEET = 'CloudState';
const BRANCHES = ['HO', 'HO Life', 'WHD', 'WHD Life'];
const CHUNK_SIZE = 35000;
const HEADERS = [
  'Branch','Bill No.','Sr. No','Code','Name','HSN','Unit','Qty',
  'Rate','Discount %','Discount Amount','Scheme Discount %',
  'Scheme ₹','CD %','Limit ₹','T-Fit Limit ₹','Status','Approvals',
  'Original Data Changed','Part Number (Other)','Remark (Other)','Updated At'
];
const WHITE = '#ffffff';
const CHANGED = '#ffe1e1'; // light-red when value differs from software paste
const MISSING = '#ffd1d1'; // required cell blank
const APPROVAL_COLORS = {
  1: '#b9ebcb', // first: green
  2: '#ffddb0', // second: light orange
  3: '#ffb768', // third: darker orange
  4: '#f28a35'  // fourth and higher: deep orange
};
const EDITABLE_COLS = { hsn:6, unit:7, rate:9, disc:10, discAmt:11, schDisc:12, schRs:13, cd:14 };
const REQUIRED = ['hsn','unit','rate','disc'];

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Spreadsheet missing: bind the script to the Google Sheet or set SHEET_ID.');
  return book;
}
function sheet_(name) {
  const ss = book_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function blank_() {
  return { version:4, branches:Object.fromEntries(BRANCHES.map(b=>[b,{bills:[]}])), other:[] };
}
function read_() {
  const sh = sheet_(STATE_SHEET);
  if (!sh.getLastRow()) return {data:blank_(),token:''};
  const values = sh.getRange(1,1,sh.getLastRow(),2).getValues();
  const token = String(values[0][0] || '');
  const json = values.map(x=>String(x[1] || '')).join('');
  return {data:json ? JSON.parse(json) : blank_(),token:token};
}
function numeric_(v) {
  const n = parseFloat(String(v == null ? '' : v).replace(/[,₹%\s]/g,''));
  return Number.isFinite(n) ? n : 0;
}
function money_(v) { return Math.max(0,Math.round((v + Number.EPSILON)*100)/100); }
function calculate_(r) {
  let value = Math.max(0,numeric_(r.rate));
  value = Math.max(0,value*(1-numeric_(r.disc)/100));
  value = Math.max(0,value-numeric_(r.discAmt));
  value = Math.max(0,value*(1-numeric_(r.schDisc)/100));
  value = Math.max(0,value-numeric_(r.schRs));
  value = money_(Math.max(0,value*(1-numeric_(r.cd)/100)));
  return [value,money_(value-.10)];
}
function approvalBackground_(r) {
  if (!r.status) return '#ffdce2'; // pending (red)
  return APPROVAL_COLORS[Math.min(4,Math.max(1,Number(r.approvals)||1))];
}
function write_(data,token) {
  const state = sheet_(STATE_SHEET);
  const json = JSON.stringify(data);
  const chunks = json.match(/.{1,35000}/gs) || [''];
  state.clearContents();
  state.getRange(1,1,chunks.length,2).setValues(chunks.map((part,i)=>[i===0?token:'',part]));
  if (!state.isSheetHidden()) state.hideSheet();

  const rows = [], backgrounds = [], notes = [];
  const stamp = new Date();
  BRANCHES.forEach(branch=>{
    const bills = data.branches?.[branch]?.bills || [];
    bills.forEach((bill, billIndex)=>(bill.rows||[]).forEach(r=>{
      const [limit,tFit] = calculate_(r);
      const original = r.original || {};
      const changedKeys = Object.keys(EDITABLE_COLS).filter(k=>String(r[k]??'')!==String(original[k]??''));
      const cells = [branch,billIndex+1,r.sr??'',r.code??'',r.name??'',r.hsn??'',r.unit??'',r.qty??'',
        r.rate??'',r.disc??'',r.discAmt??'',r.schDisc??'',r.schRs??'',r.cd??'',
        limit,tFit,r.status?'YES':'NO',Number(r.approvals)||0,changedKeys.length?'YES':'NO','','',stamp];
      rows.push(cells);
      const bg = Array(HEADERS.length).fill(WHITE);
      const note = Array(HEADERS.length).fill('');
      Object.keys(EDITABLE_COLS).forEach(k=>{
        const index = EDITABLE_COLS[k]-1;
        const oldValue = String(original[k]??'');
        const newValue = String(r[k]??'');
        if (REQUIRED.includes(k) && !newValue.trim()) {
          bg[index] = MISSING;
          note[index] = 'Required field is blank';
        } else if (newValue !== oldValue) {
          bg[index] = CHANGED;
          note[index] = 'Manually changed from: '+(oldValue || '(blank)');
        }
      });
      bg[16] = r.status ? '#daf4e6' : '#ffe5e8'; // Q: YES or NO
      bg[17] = approvalBackground_(r);             // R: 1 green, 2 orange, 3 darker, 4+ darkest
      backgrounds.push(bg);
      notes.push(note);
    }));
  });
  (data.other||[]).forEach(r=>{
    rows.push(['Other','','','','','','','','','','','','','','','',
      r.status?'YES':'NO',Number(r.approvals)||0,'',r.part??'',r.remark??'',stamp]);
    const bg = Array(HEADERS.length).fill(WHITE);
    bg[16]=r.status?'#daf4e6':'#ffe5e8';
    bg[17]=approvalBackground_(r);
    backgrounds.push(bg);
    notes.push(Array(HEADERS.length).fill(''));
  });

  const table = sheet_(ENTRIES_SHEET);
  const width = HEADERS.length;
  // Clear previous values AND backgrounds/notes: deleting a bill / All Clear removes old records.
  const oldLastRow = Math.max(table.getLastRow(),1);
  if (oldLastRow > 1) {
    const old = table.getRange(2,1,oldLastRow-1,width);
    old.clearContent();old.setBackground(WHITE);old.clearNote();
  }
  table.getRange(1,1,1,width).setValues([HEADERS])
    .setBackground('#165bb7').setFontColor(WHITE).setFontWeight('bold');
  if (rows.length) {
    const block = table.getRange(2,1,rows.length,width);
    block.setValues(rows);
    block.setBackgrounds(backgrounds);
    block.setNotes(notes);
  }
  table.setFrozenRows(1);
  // Entries is a report; CloudState is the source of truth.
}
function reply_(value,callback) {
  let body=JSON.stringify(value);
  if (callback && /^[a-zA-Z_$][\w$]*$/.test(callback)) body=callback+'('+body+');';
  return ContentService.createTextOutput(body)
    .setMimeType(callback?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);
}
function doGet(e) {
  const callback=String(e?.parameter?.callback||'');
  try {
    if (callback && !/^[a-zA-Z_$][\w$]*$/.test(callback)) throw Error('Invalid callback');
    const result=read_();
    return reply_({ok:true,data:result.data,token:result.token},callback);
  } catch(err) { return reply_({ok:false,error:String(err.message||err)},callback); }
}
function doPost(e) {
  const lock=LockService.getScriptLock();
  try {
    const body=JSON.parse(e.postData.contents);
    if (body.action!=='save'||!body.token||!body.data?.branches) throw Error('Invalid save request');
    lock.waitLock(25000);
    write_(body.data,String(body.token));
    SpreadsheetApp.flush();
    return reply_({ok:true,token:body.token});
  } catch(err) { return reply_({ok:false,error:String(err.message||err)}); }
  finally { if (lock.hasLock()) lock.releaseLock(); }
}
