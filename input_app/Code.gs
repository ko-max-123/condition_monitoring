/**
 * コンディション記録 - Google Apps Script server
 * Deploy: Web app > Execute as me > Only myself.
 * 最初に setupApp() をエディタから1回実行してください。
 */
const CONFIG = Object.freeze({
  PROPERTY_KEY: 'CONDITION_SPREADSHEET_ID',
  SHEET_NAME: 'records',
  SPREADSHEET_NAME: 'コンディション観測ログ（非公開）',
  TIME_ZONE: 'Asia/Tokyo'
});
const RAW_FIELDS = Object.freeze([
  'date','mood','physical','anxiety','irritability','fatigue','focus','motivation','security',
  'sleepHours','sleepQuality','wakeFatigue','meals','exerciseMinutes','sunlightMinutes',
  'workload','socialStress','caffeineCups','medication','note','insight'
]);
const HEADERS = RAW_FIELDS.concat(['internalScore','externalScore','totalScore','sleepFlag','updatedAt']);
const SCORED = Object.freeze(['mood','physical','anxiety','irritability','fatigue','focus','motivation','security',
                              'sleepQuality','wakeFatigue','meals','workload','socialStress']);
const COUNTS = Object.freeze(['sleepHours','exerciseMinutes','sunlightMinutes','caffeineCups']);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('コンディション記録')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function setupApp() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(CONFIG.PROPERTY_KEY);
  if (existingId) {
    const existing = SpreadsheetApp.openById(existingId);
    ensureSheet_(existing);
    return {url:existing.getUrl(), created:false};
  }
  const book = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  const sheet = book.getSheets()[0];
  sheet.setName(CONFIG.SHEET_NAME);
  sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,HEADERS.length).setBackground('#17475b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.autoResizeColumns(1,HEADERS.length);
  props.setProperty(CONFIG.PROPERTY_KEY,book.getId());
  return {url:book.getUrl(), created:true};
}

function ensureSheet_(book) {
  let sheet=book.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet=book.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  const observed=sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  if (observed.some((h,i)=>h!==HEADERS[i])) {
    throw new Error('シートの列構成が想定と異なります。手動で列を変更していないか確認してください。');
  }
  return sheet;
}
function getSheet_() {
  const id=PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTY_KEY);
  if (!id) throw new Error('初期設定が未完了です。Apps Scriptエディタで setupApp() を実行してください。');
  return ensureSheet_(SpreadsheetApp.openById(id));
}
function today_(){return Utilities.formatDate(new Date(),CONFIG.TIME_ZONE,'yyyy-MM-dd');}
function isDate_(s){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d=new Date(s+'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0,10)===s;
}
function numeric_(value,min,max,integer,label) {
  if (value===null || value===undefined || value==='') return '';
  const n=Number(value);
  if(!Number.isFinite(n)||n<min||n>max||(integer&&!Number.isInteger(n))) {
    throw new Error(label+'の値が不正です（'+min+'〜'+max+'）。');
  }
  return n;
}
function validate_(input) {
  if(!input || typeof input!=='object') throw new Error('入力データが正しくありません。');
  const r={};
  r.date=String(input.date||'');
  if(!isDate_(r.date)) throw new Error('日付の形式が正しくありません。');
  SCORED.forEach(k=>r[k]=numeric_(input[k],1,10,true,k));
  r.sleepHours=numeric_(input.sleepHours,0,24,false,'睡眠時間');
  r.exerciseMinutes=numeric_(input.exerciseMinutes,0,1440,true,'運動時間');
  r.sunlightMinutes=numeric_(input.sunlightMinutes,0,1440,true,'日光時間');
  r.caffeineCups=numeric_(input.caffeineCups,0,50,false,'カフェイン');
  ['medication','note','insight'].forEach(k=>{
    r[k]=String(input[k]??'').slice(0,k==='medication'?120:2000);
    if(/^[=+@\-\t\r]/.test(r[k])) r[k]="'"+r[k]; // スプレッドシートの数式注入を防ぐ
  });
  return r;
}
const round1_=n=>Math.round((n+1e-9)*10)/10;
const present_=v=>v!==''&&v!==null&&v!==undefined&&Number.isFinite(Number(v));
function calc_(r){
  const intKeys=['mood','physical','anxiety','irritability','fatigue','focus','motivation','security'];
  const inversed=['anxiety','irritability','fatigue'];
  const ints=intKeys.filter(k=>present_(r[k])).map(k=>inversed.includes(k)?11-Number(r[k]):Number(r[k]));
  const ext=[];
  const push=(k,fn)=>{if(present_(r[k]))ext.push(fn(Number(r[k])));};
  push('sleepHours',n=>Math.max(0,10-Math.abs(n-7.5)*2));
  push('sleepQuality',n=>n);
  push('wakeFatigue',n=>11-n);
  push('meals',n=>n);
  push('exerciseMinutes',n=>Math.min(10,n/30*10));
  push('sunlightMinutes',n=>Math.min(10,n/15*10));
  push('workload',n=>11-n);
  push('socialStress',n=>11-n);
  push('caffeineCups',n=>Math.max(0,10-2*n));
  const internalScore=ints.length?round1_(ints.reduce((a,b)=>a+b,0)/ints.length):'';
  const externalScore=ext.length?round1_(ext.reduce((a,b)=>a+b,0)/ext.length):'';
  const available=[internalScore,externalScore].filter(present_);
  const totalScore=available.length?round1_(available.reduce((a,b)=>a+Number(b),0)/available.length):'';
  let sleepFlag='';
  if(present_(r.sleepHours)) sleepFlag=r.sleepHours<6?'不足':r.sleepHours>9?'過多':'OK';
  return {internalScore,externalScore,totalScore,sleepFlag};
}
function toSheetRow_(r) {
  const calc=calc_(r);
  return HEADERS.map(k=>k==='updatedAt'?new Date().toISOString():
    RAW_FIELDS.includes(k)?(r[k]??''):calc[k]);
}
function readAll_(){
  const sheet=getSheet_();
  const count=sheet.getLastRow()-1;
  if(count<=0)return [];
  return sheet.getRange(2,1,count,HEADERS.length).getValues().map(row=>{
    const obj={};HEADERS.forEach((k,i)=>obj[k]=row[i] instanceof Date?Utilities.formatDate(row[i],CONFIG.TIME_ZONE,'yyyy-MM-dd'):row[i]);
    return obj;
  }).filter(r=>isDate_(String(r.date))).sort((a,b)=>b.date.localeCompare(a.date));
}
function getInitialData(){
  const all=readAll_();
  return {today:today_(),recent:all.slice(0,160),count:all.length,
          spreadsheetUrl:SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(CONFIG.PROPERTY_KEY)).getUrl()};
}
function getRecord(date){
  if(!isDate_(String(date))) throw new Error('日付が不正です。');
  return readAll_().find(r=>r.date===date)||null;
}
function saveRecord(input){
  const record=validate_(input);
  const lock=LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet=getSheet_();
    const last=sheet.getLastRow();
    let target=last+1;
    if(last>1){
      const dates=sheet.getRange(2,1,last-1,1).getDisplayValues();
      const idx=dates.findIndex(row=>row[0]===record.date);
      if(idx>=0)target=idx+2;
    }
    sheet.getRange(target,1,1,HEADERS.length).setValues([toSheetRow_(record)]);
    SpreadsheetApp.flush();
    return {saved:true,date:record.date,updated:target<=last,record:Object.assign({},record,calc_(record))};
  }finally{lock.releaseLock();}
}
function importRecords(rows){
  if(!Array.isArray(rows)||rows.length>2000) throw new Error('一度に読み込めるのは2000件までです。');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const sheet=getSheet_();
    const existingDates=new Set(sheet.getLastRow()>1?
      sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues().map(r=>r[0]):[]);
    const incoming=new Set(), append=[];
    let skipped=0;
    for (const raw of rows){
      const r=validate_(raw);
      if(existingDates.has(r.date)||incoming.has(r.date)){skipped++;continue;}
      incoming.add(r.date);append.push(toSheetRow_(r));
    }
    if(append.length)sheet.getRange(sheet.getLastRow()+1,1,append.length,HEADERS.length).setValues(append);
    return {imported:append.length,skipped};
  }finally{lock.releaseLock();}
}
function exportRecords(){ return readAll_().sort((a,b)=>a.date.localeCompare(b.date)); }
