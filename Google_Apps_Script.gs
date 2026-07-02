/**
 * URD Simulator — Google Sheets Backend (v4)
 * Adds: waves, attempt tracking, in-progress upserts, trainer names,
 * per-question time, per-question timings in decisions.
 *
 * Tabs:
 *   Scores    — one row per attempt (in-progress or completed).
 *               Uniquely keyed by `attemptId`; POSTs upsert.
 *   Batches   — one row per trainer-created batch.
 *   Waves     — one row per trainer-created wave.
 *   Day6_<ts> — global Day 6 CSV upload (legacy).
 *   Batch_<id>— per-batch question set.
 */

const SCORES_SHEET  = 'Scores';
const BATCHES_SHEET = 'Batches';
const WAVES_SHEET   = 'Waves';
const DAY6_TAB_PREFIX  = 'Day6_';
const BATCH_TAB_PREFIX = 'Batch_';
const ACTIVE_DAY6_PROP = 'ACTIVE_DAY6_TAB';

// Scores columns (order preserved; missing ones auto-added at end).
const SCORE_HEADERS = [
  'attemptId','datetime','startedAt','completedAt','status',
  'empId','name','country','trainer','wave','batch','batchId',
  'day','dayName','attemptNumber',
  'score','pct','correct','total','result','htqSeconds','totalSeconds',
  'decisions'
];

const BATCH_HEADERS = [
  'batchId','batchName','trainerName','waveCode',
  'startDate','endDate','createdAt','fileName',
  'questionCount','timePerQuestionSec','activeTab','createdBy'
];

const WAVE_HEADERS = [
  'waveId','waveCode','waveName','trainerName','startDate','endDate','createdAt','createdBy'
];

const QUESTION_HEADERS = [
  'id','title','body','correct_action','correct_reason',
  'explanation_correct','explanation_wrong','rating',
  'property_name','experience_type','date','policy_name','owner_response'
];

/* ═══ ROUTING ═══ */
function doPost(e){
  try{
    const body=JSON.parse(e.postData.contents);
    if(body&&body.type==='upload_day6') return handleDay6Upload_(body);
    if(body&&body.type==='create_batch') return handleCreateBatch_(body);
    if(body&&body.type==='create_wave')  return handleCreateWave_(body);
    return handleScoreSave_(body);
  }catch(err){return jsonResponse_({ok:false,error:String(err)});}
}

function doGet(e){
  try{
    const action=(e&&e.parameter&&e.parameter.action)||'';
    if(action==='get_day6')       return handleDay6Fetch_();
    if(action==='day6_status')    return handleDay6Status_();
    if(action==='list_batches')   return handleListBatches_();
    if(action==='get_batch')      return handleGetBatch_(e.parameter.id);
    if(action==='list_waves')     return handleListWaves_();
    if(action==='list_trainers')  return handleListTrainers_();
    return handleScoresFetch_();
  }catch(err){return jsonResponse_({error:String(err)});}
}

/* ═══ SCORES (with upsert-by-attemptId) ═══ */
function handleScoreSave_(body){
  const sheet=getOrCreateScoresSheet_();
  ensureColumns_(sheet,SCORE_HEADERS);
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const attemptId=body.attemptId||'';
  const status=body.status||'completed';
  // Attempt number: count PRIOR completed attempts for same empId+day+batchId
  let attemptNumber=Number(body.attemptNumber)||0;
  if(!attemptNumber){
    attemptNumber=computeAttemptNumber_(sheet,body.empId,body.day,body.batchId,attemptId);
  }

  const rowData=headers.map(function(h){
    switch(h){
      case 'attemptId':     return attemptId;
      case 'datetime':      return body.datetime||new Date().toLocaleString();
      case 'startedAt':     return body.startedAt||'';
      case 'completedAt':   return status==='completed'?(body.completedAt||new Date().toLocaleString()):'';
      case 'status':        return status;
      case 'empId':         return body.empId||'';
      case 'name':          return body.name||'';
      case 'country':       return body.country||'';
      case 'trainer':       return body.trainer||'';
      case 'wave':          return body.wave||'';
      case 'batch':         return body.batch||'';
      case 'batchId':       return body.batchId||'';
      case 'day':           return body.day||'';
      case 'dayName':       return body.dayName||'';
      case 'attemptNumber': return attemptNumber;
      case 'score':         return Number(body.score)||0;
      case 'pct':           return Number(body.pct)||0;
      case 'correct':       return Number(body.correct)||0;
      case 'total':         return Number(body.total)||0;
      case 'result':        return status==='completed'?((Number(body.pct)>=80)?'PASSED':'FAILED'):'IN_PROGRESS';
      case 'htqSeconds':    return Number(body.htqSeconds)||0;
      case 'totalSeconds':  return Number(body.totalSeconds)||0;
      case 'decisions':     return JSON.stringify(body.decisions||[]);
      default:              return '';
    }
  });

  // Upsert on attemptId
  const rowIndex=findRowByAttemptId_(sheet,headers,attemptId);
  if(rowIndex>0){
    sheet.getRange(rowIndex,1,1,rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return jsonResponse_({ok:true,attemptId:attemptId,attemptNumber:attemptNumber});
}

function findRowByAttemptId_(sheet,headers,attemptId){
  if(!attemptId)return -1;
  const idxCol=headers.indexOf('attemptId');
  if(idxCol<0)return -1;
  const lastRow=sheet.getLastRow();
  if(lastRow<=1)return -1;
  const ids=sheet.getRange(2,idxCol+1,lastRow-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(ids[i][0]===attemptId)return i+2;
  }
  return -1;
}

function computeAttemptNumber_(sheet,empId,day,batchId,currentAttemptId){
  if(!empId)return 1;
  const lastRow=sheet.getLastRow();
  if(lastRow<=1)return 1;
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const empIdCol=headers.indexOf('empId');
  const dayCol=headers.indexOf('day');
  const batchIdCol=headers.indexOf('batchId');
  const attemptIdCol=headers.indexOf('attemptId');
  const statusCol=headers.indexOf('status');
  if(empIdCol<0)return 1;
  const data=sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
  let count=0;
  for(let i=0;i<data.length;i++){
    const row=data[i];
    if(String(row[empIdCol])!==String(empId))continue;
    if(String(row[attemptIdCol])===String(currentAttemptId))continue;
    // Match by batchId if this is a batch, else by day
    if(batchId){if(String(row[batchIdCol])!==String(batchId))continue;}
    else{if(String(row[dayCol])!==String(day))continue;if(String(row[batchIdCol])!=='')continue;}
    // Only count completed prior attempts
    if(String(row[statusCol])==='completed')count++;
  }
  return count+1;
}

function handleScoresFetch_(){
  const sheet=getOrCreateScoresSheet_();
  ensureColumns_(sheet,SCORE_HEADERS);
  const data=sheet.getDataRange().getValues();
  if(data.length<=1)return jsonResponse_([]);
  const headers=data[0];
  const rows=data.slice(1).map(function(row){
    const obj={};
    headers.forEach(function(h,i){obj[h]=row[i];});
    if(obj.decisions&&typeof obj.decisions==='string'){
      try{obj.decisions=JSON.parse(obj.decisions);}catch(e){obj.decisions=[];}
    }
    ['day','score','pct','correct','total','attemptNumber','htqSeconds','totalSeconds'].forEach(function(k){
      obj[k]=Number(obj[k])||0;
    });
    return obj;
  });
  return jsonResponse_(rows);
}

function getOrCreateScoresSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(SCORES_SHEET);
  if(!sheet){
    sheet=ss.insertSheet(SCORES_SHEET);
    sheet.appendRow(SCORE_HEADERS);
    sheet.getRange(1,1,1,SCORE_HEADERS.length).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,SCORE_HEADERS.length,140);
  }
  return sheet;
}

function ensureColumns_(sheet,expectedHeaders){
  if(sheet.getLastRow()===0){
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1,1,1,expectedHeaders.length).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return;
  }
  const lastCol=sheet.getLastColumn();
  const existing=sheet.getRange(1,1,1,Math.max(1,lastCol)).getValues()[0];
  expectedHeaders.forEach(function(h){
    if(existing.indexOf(h)===-1){
      const newCol=sheet.getLastColumn()+1;
      sheet.getRange(1,newCol).setValue(h);
      sheet.getRange(1,newCol).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
      existing.push(h);
    }
  });
}

/* ═══ DAY 6 (legacy) ═══ */
function handleDay6Upload_(body){
  const result=createQuestionTab_(body.questions||[],DAY6_TAB_PREFIX,null,null);
  if(result&&result.tab){PropertiesService.getScriptProperties().setProperty(ACTIVE_DAY6_PROP,result.tab);}
  return jsonResponse_(result||{ok:false,error:'Failed'});
}
function handleDay6Fetch_(){
  const t=PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  return jsonResponse_(t?readQuestionTab_(t):[]);
}
function handleDay6Status_(){
  const t=PropertiesService.getScriptProperties().getProperty(ACTIVE_DAY6_PROP);
  if(!t)return jsonResponse_({active:false,count:0});
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t);
  if(!sh)return jsonResponse_({active:false,count:0});
  return jsonResponse_({active:true,tab:t,count:Math.max(0,sh.getLastRow()-1)});
}

/* ═══ BATCH MANAGEMENT ═══ */
function handleCreateBatch_(body){
  const questions=(body&&body.questions)||[];
  if(!Array.isArray(questions)||questions.length===0){
    return jsonResponse_({ok:false,error:'No questions provided.'});
  }
  const batchName=String(body.batchName||'').trim();
  if(!batchName)return jsonResponse_({ok:false,error:'Batch name is required.'});
  const batchId='BATCH-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  const result=createQuestionTab_(questions,BATCH_TAB_PREFIX,null,batchId);
  if(!result||!result.tab)return jsonResponse_({ok:false,error:'Failed to create batch tab.'});
  const batchesSheet=getOrCreateBatchesSheet_();
  ensureColumns_(batchesSheet,BATCH_HEADERS);
  const headers=batchesSheet.getRange(1,1,1,batchesSheet.getLastColumn()).getValues()[0];
  const row=headers.map(function(h){
    switch(h){
      case 'batchId':            return batchId;
      case 'batchName':          return batchName;
      case 'trainerName':        return body.trainerName||'';
      case 'waveCode':           return body.waveCode||'';
      case 'startDate':          return body.startDate||'';
      case 'endDate':            return body.endDate||'';
      case 'createdAt':          return new Date().toLocaleString();
      case 'fileName':           return body.fileName||'';
      case 'questionCount':      return questions.length;
      case 'timePerQuestionSec': return Number(body.timePerQuestionSec)||300;
      case 'activeTab':          return result.tab;
      case 'createdBy':          return body.createdBy||'Admin';
      default:                   return '';
    }
  });
  batchesSheet.appendRow(row);
  return jsonResponse_({ok:true,batchId:batchId,tab:result.tab,count:questions.length});
}

function handleListBatches_(){
  const sheet=getOrCreateBatchesSheet_();
  ensureColumns_(sheet,BATCH_HEADERS);
  const data=sheet.getDataRange().getValues();
  if(data.length<=1)return jsonResponse_([]);
  const headers=data[0];
  const rows=data.slice(1).map(function(row){
    const obj={};
    headers.forEach(function(h,i){obj[h]=row[i];});
    obj.questionCount=Number(obj.questionCount)||0;
    obj.timePerQuestionSec=Number(obj.timePerQuestionSec)||300;
    ['startDate','endDate'].forEach(function(k){
      if(obj[k] instanceof Date){obj[k]=Utilities.formatDate(obj[k],Session.getScriptTimeZone(),'yyyy-MM-dd');}
    });
    return obj;
  }).filter(function(b){return b.batchId;});
  return jsonResponse_(rows);
}

function handleGetBatch_(batchId){
  if(!batchId)return jsonResponse_({questions:[],batch:null});
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const batchesSheet=ss.getSheetByName(BATCHES_SHEET);
  if(!batchesSheet)return jsonResponse_({questions:[],batch:null});
  const data=batchesSheet.getDataRange().getValues();
  if(data.length<=1)return jsonResponse_({questions:[],batch:null});
  const headers=data[0];
  const idIdx=headers.indexOf('batchId');
  const tabIdx=headers.indexOf('activeTab');
  const timeIdx=headers.indexOf('timePerQuestionSec');
  let tabName=null,timeSec=300,batchMeta=null;
  for(let i=1;i<data.length;i++){
    if(data[i][idIdx]===batchId){
      tabName=data[i][tabIdx];
      timeSec=Number(data[i][timeIdx])||300;
      batchMeta={};
      headers.forEach(function(h,j){batchMeta[h]=data[i][j];});
      break;
    }
  }
  if(!tabName)return jsonResponse_({questions:[],batch:batchMeta});
  return jsonResponse_({questions:readQuestionTab_(tabName),batch:batchMeta,timePerQuestionSec:timeSec});
}

function getOrCreateBatchesSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(BATCHES_SHEET);
  if(!sheet){
    sheet=ss.insertSheet(BATCHES_SHEET);
    sheet.appendRow(BATCH_HEADERS);
    sheet.getRange(1,1,1,BATCH_HEADERS.length).setFontWeight('bold').setBackground('#00AF87').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,200);sheet.setColumnWidth(2,200);
    sheet.setColumnWidths(3,BATCH_HEADERS.length-2,130);
  }
  return sheet;
}

/* ═══ WAVE MANAGEMENT ═══ */
function handleCreateWave_(body){
  const waveCode=String(body.waveCode||'').trim();
  if(!waveCode)return jsonResponse_({ok:false,error:'Wave code is required.'});
  const waveId='WAVE-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  const sheet=getOrCreateWavesSheet_();
  ensureColumns_(sheet,WAVE_HEADERS);
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const row=headers.map(function(h){
    switch(h){
      case 'waveId':      return waveId;
      case 'waveCode':    return waveCode;
      case 'waveName':    return body.waveName||'';
      case 'trainerName': return body.trainerName||'';
      case 'startDate':   return body.startDate||'';
      case 'endDate':     return body.endDate||'';
      case 'createdAt':   return new Date().toLocaleString();
      case 'createdBy':   return body.createdBy||'Admin';
      default:            return '';
    }
  });
  sheet.appendRow(row);
  return jsonResponse_({ok:true,waveId:waveId,waveCode:waveCode});
}

function handleListWaves_(){
  const sheet=getOrCreateWavesSheet_();
  ensureColumns_(sheet,WAVE_HEADERS);
  const data=sheet.getDataRange().getValues();
  if(data.length<=1)return jsonResponse_([]);
  const headers=data[0];
  const rows=data.slice(1).map(function(row){
    const obj={};
    headers.forEach(function(h,i){obj[h]=row[i];});
    ['startDate','endDate'].forEach(function(k){
      if(obj[k] instanceof Date){obj[k]=Utilities.formatDate(obj[k],Session.getScriptTimeZone(),'yyyy-MM-dd');}
    });
    return obj;
  }).filter(function(w){return w.waveId;});
  return jsonResponse_(rows);
}

function getOrCreateWavesSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(WAVES_SHEET);
  if(!sheet){
    sheet=ss.insertSheet(WAVES_SHEET);
    sheet.appendRow(WAVE_HEADERS);
    sheet.getRange(1,1,1,WAVE_HEADERS.length).setFontWeight('bold').setBackground('#7B2D8B').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1,WAVE_HEADERS.length,140);
  }
  return sheet;
}

/* ═══ TRAINERS ═══ */
function handleListTrainers_(){
  // Extract distinct trainer names from Batches AND Waves.
  const seen={};
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ['Batches','Waves'].forEach(function(name){
    const sh=ss.getSheetByName(name);
    if(!sh)return;
    const data=sh.getDataRange().getValues();
    if(data.length<=1)return;
    const headers=data[0];
    const idx=headers.indexOf('trainerName');
    if(idx<0)return;
    for(let i=1;i<data.length;i++){
      const t=String(data[i][idx]||'').trim();
      if(t)seen[t]=true;
    }
  });
  return jsonResponse_(Object.keys(seen).sort());
}

/* ═══ QUESTION-TAB HELPERS ═══ */
function createQuestionTab_(questions,prefix,_after,customId){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const stamp=customId||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm-ss');
  const tabName=prefix+stamp;
  const sheet=ss.insertSheet(tabName);
  sheet.appendRow(QUESTION_HEADERS);
  sheet.getRange(1,1,1,QUESTION_HEADERS.length).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  const rows=questions.map(function(q,i){
    return [
      (prefix===BATCH_TAB_PREFIX?7000:6000)+i+1,
      q.title||'',q.body||'',q.correct_action||'',q.correct_reason||'',
      q.explanation_correct||'',q.explanation_wrong||'',
      Number(q.rating)||3,q.property_name||'',q.experience_type||'',
      q.date||'',q.policy_name||'',q.owner_response||''
    ];
  });
  if(rows.length>0)sheet.getRange(2,1,rows.length,QUESTION_HEADERS.length).setValues(rows);
  [50,250,400,130,180,300,300,60,180,130,100,180,300].forEach(function(w,i){sheet.setColumnWidth(i+1,w);});
  return {ok:true,tab:tabName,count:questions.length};
}
function readQuestionTab_(tabName){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  if(!sh)return [];
  const data=sh.getDataRange().getValues();
  if(data.length<=1)return [];
  const headers=data[0];
  return data.slice(1).map(function(row){
    const q={};headers.forEach(function(h,i){q[h]=row[i];});return q;
  }).filter(function(q){return q.title&&q.body&&q.correct_action;});
}
function jsonResponse_(data){
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
