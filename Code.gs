/**
 * ZIET Mumbai Workshop Deliverables & Classroom Impact Hub
 * Apps Script backend. Set deploy URL in the GitHub frontend after deployment.
 */
const CONFIG = {
  SPREADSHEET_ID: '1y8BQ3UsRE7xlHQFPyRVops7vOh8n7ZOvsEuI4sLXMZQ',
  DRIVE_ROOT_ID: '1pFe84ogk9ihbPDTmwscOVPHvIyUcYC7q',
  // Set to the Nameplates spreadsheet ID if it is not the master spreadsheet.
  NAMEPLATES_SPREADSHEET_ID: '1y8BQ3UsRE7xlHQFPyRVops7vOh8n7ZOvsEuI4sLXMZQ',
  NAMEPLATES_SHEET_NAME: '', // Leave blank to auto-detect a sheet with participant/name and Vidyalaya headers.
  ADMIN_EMAILS: [] // Add authorised ZIET admin email addresses.
};
const SHEETS = {
  Workshops: ['Workshop ID','Workshop Title','Start Date','End Date','Region','Status','Created At'],
  Participants: ['Participant ID','Workshop ID','Name','Designation','Vidyalaya','Region','Employee Code','DOJ','DOB','Gender','Social Category','PwD Status','WhatsApp','Individual Email','Vidyalaya Email','KVS Recognitions','Drive Folder ID','Drive Folder URL','Created At','Updated At'],
  Submissions: ['Submission ID','Workshop ID','Participant ID','Phase','Category','Submission Date','Title','Description','Link','File ID','Status','Updated At'],
  Daily_Proceedings: ['Entry ID','Workshop ID','Participant ID','Day/Date','Proceedings','Document Link','Studio Details','Studio Link','Group Work','Group Link','Assignment','Assignment Link','Takeaways','Resource Links','Updated At'],
  Classroom_Impact: ['Entry ID','Workshop ID','Participant ID','Implementation Date','Class/Subject','Challenge','Intervention','Lesson Link','Student Response','Student Work Link','Evidence Link','Peer Feedback','School Sharing','Updated At'],
  Action_Research: ['Entry ID','Workshop ID','Participant ID','Title','Problem','Research Question','Baseline','Intervention','Method','Findings','Analysis','Reflection','Future CPD Need','Updated At'],
  Monitoring: ['Entry ID','Workshop ID','Participant ID','Milestone','Status','Progress Update','Evidence Link','Next Steps','Updated At'],
  Audit_Log: ['Timestamp','User Email','Action','Workshop ID','Participant ID','Details']
};
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Workshop Deliverables & Classroom Impact Hub | ZIET Mumbai')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// Called by the participant interface through google.script.run when served by Apps Script.
function apiCall(payload) {
  return route_(payload || {});
}
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = route_(body);
    return json_({ok:true, ...result});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ss_(){ return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function root_(){ return DriveApp.getFolderById(CONFIG.DRIVE_ROOT_ID); }
function setup(){
  const ss=ss_();
  Object.keys(SHEETS).forEach(name=>{
    let sh=ss.getSheetByName(name);
    if(!sh) sh=ss.insertSheet(name);
    const headers=SHEETS[name];
    if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    else {
      const existing=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);
      headers.forEach(h=>{if(existing.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});
    }
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#173c70').setFontColor('#ffffff');
  });
  ['01_Participant_Profiles','02_Pre_Workshop_Deliverables','03_Ongoing_Deliverables','04_Classroom_Implementation','05_Action_Research','06_PDF_Workbooks','07_Monitoring_30_60_90_Days','08_Administrative_Reports'].forEach(n=>{const it=root_().getFoldersByName(n);if(!it.hasNext())root_().createFolder(n);});
  return {spreadsheet:ss.getName(),folder:root_().getName(),sheets:Object.keys(SHEETS)};
}
function route_(b){
  switch(b.action){
    case 'setup': return setup();
    case 'listSheets': return {sheets:ss_().getSheets().map(s=>s.getName())};
    case 'createWorkshop': return createWorkshop_(b);
    case 'syncNameplates': return syncNameplates_(b);
    case 'listWorkshops': return {workshops:listRows_('Workshops').filter(r=>String(r['Workshop Title']||'').trim()).map(r=>({id:r['Workshop ID'],title:r['Workshop Title']}))};
    case 'listParticipants': return {participants:listRows_('Participants').filter(r=>String(r['Workshop ID'])===String(b.workshopId)).map(r=>({id:r['Participant ID'],name:r.Name,designation:r.Designation,vidyalaya:r.Vidyalaya}))};
    case 'saveSubmission': return saveSubmission_(b);
    case 'adminSummary': return adminSummary_(b);
    case 'createWorkbookPdf': return createWorkbookPdf_(b);
    case 'exportAdminCsv': return exportAdminCsv_(b);
    default: throw new Error('Unknown action: '+b.action);
  }
}

function createWorkshop_(b){
  const title=String(b.title||'').trim();if(!title)throw new Error('Enter a workshop title');
  const id='WS-'+Utilities.getUuid().slice(0,8).toUpperCase();
  appendObject_('Workshops',{'Workshop ID':id,'Workshop Title':title,'Start Date':b.startDate||'','End Date':b.endDate||'','Region':b.region||'','Status':'Active','Created At':new Date()});
  return {id,title};
}
function syncNameplates_(b){
  const sourceId=CONFIG.NAMEPLATES_SPREADSHEET_ID||CONFIG.SPREADSHEET_ID;
  const source=SpreadsheetApp.openById(sourceId);
  const wanted=String(b.sheetName||CONFIG.NAMEPLATES_SHEET_NAME||'').trim();
  let sh=wanted?source.getSheetByName(wanted):null;
  if(!sh){
    const aliases={
      name:['name','participant name','teacher name','full name'],
      designation:['designation','post','category'],
      school:['vidyalaya','school','kv name','institution'],
      region:['region','ro','regional office'],
      code:['employee code','emp code','employee id','staff id']
    };
    for(const candidate of source.getSheets()){
      const vals=candidate.getRange(1,1,1,Math.max(1,candidate.getLastColumn())).getDisplayValues()[0].map(x=>String(x).trim().toLowerCase());
      const hasName=vals.some(x=>aliases.name.includes(x));
      const hasSchool=vals.some(x=>aliases.school.includes(x));
      if(hasName&&hasSchool){sh=candidate;break;}
    }
  }
  if(!sh)throw new Error('Could not identify the Nameplates sheet. Select a sheet name in the setup and confirm it has Name and Vidyalaya/School column headings.');
  const rows=sh.getDataRange().getDisplayValues();if(rows.length<2)throw new Error('The selected Nameplates sheet has no participant rows.');
  const heads=rows.shift().map(x=>String(x).trim().toLowerCase());
  const find=names=>heads.findIndex(h=>names.includes(h));
  const ix={name:find(['name','participant name','teacher name','full name']),designation:find(['designation','post','category']),school:find(['vidyalaya','school','kv name','institution']),region:find(['region','ro','regional office']),code:find(['employee code','emp code','employee id','staff id'])};
  if(ix.name<0||ix.school<0)throw new Error('Nameplates needs recognizable Name and Vidyalaya/School column headings.');
  const workshopId=String(b.workshopId||'');if(!workshopId)throw new Error('Select or create a workshop first.');
  const shOut=ss_().getSheetByName('Participants');let count=0;
  rows.forEach(row=>{
    const name=String(row[ix.name]||'').trim();if(!name)return;
    const code=ix.code>=0?String(row[ix.code]||'').trim():'';
    const existing=listRows_('Participants').find(p=>String(p['Workshop ID'])===workshopId&&((code&&String(p['Employee Code'])===code)||(!code&&String(p.Name)===name)));
    if(existing)return;
    appendObject_('Participants',{'Participant ID':Utilities.getUuid(),'Workshop ID':workshopId,'Name':name,'Designation':ix.designation>=0?row[ix.designation]:'','Vidyalaya':row[ix.school],'Region':ix.region>=0?row[ix.region]:'','Employee Code':code,'Created At':new Date(),'Updated At':new Date()});
    count++;
  });
  return {sheet:sh.getName(),imported:count};
}

function listRows_(name){
  const sh=ss_().getSheetByName(name); if(!sh||sh.getLastRow()<2)return [];
  const vals=sh.getDataRange().getValues(), heads=vals.shift();
  return vals.filter(r=>r.some(v=>v!=='')).map(row=>Object.fromEntries(heads.map((h,i)=>[h,row[i]])));
}
function saveSubmission_(b){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const section=String(b.section||'');const v=b.values||{};
    const mapping={profile:['Participants',{'Employee Code':'employeeCode','DOJ':'doj','DOB':'dob','Gender':'gender','Social Category':'socialCategory','PwD Status':'pwdStatus','WhatsApp':'whatsapp','Individual Email':'individualEmail','Vidyalaya Email':'vidyalayaEmail','KVS Recognitions':'recognitions'}],pre:['Submissions',{}],ongoing:['Daily_Proceedings',{}],post:['Classroom_Impact',{}],research:['Action_Research',{}],monitor:['Monitoring',{}]};
    const tab=mapping[section];if(!tab)throw new Error('Invalid section');
    const sh=ss_().getSheetByName(tab[0]);if(!sh)throw new Error('Run setup() first');
    const recordId=Utilities.getUuid(), now=new Date();
    if(section==='profile'){
      upsertParticipant_(b.participantId,v);audit_('save profile',b);
    }else if(section==='pre'){
      const preItems=[['Best Practice','bestPractice','bestPracticeLink'],['Lesson Plan / Unit Plan','lessonPlan','lessonPlanLink'],['Teaching Learning Material','tlm','tlmLink'],['Assessment Tool / Rubric','assessment','assessmentLink'],['Self-assessment & Learning Needs','tna',''],['Workshop Expectations','expectations','']];
      preItems.forEach(([title,descKey,linkKey])=>{const desc=String(v[descKey]||'');const link=linkKey?String(v[linkKey]||''):'';if(desc||link)appendObject_('Submissions',{'Submission ID':Utilities.getUuid(),'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Phase':'Pre','Category':title,'Submission Date':now,'Title':title,'Description':desc,'Link':link,'Status':'Submitted','Updated At':now});});audit_('save pre',b);
    }else if(section==='ongoing'){
      appendObject_('Daily_Proceedings',{'Entry ID':recordId,'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Day/Date':v.date,'Proceedings':v.proceedings,'Document Link':v.proceedingsLink,'Studio Details':v.studio,'Studio Link':v.studioLink,'Group Work':v.groupTask,'Group Link':v.groupLink,'Assignment':v.assignment,'Assignment Link':v.assignmentLink,'Takeaways':v.takeaways,'Resource Links':v.resourceLinks,'Updated At':now});audit_('save ongoing',b);
    }else if(section==='post'){
      appendObject_('Classroom_Impact',{'Entry ID':recordId,'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Implementation Date':v.implementationDate,'Class/Subject':v.classSubject,'Challenge':v.challenge,'Intervention':v.implementation,'Lesson Link':v.lessonLink,'Student Response':v.studentResponse,'Student Work Link':v.studentWorkLink,'Evidence Link':v.evidenceLink,'Peer Feedback':v.peerFeedback,'School Sharing':v.sharing,'Updated At':now});audit_('save post',b);
    }else if(section==='research'){
      appendObject_('Action_Research',{'Entry ID':recordId,'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Title':v.researchTitle,'Problem':v.problem,'Research Question':v.question,'Baseline':v.baseline,'Intervention':v.intervention,'Method':v.method,'Findings':v.findings,'Analysis':v.analysis,'Reflection':v.reflection,'Future CPD Need':v.futureNeed,'Updated At':now});audit_('save research',b);
    }else if(section==='monitor'){
      appendObject_('Monitoring',{'Entry ID':recordId,'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Milestone':v.milestone,'Status':v.status,'Progress Update':v.progress,'Evidence Link':v.monitorLink,'Next Steps':v.nextSteps,'Updated At':now});audit_('save monitoring',b);
    }
    return {saved:true,id:recordId};
  }finally{lock.releaseLock();}
}
function appendObject_(sheet,obj){
  const sh=ss_().getSheetByName(sheet);const heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(heads.map(h=>obj[h]!==undefined?obj[h]:''));
}
function upsertParticipant_(id,v){
  const sh=ss_().getSheetByName('Participants'), rows=listRows_('Participants'), index=rows.findIndex(r=>String(r['Participant ID'])===String(id));
  if(index<0)throw new Error('Participant not found in roster');
  const row=index+2, heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], map={'Employee Code':v.employeeCode,'DOJ':v.doj,'DOB':v.dob,'Gender':v.gender,'Social Category':v.socialCategory,'PwD Status':v.pwdStatus,'WhatsApp':v.whatsapp,'Individual Email':v.individualEmail,'Vidyalaya Email':v.vidyalayaEmail,'KVS Recognitions':v.recognitions,'Updated At':new Date()};
  heads.forEach((h,i)=>{if(map[h]!==undefined)sh.getRange(row,i+1).setValue(map[h]);});
}
function adminSummary_(b){
  const w=String(b.workshopId||'');const filt=tab=>listRows_(tab).filter(r=>String(r['Workshop ID'])===w).length;
  return {totalParticipants:listRows_('Participants').filter(r=>String(r['Workshop ID'])===w).length,pre:filt('Submissions'),ongoing:filt('Daily_Proceedings'),post:filt('Classroom_Impact'),research:filt('Action_Research')};
}
function audit_(action,b){try{appendObject_('Audit_Log',{'Timestamp':new Date(),'User Email':Session.getActiveUser().getEmail(),'Action':action,'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Details':JSON.stringify(b.values||{})});}catch(e){}}
function createWorkbookPdf_(b){
  const p=listRows_('Participants').find(r=>String(r['Participant ID'])===String(b.participantId));if(!p)throw new Error('Participant not found');
  const doc=DocumentApp.create('Workshop Impact Workbook - '+p.Name);const body=doc.getBody();
  body.appendParagraph('ZIET MUMBAI | KENDRIYA VIDYALAYA SANGATHAN').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('WORKSHOP DELIVERABLES & CLASSROOM IMPACT WORKBOOK').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Participant: '+p.Name+' | '+(p.Designation||'')+' | '+(p.Vidyalaya||''));
  body.appendParagraph('Employee Code: '+(p['Employee Code']||'')+' | Region: '+(p.Region||''));
  ['Submissions','Daily_Proceedings','Classroom_Impact','Action_Research','Monitoring'].forEach(tab=>{
    body.appendParagraph(tab.replace(/_/g,' ')).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    listRows_(tab).filter(r=>String(r['Participant ID'])===String(b.participantId)&&String(r['Workshop ID'])===String(b.workshopId)).forEach(r=>{
      Object.keys(r).forEach(k=>{if(r[k]!==''&&r[k]!==null)body.appendParagraph(k+': '+r[k]);});body.appendParagraph('');
    });
  });
  doc.saveAndClose();const pdf=DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName('Workshop Impact Workbook - '+p.Name+'.pdf');
  const folder=root_().getFoldersByName('06_PDF_Workbooks').next();const file=folder.createFile(pdf);return {url:file.getUrl()};
}
function exportAdminCsv_(b){
  const data=listRows_('Participants').filter(r=>String(r['Workshop ID'])===String(b.workshopId));
  const csv=[Object.keys(data[0]||{}).join(','),...data.map(r=>Object.values(r).map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(','))].join('\n');
  const file=root_().getFoldersByName('08_Administrative_Reports').next().createFile('Workshop-participants-export.csv',csv,MimeType.CSV);
  return {url:file.getUrl()};
}
/** Run once after checking the configured source. */
function initialiseWorkshopImpactHub(){return setup();}
