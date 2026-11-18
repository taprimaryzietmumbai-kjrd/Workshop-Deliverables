/**
 * ZIET Mumbai Workshop Deliverables & Classroom Impact Hub
 * Apps Script backend. Set deploy URL in the GitHub frontend after deployment.
 */
const CONFIG = {
  SPREADSHEET_ID: '1y8BQ3UsRE7xlHQFPyRVops7vOh8n7ZOvsEuI4sLXMZQ',
  DRIVE_ROOT_ID: '1pFe84ogk9ihbPDTmwscOVPHvIyUcYC7q',
  CORRECTIONS_SPREADSHEET_ID: '1W1N0mDY6t2d_GGJrqgkHgGeJrK9-kg0A2HMQAojkhr0',
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
  Stage_Status: ['Status ID','Workshop ID','Participant ID','Stage','Completion Status','Submitted Items','Last Submitted At','Admin Verification','Verification Notes','Verified At'],
  Group_Deliverables: ['Entry ID','Workshop ID','Group Name','Contribution Type','Title','Description','Evidence Link','Member Participant IDs','Submitted By','Status','Updated At'],
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
    case 'syncCorrectionsRoster': return syncCorrectionsRoster_(b);
    case 'listWorkshops': return {workshops:listRows_('Workshops').filter(r=>String(r['Workshop Title']||'').trim()).map(r=>({id:r['Workshop ID'],title:r['Workshop Title']}))};
    case 'listParticipants': return {participants:listRows_('Participants').filter(r=>String(r['Workshop ID'])===String(b.workshopId)).map(r=>({id:r['Participant ID'],name:r.Name,designation:r.Designation,vidyalaya:r.Vidyalaya}))};
    case 'saveSubmission': return saveSubmission_(b);
    case 'getParticipantSubmissions': return getParticipantSubmissions_(b);
    case 'saveGroupDeliverable': return saveGroupDeliverable_(b);
    case 'verifySubmission': return verifySubmission_(b);
    case 'adminSummary': return adminSummary_(b);
    case 'createWorkbookPdf': return createWorkbookPdf_(b);
    case 'exportAdminCsv': return exportAdminCsv_(b);
    default: throw new Error('Unknown action: '+b.action);
  }
}

function createWorkshop_(b){
  setup(); // Ensure the destination database tabs and Drive folders exist; source nominations remain read-only.
  const title=String(b.title||'').trim();if(!title)throw new Error('Enter a workshop title');
  const id='WS-'+Utilities.getUuid().slice(0,8).toUpperCase();
  appendObject_('Workshops',{'Workshop ID':id,'Workshop Title':title,'Start Date':b.startDate||'','End Date':b.endDate||'','Region':b.region||'','Status':'Active','Created At':new Date()});
  return {id,title};
}
function syncCorrectionsRoster_(b){
  const workshopId=String(b.workshopId||'').trim();
  if(!workshopId) throw new Error('Create or select a workshop first.');
  const workshop=listRows_('Workshops').find(w=>String(w['Workshop ID'])===workshopId);
  if(!workshop) throw new Error('Workshop record not found.');
  const sourceName=String(b.sourceTabName||'PRTM MUMBAI').trim();
  const book=SpreadsheetApp.openById(CONFIG.CORRECTIONS_SPREADSHEET_ID);
  const source=book.getSheetByName(sourceName);
  if(!source) throw new Error('Source nominations tab "'+sourceName+'" was not found.');
  const sourceLastRow=source.getLastRow(), sourceLastCol=source.getLastColumn();
  if(sourceLastRow<2) throw new Error('The source nominations tab has no participant rows.');
  const headers=source.getRange(1,1,1,sourceLastCol).getDisplayValues()[0];
  const norm=s=>String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
  const indexOf=names=>headers.findIndex(h=>names.includes(norm(h)));
  const ix={
    name:indexOf(['name of the participant','participant name','name']),
    designation:indexOf(['designation']),school:indexOf(['school','vidyalaya']),
    employeeCode:indexOf(['employee code']),doj:indexOf(['dojkvs','date of joining','doj']),
    region:indexOf(['region']),gender:indexOf(['gender']),
    category:indexOf(['category','social category']),
    whatsapp:indexOf(['whatsapp','whatsapp no','mobile']),
    email:indexOf(['email','individual email']),
    vidyalayaEmail:indexOf(['vidyalayamail','vidyalaya email','school email']),
    dob:indexOf(['dob','date of birth'])
  };
  if(ix.name<0||ix.school<0) throw new Error('Source tab must contain Name of the Participant and School columns.');
  let target=workshop['Corrections Tab']?book.getSheetByName(String(workshop['Corrections Tab'])):null;
  let created=false;
  if(!target){
    target=book.insertSheet(uniqueWorkshopTabName_(book,workshop['Workshop Title']));
    // Retain all source headings. Copy nomination columns A:AH only;
    // correction/audit columns AI:AV are intentionally blank for this workshop.
    target.getRange(1,1,1,sourceLastCol).setValues([source.getRange(1,1,1,sourceLastCol).getValues()[0]]);
    const sourceRows=source.getRange(2,1,sourceLastRow-1,sourceLastCol).getValues();
    const output=sourceRows.map(row=>{
      const newRow=new Array(sourceLastCol).fill('');
      for(let col=0;col<Math.min(34,sourceLastCol);col++)newRow[col]=row[col];
      return newRow;
    }).filter(row=>String(row[ix.name]||'').trim()!=='');
    if(output.length)target.getRange(2,1,output.length,sourceLastCol).setValues(output);
    target.setFrozenRows(1);
    target.getRange(1,1,1,sourceLastCol).setFontWeight('bold').setBackground('#173c70').setFontColor('#ffffff').setWrap(true);
    for(let col=1;col<=sourceLastCol;col++){try{target.setColumnWidth(col,source.getColumnWidth(col));}catch(e){}}
    created=true;
  }
  const targetLastRow=target.getLastRow();
  const targetValues=targetLastRow>1?target.getRange(2,1,targetLastRow-1,target.getLastColumn()).getValues():[];
  const targetHeaders=target.getRange(1,1,1,target.getLastColumn()).getDisplayValues()[0];
  const targetIx={};
  ['name','designation','school','employeeCode','doj','region','gender','category','whatsapp','email','vidyalayaEmail','dob'].forEach(k=>{
    const aliases={name:['name of the participant','participant name','name'],designation:['designation'],school:['school','vidyalaya'],employeeCode:['employee code'],doj:['dojkvs','date of joining','doj'],region:['region'],gender:['gender'],category:['category','social category'],whatsapp:['whatsapp','whatsapp no','mobile'],email:['email','individual email'],vidyalayaEmail:['vidyalayamail','vidyalaya email','school email'],dob:['dob','date of birth']};
    targetIx[k]=targetHeaders.findIndex(h=>aliases[k].includes(norm(h)));
  });
  if(!ss_().getSheetByName('Participants'))throw new Error('Workshop Impact database is not initialised. Run initialiseWorkshopImpactHub() once.');
  const existing=listRows_('Participants').filter(p=>String(p['Workshop ID'])===workshopId);
  const seenCodes=new Set(existing.map(p=>String(p['Employee Code']||'').trim()).filter(Boolean));
  const seenNames=new Set(existing.map(p=>norm(p.Name)));
  const profileRoot=root_().getFoldersByName('01_Participant_Profiles');
  if(!profileRoot.hasNext())throw new Error('Participant profile folder is missing. Run initialiseWorkshopImpactHub() once.');
  const profileFolder=profileRoot.next();
  let imported=0;
  targetValues.forEach(row=>{
    const val=k=>targetIx[k]>=0?row[targetIx[k]]:'';
    const name=String(val('name')||'').trim();if(!name)return;
    const code=String(val('employeeCode')||'').trim();
    if((code&&seenCodes.has(code))||(!code&&seenNames.has(norm(name))))return;
    const participantId=Utilities.getUuid();
    const safeName=name.replace(/[^\p{L}\p{N}\s._-]/gu,'-').slice(0,70);
    const folder=profileFolder.createFolder(safeName+' - '+participantId.slice(0,8));
    appendObject_('Participants',{
      'Participant ID':participantId,'Workshop ID':workshopId,'Name':name,
      'Designation':val('designation'),'Vidyalaya':val('school'),'Region':val('region'),
      'Employee Code':code,'DOJ':val('doj'),'DOB':val('dob'),'Gender':val('gender'),
      'Social Category':val('category'),'WhatsApp':val('whatsapp'),
      'Individual Email':val('email'),'Vidyalaya Email':val('vidyalayaEmail'),
      'Drive Folder ID':folder.getId(),'Drive Folder URL':folder.getUrl(),
      'Created At':new Date(),'Updated At':new Date()
    });
    if(code)seenCodes.add(code);seenNames.add(norm(name));imported++;
  });
  ensureWorkshopMetadataColumns_();
  const nominatedCount=targetValues.filter(row=>String(row[targetIx.name]||'').trim()).length;
  setWorkshopMetadata_(workshopId,target.getName(),sourceName,nominatedCount);
  return {sourceSpreadsheet:book.getName(),sourceTab:sourceName,createdTab:target.getName(),created:created,nominationRows:nominatedCount,imported:imported,tabUrl:book.getUrl()+'#gid='+target.getSheetId()};
}
function uniqueWorkshopTabName_(book,title){
  let base=String(title||'Workshop').replace(/[^\p{L}\p{N}\s._-]/gu,' ').replace(/\s+/g,' ').trim();
  if(!base)base='Workshop';
  base=base.slice(0,90);
  let name=base,n=2;
  while(book.getSheetByName(name)){const suffix=' ('+n+')';name=base.slice(0,100-suffix.length)+suffix;n++;}
  return name;
}
function ensureWorkshopMetadataColumns_(){
  const sh=ss_().getSheetByName('Workshops');
  ['Corrections Tab','Nomination Source Tab','Nominated Count'].forEach(h=>{
    const heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    if(heads.indexOf(h)<0)sh.getRange(1,sh.getLastColumn()+1).setValue(h);
  });
}
function setWorkshopMetadata_(id,tabName,sourceName,count){
  const sh=ss_().getSheetByName('Workshops'),data=sh.getDataRange().getValues(),heads=data[0];
  const idCol=heads.indexOf('Workshop ID');
  const rowIndex=data.findIndex((r,i)=>i>0&&String(r[idCol])===String(id));
  if(rowIndex<1)return;
  ['Corrections Tab','Nomination Source Tab','Nominated Count'].forEach((h,i)=>sh.getRange(rowIndex+1,heads.indexOf(h)+1).setValue([tabName,sourceName,count][i]));
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
    const mapping={profile:'Participants',pre:'Submissions',ongoing:'Daily_Proceedings',post:'Classroom_Impact',research:'Action_Research',monitor:'Monitoring'};
    const sheetName=mapping[section];if(!sheetName)throw new Error('Invalid section');
    if(!ss_().getSheetByName(sheetName))throw new Error('Run setup() first');
    const now=new Date();let recordId=String(b.recordId||'');
    if(section==='profile'){
      upsertParticipant_(b.participantId,v);audit_('save profile',b);return {saved:true,id:b.participantId};
    }
    if(section==='pre'){
      const items=[['Best Practice','bestPractice','bestPracticeLink'],['Lesson Plan / Unit Plan','lessonPlan','lessonPlanLink'],['Teaching Learning Material','tlm','tlmLink'],['Assessment Tool / Rubric','assessment','assessmentLink'],['Self-assessment & Learning Needs','tna',''],['Workshop Expectations','expectations','']];
      let saved=0;
      items.forEach(([title,dk,lk])=>{
        const desc=String(v[dk]||'').trim(),link=lk?String(v[lk]||'').trim():'';
        if(!desc&&!link)return;
        const existing=listRows_('Submissions').find(r=>String(r['Workshop ID'])===String(b.workshopId)&&String(r['Participant ID'])===String(b.participantId)&&r.Phase==='Pre'&&r.Category===title);
        const obj={'Submission ID':existing?existing['Submission ID']:Utilities.getUuid(),'Workshop ID':b.workshopId,'Participant ID':b.participantId,'Phase':'Pre','Category':title,'Submission Date':existing?existing['Submission Date']:now,'Title':title,'Description':desc,'Link':link,'Status':'Submitted','Updated At':now};
        if(existing)updateObjectById_('Submissions','Submission ID',existing['Submission ID'],obj);else appendObject_('Submissions',obj);
        saved++;
      });
      updateStageStatus_(b.workshopId,b.participantId,'Pre',saved);
      audit_('save pre',b);return {saved:true,id:'pre-'+b.participantId,itemsSaved:saved};
    }
    const defs={
      ongoing:{idCol:'Entry ID',obj:{'Day/Date':v.date,'Proceedings':v.proceedings,'Document Link':v.proceedingsLink,'Studio Details':v.studio,'Studio Link':v.studioLink,'Group Work':v.groupTask,'Group Link':v.groupLink,'Assignment':v.assignment,'Assignment Link':v.assignmentLink,'Takeaways':v.takeaways,'Resource Links':v.resourceLinks}},
      post:{idCol:'Entry ID',obj:{'Implementation Date':v.implementationDate,'Class/Subject':v.classSubject,'Challenge':v.challenge,'Intervention':v.implementation,'Lesson Link':v.lessonLink,'Student Response':v.studentResponse,'Student Work Link':v.studentWorkLink,'Evidence Link':v.evidenceLink,'Peer Feedback':v.peerFeedback,'School Sharing':v.sharing}},
      research:{idCol:'Entry ID',obj:{'Title':v.researchTitle,'Problem':v.problem,'Research Question':v.question,'Baseline':v.baseline,'Intervention':v.intervention,'Method':v.method,'Findings':v.findings,'Analysis':v.analysis,'Reflection':v.reflection,'Future CPD Need':v.futureNeed}},
      monitor:{idCol:'Entry ID',obj:{'Milestone':v.milestone,'Status':v.status,'Progress Update':v.progress,'Evidence Link':v.monitorLink,'Next Steps':v.nextSteps}}
    };
    const d=defs[section];
    const obj={'Workshop ID':b.workshopId,'Participant ID':b.participantId,...d.obj,'Updated At':now};
    if(section==='ongoing'&&!recordId)recordId=Utilities.getUuid();
    if(section!=='ongoing'&&!recordId){
      const match=listRows_(sheetName).find(r=>String(r['Workshop ID'])===String(b.workshopId)&&String(r['Participant ID'])===String(b.participantId)&&(section!=='monitor'||String(r.Milestone)===String(v.milestone)));
      if(match)recordId=String(match[d.idCol]);
    }
    obj[d.idCol]=recordId||Utilities.getUuid();
    if(recordId&&rowById_(sheetName,d.idCol,recordId))updateObjectById_(sheetName,d.idCol,recordId,obj);else appendObject_(sheetName,obj);
    updateStageStatus_(b.workshopId,b.participantId,section==='ongoing'?'Ongoing':section==='post'?'Post':section==='research'?'Action Research':'Monitoring',1);
    audit_('save '+section,b);return {saved:true,id:obj[d.idCol],updated:!!recordId};
  }finally{lock.releaseLock();}
}
function rowById_(sheet,idCol,id){return listRows_(sheet).find(r=>String(r[idCol])===String(id));}
function updateObjectById_(sheet,idCol,id,obj){
  const sh=ss_().getSheetByName(sheet),data=sh.getDataRange().getValues(),heads=data[0],ix=heads.indexOf(idCol);
  const ri=data.findIndex((r,i)=>i>0&&String(r[ix])===String(id));if(ri<1)throw new Error('Record to update was not found.');
  heads.forEach((h,i)=>{if(Object.prototype.hasOwnProperty.call(obj,h))sh.getRange(ri+1,i+1).setValue(obj[h]);});
}
function updateStageStatus_(workshopId,participantId,stage,count){
  const sh=ss_().getSheetByName('Stage_Status'),rows=listRows_('Stage_Status');
  const old=rows.find(r=>String(r['Workshop ID'])===String(workshopId)&&String(r['Participant ID'])===String(participantId)&&r.Stage===stage);
  const obj={'Status ID':old?old['Status ID']:Utilities.getUuid(),'Workshop ID':workshopId,'Participant ID':participantId,'Stage':stage,'Completion Status':'Submitted','Submitted Items':count,'Last Submitted At':new Date()};
  if(old)updateObjectById_('Stage_Status','Status ID',old['Status ID'],obj);else appendObject_('Stage_Status',obj);
}
function getParticipantSubmissions_(b){
  const match=r=>String(r['Workshop ID'])===String(b.workshopId)&&String(r['Participant ID'])===String(b.participantId);
  return {pre:listRows_('Submissions').filter(r=>match(r)&&r.Phase==='Pre'),ongoing:listRows_('Daily_Proceedings').filter(match),post:listRows_('Classroom_Impact').filter(match),research:listRows_('Action_Research').filter(match),monitor:listRows_('Monitoring').filter(match),stages:listRows_('Stage_Status').filter(match)};
}
function saveGroupDeliverable_(b){
  const v=b.values||{},id=String(b.recordId||Utilities.getUuid());
  const obj={'Entry ID':id,'Workshop ID':b.workshopId,'Group Name':v.groupName,'Contribution Type':v.contributionType,'Title':v.title,'Description':v.description,'Evidence Link':v.evidenceLink,'Member Participant IDs':(v.memberParticipantIds||[]).join('|'),'Submitted By':b.participantId,'Status':'Submitted','Updated At':new Date()};
  if(rowById_('Group_Deliverables','Entry ID',id))updateObjectById_('Group_Deliverables','Entry ID',id,obj);else appendObject_('Group_Deliverables',obj);
  audit_('save group deliverable',b);return {saved:true,id:id};
}
function verifySubmission_(b){
  const allowed=['Pending','Verified','Needs Revision','Implemented','Impact Documented'];
  if(allowed.indexOf(String(b.status))<0)throw new Error('Invalid verification status');
  const sh=ss_().getSheetByName('Stage_Status'),rows=listRows_('Stage_Status');
  const r=rows.find(x=>String(x['Workshop ID'])===String(b.workshopId)&&String(x['Participant ID'])===String(b.participantId)&&x.Stage===b.stage);
  if(!r)throw new Error('Stage submission record not found');
  updateObjectById_('Stage_Status','Status ID',r['Status ID'],{'Admin Verification':b.status,'Verification Notes':b.notes||'','Verified At':new Date()});
  audit_('verify '+b.stage,{...b,values:{status:b.status,notes:b.notes||''}});return {saved:true};
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
  const doc=DocumentApp.create('Workshop Impact Workbook - '+p.Name);const body=doc.getBody();body.setPageWidth(595.28).setPageHeight(841.89).setMarginTop(36).setMarginBottom(36).setMarginLeft(42).setMarginRight(42);
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
  const csv=[Object.keys(data[0]||{}).join(','),...data.map(r=>Object.values(r).map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(','))].join('
');
  const file=root_().getFoldersByName('08_Administrative_Reports').next().createFile('Workshop-participants-export.csv',csv,MimeType.CSV);
  return {url:file.getUrl()};
}

/**
 * Read-only diagnostic for the supplied workshop corrections spreadsheet.
 * Run this from the Apps Script editor. It does not create, edit or delete tabs.
 */
function inspectCorrectionsWorkbook(){
  const book=SpreadsheetApp.openById(CONFIG.CORRECTIONS_SPREADSHEET_ID);
  const report=book.getSheets().map(sh=>{
    const width=Math.max(sh.getLastColumn(),1);
    const headers=sh.getLastRow()?sh.getRange(1,1,1,width).getDisplayValues()[0]:[];
    return {sheetName:sh.getName(),lastRow:sh.getLastRow(),lastColumn:sh.getLastColumn(),headers:headers};
  });
  Logger.log(JSON.stringify({spreadsheetName:book.getName(),spreadsheetUrl:book.getUrl(),sheets:report},null,2));
  return report;
}

/** Run once after checking the configured source. */
function initialiseWorkshopImpactHub(){return setup();}
