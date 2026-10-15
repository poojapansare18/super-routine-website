/** Super Routine waitlist backend. Bound to the existing spreadsheet or set SPREADSHEET_ID. */
const SPREADSHEET_ID = '1MfYqSSiV5WDnqlEnwz54Qc3GsdC5dOtVHbrc_nbFDZk';
const SHEET_NAME = 'Responses';
const BASELINE = 357;
const HEADERS = ['waitlist_number','submitted_at','first_name','whatsapp_number','primary_goal','current_pain_point','packaging_preference','supplement_format','sachet_preference','monthly_price','purchase_intent','discovery_source','purchase_barrier','marketing_consent','updates_opt_in','page_source'];
function doGet(e){
  const p=(e&&e.parameter)||{};
  let result;
  if(String(p.action||'')==='submit'){
    try { result=savePayload_(JSON.parse(String(p.payload||''))); }
    catch(err) { result={ok:false,reason:'server_error',message:'The response could not be saved. Please try again.'}; }
  } else {
    const people=uniquePeople_(getSheet_()).length;
    result={ok:true,count:people,displayNumber:BASELINE+people};
  }
  const callback=String(p.callback||'');
  if(callback){
    if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback))return ContentService.createTextOutput('/* invalid callback */').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}
function doPost(e){
  try {
    const raw=e&&e.parameter&&e.parameter.payload;
    if(!raw)return json_({ok:false,reason:'missing_payload'});
    return json_(savePayload_(JSON.parse(raw)));
  } catch(err) {
    return json_({ok:false,reason:'server_error',message:'The response could not be saved. Please try again.'});
  }
}
function savePayload_(p){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    const name=clean_(p.first_name,80),phone=clean_(p.whatsapp_number||p.phone,32),digits=phone.replace(/\D/g,'');
    if(name.length<2||digits.length<8||p.marketing_consent!==true)return {ok:false,reason:'invalid_fields',message:'Please enter your name, a valid WhatsApp number and consent.'};
    const sh=getSheet_(),last=sh.getLastRow(),normalized=phoneKey_(phone);
    if(last>1){const phones=sh.getRange(2,4,last-1,1).getDisplayValues().flat();if(phones.some(v=>phoneKey_(v)===normalized))return {ok:false,reason:'already_registered'};}
    const peopleBefore=uniquePeople_(sh).length;
    const waitlistNo=BASELINE+peopleBefore+1;
    const min=Number(p.monthly_price_min_inr||0),max=Number(p.monthly_price_max_inr||0);
    const price=min&&max?'₹'+min.toLocaleString('en-IN')+'–₹'+max.toLocaleString('en-IN'):'';
    sh.appendRow([waitlistNo,new Date(),safe_(name),safe_(phone),clean_(p.primary_goal,120),clean_(p.current_pain_point,160),clean_(p.packaging_preference,80),clean_(p.supplement_format,80),clean_(p.sachet_preference,80),price,clean_(p.purchase_intent,100),clean_(p.discovery_source,100),'',true,false,clean_(p.page,1000)]);
    SpreadsheetApp.flush();
    try{refreshDashboard_();}catch(err){console.error('Dashboard refresh warning: '+err);}
    return {ok:true,waitlist_number:waitlistNo,countAfterSave:peopleBefore+1};
  } finally {
    try{lock.releaseLock();}catch(ignore){}
  }
}
function setupSheet(){const sh=getSheet_();refreshDashboard_();return 'Ready: Responses and Founder Dashboard are set up.';}
function refreshDashboard_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=getSheet_(),rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,HEADERS.length).getValues().filter(r=>r.some(v=>v!==''&&v!==null)):[];
  let d=ss.getSheetByName('Founder Dashboard');if(!d)d=ss.insertSheet('Founder Dashboard');d.clear();
  d.getRange('A1:H1').merge().setValue('SUPER ROUTINE  |  FOUNDER DASHBOARD');
  d.getRange('A2:H2').merge().setValue('Live customer research snapshot');
  const people=uniquePeople_(sh).length, opt=rows.filter(r=>r[14]===true).length;
  const cards=[['A4:B4','A5:B6','SAVED RESPONSES',people,'#DCEBFF','#174A8B'],['C4:D4','C5:D6','DISPLAY COUNTER',BASELINE+people,'#DDF7E8','#176B45'],['E4:F4','E5:F6','UPDATE OPT-IN','—','#FCE8D5','#9A4D0A'],['G4:H4','G5:H6','SAMPLE',people<10?'Early':'Growing','#EDE3FF','#6336A3']];
  cards.forEach(c=>{d.getRange(c[0]).merge().setValue(c[2]).setBackground(c[4]).setFontColor(c[5]).setFontWeight('bold');d.getRange(c[1]).merge().setValue(c[3]).setBackground(c[4]).setFontColor(c[5]).setFontSize(22).setFontWeight('bold');});
  d.getRange('A8:H8').merge().setValue('QUICK INSIGHTS').setBackground('#17365D').setFontColor('#FFFFFF').setFontWeight('bold');
  const fields=[['PRIMARY GOAL',5],['TOP PAIN POINT',6],['PACKAGING',7],['FORMAT',8],['PRICE RANGE',10],['PURCHASE INTENT',11],['DISCOVERY',12]];
  const lines=fields.map(([label,col])=>{const counts={};rows.forEach(r=>{const v=String(r[col-1]||'').trim();if(v)counts[v]=(counts[v]||0)+1;});const top=Object.keys(counts).sort((a,b)=>counts[b]-counts[a])[0];return [label,top?top+' · '+counts[top]:'No responses yet'];});
  d.getRange(9,1,lines.length,2).setValues(lines);d.getRange('A9:A15').setBackground('#EAF1FB').setFontWeight('bold');d.getRange('B9:B15').setBackground('#F7F9FC');
  d.setFrozenRows(2);d.setColumnWidths(1,8,130);d.setRowHeight(1,42);d.setRowHeight(2,26);d.getRange('A1:H1').setBackground('#17365D').setFontColor('#FFFFFF').setFontSize(17).setFontWeight('bold');d.getRange('A2:H2').setBackground('#EAF1F8').setFontColor('#52657A');d.getRange('A4:H6').setVerticalAlignment('middle');d.setRowHeights(9,7,28);d.getRange('A9:B15').setWrap(true);ss.setActiveSheet(d);ss.moveActiveSheet(1);
}
function getSheet_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#DCEBFF');}return sh;}
function uniquePeople_(sh){if(sh.getLastRow()<2)return [];const values=sh.getRange(2,4,sh.getLastRow()-1,1).getDisplayValues().flat(),seen={};values.forEach(v=>{const k=phoneKey_(v);if(k)seen[k]=true;});return Object.keys(seen);}
function phoneKey_(v){return String(v||'').replace(/\D/g,'').slice(-10);}
function clean_(v,n){return String(v==null?'':v).trim().slice(0,n);}
function safe_(v){return /^[=+@-]/.test(String(v))?"'"+v:String(v);}
function json_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
