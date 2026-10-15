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
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh=getSheet_();
  const last=sh.getLastRow();
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(h=>String(h).trim());
  const rows=last>1?sh.getRange(2,1,last-1,sh.getLastColumn()).getValues().filter(r=>r.some(v=>v!==''&&v!==null)):[];
  const ix={};headers.forEach((h,i)=>ix[h]=i);
  const val=(row,key)=>ix[key]===undefined?'':String(row[ix[key]]==null?'':row[ix[key]]).trim();
  const splitAnswers=v=>String(v||'').split(/[,;\\n]+/).map(s=>s.trim()).filter(Boolean);
  const countAnswers=(key,multi)=>{
    const counts={};
    rows.forEach(row=>{
      const raw=val(row,key);
      const answers=multi?splitAnswers(raw):(raw?[raw]:[]);
      const seen={};
      answers.forEach(answer=>{
        const normalized=answer.replace(/\\s+/g,' ').trim();
        if(!normalized)return;
        const dedupe=normalized.toLowerCase();
        if(seen[dedupe])return;
        seen[dedupe]=true;
        counts[normalized]=(counts[normalized]||0)+1;
      });
    });
    return Object.keys(counts).map(label=>[label,counts[label],rows.length?counts[label]/rows.length:0]).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]));
  };

  let d=ss.getSheetByName('Founder Dashboard');if(!d)d=ss.insertSheet('Founder Dashboard');
  let a=ss.getSheetByName('Analysis Data');if(!a)a=ss.insertSheet('Analysis Data');
  d.getCharts().forEach(chart=>d.removeChart(chart));
  [d,a].forEach(s=>s.getDataRange().getMergedRanges().forEach(r=>r.breakApart()));
  d.clear();a.clear();

  const people=uniquePeople_(sh).length;
  const optIndex=ix.updates_opt_in;
  const opt=optIndex===undefined?0:rows.filter(r=>r[optIndex]===true||String(r[optIndex]).toLowerCase()==='true').length;
  const consentIndex=ix.marketing_consent;
  const consent=consentIndex===undefined?0:rows.filter(r=>r[consentIndex]===true||String(r[consentIndex]).toLowerCase()==='true').length;
  const optRate=rows.length?Math.round(opt/rows.length*100)+'%':'—';
  const consentRate=rows.length?Math.round(consent/rows.length*100)+'%':'—';

  d.getRange('A1:H1').merge().setValue('SUPER ROUTINE  |  FOUNDER DASHBOARD');
  d.getRange('A2:H2').merge().setValue('Customer research snapshot  •  Updated '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy, HH:mm'));
  const cards=[
    ['A4:B4','A5:B6','SAVED RESPONSES',rows.length,'#DCEBFF','#174A8B'],
    ['C4:D4','C5:D6','DISPLAY COUNTER',BASELINE+people,'#DDF7E8','#176B45'],
    ['E4:F4','E5:F6','UPDATE OPT-IN',optRate,'#FCE8D5','#9A4D0A'],
    ['G4:H4','G5:H6','SAMPLE',rows.length<10?'Early':'Growing','#EDE3FF','#6336A3']
  ];
  cards.forEach(c=>{
    d.getRange(c[0]).merge().setValue(c[2]).setBackground(c[4]).setFontColor(c[5]).setFontWeight('bold');
    d.getRange(c[1]).merge().setValue(c[3]).setBackground(c[4]).setFontColor(c[5]).setFontSize(22).setFontWeight('bold').setVerticalAlignment('middle');
  });

  d.getRange('A8:H8').merge().setValue('QUICK INSIGHTS');
  const insights=[
    ['TOP GOAL',countAnswers('primary_goal',true)],
    ['TOP PAIN POINT',countAnswers('current_pain_point',true)],
    ['TOP PACKAGING',countAnswers('packaging_preference',false)],
    ['TOP FORMAT',countAnswers('supplement_format',false)],
    ['TOP PRICE RANGE',countAnswers('monthly_price',false)],
    ['PURCHASE INTENT',countAnswers('purchase_intent',false)],
    ['DISCOVERY CHANNEL',countAnswers('discovery_source',false)],
    ['PURCHASE BARRIER',countAnswers('purchase_barrier',true)]
  ];
  const insightRows=insights.map(([label,list])=>[label,list.length?list[0][0]+' · '+list[0][1]+' ('+Math.round(list[0][2]*100)+'%)':'No responses yet']);
  d.getRange(9,1,insightRows.length,2).setValues(insightRows);
  d.getRange('A9:A16').setBackground('#EAF1FB').setFontColor('#244A73').setFontWeight('bold');
  d.getRange('B9:B16').setBackground('#F7F9FC');
  d.getRange('A9:B16').setWrap(true).setVerticalAlignment('middle');
  d.getRange('D9:H12').merge().setValue(rows.length<10
    ? 'Early sample: directional signals only. More responses are needed before product, pricing or marketing decisions.'
    : 'Percentages show the share of respondents. Multi-select answers may add up to more than 100%.');
  d.getRange('D9:H12').setBackground('#FFF1D8').setFontColor('#784B0B').setWrap(true).setVerticalAlignment('middle');

  const charts=[
    ['primary_goal','Customer goals',true,'#3977D5'],
    ['current_pain_point','Pain points',true,'#23A58A'],
    ['monthly_price','Price preference',false,'#F29E4C'],
    ['purchase_intent','Purchase intent',false,'#8A63D2'],
    ['discovery_source','Discovery channels',false,'#E36C82'],
    ['packaging_preference','Packaging preference',false,'#32A6B8']
  ];
  const ranges={};
  let ar=1;
  charts.forEach(item=>{
    const data=countAnswers(item[0],item[2]);
    a.getRange(ar,1,1,3).setValues([[item[1],'Responses','Share']]).setFontWeight('bold').setBackground('#17365D').setFontColor('#FFFFFF');
    if(data.length){
      a.getRange(ar+1,1,data.length,3).setValues(data);
      ranges[item[0]]={row:ar,rows:data.length+1};
    } else {
      a.getRange(ar+1,1,1,3).setValues([['No responses',0,0]]);
      ranges[item[0]]={row:ar,rows:2};
    }
    ar+=Math.max(data.length,1)+3;
  });
  const positions=[[15,1],[15,5],[32,1],[32,5],[49,1],[49,5]];
  charts.forEach((item,i)=>{
    const range=ranges[item[0]];
    const chart=d.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(a.getRange(range.row,1,range.rows,2)).setNumHeaders(1)
      .setOption('title',item[1]).setOption('legend',{position:'none'})
      .setOption('colors',[item[3]]).setOption('backgroundColor','#FFFFFF')
      .setOption('chartArea',{backgroundColor:'#FFFFFF',width:'78%',height:'75%'})
      .setOption('hAxis',{minValue:0,title:'Respondents',textStyle:{color:'#5B6677',fontSize:9}})
      .setOption('vAxis',{textStyle:{color:'#344054',fontSize:10}})
      .setOption('titleTextStyle',{color:'#183153',fontSize:13,bold:true})
      .setOption('width',520).setOption('height',290)
      .setPosition(positions[i][0],positions[i][1],0,0).build();
    d.insertChart(chart);
  });

  d.setFrozenRows(2);d.setColumnWidths(1,8,130);
  d.setRowHeight(1,42);d.setRowHeight(2,26);d.setRowHeights(4,3,28);d.setRowHeights(9,8,30);d.setRowHeight(8,28);
  d.getRange('A1:H1').setBackground('#17365D').setFontColor('#FFFFFF').setFontSize(17).setFontWeight('bold').setVerticalAlignment('middle');
  d.getRange('A2:H2').setBackground('#EAF1F8').setFontColor('#52657A').setFontSize(10);
  d.getRange('A8:H8').setBackground('#17365D').setFontColor('#FFFFFF').setFontWeight('bold');
  a.setFrozenRows(1);a.setColumnWidth(1,260);a.setColumnWidth(2,120);a.setColumnWidth(3,100);
  if(a.getLastRow()>0)a.getDataRange().setVerticalAlignment('middle').setWrap(true);
  a.hideSheet();ss.setActiveSheet(d);ss.moveActiveSheet(1);
}

function getSheet_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#DCEBFF');}return sh;}
function uniquePeople_(sh){if(sh.getLastRow()<2)return [];const values=sh.getRange(2,4,sh.getLastRow()-1,1).getDisplayValues().flat(),seen={};values.forEach(v=>{const k=phoneKey_(v);if(k)seen[k]=true;});return Object.keys(seen);}
function phoneKey_(v){return String(v||'').replace(/\D/g,'').slice(-10);}
function clean_(v,n){return String(v==null?'':v).trim().slice(0,n);}
function safe_(v){return /^[=+@-]/.test(String(v))?"'"+v:String(v);}
function json_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON);}
