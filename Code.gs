/** Technology Battle — Google Sheets bound Apps Script. */
const QUESTS = [
 'ไฟดับทั้งโรงเรียน: มีเทคโนโลยีอะไรช่วยแก้ปัญหาได้บ้าง?',
 'ติดเกาะร้าง: มีเทคโนโลยีอะไรช่วยให้เอาตัวรอดได้บ้าง?',
 'โรงเรียนเจอน้ำท่วม: มีเทคโนโลยีอะไรช่วยได้บ้าง?'
];
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('Technology Battle | ม.2').addMetaTag('viewport','width=device-width,initial-scale=1');}
function setupGame(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss) throw Error('ต้องเปิด Apps Script จาก Google Sheets (Extensions > Apps Script)');
 const props=PropertiesService.getScriptProperties(); props.setProperty('SHEET_ID',ss.getId());
 let state=ss.getSheetByName('Game');if(!state)state=ss.insertSheet('Game');
 let answers=ss.getSheetByName('Answers');if(!answers)answers=ss.insertSheet('Answers');
 if(state.getLastRow()===0) state.appendRow(['Round','Status','Question','DurationSeconds','StartMs','EndMs','Code1','Code2','Code3','Reveal']);
 if(answers.getLastRow()===0) answers.appendRow(['Timestamp','Round','Team','Answer','Normalized','Verdict','RowID']);
 if(state.getLastRow()<2)state.appendRow([1,'waiting',QUESTS[0],90,0,0,newCode_(),newCode_(),newCode_(),false]);
 return 'เตรียม Google Sheets พร้อมแล้ว';
}
function newCode_(){
 const chars='ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
 let out='';
 for(let i=0;i<4;i++) out+=chars.charAt(Math.floor(Math.random()*chars.length));
 return out;
}
function ss_(){const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');if(!id)throw Error('กรุณารัน setupGame() ก่อน');return SpreadsheetApp.openById(id);}
function game_(){const v=ss_().getSheetByName('Game').getRange(2,1,1,10).getValues()[0];return {round:Number(v[0]),status:String(v[1]),question:String(v[2]),duration:Number(v[3]),start:Number(v[4]),end:Number(v[5]),codes:v.slice(6,9).map(String),reveal:v[9]===true};}
function save_(g){ss_().getSheetByName('Game').getRange(2,1,1,10).setValues([[g.round,g.status,g.question,g.duration,g.start,g.end,...g.codes,g.reveal]]);}
function authTeacher_(pin){const configured=PropertiesService.getScriptProperties().getProperty('TEACHER_PIN');if(!configured)throw Error('ตั้ง TEACHER_PIN ใน Script properties ก่อน');if(!pin||String(pin)!==configured)throw Error('รหัสครูไม่ถูกต้อง');}
function team_(code,g){const i=g.codes.indexOf(String(code||'').trim().toUpperCase());if(i<0)throw Error('รหัสกลุ่มไม่ถูกต้อง');return i+1;}
function clean_(value){return String(value||'').normalize('NFKC').trim().replace(/\s+/g,' ').slice(0,70);}
function norm_(s){return clean_(s).toLocaleLowerCase().replace(/[\s\-_.!,，。!?]/g,'');}
function answerRows_(){const sh=ss_().getSheetByName('Answers');if(sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,7).getValues().map((r,i)=>({row:i+2,round:Number(r[1]),team:Number(r[2]),answer:String(r[3]),normalized:String(r[4]),verdict:String(r[5]),id:String(r[6])}));}
function view_(g,rows,role,team){
 const now=Date.now();const over=g.status==='running' && now>=g.end;
 const reveal=g.reveal && (g.status==='review'||over||g.status==='finished');
 const roundRows=rows.filter(r=>r.round===g.round);
 const groups=[1,2,3].map(t=>{
   const ar=roundRows.filter(r=>r.team===t);
   const visible=role==='teacher'?(reveal):(team===t);
   return {team:t,count:ar.length,score:ar.filter(r=>r.verdict==='yes').length,answers:ar.map(r=>({id:visible?r.id:null,text:visible?r.answer:'***',verdict:visible?r.verdict:null}))};
 });
 return {round:g.round,status:over?'review':g.status,question:g.question,duration:g.duration,remaining:g.status==='running'?Math.max(0,g.end-now):0,deadline:g.end,serverNow:now,reveal,groups,totals:role==='teacher'?([1,2,3].map(t=>rows.filter(r=>r.team===t&&r.verdict==='yes').length)):undefined,questions:role==='teacher'?QUESTS:undefined,codes:role==='teacher'?g.codes:undefined};
}
function teacherView(pin){authTeacher_(pin);const g=game_();return view_(g,answerRows_(),'teacher',0);}
function studentView(code){const g=game_();const t=team_(code,g);return {team:t,...view_(g,answerRows_(),'student',t)};}
function teacherAction(pin,action,payload){authTeacher_(pin);const lock=LockService.getScriptLock();lock.waitLock(15000);try{
 const g=game_(),now=Date.now();payload=payload||{};
 if(action==='start'){
   if(g.status==='running'&&now<g.end)throw Error('กำลังแข่งขันอยู่');
   if(g.status!=='waiting')throw Error('กดรอบใหม่ก่อนเริ่มการแข่งขันรอบถัดไป');
   const d=Number(payload.duration||90);if(![60,90,120].includes(d))throw Error('เวลาไม่ถูกต้อง');g.duration=d;g.start=now;g.end=now+d*1000;g.status='running';g.reveal=false;
 } else if(action==='reveal'){
   if(g.status==='running'&&now<g.end)throw Error('ต้องรอหมดเวลาก่อนเปิดเฉลย');g.status='review';g.reveal=true;
 } else if(action==='next'){
   if(g.status==='running'&&now<g.end)throw Error('ไม่สามารถเปลี่ยนรอบระหว่างแข่งขัน');
   if(g.round>=3)throw Error('ครบ 3 รอบแล้ว กดเริ่มเกมใหม่');g.round++;g.question=QUESTS[g.round-1];g.status='waiting';g.start=0;g.end=0;g.reveal=false;
 } else if(action==='reset'){
   if(g.status==='running'&&now<g.end)throw Error('ไม่สามารถล้างเกมระหว่างแข่งขัน');
   ss_().getSheetByName('Answers').getRange(2,1,Math.max(1,ss_().getSheetByName('Answers').getMaxRows()-1),7).clearContent();
   g.round=1;g.question=QUESTS[0];g.status='waiting';g.start=0;g.end=0;g.reveal=false; // เก็บรหัสกลุ่มเดิมไว้เมื่อเริ่มเกมใหม่
 } else if(action==='judge'){
   if(!g.reveal)throw Error('ต้องเปิดเฉลยก่อนตรวจ');const rows=answerRows_();const row=rows.find(r=>r.id===String(payload.id)&&r.round===g.round);
   if(!row)throw Error('ไม่พบคำตอบ');const value=payload.verdict==='yes'?'yes':payload.verdict==='no'?'no':'pending';ss_().getSheetByName('Answers').getRange(row.row,6).setValue(value);
 } else {throw Error('คำสั่งไม่ถูกต้อง');}
 if(action!=='judge')save_(g);return view_(g,answerRows_(),'teacher',0);
 } finally{lock.releaseLock();}}
function sendAnswer(code,raw){const lock=LockService.getScriptLock();lock.waitLock(15000);try{
 const g=game_(),t=team_(code,g),now=Date.now();if(g.status!=='running'||now>=g.end)throw Error('หมดเวลาหรือยังไม่เริ่มรอบนี้');
 const value=clean_(raw),n=norm_(value);if(value.length<2||value.length>70)throw Error('คำตอบต้องมี 2–70 ตัวอักษร');
 const rows=answerRows_();if(rows.some(r=>r.round===g.round&&r.team===t&&r.normalized===n))throw Error('กลุ่มนี้ส่งคำตอบซ้ำแล้ว');
 if(rows.filter(r=>r.round===g.round&&r.team===t).length>=50)throw Error('กลุ่มนี้ส่งครบจำนวนสูงสุดแล้ว');
 ss_().getSheetByName('Answers').appendRow([new Date(),g.round,t,value,n,'pending',Utilities.getUuid()]);return {ok:true};
 }finally{lock.releaseLock();}}
