const ADMIN_EMAIL = 'info@happysugar.online';
const TYPES = ['商品について','お取引について','業務システム開発について','業務提携について','その他'];
const HEADERS = ['受付ID','受付日時','種別','会社名','氏名','メール','電話','内容','対応状況','担当者','対応メモ','通知結果','自動返信結果','送信ID'];
function setup() {
  const p=PropertiesService.getScriptProperties();
  if(p.getProperty('CONTACT_SHEET_ID')) { console.log(SpreadsheetApp.openById(p.getProperty('CONTACT_SHEET_ID')).getUrl()); return; }
  const ss=SpreadsheetApp.create('Happysugar 公式HP お問い合わせ管理');
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  const s=ss.getSheets()[0]; s.setName('受付台帳');
  s.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setBackground('#20432c').setFontColor('#ffffff').setFontWeight('bold');
  s.setFrozenRows(1); s.setColumnWidths(1,14,145); s.setColumnWidth(8,440); s.setColumnWidth(11,300);
  s.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['未対応','対応中','対応済み'],true).setAllowInvalid(false).build());
  s.getRange('B2:B').setNumberFormat('yyyy/mm/dd hh:mm:ss'); s.getRange('G2:G').setNumberFormat('@');
  p.setProperty('CONTACT_SHEET_ID',ss.getId()); console.log(ss.getUrl());
}
function doGet(){return HtmlService.createHtmlOutputFromFile('Contact').setTitle('お問い合わせ | 合同会社Happysugar').addMetaTag('viewport','width=device-width, initial-scale=1');}
function submitContact(d){
  if(!d || typeof d !== 'object' || d.trap || d.consent!==true) throw new Error('入力内容をご確認ください。');
  const str=(k,max,required)=>{const v=String(d[k]||'').trim();if(v.length>max||(required&&!v))throw new Error('入力内容をご確認ください。');return v;};
  const type=str('type',40,true),company=str('company',150,false),name=str('name',100,true),email=str('email',254,true).toLowerCase(),phone=str('phone',40,false),message=str('message',3000,true),token=str('token',80,true);
  if(!TYPES.includes(type)||!/^\S+@[^\s@]+\.[^\s@]+$/.test(email)||/[\r\n]/.test(email)||! /^[a-zA-Z0-9-]{20,80}$/.test(token)) throw new Error('入力内容をご確認ください。');
  const lock=LockService.getScriptLock(); if(!lock.tryLock(10000))throw new Error('混み合っています。少し時間をおいて再度お試しください。');
  let s,row,id;
  try {
    const sheetId=PropertiesService.getScriptProperties().getProperty('CONTACT_SHEET_ID');
    if(!sheetId)throw new Error('受付準備中です。メールでお問い合わせください。');
    s=SpreadsheetApp.openById(sheetId).getSheetByName('受付台帳');
    const last=s.getLastRow();
    if(last>1){const hit=s.getRange(2,14,last-1,1).createTextFinder(token).matchEntireCell(true).findNext();if(hit)return {id:s.getRange(hit.getRow(),1).getDisplayValue()};}
    const cache=CacheService.getScriptCache();const key='email_'+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,email)).slice(0,40);
    if(cache.get(key))throw new Error('連続送信は時間をおいてお試しください。');
    const now=new Date();const recent=last>1?s.getRange(Math.max(2,last-99),2,Math.min(last-1,100),1).getValues():[];
    if(recent.filter(x=>x[0] instanceof Date && now-x[0]<3600000).length>=30)throw new Error('受付が混み合っています。メールでお問い合わせください。');
    id='HS-'+Utilities.getUuid().slice(0,12).toUpperCase();row=last+1;
    const safe=v=>/^[\s]*[=+\-@]/.test(v)?"'"+v:v;
    s.getRange(row,1,1,14).setValues([[id,now,type,safe(company),safe(name),safe(email),safe(phone),safe(message),'未対応','','','未送信','未送信',token]]);
    SpreadsheetApp.flush();cache.put(key,'1',60);
  }finally{lock.releaseLock();}
  let notice='成功',reply='成功';
  try{MailApp.sendEmail({to:ADMIN_EMAIL,subject:'【HP問い合わせ】'+id+' '+type,body:'受付ID: '+id+'\n種別: '+type+'\n会社: '+company+'\n氏名: '+name+'\nメール: '+email+'\n電話: '+phone+'\n\n'+message,replyTo:email,name:'Happysugar 受付'});}catch(e){notice='失敗';}
  try{MailApp.sendEmail({to:email,subject:'【Happysugar】お問い合わせを受け付けました',body:name+' 様\n\nこの度は合同会社Happysugarへお問い合わせいただき、ありがとうございます。\n受付ID：'+id+'\n内容を確認のうえ担当者より順次ご連絡いたします。\n\n合同会社Happysugar\n小さな気づきから、大きな幸せへ。\nhttps://happysugar.online/\ninfo@happysugar.online\n※本メールは自動送信です。',replyTo:ADMIN_EMAIL,name:'合同会社Happysugar'});}catch(e){reply='失敗';}
  try{s.getRange(row,12,1,2).setValues([[notice,reply]]);}catch(e){console.error('Mail status update failed '+id);}
  return {id:id};
}

/** GitHub Pagesからの通常のHTMLフォームPOST。CORSやiframeの応答に依存しない。 */
function doPost(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    result = submitContact({
      type:p.type, company:p.company, name:p.name, email:p.email,
      phone:p.phone, message:p.message, trap:p.trap,
      consent:p.consent==='yes', token:p.token
    });
  } catch (err) {
    console.error('Contact submission failed');
    return contactResultPage_('受付結果を確認できませんでした','再送する前に、info@happysugar.online へお問い合わせください。','');
  }
  return contactResultPage_('お問い合わせを受け付けました','内容を確認のうえ、担当者より順次ご連絡いたします。',result.id);
}
function contactResultPage_(title,message,id) {
  const esc = v => String(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><base target="_top"><meta charset="utf-8"><style>body{font-family:system-ui,sans-serif;background:#fff4f2;color:#20432c;line-height:1.8;margin:0}main{max-width:650px;margin:40px auto;padding:28px;background:white;border-radius:12px}img{width:230px;max-width:100%}a.button{display:inline-block;background:#20432c;color:white;padding:12px 20px;border-radius:5px;text-decoration:none}@media(max-width:700px){main{margin:15px}}</style></head><body><main><img src="https://happysugar.online/assets/logo-horizontal.webp" alt="Happysugar"><h1>'+esc(title)+'</h1><p>'+esc(message)+'</p>'+(id?'<p>受付ID：<strong>'+esc(id)+'</strong></p>':'')+'<a class="button" href="https://happysugar.online/">会社HPに戻る</a><p><a href="mailto:info@happysugar.online">info@happysugar.online</a></p></main></body></html>').setTitle(title).addMetaTag('viewport','width=device-width, initial-scale=1');
}
