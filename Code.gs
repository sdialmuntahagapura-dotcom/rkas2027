/**
 * SIP-RKAS V5 — Google Apps Script backend
 * Spreadsheet tetap menjadi satu-satunya database utama.
 */
const CFG={TZ:'Asia/Jakarta',SESSION:21600,FOLDER:'ARSIP SIP-RKAS',SHEETS:{
SETTING:['KEY','VALUE'],USERS:['ID','USERNAME','PASSWORD_HASH','NAME','ROLE','ACTIVE','CREATED_AT'],
SUPPLIERS:['ID','NAME','ADDRESS','PHONE','TAX_ID','NOTES','CREATED_AT'],
ACTIVITIES:['ID','CODE','NAME','ACCOUNT_CODE','FUND_SOURCE','BUDGET','NOTES','CREATED_AT'],
TRANSACTIONS:['ID','DOC_NO','DATE','SUPPLIER_ID','ACTIVITY_ID','DESCRIPTION','PAYMENT_METHOD','PAYMENT_REF','SUBTOTAL','TAX','TOTAL','STATUS','CREATED_BY','CREATED_AT'],
ITEMS:['ID','TRANSACTION_ID','NAME','SPEC','UNIT','QTY','PRICE','TOTAL'],
DOCUMENTS:['ID','TRANSACTION_ID','ACTIVITY_ID','TYPE','FILE_ID','FILE_URL','CREATED_AT'],
ATTACHMENTS:['ID','TRANSACTION_ID','ACTIVITY_ID','CATEGORY','NAME','MIME_TYPE','FILE_ID','FILE_URL','CREATED_AT'],
AUDIT_LOG:['ID','USER','ACTION','DETAIL','CREATED_AT'],
TEMPLATES:['TYPE','TITLE','INTRO','FOOTER','SIGN_LEFT','SIGN_RIGHT'],
CHECKLIST:['ID','TRANSACTION_ID','KEY','LABEL','REQUIRED','DONE','NOTE','UPDATED_AT'],
BUDGET_IMPORTS:['ID','PERIOD','FILE_NAME','ROWS_IMPORTED','IMPORTED_BY','IMPORTED_AT'],
TAX_LEDGER:['ID','TRANSACTION_ID','TAX_TYPE','TAX_BASE','RATE','AMOUNT','REFERENCE_NO','NOTE','CREATED_AT'],
APPROVALS:['ID','TRANSACTION_ID','STAGE','DECISION','NOTE','BY_USER','AT'],
PERIODS:['ID','PERIOD','STATUS','OPENED_AT','CLOSED_AT','CLOSED_BY']
}};
function ss_(){let id=PropertiesService.getScriptProperties().getProperty('SID');return id?SpreadsheetApp.openById(id):SpreadsheetApp.getActiveSpreadsheet()}
function sh_(n){let s=ss_().getSheetByName(n);if(!s)throw Error('Sheet '+n+' belum dibuat');return s}
function all_(n){let v=sh_(n).getDataRange().getValues();if(v.length<2)return[];return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(r=>Object.fromEntries(v[0].map((h,i)=>[h,r[i]])))}
function add_(n,o){sh_(n).appendRow(CFG.SHEETS[n].map(k=>o[k]===undefined?'':o[k]))}
function uid_(){return Utilities.getUuid()}function now_(){return Utilities.formatDate(new Date(),CFG.TZ,'yyyy-MM-dd HH:mm:ss')}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('')}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function ok_(d){return{ok:true,data:d}}function fail_(e){return{ok:false,error:String(e.message||e)}}
function setupSIPRKAS(){let s=SpreadsheetApp.getActiveSpreadsheet();if(!s)throw Error('Buka Apps Script dari Spreadsheet utama');PropertiesService.getScriptProperties().setProperty('SID',s.getId());
Object.keys(CFG.SHEETS).forEach(n=>{let t=s.getSheetByName(n)||s.insertSheet(n);if(!t.getLastRow())t.appendRow(CFG.SHEETS[n]);t.setFrozenRows(1);t.getRange(1,1,1,CFG.SHEETS[n].length).setFontWeight('bold').setBackground('#dbeafe')});
let it=DriveApp.getFoldersByName(CFG.FOLDER),folder=it.hasNext()?it.next():DriveApp.createFolder(CFG.FOLDER);PropertiesService.getScriptProperties().setProperty('FID',folder.getId());
if(!all_('SETTING').length)[['SCHOOL_NAME','Nama Sekolah'],['NPSN',''],['SCHOOL_ADDRESS',''],['HEAD_NAME','Nama Kepala Sekolah'],['HEAD_NIP',''],['TREASURER_NAME','Nama Bendahara'],['TREASURER_NIP',''],['FISCAL_YEAR',String(new Date().getFullYear())],['CITY','Kabupaten/Kota'],['DOC_PREFIX','RKAS'],['SCHOOL_LOGO_URL',''],['DEFAULT_FUND','BOSP Reguler']].forEach(([KEY,VALUE])=>add_('SETTING',{KEY,VALUE}));
if(!all_('USERS').length)resetInitialUsers();
if(!all_('TEMPLATES').length)[['KUITANSI','KUITANSI PEMBELIAN','Telah diterima pembayaran untuk pembelian barang berikut:','Pastikan bukti pembayaran sah dilampirkan.','Penerima/Penyedia','Mengetahui, Kepala Sekolah'],['SURAT_PESANAN','SURAT PESANAN BARANG','Dengan ini dipesan barang-barang berikut:','Pesanan sesuai spesifikasi dan kesepakatan.','Penyedia','Pihak Sekolah'],['PEMERIKSAAN','BERITA ACARA PEMERIKSAAN BARANG','Telah dilakukan pemeriksaan atas barang berikut:','Hasil pemeriksaan harus sesuai kondisi sebenarnya.','Petugas Pemeriksa','Penanggung Jawab'],['BAST','BERITA ACARA SERAH TERIMA BARANG','Telah dilakukan serah terima barang berikut:','Ditandatangani setelah barang benar-benar diterima.','Pihak Penyedia','Pihak Sekolah'],['REKAP','REKAPITULASI BELANJA','Rekapitulasi transaksi belanja berikut:','Rekap berdasarkan data transaksi tersimpan.','Bendahara','Kepala Sekolah']].forEach(r=>add_('TEMPLATES',Object.fromEntries(CFG.SHEETS.TEMPLATES.map((k,i)=>[k,r[i]]))));
return 'Setup V3 selesai: '+s.getUrl()+' | Folder: '+folder.getUrl()}
function resetInitialUsers(){let s=sh_('USERS');if(s.getLastRow()>1)s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).clearContent();[['kepsek','Kepala Sekolah','KEPALA_SEKOLAH'],['bendahara','Bendahara','BENDAHARA']].forEach(x=>add_('USERS',{ID:uid_(),USERNAME:x[0],PASSWORD_HASH:hash_('GantiSaya123!'),NAME:x[1],ROLE:x[2],ACTIVE:true,CREATED_AT:now_()}))}
function login_(p){let u=all_('USERS').find(x=>x.USERNAME===p.username&&x.ACTIVE===true&&x.PASSWORD_HASH===hash_(p.password));if(!u)throw Error('Username/password salah');let t=uid_()+uid_();CacheService.getScriptCache().put('S_'+t,JSON.stringify({id:u.ID,username:u.USERNAME,name:u.NAME,role:u.ROLE}),CFG.SESSION);return{token:t,user:{username:u.USERNAME,name:u.NAME,role:u.ROLE}}}
function auth_(p){let v=CacheService.getScriptCache().get('S_'+p.token);if(!v)throw Error('Sesi berakhir, silakan login kembali');return JSON.parse(v)}
function guard_(u){if(!['BENDAHARA','KEPALA_SEKOLAH'].includes(u.role))throw Error('Akses ditolak')}
function settings_(){return Object.fromEntries(all_('SETTING').map(x=>[x.KEY,x.VALUE]))}
function setSettings_(o){let s=sh_('SETTING'),v=s.getDataRange().getValues();Object.keys(o).forEach(k=>{let i=v.findIndex((r,n)=>n>0&&r[0]===k);if(i>0)s.getRange(i+1,2).setValue(o[k]);else add_('SETTING',{KEY:k,VALUE:o[k]})})}
function log_(u,a,d){add_('AUDIT_LOG',{ID:uid_(),USER:u.username,ACTION:a,DETAIL:JSON.stringify(d||{}),CREATED_AT:now_()})}
function upsert_(n,o){let s=sh_(n),v=s.getDataRange().getValues(),ix=v.findIndex((r,i)=>i>0&&r[0]===o.ID),row=CFG.SHEETS[n].map(k=>o[k]===undefined?'':o[k]);if(ix>0)s.getRange(ix+1,1,1,row.length).setValues([row]);else s.appendRow(row);return all_(n)}
function saveTx_(t,u){if(!t.date||!t.supplierId||!t.activityId||!t.items?.length)throw Error('Tanggal, penyedia, kegiatan dan barang wajib diisi');
let items=t.items.map(x=>({name:String(x.name||'').trim(),spec:x.spec||'',unit:x.unit||'unit',qty:Number(x.qty),price:Number(x.price)}));if(items.some(x=>!x.name||x.qty<=0||x.price<0||!isFinite(x.qty+x.price)))throw Error('Rincian barang tidak valid');
let lock=LockService.getScriptLock();lock.waitLock(20000);try{let st=settings_(),seq=all_('TRANSACTIONS').length+1,id=uid_(),subtotal=items.reduce((a,x)=>a+x.qty*x.price,0),tax=Number(t.tax||0),no=(st.DOC_PREFIX||'RKAS')+'/'+(st.FISCAL_YEAR||new Date().getFullYear())+'/'+String(seq).padStart(4,'0');
add_('TRANSACTIONS',{ID:id,DOC_NO:no,DATE:t.date,SUPPLIER_ID:t.supplierId,ACTIVITY_ID:t.activityId,DESCRIPTION:t.description||'',PAYMENT_METHOD:t.paymentMethod||'Transfer',PAYMENT_REF:t.paymentRef||'',SUBTOTAL:subtotal,TAX:tax,TOTAL:subtotal+tax,STATUS:'DRAFT',CREATED_BY:u.username,CREATED_AT:now_()});
items.forEach(x=>add_('ITEMS',{ID:uid_(),TRANSACTION_ID:id,NAME:x.name,SPEC:x.spec,UNIT:x.unit,QTY:x.qty,PRICE:x.price,TOTAL:x.qty*x.price}));
['invoice','payment','order','inspection','handover','photo'].forEach((k,i)=>add_('CHECKLIST',{ID:uid_(),TRANSACTION_ID:id,KEY:k,LABEL:['Nota/faktur penyedia','Bukti pembayaran','Surat pesanan','Pemeriksaan barang','Serah terima barang','Foto dokumentasi'][i],REQUIRED:i<5,DONE:false,NOTE:'',UPDATED_AT:now_()}));
if(tax>0)add_('TAX_LEDGER',{ID:uid_(),TRANSACTION_ID:id,TAX_TYPE:t.taxType||'PAJAK/LAINNYA',TAX_BASE:subtotal,RATE:Number(t.taxRate||0),AMOUNT:tax,REFERENCE_NO:t.taxReference||'',NOTE:t.taxNote||'Diinput manual; verifikasi sesuai bukti dan ketentuan berlaku',CREATED_AT:now_()});
log_(u,'CREATE_TX',{id,no,total:subtotal+tax});return{id,docNo:no,total:subtotal+tax}
}finally{lock.releaseLock()}}
function money_(n){return'Rp '+Number(n||0).toLocaleString('id-ID')}function date_(s){let p=String(s||'').slice(0,10).split('-');return p.length===3?p[2]+'/'+p[1]+'/'+p[0]:s}
function esc_(s){return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function terbilang_(n){n=Math.floor(Math.abs(n));let a=['','satu','dua','tiga','empat','lima','enam','tujuh','delapan','sembilan','sepuluh','sebelas'];if(n<12)return a[n];if(n<20)return terbilang_(n-10)+' belas';if(n<100)return terbilang_(Math.floor(n/10))+' puluh '+terbilang_(n%10);if(n<200)return'seratus '+terbilang_(n-100);if(n<1000)return terbilang_(Math.floor(n/100))+' ratus '+terbilang_(n%100);if(n<2000)return'seribu '+terbilang_(n-1000);if(n<1e6)return terbilang_(Math.floor(n/1000))+' ribu '+terbilang_(n%1000);if(n<1e9)return terbilang_(Math.floor(n/1e6))+' juta '+terbilang_(n%1e6);return String(n)}
function docHtml_(tx,type){let st=settings_(),sp=all_('SUPPLIERS').find(x=>x.ID===tx.SUPPLIER_ID)||{},ac=all_('ACTIVITIES').find(x=>x.ID===tx.ACTIVITY_ID)||{},items=all_('ITEMS').filter(x=>x.TRANSACTION_ID===tx.ID),tpl=all_('TEMPLATES').find(x=>x.TYPE===type)||{};
let rows=items.map((x,i)=>`<tr><td>${i+1}</td><td>${esc_(x.NAME)}</td><td>${esc_(x.SPEC)}</td><td>${esc_(x.UNIT)}</td><td>${x.QTY}</td><td>${money_(x.PRICE)}</td><td>${money_(x.TOTAL)}</td></tr>`).join('');
let total=type==='KUITANSI'?`<p><b>Jumlah: ${money_(tx.TOTAL)}</b></p><p>Terbilang: <b>${terbilang_(tx.TOTAL)} rupiah</b></p>`:`<p style="text-align:right"><b>Subtotal ${money_(tx.SUBTOTAL)}<br>Pajak/penyesuaian ${money_(tx.TAX)}<br>Total ${money_(tx.TOTAL)}</b></p>`;
return `<html><head><meta charset="utf-8"><style>body{font:11pt Arial;margin:35px}.head{text-align:center;border-bottom:3px double #111;padding-bottom:12px;margin-bottom:20px}h2{text-align:center;font-size:14pt}.meta{line-height:1.8}table{width:100%;border-collapse:collapse;margin:16px 0}td,th{border:1px solid #333;padding:6px}th{background:#eee}.sign{display:flex;justify-content:space-between;text-align:center;margin-top:55px}.sign div{width:42%;line-height:1.5}.foot{font-size:9pt;margin-top:28px;color:#555}</style></head><body><div class="head"><b>${esc_(st.SCHOOL_NAME)}</b><br>${esc_(st.SCHOOL_ADDRESS)}<br>NPSN ${esc_(st.NPSN)}</div><h2>${esc_(tpl.TITLE||type)}</h2><p style="text-align:center">Nomor: ${esc_(tx.DOC_NO)}</p><p>${esc_(tpl.INTRO)}</p><div class="meta">Tanggal: ${date_(tx.DATE)}<br>Penyedia: ${esc_(sp.NAME)}<br>Alamat: ${esc_(sp.ADDRESS)}<br>Kegiatan: ${esc_(ac.NAME)}<br>Kode rekening: ${esc_(ac.ACCOUNT_CODE)}<br>Sumber dana: BOSP Reguler</div><table><tr><th>No</th><th>Barang</th><th>Spesifikasi</th><th>Satuan</th><th>Qty</th><th>Harga</th><th>Jumlah</th></tr>${rows}</table>${total}<p>Metode pembayaran: ${esc_(tx.PAYMENT_METHOD)} ${esc_(tx.PAYMENT_REF)}</p><div class="sign"><div>${esc_(tpl.SIGN_LEFT)}<br><br><br><b>${esc_(sp.NAME)}</b></div><div>${esc_(tpl.SIGN_RIGHT)}<br><br><br><b>${esc_(st.HEAD_NAME)}</b></div></div><div class="foot">${esc_(tpl.FOOTER)} Dokumen harus diverifikasi dengan kejadian dan bukti yang sebenarnya.</div></body></html>`}
function pdfFor_(tx,type){return Utilities.newBlob(docHtml_(tx,type),'text/html','doc.html').getAs(MimeType.PDF).setName(type+'_'+tx.DOC_NO.replace(/[\/\\]/g,'-')+'.pdf')}
function makeDoc_(txid,type,u){let tx=all_('TRANSACTIONS').find(x=>x.ID===txid);if(!tx)throw Error('Transaksi tidak ditemukan');let f=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FID')).createFile(pdfFor_(tx,type));add_('DOCUMENTS',{ID:uid_(),TRANSACTION_ID:txid,ACTIVITY_ID:tx.ACTIVITY_ID,TYPE:type,FILE_ID:f.getId(),FILE_URL:f.getUrl(),CREATED_AT:now_()});log_(u,'GENERATE_DOC',{txid,type});return{name:f.getName(),url:f.getUrl(),type}}
function generateAll_(txid,u){return['KUITANSI','SURAT_PESANAN','PEMERIKSAAN','BAST','REKAP'].map(t=>makeDoc_(txid,t,u))}
function bundleTx_(txid,u){let tx=all_('TRANSACTIONS').find(x=>x.ID===txid);if(!tx)throw Error('Transaksi tidak ditemukan');let files=['KUITANSI','SURAT_PESANAN','PEMERIKSAAN','BAST','REKAP'].map(t=>pdfFor_(tx,t));let zip=Utilities.zip(files,'SPJ_'+tx.DOC_NO.replace(/[\/\\]/g,'-')+'.zip');let f=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FID')).createFile(zip);add_('DOCUMENTS',{ID:uid_(),TRANSACTION_ID:txid,ACTIVITY_ID:tx.ACTIVITY_ID,TYPE:'PAKET_ZIP',FILE_ID:f.getId(),FILE_URL:f.getUrl(),CREATED_AT:now_()});log_(u,'BUNDLE_TX',{txid});return{url:f.getUrl(),name:f.getName()}}
function upload_(p,u){let tx=all_('TRANSACTIONS').find(x=>x.ID===p.transactionId);if(!tx)throw Error('Transaksi tidak ditemukan');let bytes=Utilities.base64Decode(p.base64),blob=Utilities.newBlob(bytes,p.mimeType||'application/octet-stream',p.name||'lampiran');if(bytes.length>10*1024*1024)throw Error('Maksimal 10 MB per lampiran');let folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FID'));let f=folder.createFile(blob);add_('ATTACHMENTS',{ID:uid_(),TRANSACTION_ID:tx.ID,ACTIVITY_ID:tx.ACTIVITY_ID,CATEGORY:p.category||'LAINNYA',NAME:f.getName(),MIME_TYPE:f.getMimeType(),FILE_ID:f.getId(),FILE_URL:f.getUrl(),CREATED_AT:now_()});let key=p.category||'';if(key){let row=all_('CHECKLIST').find(x=>x.TRANSACTION_ID===tx.ID&&x.KEY===key);if(row)updateChecklist_(row.ID,{DONE:true})}log_(u,'UPLOAD',{tx:tx.ID,name:f.getName(),category:key});return{url:f.getUrl(),name:f.getName()}}
function updateChecklist_(id,o){let s=sh_('CHECKLIST'),v=s.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===id);if(i<1)throw Error('Checklist tidak ditemukan');let h=v[0];Object.keys(o).forEach(k=>{let c=h.indexOf(k);if(c>=0)s.getRange(i+1,c+1).setValue(o[k])});s.getRange(i+1,h.indexOf('UPDATED_AT')+1).setValue(now_());return true}
function checklist_(txid){return all_('CHECKLIST').filter(x=>x.TRANSACTION_ID===txid)}
function report_(p){let from=p.from||'0000-01-01',to=p.to||'9999-12-31',tx=all_('TRANSACTIONS').filter(x=>String(x.DATE)>=from&&String(x.DATE)<=to);if(p.activityId)tx=tx.filter(x=>x.ACTIVITY_ID===p.activityId);return tx}
function summary_(){let tx=all_('TRANSACTIONS'),acts=all_('ACTIVITIES');return acts.map(a=>{let list=tx.filter(t=>t.ACTIVITY_ID===a.ID),real=list.reduce((s,t)=>s+Number(t.TOTAL||0),0);return{activityId:a.ID,code:a.CODE,name:a.NAME,accountCode:a.ACCOUNT_CODE,budget:Number(a.BUDGET||0),realization:real,remaining:Number(a.BUDGET||0)-real,transactions:list.length}})}
function importBudget_(p,u){let rows=p.rows||[];if(!rows.length)throw Error('Tidak ada baris impor');let count=0;rows.forEach(r=>{if(!r.name) return;let exists=all_('ACTIVITIES').find(a=>a.CODE===String(r.code||'')&&a.ACCOUNT_CODE===String(r.accountCode||''));let o={ID:exists?exists.ID:uid_(),CODE:String(r.code||''),NAME:String(r.name),ACCOUNT_CODE:String(r.accountCode||''),FUND_SOURCE:'BOSP Reguler',BUDGET:Number(r.budget||0),NOTES:r.notes||'',CREATED_AT:now_()};upsert_('ACTIVITIES',o);count++});add_('BUDGET_IMPORTS',{ID:uid_(),PERIOD:p.period||settings_().FISCAL_YEAR,FILE_NAME:p.fileName||'Input manual/CSV',ROWS_IMPORTED:count,IMPORTED_BY:u.username,IMPORTED_AT:now_()});log_(u,'IMPORT_BUDGET',{count});return{count,activities:all_('ACTIVITIES')}}
function exportActivityZip_(activityId,u){let act=all_('ACTIVITIES').find(x=>x.ID===activityId);if(!act)throw Error('Kegiatan tidak ditemukan');let tx=all_('TRANSACTIONS').filter(x=>x.ACTIVITY_ID===activityId),blobs=[];tx.forEach(t=>{['KUITANSI','SURAT_PESANAN','PEMERIKSAAN','BAST','REKAP'].forEach(type=>blobs.push(pdfFor_(t,type)))});all_('ATTACHMENTS').filter(x=>x.ACTIVITY_ID===activityId).forEach(a=>{try{blobs.push(DriveApp.getFileById(a.FILE_ID).getBlob().setName('LAMPIRAN_'+a.NAME))}catch(e){}});if(!blobs.length)throw Error('Belum ada dokumen/lampiran untuk kegiatan ini');let zip=Utilities.zip(blobs,'SPJ_'+act.CODE+'_'+act.NAME.replace(/[^\w-]/g,'_')+'.zip'),f=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FID')).createFile(zip);log_(u,'EXPORT_ACTIVITY',{activityId,count:tx.length});return{url:f.getUrl(),name:f.getName(),transactions:tx.length}}
function changePass_(u,p){if(String(p.newPassword||'').length<10)throw Error('Password minimal 10 karakter');let x=all_('USERS').find(a=>a.ID===u.id);if(!x||x.PASSWORD_HASH!==hash_(p.oldPassword))throw Error('Password lama salah');let s=sh_('USERS'),v=s.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===u.id);s.getRange(i+1,3).setValue(hash_(p.newPassword));return true}
function backup_(u){let source=ss_(),copy=DriveApp.getFileById(source.getId()).makeCopy('BACKUP_SIP_RKAS_'+Utilities.formatDate(new Date(),CFG.TZ,'yyyyMMdd_HHmmss'));log_(u,'BACKUP',{fileId:copy.getId()});return{url:copy.getUrl(),name:copy.getName()}}

function approveTx_(p,u){
 let tx=all_('TRANSACTIONS').find(x=>x.ID===p.transactionId);if(!tx)throw Error('Transaksi tidak ditemukan');
 if(u.role!=='KEPALA_SEKOLAH')throw Error('Hanya Kepala Sekolah yang dapat memberi keputusan');
 if(tx.STATUS==='VOID')throw Error('Transaksi dibatalkan');
 let decision=p.decision;if(!['APPROVED','REJECTED'].includes(decision))throw Error('Keputusan tidak valid');
 let s=sh_('TRANSACTIONS'),v=s.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===tx.ID),c=v[0].indexOf('STATUS');
 s.getRange(i+1,c+1).setValue(decision);
 add_('APPROVALS',{ID:uid_(),TRANSACTION_ID:tx.ID,STAGE:'KEPALA_SEKOLAH',DECISION:decision,NOTE:p.note||'',BY_USER:u.username,AT:now_()});
 log_(u,'APPROVAL',{transactionId:tx.ID,decision,note:p.note||''});return{status:decision};
}
function voidTx_(p,u){
 if(u.role!=='KEPALA_SEKOLAH')throw Error('Hanya Kepala Sekolah yang dapat membatalkan transaksi');
 let tx=all_('TRANSACTIONS').find(x=>x.ID===p.transactionId);if(!tx)throw Error('Transaksi tidak ditemukan');
 if(tx.STATUS==='VOID')throw Error('Transaksi sudah dibatalkan');
 let s=sh_('TRANSACTIONS'),v=s.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===tx.ID),c=v[0].indexOf('STATUS');
 s.getRange(i+1,c+1).setValue('VOID');
 add_('APPROVALS',{ID:uid_(),TRANSACTION_ID:tx.ID,STAGE:'PEMBATALAN',DECISION:'VOID',NOTE:p.note||'Pembatalan oleh Kepala Sekolah',BY_USER:u.username,AT:now_()});
 log_(u,'VOID_TX',{transactionId:tx.ID,note:p.note||''});return{status:'VOID'};
}
function auditReport_(p){
 let rows=all_('AUDIT_LOG');if(p.from)rows=rows.filter(x=>String(x.CREATED_AT)>=p.from);
 if(p.to)rows=rows.filter(x=>String(x.CREATED_AT)<=p.to+' 23:59:59');
 if(p.username)rows=rows.filter(x=>x.USER===p.username);return rows.slice(-500).reverse();
}
function taxSummary_(){
 let rows=all_('TAX_LEDGER');let groups={};rows.forEach(x=>{let k=x.TAX_TYPE||'PAJAK/LAINNYA';groups[k]=(groups[k]||0)+Number(x.AMOUNT||0)});
 return Object.keys(groups).map(k=>({type:k,amount:groups[k],count:rows.filter(x=>(x.TAX_TYPE||'PAJAK/LAINNYA')===k).length}));
}
function closePeriod_(p,u){
 if(u.role!=='KEPALA_SEKOLAH')throw Error('Hanya Kepala Sekolah yang dapat menutup periode');
 let period=String(p.period||'').trim();if(!period)throw Error('Periode wajib diisi');
 let old=all_('PERIODS').find(x=>x.PERIOD===period);
 if(old&&old.STATUS==='CLOSED')throw Error('Periode sudah ditutup');
 let obj={ID:old?old.ID:uid_(),PERIOD:period,STATUS:'CLOSED',OPENED_AT:old?old.OPENED_AT:now_(),CLOSED_AT:now_(),CLOSED_BY:u.username};
 upsert_('PERIODS',obj);log_(u,'CLOSE_PERIOD',{period});return obj;
}
function reopenPeriod_(p,u){
 if(u.role!=='KEPALA_SEKOLAH')throw Error('Hanya Kepala Sekolah yang dapat membuka kembali periode');
 let x=all_('PERIODS').find(a=>a.PERIOD===p.period);if(!x)throw Error('Periode tidak ditemukan');
 x.STATUS='OPEN';x.CLOSED_AT='';x.CLOSED_BY='';upsert_('PERIODS',x);log_(u,'REOPEN_PERIOD',{period:p.period});return x;
}
function assertPeriodOpen_(date){
 let period=String(date||'').slice(0,4),x=all_('PERIODS').find(a=>a.PERIOD===period&&a.STATUS==='CLOSED');
 if(x)throw Error('Periode '+period+' sudah ditutup. Minta Kepala Sekolah membuka kembali periode.');
}

function route_(a,p,u){guard_(u);switch(a){
case'bootstrap':return{user:u,settings:settings_(),suppliers:all_('SUPPLIERS'),activities:all_('ACTIVITIES'),transactions:all_('TRANSACTIONS'),items:all_('ITEMS'),documents:all_('DOCUMENTS'),attachments:all_('ATTACHMENTS'),templates:all_('TEMPLATES'),checklists:all_('CHECKLIST'),taxLedger:all_('TAX_LEDGER'),approvals:all_('APPROVALS'),periods:all_('PERIODS'),audit:all_('AUDIT_LOG').slice(-250).reverse(),summary:summary_()};
case'saveSettings':setSettings_(p.settings||{});log_(u,'SETTINGS',p.settings);return settings_();
case'saveSupplier':let s=p.supplier;if(!s.NAME)throw Error('Nama penyedia wajib');return upsert_('SUPPLIERS',{ID:s.ID||uid_(),NAME:s.NAME,ADDRESS:s.ADDRESS||'',PHONE:s.PHONE||'',TAX_ID:s.TAX_ID||'',NOTES:s.NOTES||'',CREATED_AT:now_()});
case'saveActivity':let a=p.activity;if(!a.NAME)throw Error('Nama kegiatan wajib');return upsert_('ACTIVITIES',{ID:a.ID||uid_(),CODE:a.CODE||'',NAME:a.NAME,ACCOUNT_CODE:a.ACCOUNT_CODE||'',FUND_SOURCE:'BOSP Reguler',BUDGET:Number(a.BUDGET||0),NOTES:a.NOTES||'',CREATED_AT:now_()});
case'saveTransaction':assertPeriodOpen_(p.transaction&&p.transaction.date);return saveTx_(p.transaction,u);
case'generateAll':return generateAll_(p.transactionId,u);
case'bundleTransaction':return bundleTx_(p.transactionId,u);
case'uploadAttachment':return upload_(p,u);
case'updateChecklist':updateChecklist_(p.id,{DONE:!!p.done,NOTE:p.note||''});return checklist_(p.transactionId);
case'report':return report_(p);
case'budgetSummary':return summary_();
case'importBudget':return importBudget_(p,u);
case'exportActivityZip':return exportActivityZip_(p.activityId,u);
case'saveTemplate':let t=p.template,sh=sh_('TEMPLATES'),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===t.TYPE),row=CFG.SHEETS.TEMPLATES.map(k=>t[k]||'');if(i>0)sh.getRange(i+1,1,1,row.length).setValues([row]);else sh.appendRow(row);return all_('TEMPLATES');
case'changePassword':return changePass_(u,p);
case'backup':return backup_(u);
case'approveTransaction':return approveTx_(p,u);
case'voidTransaction':return voidTx_(p,u);
case'auditReport':return auditReport_(p);
case'taxSummary':return taxSummary_();
case'closePeriod':return closePeriod_(p,u);
case'reopenPeriod':return reopenPeriod_(p,u);
case'logout':CacheService.getScriptCache().remove('S_'+p.token);return true;
default:throw Error('Aksi tidak dikenal: '+a)}}
function doGet(e){return out_(fail_('Gunakan POST dari aplikasi SIP-RKAS'))}
function doPost(e){try{let p=JSON.parse(e.postData.contents||'{}'),a=p.action;if(a==='login')return out_(ok_(login_(p)));let u=auth_(p);return out_(ok_(route_(a,p,u)))}catch(x){return out_(fail_(x))}}
