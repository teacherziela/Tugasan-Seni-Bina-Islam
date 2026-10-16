// HISTORYVERSE 360 — backend tugasan dan dashboard guru dalam Google Sheets.
// 1. Cipta Google Spreadsheet kosong dan folder Google Drive.
// 2. Isi ID kedua-duanya di bawah, kemudian jalankan setupDashboard() sekali.
// 3. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
// 4. Salin URL /exec ke SCRIPT_URL dalam index.html dan muat naik ke GitHub Pages.
const SHEET_ID = '1i_IyuupzGqXOMqqB5OIA5Bdr9Y9IFtHK_Lbompj5itU';
const FOLDER_ID = '1qDqfOSKjJjJXE2KIOAZKKCnxtNCT0NWa';
const TAB = 'Penghantaran';
const HEADERS = ['Tarikh Hantar','ID','Nama Murid','Kelas','Nama Rekaan','Fungsi Bangunan','Ciri Seni Bina','Pautan Fail','Status Semakan','TP PBD','Ulasan Guru'];
function setupDashboard(){
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sh=ss.getSheetByName(TAB)||ss.insertSheet(TAB);
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,HEADERS.length).setBackground('#56338e').setFontColor('white').setFontWeight('bold');
  sh.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Belum Semak','Sudah Semak'],true).build());
  sh.getRange('J2:J').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TP1','TP2','TP3','TP4','TP5','TP6'],true).build());
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.autoResizeColumns(1,11);
  if(!sh.getFilter())sh.getRange(1,1,Math.max(sh.getMaxRows(),2),11).createFilter();
  let summary=ss.getSheetByName('Ringkasan')||ss.insertSheet('Ringkasan');
  summary.clear();
  summary.getRange('A1:B1').setValues([['HISTORYVERSE 360 — ARKITEK MUDA','Bilangan']]).setBackground('#56338e').setFontColor('white').setFontWeight('bold');
  summary.getRange('A2:A5').setValues([['Jumlah penghantaran'],['Belum Semak'],['Sudah Semak'],['Telah diberi TP']]);
  summary.getRange('B2').setFormula("=COUNTA('Penghantaran'!B2:B)");
  summary.getRange('B3').setFormula("=COUNTIF('Penghantaran'!I2:I,\"Belum Semak\")");
  summary.getRange('B4').setFormula("=COUNTIF('Penghantaran'!I2:I,\"Sudah Semak\")");
  summary.getRange('B5').setFormula("=COUNTIF('Penghantaran'!J2:J,\"TP*\")");
  summary.getRange('A7:B7').setValues([['Kelas','Bilangan Hantar']]).setBackground('#56338e').setFontColor('white').setFontWeight('bold');
  const classes=['1 Adil','1 Gigih','1 Cendekia','1 Yakin','1 Bestari','1 Tekun','1 Usaha'];
  summary.getRange(8,1,classes.length,1).setValues(classes.map(x=>[x]));
  classes.forEach((_,i)=>summary.getRange(8+i,2).setFormula(`=COUNTIF('Penghantaran'!D2:D,A${8+i})`));
  summary.autoResizeColumns(1,2);
}
function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const p=JSON.parse(e.postData.contents);
    for(const k of ['nama','kelas','tajuk','fungsi','ciri','base64','fileName','mimeType'])if(!p[k])throw Error('Maklumat tidak lengkap: '+k);
    if(!['application/pdf','image/jpeg','image/png'].includes(p.mimeType))throw Error('Jenis fail tidak dibenarkan');
    if(!/^1 (Adil|Gigih|Cendekia|Yakin|Bestari|Tekun|Usaha)$/.test(p.kelas))throw Error('Kelas tidak sah');
    const bytes=Utilities.base64Decode(p.base64);if(bytes.length>8*1024*1024)throw Error('Fail terlalu besar');
    const clean=s=>String(s).replace(/[\r\n\t]/g,' ').slice(0,120).replace(/^[=+@-]/,"'");
    const id=Utilities.getUuid().slice(0,8).toUpperCase();
    const ext=p.mimeType==='application/pdf'?'.pdf':p.mimeType==='image/png'?'.png':'.jpg';
    const file=DriveApp.getFolderById(FOLDER_ID).createFile(Utilities.newBlob(bytes,p.mimeType,id+'_'+clean(p.nama)+ext));
    const sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
    if(!sh)throw Error('Jalankan setupDashboard dahulu');
    sh.appendRow([new Date(),id,clean(p.nama),clean(p.kelas),clean(p.tajuk),clean(p.fungsi),clean(p.ciri),file.getUrl(),'Belum Semak','','']);
    return out({ok:true,id});
  }catch(err){return out({ok:false,error:String(err.message||err)});}finally{lock.releaseLock();}
}
function out(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
