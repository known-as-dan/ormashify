import {parseNativeDocx,generateReport} from './report.js';
import {getDocument,GlobalWorkerOptions} from './vendor/pdf.min.mjs';
GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.min.mjs',import.meta.url).href;
const $=id=>document.getElementById(id);
let records=[],previewUrl=null,busy=false,sourceVersion=0,previewPdf=null,previewPage=1,renderTask=null;
function message(text,success=false){$('message').textContent=text;$('message').hidden=!text;$('message').classList.toggle('success',success);}
function displayError(error){
 const text=String(error?.message||'');
 const messages=[
  [/100 MB/,'הקובץ גדול מ־100 מגה־בייט. יש לפצל אותו לדוחות קטנים יותר ב־SmartView.'],
  [/extra comments|header\/footer/,'הקובץ כולל הערות או תוכן נוסף בכותרת העליונה או התחתונה. יש לייצא דוח Basic ללא תוספות אלה.'],
  [/tracked changes/,'הקובץ כולל מעקב שינויים. יש לאשר או לדחות את השינויים בעותק של הקובץ.'],
  [/external|linked rather|image link/,'הקובץ מפנה לתמונות חיצוניות. יש לייצא דוח שבו התמונות כלולות בקובץ.'],
  [/characters this formatter cannot render/,'בקובץ יש תווים שאינם נתמכים. יש לשמור את המקור ולהעביר אותו לעיצוב ידני.'],
  [/site name is too long/,'שם האתר ארוך מדי לכותרת. יש להזין את שם האתר ללא הערות נוספות.'],
  [/asset could not load/,'לא ניתן לטעון רכיב נדרש. יש לרענן את העמוד ולנסות שוב.'],
  [/missing its thermal or ordinary/,'חסר צילום תרמי או צילום רגיל. יש לכלול את שניהם בייצוא מ־SmartView.'],
  [/too long|too large|400 image/,'תוכן הקובץ חורג מהפורמט הנתמך. יש לפצל את דוח המקור או להעביר אותו לעיצוב ידני.'],
  [/valid inspection date/,'יש להזין שם אתר, קוד מערכת ותאריך בדיקה תקין.']
 ];
 return messages.find(([pattern])=>pattern.test(text))?.[1]||'לא ניתן לעצב את הקובץ הזה. יש לייצא מחדש מ־SmartView בתבנית Basic לפי ההנחיות.';
}
function sync(){const complete=records.length&&$('site').value.trim()&&$('system').value.trim()&&$('inspection-date').validity.valid&&$('inspection-date').value;$('generate').disabled=busy||!complete;for(const id of ['source-file','site','system','inspection-date'])$(id).disabled=busy;}
function invalidate(){if(renderTask){renderTask.cancel();renderTask=null;}if(previewPdf){previewPdf.destroy();previewPdf=null;}if(previewUrl){URL.revokeObjectURL(previewUrl);previewUrl=null;}$('pdf-preview').hidden=true;$('placeholder').hidden=false;$('download').hidden=true;$('download').removeAttribute('href');$('page-count').textContent='A4';}
async function renderPreview(){if(!previewPdf)return;const version=previewPdf,requested=previewPage,page=await version.getPage(requested);if(version!==previewPdf||requested!==previewPage)return;if(renderTask)renderTask.cancel();const viewport=page.getViewport({scale:1.6}),canvas=$('pdf-canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);const task=page.render({canvasContext:canvas.getContext('2d'),viewport});renderTask=task;try{await task.promise;}catch(e){if(e.name==='RenderingCancelledException')return;throw e;}finally{if(renderTask===task)renderTask=null;}if(version!==previewPdf||requested!==previewPage)return;$('current-page').textContent=`עמוד ${requested} מתוך ${version.numPages}`;$('previous-page').disabled=requested===1;$('next-page').disabled=requested===version.numPages;}
async function navigate(step){if(!previewPdf||renderTask)return;previewPage=Math.max(1,Math.min(previewPdf.numPages,previewPage+step));try{await renderPreview();}catch{message('לא ניתן להציג את העמוד. יש ליצור את הדוח מחדש.');}}
$('previous-page').addEventListener('click',()=>navigate(-1));
$('next-page').addEventListener('click',()=>navigate(1));
async function selectFile(file){if(!file||busy)return;const version=++sourceVersion;records=[];invalidate();sync();message('');$('source-summary').hidden=true;
 if(!/\.docx$/i.test(file.name)){message('יש לבחור קובץ Word שיוצא מ־SmartView בתבנית Basic.');$('source-file').value='';return;}
 $('drop-title').textContent='קריאת הקובץ…';$('drop-subtitle').textContent=file.name;busy=true;sync();
 try{const parsed=await parseNativeDocx(await file.arrayBuffer());if(version!==sourceVersion)return;records=parsed;$('drop-title').textContent='בחירת קובץ אחר';$('source-summary').textContent=`${records.length} צילומים תרמיים`;$('source-summary').hidden=false;}
 catch(error){message(displayError(error));$('drop-title').textContent='בחירת קובץ Word';$('drop-subtitle').textContent='או גרירת קובץ DOCX לכאן';$('source-file').value='';}
 finally{busy=false;sync();}
}
$('source-file').addEventListener('change',e=>{const file=e.target.files[0];e.target.value='';selectFile(file);});
for(const event of ['dragenter','dragover'])$('drop-zone').addEventListener(event,e=>{e.preventDefault();$('drop-zone').classList.add('dragging');});
for(const event of ['dragleave','drop'])$('drop-zone').addEventListener(event,e=>{e.preventDefault();$('drop-zone').classList.remove('dragging');});
$('drop-zone').addEventListener('drop',e=>selectFile(e.dataTransfer.files[0]));
document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
for(const id of ['site','system','inspection-date'])$(id).addEventListener('input',()=>{invalidate();message('');sync();});
$('report-form').addEventListener('submit',async e=>{e.preventDefault();if(!records.length||busy||!e.target.reportValidity())return;busy=true;sync();message('יצירת הדוח…');
 try{const result=await generateReport(records,{site:$('site').value.trim(),system:$('system').value.trim(),date:$('inspection-date').value},(current,total)=>message(`יצירת דוח: ${current} מתוך ${total}`));
 invalidate();previewUrl=URL.createObjectURL(new Blob([result.bytes],{type:'application/pdf'}));previewPdf=await getDocument({data:result.bytes.slice(),isEvalSupported:false}).promise;previewPage=1;$('pdf-preview').hidden=false;$('placeholder').hidden=true;await renderPreview();$('download').href=previewUrl;$('download').download=result.filename;$('download').hidden=false;$('page-count').textContent=`${result.pageCount} עמודים · A4`;
 message(`הדוח מוכן.${result.pageCount>result.imageCount?' טבלאות ארוכות והערות נמשכות בעמודים נוספים.':''}`,true);
 }catch(error){message(displayError(error));}
 finally{busy=false;sync();}
});
window.addEventListener('pagehide',()=>{if(previewUrl)URL.revokeObjectURL(previewUrl);});
sync();
