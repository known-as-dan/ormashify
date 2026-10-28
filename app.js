import {parseNativeDocx,generateReport} from './report.js';
import {getDocument,GlobalWorkerOptions} from './vendor/pdf.min.mjs';
GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.min.mjs',import.meta.url).href;
const $=id=>document.getElementById(id);
let records=[],previewUrl=null,busy=false,sourceVersion=0,previewPdf=null,previewPage=1,renderTask=null;
function message(text,success=false){$('message').textContent=text;$('message').hidden=!text;$('message').classList.toggle('success',success);}
function sync(){const complete=records.length&&$('site').value.trim()&&$('system').value.trim()&&$('inspection-date').validity.valid&&$('inspection-date').value;$('generate').disabled=busy||!complete;for(const id of ['source-file','site','system','inspection-date'])$(id).disabled=busy;}
function invalidate(){if(renderTask){renderTask.cancel();renderTask=null;}if(previewPdf){previewPdf.destroy();previewPdf=null;}if(previewUrl){URL.revokeObjectURL(previewUrl);previewUrl=null;}$('pdf-preview').hidden=true;$('placeholder').hidden=false;$('download').hidden=true;$('download').removeAttribute('href');$('page-count').textContent='A4 · Ormash format';}
async function renderPreview(){if(!previewPdf)return;const version=previewPdf,requested=previewPage,page=await version.getPage(requested);if(version!==previewPdf||requested!==previewPage)return;if(renderTask)renderTask.cancel();const viewport=page.getViewport({scale:1.6}),canvas=$('pdf-canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);const task=page.render({canvasContext:canvas.getContext('2d'),viewport});renderTask=task;try{await task.promise;}catch(e){if(e.name==='RenderingCancelledException')return;throw e;}finally{if(renderTask===task)renderTask=null;}if(version!==previewPdf||requested!==previewPage)return;$('current-page').textContent=`Page ${requested} of ${version.numPages}`;$('previous-page').disabled=requested===1;$('next-page').disabled=requested===version.numPages;}
async function navigate(step){if(!previewPdf||renderTask)return;previewPage=Math.max(1,Math.min(previewPdf.numPages,previewPage+step));try{await renderPreview();}catch{message('This preview page could not be displayed. Rebuild the report and try again.');}}
$('previous-page').addEventListener('click',()=>navigate(-1));
$('next-page').addEventListener('click',()=>navigate(1));
async function selectFile(file){if(!file||busy)return;const version=++sourceVersion;records=[];invalidate();sync();message('');$('source-summary').hidden=true;
 if(!/\.docx$/i.test(file.name)){message('Choose a Basic Word report (.docx) exported from SmartView.');$('source-file').value='';return;}
 $('drop-title').textContent='Reading your report…';$('drop-subtitle').textContent=file.name;busy=true;sync();
 try{const parsed=await parseNativeDocx(await file.arrayBuffer());if(version!==sourceVersion)return;records=parsed;$('drop-title').textContent='Change Word report';$('source-summary').textContent=`${records.length} image records · thermal and ordinary photos included`;$('source-summary').hidden=false;}
 catch(error){message(error.message||'The report could not be read.');$('drop-title').textContent='Choose a Word report';$('drop-subtitle').textContent='or drop your .docx here';$('source-file').value='';}
 finally{busy=false;sync();}
}
$('source-file').addEventListener('change',e=>{const file=e.target.files[0];e.target.value='';selectFile(file);});
for(const event of ['dragenter','dragover'])$('drop-zone').addEventListener(event,e=>{e.preventDefault();$('drop-zone').classList.add('dragging');});
for(const event of ['dragleave','drop'])$('drop-zone').addEventListener(event,e=>{e.preventDefault();$('drop-zone').classList.remove('dragging');});
$('drop-zone').addEventListener('drop',e=>selectFile(e.dataTransfer.files[0]));
document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
for(const id of ['site','system','inspection-date'])$(id).addEventListener('input',()=>{invalidate();message('');sync();});
$('report-form').addEventListener('submit',async e=>{e.preventDefault();if(!records.length||busy||!e.target.reportValidity())return;busy=true;sync();message('Building your report…');
 try{const result=await generateReport(records,{site:$('site').value.trim(),system:$('system').value.trim(),date:$('inspection-date').value},(current,total)=>message(`Building report · image ${current} of ${total}`));
 invalidate();previewUrl=URL.createObjectURL(new Blob([result.bytes],{type:'application/pdf'}));previewPdf=await getDocument({data:result.bytes.slice(),isEvalSupported:false}).promise;previewPage=1;$('pdf-preview').hidden=false;$('placeholder').hidden=true;await renderPreview();$('download').href=previewUrl;$('download').download=result.filename;$('download').hidden=false;$('page-count').textContent=`${result.pageCount} pages · A4`;
 message(`${result.imageCount} image records formatted. Check the preview, then download your PDF.${result.pageCount>result.imageCount?' Longer source tables continue onto additional pages so nothing is cut off.':''}`,true);
 }catch(error){message(error.message||'The PDF could not be generated. Try a smaller Basic Word export.');}
 finally{busy=false;sync();}
});
window.addEventListener('pagehide',()=>{if(previewUrl)URL.revokeObjectURL(previewUrl);});
sync();
