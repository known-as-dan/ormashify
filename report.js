/* Browser-only formatter for native SmartView Basic DOCX exports. */
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ink='#283B47', muted='#5D707B', rule='#D2DCE1', accent='#E34B1C';
const bidi=globalThis.bidi_js();
const descendants=(node,ns,tag)=>Array.from(node.getElementsByTagNameNS(ns,tag));
const children=(node,tag)=>Array.from(node.children||node.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName===tag);
const hasText=node=>descendants(node,W,'t').some(n=>n.textContent.trim());
function inline(node){let out='';for(const n of Array.from(node.childNodes)){if(n.nodeType!==1)continue;if(n.namespaceURI===W&&n.localName==='t')out+=n.textContent;else if(n.namespaceURI===W&&n.localName==='tab')out+='\t';else if(n.namespaceURI===W&&n.localName==='br')out+='\n';else out+=inline(n);}return out;}
const cellText=cell=>children(cell,'p').map(inline).join('\n');
const rowsOf=table=>children(table,'tr').map(row=>children(row,'tc').map(cellText));
function xml(bytes,name){const s=new TextDecoder().decode(bytes);if(/<!DOCTYPE|<!ENTITY/i.test(s))throw new Error('This document uses an unsupported XML format. Please export a fresh Basic Word report from SmartView.');const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('The Word document could not be read. Export it again from SmartView.');return d;}
function sourcePart(target){if(!target||/^[a-z]+:/i.test(target)||target.startsWith('/'))throw new Error('Linked external images aren’t supported. Include the images in the Word export.');const parts=['word'];for(const bit of target.replaceAll('\\','/').split('/')){if(bit==='..')parts.pop();else if(bit!=='.')parts.push(bit);}const result=parts.join('/');if(!result.startsWith('word/'))throw new Error('The document contains an unsupported image link.');return result;}
function imageType(bytes){if(bytes[0]===255&&bytes[1]===216)return'image/jpeg';if(bytes[0]===137&&bytes[1]===80&&bytes[2]===78&&bytes[3]===71)return'image/png';throw new Error('A report image uses an unsupported format. Export a Basic Word report with its original photos.');}
function extraPartContent(doc){
  if(descendants(doc,A,'blip').length||descendants(doc,W,'pict').length)return true;
  let field=null,substantive=false;
  for(const n of Array.from(doc.getElementsByTagName('*'))){if(n.namespaceURI!==W)continue;
    if(n.localName==='fldChar'){const type=n.getAttributeNS(W,'fldCharType');if(type==='begin'){if(field)return true;field={instruction:'',text:''};}else if(type==='end'&&field){if(!/^\s*(PAGE|NUMPAGES)\s*$/i.test(field.instruction)&&field.text.trim())substantive=true;field=null;}}
    else if(n.localName==='instrText'&&field)field.instruction+=n.textContent;
    else if(n.localName==='t'){if(field)field.text+=n.textContent;else if(n.textContent.trim())substantive=true;}
  }
  return substantive||!!field;
}
function assertTextPreserved(block,record){const count=s=>{const m=new Map();for(const c of s.replace(/\s/g,''))m.set(c,(m.get(c)||0)+1);return m;};const original=count(block.flatMap(n=>descendants(n,W,'t').map(t=>t.textContent)).join(''));const mapped=count([...record.images.main.captions,...record.images.visible.captions,...record.tables.flat(2),...record.notes,'Image Info','Main Image Markers'].join(''));for(const[c,n]of original)if((mapped.get(c)||0)<n)throw new Error('Some source text does not fit the supported Basic export structure. Formatting stopped so it is not left out.');}
export async function parseNativeDocx(buffer){
  const bytes=new Uint8Array(buffer);if(bytes.byteLength>100*1024*1024)throw new Error('This report is larger than 100 MB. Split it into smaller Basic Word reports in SmartView.');
  let entries,total=0;try{entries=globalThis.fflate.unzipSync(bytes,{filter:file=>{total+=file.originalSize;if(total>350*1024*1024)throw new Error('expanded size');return true;}});}catch{throw new Error('This file could not be opened as a Word report, or it is too large. Choose an original SmartView .docx export.');}
  for(const k of ['word/document.xml','word/_rels/document.xml.rels'])if(!entries[k])throw new Error('Choose the Basic Word (.docx) report exported by SmartView.');
  for(const [name,raw]of Object.entries(entries)){if(/^word\/(comments|footnotes|endnotes|header\d+|footer\d+)\.xml$/.test(name)&&extraPartContent(xml(raw,name)))throw new Error('This export has extra comments, footnotes or header/footer content. Use a Basic export without these additions so no source content is lost.');if(/vbaProject|embeddings\//i.test(name))throw new Error('Embedded documents or macros aren’t supported. Choose an original Basic export.');}
  const document=xml(entries['word/document.xml']),rels=xml(entries['word/_rels/document.xml.rels']);
  if(descendants(document,W,'del').length||descendants(document,W,'ins').length)throw new Error('This document has tracked changes. Resolve them in a copy before formatting.');
  const links=new Map(Array.from(rels.documentElement.children).map(n=>[n.getAttribute('Id'),{target:n.getAttribute('Target'),mode:n.getAttribute('TargetMode')}]));
  const body=descendants(document,W,'body')[0],blocks=[];if(!body)throw new Error('This Word document has no report pages.');let current=[];
  for(const n of Array.from(body.children)){if(n.localName==='sectPr')continue;current.push(n);if(descendants(n,W,'br').some(b=>b.getAttributeNS(W,'type')==='page')){blocks.push(current);current=[];}}
  if(current.some(n=>hasText(n)||descendants(n,A,'blip').length))blocks.push(current);
  if(!blocks.length||blocks.length>400)throw new Error('Use a Basic Word report containing between 1 and 400 image records.');
  const records=[],seen=new Set();
  for(const [index,block]of blocks.entries()){
    if(block.some(n=>!['p','tbl'].includes(n.localName)&&(hasText(n)||descendants(n,A,'blip').length))||block.some(n=>descendants(n,W,'pict').length||descendants(n,W,'object').length||n.localName==='altChunk'))throw new Error('This export has extra content outside the supported Basic layout. Keep the source and use the full formatting workflow for this report.');
    const pictureTables=block.filter(n=>n.localName==='tbl'&&descendants(n,A,'blip').length);
    if(pictureTables.length!==1||block.some(n=>n.localName!=='tbl'&&descendants(n,A,'blip').length))throw new Error(`Image record ${index+1} has a different layout. Export the report using SmartView’s Basic template.`);
    const pictureTable=pictureTables[0],images={};
    for(const cell of descendants(pictureTable,W,'tc')){
      const blips=descendants(cell,A,'blip');if(!blips.length){if(hasText(cell))throw new Error('The image table includes extra text that needs a different layout. Use an original Basic export.');continue;}
      if(blips.length!==1)throw new Error('Multiple photos share an image cell. Use the original Basic Word export.');
      const link=links.get(blips[0].getAttributeNS(R,'embed'));if(!link||link.mode==='External')throw new Error('A report photo is linked rather than included. Export with embedded photos.');
      const media=sourcePart(link.target),kind=/MAIN_IMAGE_/i.test(media)?'main':/VISIBLE_/i.test(media)?'visible':null;
      if(!kind||images[kind]||!entries[media])throw new Error('The thermal and ordinary photos could not be identified. Use SmartView’s original Basic Word export.');
      images[kind]={bytes:entries[media],type:imageType(entries[media]),captions:children(cell,'p').map(inline).filter(t=>t.trim()),media};
    }
    if(!images.main||!images.visible)throw new Error(`Image record ${index+1} is missing its thermal or ordinary photo. Enable both in SmartView’s export options.`);
    const ids=inline(pictureTable).match(/IR_\d{5}\.IS2/gi)||[];
    if(ids.length!==1||seen.has(ids[0].toUpperCase()))throw new Error('An image identifier is missing or repeated. Check the selected source report.');seen.add(ids[0].toUpperCase());
    const tables=block.filter(n=>n.localName==='tbl'&&n!==pictureTable&&hasText(n)).map(rowsOf);
    if(tables.length!==3)throw new Error(`Image record ${index+1} has extra or missing data tables. Use a Basic export with image details and both original marker tables.`);
    if(tables.some(t=>!t.length||t.some(r=>r.length!==t[0].length||r.length>8)))throw new Error('This export has merged or irregular data tables. Its original content needs a different layout.');
    const notes=block.filter(n=>n.localName==='p'&&hasText(n)).map(inline).filter(t=>!['Image Info','Main Image Markers'].includes(t.trim()));
    const record={id:ids[0],images,tables,notes};assertTextPreserved(block,record);records.push(record);
  }
  return records;
}
function visual(text,dir){if(!/[\u0590-\u08ff]/.test(text))return text;const levels=bidi.getEmbeddingLevels(text,dir||'rtl');return bidi.getReorderedString(text,levels);}
function rgb(hex){return globalThis.PDFLib.rgb(...hex.slice(1).match(/../g).map(x=>parseInt(x,16)/255));}
export function reportFilename({site,system,date}){const d=date.split('-').reverse().join('-');return `בדיקה תרמוגרפית - ${site} ${system} - ${d}.pdf`.replace(/[<>:"/\\|?*\u0000-\u001f]/g,'-');}
export async function generateReport(records,identity,onProgress=()=>{}){
  if(!records.length)throw new Error('Choose a source report first.');
  const dateParts=identity.date.split('-').map(Number), dateCheck=new Date(Date.UTC(dateParts[0],dateParts[1]-1,dateParts[2]));
  if(!identity.site.trim()||!identity.system.trim()||dateCheck.toISOString().slice(0,10)!==identity.date)throw new Error('Enter the site name, system code and a valid inspection date.');
  const {PDFDocument}=globalThis.PDFLib,pdf=await PDFDocument.create();pdf.registerFontkit(globalThis.fontkit);
  const asset=async path=>{const r=await fetch(new URL(path,import.meta.url));if(!r.ok)throw new Error('A required app asset could not load. Reload this page and try again.');return new Uint8Array(await r.arrayBuffer());};
  const [fontBytes,boldBytes,latinBytes,latinBoldBytes,logoBytes]=await Promise.all(['./assets/NotoSansHebrew-Regular.ttf','./assets/NotoSansHebrew-Bold.ttf','./assets/NotoSans-Regular.ttf','./assets/NotoSans-Bold.ttf','./assets/ormash-logo.png'].map(asset));
  const regular=await pdf.embedFont(fontBytes,{subset:true}),bold=await pdf.embedFont(boldBytes,{subset:true});
  const latin=await pdf.embedFont(latinBytes,{subset:true}),latinBold=await pdf.embedFont(latinBoldBytes,{subset:true});
  const logo=await pdf.embedPng(logoBytes),pages=[];const WIDTH=595.276,HEIGHT=841.89,LEFT=43,RIGHT=WIDTH-43,CONTENT=RIGHT-LEFT,BOTTOM=798;
  const fontSets=new Map([regular,bold,latin,latinBold].map(f=>[f,new Set(f.getCharacterSet())]));
  const control=/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/;
  function fontFor(s,isBold=false){const fonts=/[\u0590-\u05ff]/.test(s)?[isBold?bold:regular,isBold?latinBold:latin]:[isBold?latinBold:latin,isBold?bold:regular];const f=fonts.find(f=>fontSets.get(f).has(s.codePointAt(0)));if(!f)throw new Error('Some source text uses characters this formatter cannot render. Keep the original report and use the full formatting workflow for this document.');return f;}
  function measure(s,size,isBold=false){return Array.from(visual(s)).reduce((w,c)=>w+(control.test(c)?0:fontFor(c,isBold).widthOfTextAtSize(c,size)),0);}
  function wrap(s,width,size,isBold=false){const lines=[];for(const paragraph of String(s).replaceAll('\r','').split('\n')){if(!paragraph){lines.push('');continue;}let line='';for(const word of paragraph.split(/\s+/)){const proposed=line?line+' '+word:word;if(measure(proposed,size,isBold)<=width){line=proposed;continue;}if(line){lines.push(line);line='';}if(measure(word,size,isBold)<=width){line=word;continue;}for(const char of word){if(measure(line+char,size,isBold)>width&&line){lines.push(line);line='';}line+=char;}}lines.push(line);}return lines;}
  function drawText(page,s,x,y,{size=9,bold:heavy=false,color=ink,align='left',width=CONTENT,dir}={}){for(const [i,line]of String(s).split('\n').entries()){if(!line)continue;const text=visual(line,dir),w=measure(line,size,heavy);let pos=align==='right'?x+width-w:align==='center'?x+(width-w)/2:x;for(const c of text){if(control.test(c))continue;const font=fontFor(c,heavy);page.drawText(c,{x:pos,y:HEIGHT-y-size-i*size*1.35,size,font,color:rgb(color)});pos+=font.widthOfTextAtSize(c,size);}}}
  function line(page,y,color=rule,thickness=.5){page.drawLine({start:{x:LEFT,y:HEIGHT-y},end:{x:RIGHT,y:HEIGHT-y},thickness,color:rgb(color)});}
  function header(page){page.drawImage(logo,{x:LEFT,y:HEIGHT-79,width:85,height:85*logo.height/logo.width});drawText(page,'בדיקה תרמוגרפית',LEFT+100,28,{size:23,bold:true,align:'right',width:CONTENT-100});const titles=wrap(`${identity.site}  |  ${identity.system}`,CONTENT-100,10.5,true);if(titles.length>2)throw new Error('The site name is too long for the report header. Use the full site name without extra commentary.');titles.forEach((t,i)=>drawText(page,t,LEFT+100,61+i*14,{size:10.5,bold:true,align:'right',width:CONTENT-100}));drawText(page,'תאריך בדיקה: '+identity.date.split('-').reverse().join('-'),LEFT+100,63+titles.length*14,{size:9,color:muted,align:'right',width:CONTENT-100});const bottom=86+titles.length*14;line(page,bottom,accent,1.5);return bottom;}
  let page,y;
  function newPage(continued){page=pdf.addPage([WIDTH,HEIGHT]);pages.push(page);y=header(page)+13;if(continued){drawText(page,continued,LEFT,y,{size:10,bold:true});y+=22;}return page;}
  function textBlock(s,{size=9,bold:heavy=false,color=ink}={}){const lines=wrap(s,CONTENT,size,heavy);for(const text of lines){if(y+size*1.35>BOTTOM)newPage('Source details · continued');drawText(page,text,LEFT,y,{size,bold:heavy,color});y+=size*1.35;}y+=3;}
  function title(s,reserve=52){if(y+reserve>BOTTOM)newPage('Source details · continued');drawText(page,s,LEFT,y,{size:10,bold:true});y+=17;}
  function planTable(rows,widths,size=8){return rows.map((row,index)=>{const cells=row.map((s,i)=>wrap(s,widths[i]-12,size,index===0));return{cells,height:Math.max(...cells.map(a=>a.length))*size*1.3+6};});}
  function table(rows,{x=LEFT,top=y,width=CONTENT,widths,size=8,firstColumnLabel=false,allowPages=true}={}){
    widths=widths||Array(rows[0].length).fill(width/rows[0].length);const planned=planTable(rows,widths,size);let pos=top;
    const drawRow=(row,i,at)=>{let left=x;for(let c=0;c<widths.length;c++){page.drawRectangle({x:left,y:HEIGHT-at-row.height,width:widths[c],height:row.height,color:rgb(firstColumnLabel?(c===0?'#F5F7F8':'#FFFFFF'):(i===0?'#EAF0F3':'#FFFFFF')),borderColor:rgb(rule),borderWidth:.4});row.cells[c].forEach((t,k)=>drawText(page,t,left+6,at+3+k*size*1.3,{size,bold:!firstColumnLabel&&i===0,align:c===0?'left':'center',width:widths[c]-12}));left+=widths[c];}};
    for(let i=0;i<planned.length;i++){const row=planned[i];if(row.height>BOTTOM-160)throw new Error('A table cell is too long to fit safely. Review the original Word document.');if(allowPages&&pos+row.height>BOTTOM){newPage('Source details · continued');pos=y;if(i>0&&!firstColumnLabel){drawRow(planned[0],0,pos);pos+=planned[0].height;}}drawRow(row,i,pos);pos+=row.height;}if(allowPages)y=pos+11;return pos;
  }
  for(const [index,record]of records.entries()){
    onProgress(index+1,records.length);newPage();
    drawText(page,'צילום תרמי',LEFT,y,{size:10,bold:true,align:'right'});y+=21;
    const main=record.images.main,visible=record.images.visible;
    const embed=im=>im.type==='image/png'?pdf.embedPng(im.bytes):pdf.embedJpg(im.bytes);
    const mainImage=await embed(main),visibleImage=await embed(visible);
    const mainH=Math.min(358,CONTENT*mainImage.height/mainImage.width),mainW=mainH*mainImage.width/mainImage.height;
    page.drawImage(mainImage,{x:LEFT+(CONTENT-mainW)/2,y:HEIGHT-y-mainH,width:mainW,height:mainH});y+=mainH+7;
    for(const [i,caption]of main.captions.entries())textBlock(caption,{size:i===0?10:8.5,bold:i===0,color:i===0?ink:muted});
    y+=5;const groupTop=y,photoW=200,photoH=Math.min(150,photoW*visibleImage.height/visibleImage.width);
    drawText(page,'צילום רגיל',LEFT,groupTop,{size:10,bold:true,align:'right',width:photoW});drawText(page,'Image Info',LEFT+225,groupTop,{size:10,bold:true,width:CONTENT-225});
    const metaH=planTable(record.tables[0],[145,CONTENT-225-145],8).reduce((n,r)=>n+r.height,0);
    if(record.tables[0][0].length!==2)throw new Error('The source image-details table has an unsupported layout.');
    if(groupTop+20+Math.max(metaH,photoH+30)>BOTTOM){newPage(`${record.id} · source details`);}
    const actualTop=y;
    if(actualTop!==groupTop){drawText(page,'צילום רגיל',LEFT,actualTop,{size:10,bold:true,align:'right',width:photoW});drawText(page,'Image Info',LEFT+225,actualTop,{size:10,bold:true,width:CONTENT-225});}
    page.drawImage(visibleImage,{x:LEFT,y:HEIGHT-actualTop-20-photoH,width:photoW,height:photoH});let captionY=actualTop+20+photoH+4;
    for(const caption of visible.captions){for(const t of wrap(caption,photoW,7.5)){drawText(page,t,LEFT,captionY,{size:7.5,color:muted,width:photoW});captionY+=10;}}
    const metaEnd=table(record.tables[0],{x:LEFT+225,top:actualTop+20,width:CONTENT-225,widths:[145,CONTENT-225-145],size:8,firstColumnLabel:true,allowPages:false});
    if(Math.max(metaEnd,captionY)>BOTTOM)throw new Error('The source photo/details block is too long for this layout. Its content has been kept in your original document.');
    y=Math.max(metaEnd,captionY)+14;title('Main Image Markers');
    for(const [i,rows]of record.tables.slice(1).entries()){if(i)y+=2;table(rows,{size:7.5});}
    if(record.notes.length){y-=7;title('Source notes',32);for(const note of record.notes)textBlock(note);}
    if(index%4===0)await new Promise(resolve=>setTimeout(resolve,0));
  }
  for(const [index,p]of pages.entries()){line(p,811);drawText(p,`ORMASH  |  ${identity.system}  |  Page ${index+1} of ${pages.length}`,LEFT,820,{size:7.5,color:muted,align:'right'});}
  pdf.setTitle(`בדיקה תרמוגרפית - ${identity.site} ${identity.system}`);pdf.setAuthor('Ormash');pdf.setCreator('ormashify · browser formatter');pdf.setSubject('Formatted SmartView export; source camera readings preserved.');
  const output=await pdf.save();return{bytes:output,pageCount:pages.length,imageCount:records.length,filename:reportFilename(identity)};
}
