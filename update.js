'use strict';
const $=id=>document.getElementById(id);let generated=null;
function xml(text){let d=new DOMParser().parseFromString(text,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('Excel内のXMLを読み取れません。');return d;}
const tags=(node,name)=>Array.from(node.getElementsByTagNameNS('*',name));
const cellText=node=>tags(node,'t').filter(t=>t.parentElement?.localName!=='rPh').map(t=>t.textContent).join('');
async function readWorkbook(file){
 if(file.size>20000000)throw Error('Excelは20MB以下にしてください。');
 const zip=await JSZip.loadAsync(file);let total=0;
 for(let entry of Object.values(zip.files)){if(entry._data?.uncompressedSize)total+=entry._data.uncompressedSize;}if(total>80000000)throw Error('展開後のデータが大きすぎます。');
 async function read(path,optional=false){let f=zip.file(path);if(!f){if(optional)return null;throw Error('必要なExcelデータがありません：'+path);}return xml(await f.async('string'));}
 const shared=await read('xl/sharedStrings.xml',true),strings=shared?tags(shared,'si').map(cellText):[];
 const workbook=await read('xl/workbook.xml'),rels=await read('xl/_rels/workbook.xml.rels'),links=Object.fromEntries(tags(rels,'Relationship').map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));let sheets={};
 for(let sheet of tags(workbook,'sheet')){let name=sheet.getAttribute('name');if(!['ポケモン','出現'].includes(name))continue;let id=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),target=links[id];if(!target)throw Error('シートの参照が不正です。');let path=target.startsWith('/')?target.slice(1):'xl/'+target;let doc=await read(path),rows=[];
  for(let row of tags(doc,'row')){let cells=[];for(let c of tags(row,'c')){let ref=c.getAttribute('r'),letters=ref.match(/^[A-Z]+/)[0],col=0;for(let l of letters)col=col*26+l.charCodeAt(0)-64;let type=c.getAttribute('t'),v=tags(c,'v')[0]?.textContent??'';cells[col-1]=type==='s'?strings[Number(v)]:type==='inlineStr'?cellText(c):v;}rows[Number(row.getAttribute('r'))-1]=cells;}sheets[name]=Array.from(rows,r=>r||[]);
 }
 return SleepCore.buildData(sheets);
}
$('excel').onchange=async e=>{generated=null;$('download').disabled=true;$('preview').replaceChildren();let file=e.target.files[0];if(!file)return;$('message').textContent='Excelを確認しています…';try{let data=await readWorkbook(file);generated='window.SLEEP_DATA = '+JSON.stringify(data,null,2)+';\n';$('message').textContent='確認できました。対象 '+data.pokemon.length+'種類 / 食材 '+data.ingredients.length+'種類 / 島 '+data.fields.length+'件';let p=document.createElement('p');p.className='hint';p.textContent='対象：食材得意＋ヤドン。進化系統の整理は島ごとに行います。更新日：'+data.updated;$('preview').append(p);for(let warning of data.warnings){let p=document.createElement('p');p.textContent=warning;$('preview').append(p);}$('download').disabled=false;}catch(err){$('message').textContent='作成できませんでした：'+err.message;}};
$('download').onclick=()=>{if(!generated)return;let url=URL.createObjectURL(new Blob([generated],{type:'text/javascript;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='data.js';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
