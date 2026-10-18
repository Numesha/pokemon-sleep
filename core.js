(function(root){
'use strict';
const PATTERNS=['AAA','AAB','AAC','ABA','ABB','ABC'];
const LABELS={priority:'厳選最優先',future:'将来厳選',temp:'仮採用',done:'厳選終了'};
const norm=s=>String(s||'').normalize('NFKC').toLowerCase().replace(/[ぁ-ゖ]/g,c=>String.fromCharCode(c.charCodeAt(0)+96)).replace(/\s/g,'');
function buildData(sheets){
 const source=sheets['ポケモン'],spawns=sheets['出現'];
 if(!source||!spawns)throw Error('「ポケモン」「出現」シートが必要です。');
 const head=source[0], ix=n=>{let i=head.indexOf(n);if(i<0)throw Error('必要な列がありません：'+n);return i;};
 const col=Object.fromEntries(['内部ID','名前','進化前ID','バリエーションID','得意ID','睡眠タイプID','食材A','食材B','食材C'].map(n=>[n,ix(n)]));
 const seen=new Set(),all=source.slice(1).filter(r=>r[col['内部ID']]!==null&&r[col['内部ID']]!==undefined&&r[col['内部ID']]!=='').map(r=>{
  const id=String(r[col['内部ID']]);if(seen.has(id))throw Error('内部IDが重複：'+id);seen.add(id);
  return {id,name:String(r[col['名前']]||''),parent:String(r[col['進化前ID']]||'NONE'),variant:String(r[col['バリエーションID']]||'通常'),specialty:r[col['得意ID']],sleep:r[col['睡眠タイプID']],foods:['食材A','食材B','食材C'].map(k=>String(r[col[k]]||'なし'))};
 });
 const warnings=[];
 // Supplied workbook has 149 -> 178 (ネイティオ). Correct only this known entry in generated data.
 const dragon=all.find(p=>p.id==='149');if(dragon?.parent==='178'&&all.some(p=>p.id==='148'&&p.name==='ハクリュー')){dragon.parent='148';warnings.push('カイリューの進化前ID：178 → 148（ハクリュー）として変換しました。元Excelは変更していません。');}
 const lookup=new Map(all.map(p=>[p.id,p]));
 const sh=spawns[0],si=sh.indexOf('内部ID'),sf=sh.indexOf('島名');if(si<0||sf<0)throw Error('出現シートに内部ID・島名が必要です。');
 const fields=[];for(const r of spawns.slice(1))if(r[si]!=null&&r[sf]){let p=lookup.get(String(r[si]));if(!p)throw Error('出現シートの不明な内部ID：'+r[si]);(p.fields??=[]).push(String(r[sf]));if(!fields.includes(String(r[sf])))fields.push(String(r[sf]));}
 const pokemon=all.filter(p=>p.specialty==='食材'||p.name==='ヤドン').map(p=>{
  if(!p.name||!['うとうと','すやすや','ぐっすり'].includes(p.sleep)||p.foods.slice(0,2).includes('なし'))throw Error('必須データを確認してください：'+p.id);
  if(!p.fields?.length)throw Error('出現島がありません：'+p.name);
  let ancestors=[],cursor=p,visited=new Set([p.id]);while(cursor.parent!=='NONE'){if(visited.has(cursor.parent))throw Error('進化関係が循環：'+p.name);visited.add(cursor.parent);cursor=lookup.get(cursor.parent);if(!cursor)throw Error('進化前IDが見つかりません：'+p.name);ancestors.push(cursor.id);}
  return {...p,fields:[...new Set(p.fields)],ancestors,display:p.name+(p.variant==='通常'?'':'（'+p.variant+'）')};
 });
 return {version:1,updated:new Date().toISOString().slice(0,10),pokemon,fields,warnings,ingredients:[...new Set(pokemon.flatMap(p=>p.foods))].filter(f=>f!=='なし')};
}
function compositions(p,patterns){return PATTERNS.filter(k=>patterns.includes(k)&&(!k.includes('C')||p.foods[2]!=='なし')).map(pattern=>({pattern,foods:[...pattern].map(k=>p.foods[k.charCodeAt(0)-65])}));}
function representatives(data,field){const present=data.pokemon.filter(p=>p.fields.includes(field));const ids=new Set(present.map(p=>p.id));return present.filter(p=>!p.ancestors.some(id=>ids.has(id)));}
function matches(p,filters,status){
 if(filters.sleep&&p.sleep!==filters.sleep)return [];
 if(filters.q&&!norm(p.display+' '+p.fields.join(' ')).includes(norm(filters.q)))return [];
 return compositions(p,filters.patterns).filter(c=>c.foods.some(f=>(!filters.foods.length||filters.foods.includes(f))&&(!filters.status||status[f]===filters.status)));
}
function search(data,filters,status,excluded){const allowed=new Set(data.fields.filter(f=>!filters.field||f===filters.field).flatMap(f=>representatives(data,f).map(p=>p.id)));return data.pokemon.filter(p=>allowed.has(p.id)&&!excluded.includes(p.id)).map(p=>({...p,compositions:matches(p,filters,status),shownFields:p.fields.filter(f=>(!filters.field||filters.field===f)&&representatives(data,f).some(x=>x.id===p.id))})).filter(p=>p.compositions.length);}
function islands(data,filters,status,excluded){return data.fields.filter(f=>!filters.field||f===filters.field).map(field=>{
 const pokemon=representatives(data,field).filter(p=>!excluded.includes(p.id)).map(p=>({...p,compositions:matches(p,filters,status)})).filter(p=>p.compositions.length);
 const count=list=>Object.fromEntries(['priority','future','temp','done'].map(s=>[s,list.filter(p=>p.compositions.some(c=>c.foods.some(f=>(!filters.foods.length||filters.foods.includes(f))&&(!filters.status||s===filters.status)&&status[f]===s))).length]));
 return {field,pokemon,counts:count(pokemon),details:['うとうと','すやすや','ぐっすり'].map(sleep=>{let list=pokemon.filter(p=>p.sleep===sleep);return {sleep,pokemon:list,counts:count(list)};})};
 }).sort((a,b)=>b.counts.priority-a.counts.priority||b.counts.future-a.counts.future||b.counts.temp-a.counts.temp||data.fields.indexOf(a.field)-data.fields.indexOf(b.field));}
function islandRanking(data,filters,status,excluded,detailed=false){
 const totals=islands(data,filters,status,excluded),sleeps=['うとうと','すやすや','ぐっすり'],referenceFields=['ワカクサ本島','ワカクサ本島EX'];
 const rows=detailed?totals.flatMap(r=>r.details.filter(d=>!filters.sleep||d.sleep===filters.sleep).map(d=>({field:r.field,...d}))):totals;
 const score=(a,b)=>b.counts.priority-a.counts.priority||b.counts.future-a.counts.future||b.counts.temp-a.counts.temp;
 const ranked=rows.filter(r=>!referenceFields.includes(r.field)).sort((a,b)=>score(a,b)||data.fields.indexOf(a.field)-data.fields.indexOf(b.field)||sleeps.indexOf(a.sleep)-sleeps.indexOf(b.sleep));
 const reference=rows.filter(r=>referenceFields.includes(r.field)).sort((a,b)=>referenceFields.indexOf(a.field)-referenceFields.indexOf(b.field)||sleeps.indexOf(a.sleep)-sleeps.indexOf(b.sleep));
 return {ranked,reference};
}
const api={PATTERNS,LABELS,norm,buildData,compositions,representatives,matches,search,islands,islandRanking};if(typeof module!=='undefined')module.exports=api;else root.SleepCore=api;
})(typeof window==='undefined'?globalThis:window);
