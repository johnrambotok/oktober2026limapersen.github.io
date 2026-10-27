const $=id=>document.getElementById(id);
let rows=[],pageNo=0,sortKey='code',sortDir=1,chart1=null,chart2=null;
const t=v=>v==null?'':String(v).trim();
const n=(v,blankNull=false)=>{if(v==null||t(v)===''||t(v)==='-')return blankNull?null:0;if(typeof v==='number')return v;const x=Number(String(v).replace(/,/g,'').replace(/%/g,'').trim());return Number.isFinite(x)?x:(blankNull?null:0)};
const f=v=>v==null||v===''?'—':Number(v).toLocaleString('id-ID',{maximumFractionDigits:2});
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function parse(wb,name){
 const a=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:null,raw:false});
 let h=-1;for(let i=0;i<Math.min(12,a.length);i++)if((a[i]||[]).some(v=>t(v)==='Kode Efek')){h=i;break}
 if(h<0)throw Error('Kolom Kode Efek tidak ditemukan: '+name);
 const m=t((a[0]||[])[0]).match(/tanggal\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})/i),date=m?m[1]:name;
 let out=[],lastCode='',lastIssuer='',lastInvestor='';
 for(let i=h+2;i<a.length;i++){
  const r=a[i]||[];let code=t(r[1]);
  const hasRowData=[r[3],r[4],r[5],r[6],r[9],r[10]].some(v=>t(v)!=='');
  if(code)lastCode=code;else if(hasRowData)code=lastCode;
  if(!code||code==='Kode Efek'||/^catatan/i.test(code))continue;
  let issuer=t(r[2]);if(issuer)lastIssuer=issuer;else issuer=lastIssuer;
  let investor=t(r[4]);if(investor)lastInvestor=investor;else investor=lastInvestor;
  if(!investor)continue;
  out.push({code,issuer,account:t(r[3]),investor,accountName:t(r[5]),address:t(r[6]),address2:t(r[7]),nationality:t(r[10]).toUpperCase(),dom:t(r[9]),
   prevShares:n(r[11]),prevCombined:n(r[12],true),prevPct:n(r[13],true),shares:n(r[14]),combined:n(r[15],true),pct:n(r[16],true),change:t(r[17]),date});
 }
 return {data:out,date};
}
async function read(file){return parse(XLSX.read(await file.arrayBuffer(),{type:'array'}),file.name)}
function key(r){return [r.code,r.investor,r.account,r.accountName].map(v=>t(v).toUpperCase()).join('|')}
async function load(){
 const of=$('older').files[0],nf=$('newer').files[0];if(!of&&!nf){alert('Pilih minimal satu file Excel.');return}
 $('status').textContent='Sedang membaca dan membandingkan file…';$('load').disabled=true;
 try{
  const old=of?await read(of):null,latest=nf?await read(nf):old;
  const oldMap=new Map();if(old)for(const r of old.data)oldMap.set(key(r),r);
  const latestKeys=new Set(latest.data.map(key));
  rows=latest.data.map(r=>{
   const p=oldMap.get(key(r));
   const prevPct=old&&nf?(p?p.pct:null):r.prevPct;
   const delta=(r.pct!=null&&prevPct!=null)?r.pct-prevPct:null;
   const priorShares=old&&nf?(p?p.shares:null):r.prevShares;
   const shareDelta=(r.shares!=null&&priorShares!=null)?r.shares-priorShares:null;
   const changed=old&&nf?(!!p&&(delta!=null&&Math.abs(delta)>.0001||shareDelta!=null&&shareDelta!==0)):((r.change!==''&&r.change!=='-'&&r.change!=='0')||(delta!=null&&Math.abs(delta)>.0001)||(shareDelta!=null&&shareDelta!==0));
   return {...r,prevPct,delta,priorShares,shareDelta,changed,foundInLatest:true,matchFound:!!p};
  });
  if(old&&nf){for(const p of old.data){if(!latestKeys.has(key(p)))rows.push({...p,pct:null,shares:null,combined:null,prevPct:p.pct,delta:null,priorShares:p.shares,shareDelta:null,change:'Tidak ditemukan pada laporan terbaru',changed:false,foundInLatest:false,matchFound:true})}}
  $('status').textContent=old&&nf?'Data siap · '+old.date+' → '+latest.date+' · '+rows.length.toLocaleString('id-ID')+' baris':'Data siap · '+latest.date+' · '+rows.length.toLocaleString('id-ID')+' baris';
  $('olderName').textContent=of?of.name:'Belum dipilih';$('newerName').textContent=nf?nf.name:'Belum dipilih';
  fillFilters();kpis();charts();pageNo=0;render();
 }catch(e){console.error(e);$('status').textContent='Gagal membaca file: '+e.message;alert('Gagal membaca Excel: '+e.message)}finally{$('load').disabled=false}
}
function fillFilters(){const ss=[...new Set(rows.map(r=>r.code))].sort(),ds=[...new Set(rows.map(r=>r.dom).filter(Boolean))].sort();$('stock').innerHTML='<option value="">Semua kode</option>'+ss.map(v=>'<option>'+esc(v)+'</option>').join('');$('dom').innerHTML='<option value="">Semua domisili</option>'+ds.map(v=>'<option>'+esc(v)+'</option>').join('')}
function filtered(){
 const q=$('q').value.toLowerCase(),s=$('stock').value,na=$('nationality').value,d=$('dom').value,c=$('change').value;
 return rows.filter(r=>(!q||[r.code,r.issuer,r.investor,r.account,r.accountName].join(' ').toLowerCase().includes(q))&&(!s||r.code===s)&&(!na||r.nationality===na)&&(!d||r.dom===d)&&(!c||(c==='yes'&&r.changed)||(c==='no'&&!r.changed)||(c==='up'&&r.delta!=null&&r.delta>0)||(c==='down'&&r.delta!=null&&r.delta<0))).sort((a,b)=>{const x=a[sortKey]??'',y=b[sortKey]??'';return(typeof x==='number'&&typeof y==='number'?x-y:String(x).localeCompare(String(y),'id'))*sortDir});
}
function kpis(){
 $('rows').textContent=f(rows.filter(r=>r.foundInLatest).length);
 $('stocks').textContent=f(new Set(rows.filter(r=>r.foundInLatest).map(r=>r.code)).size);
 $('investors').textContent=f(new Set(rows.filter(r=>r.foundInLatest).map(r=>r.investor.toUpperCase())).size);
 $('changed').textContent=f(rows.filter(r=>r.foundInLatest&&r.changed).length);
}
function charts(){
 const current=rows.filter(r=>r.foundInLatest),im={};current.forEach(r=>{const k=r.issuer||r.code;im[k]=(im[k]||0)+1});
 const top=Object.entries(im).sort((a,b)=>b[1]-a[1]).slice(0,10),s={'Lokal (L)':0,'Asing (A)':0,'Lain/tidak diketahui':0};
 current.forEach(r=>{if(r.nationality==='L')s['Lokal (L)']++;else if(r.nationality==='A')s['Asing (A)']++;else s['Lain/tidak diketahui']++});
 if(chart1)chart1.destroy();chart1=new Chart($('issuers'),{type:'bar',data:{labels:top.map(x=>x[0]),datasets:[{label:'Baris kepemilikan',data:top.map(x=>x[1]),backgroundColor:'#7dd3fc',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#9aabc4'},grid:{color:'#2b3a52'}},y:{ticks:{color:'#cbd5e1'},grid:{display:false}}}}});
 if(chart2)chart2.destroy();chart2=new Chart($('statuses'),{type:'doughnut',data:{labels:Object.keys(s),datasets:[{data:Object.values(s),backgroundColor:['#7dd3fc','#a7f3d0','#64748b'],borderColor:'#121d30',borderWidth:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{color:'#cbd5e1'}}}}});
}
function render(){
 const a=filtered(),pages=Math.max(1,Math.ceil(a.length/50));pageNo=Math.min(pageNo,pages-1);
 $('body').innerHTML=a.slice(pageNo*50,(pageNo+1)*50).map(r=>'<tr><td>'+esc(r.code)+'</td><td>'+esc(r.issuer)+'</td><td>'+esc(r.investor)+'</td><td>'+esc(r.account)+'</td><td>'+esc(r.nationality||'—')+'</td><td>'+esc(r.dom||'—')+'</td><td>'+f(r.shares)+'</td><td>'+f(r.combined)+'</td><td>'+f(r.pct)+'%</td><td>'+f(r.prevPct)+'%</td><td class="'+(r.delta>0?'up':r.delta<0?'down':'')+'">'+(r.delta>0?'+':'')+f(r.delta)+'</td><td>'+(r.foundInLatest?(r.changed?'Berubah':'Tidak berubah'):'Tidak ditemukan di laporan terbaru')+'</td></tr>').join('')||'<tr><td colspan="12">Tidak ada baris yang cocok.</td></tr>';
 $('count').textContent=f(a.length)+' baris cocok · '+f(rows.length)+' total termasuk baris yang tidak ditemukan';$('page').textContent='Halaman '+(pageNo+1)+' / '+pages;$('prev').disabled=pageNo===0;$('next').disabled=pageNo>=pages-1;
}
function exportCsv(){
 const a=filtered(),cols=[['Kode',r=>r.code],['Emiten',r=>r.issuer],['Pemegang Saham',r=>r.investor],['Pemegang Rekening Efek',r=>r.account],['Nama Rekening Efek',r=>r.accountName],['Status',r=>r.nationality],['Domisili',r=>r.dom],['Jumlah Saham Terbaru',r=>r.shares],['Gabungan Investor',r=>r.combined],['Kepemilikan Terbaru %',r=>r.pct],['Sebelumnya %',r=>r.prevPct],['Selisih p.p.',r=>r.delta],['Status Laporan',r=>r.foundInLatest?'Ditemukan':'Tidak ditemukan'],['Perubahan',r=>r.change]];
 const csv=[cols.map(c=>'"'+c[0]+'"').join(';'),...a.map(r=>cols.map(c=>'"'+String(c[1](r)??'').replace(/"/g,'""')+'"').join(';'))].join('\r\n');
 const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download='idx-ownership-filtered.csv';link.click();URL.revokeObjectURL(url);
}
$('load').addEventListener('click',load);
['q','stock','nationality','dom','change'].forEach(id=>$(id).addEventListener(id==='q'?'input':'change',()=>{pageNo=0;render()}));
$('prev').onclick=()=>{pageNo=Math.max(0,pageNo-1);render()};$('next').onclick=()=>{pageNo++;render()};$('export').onclick=exportCsv;
document.querySelectorAll('th[data-k]').forEach(th=>th.addEventListener('click',()=>{const k=th.dataset.k;if(k===sortKey)sortDir*=-1;else{sortKey=k;sortDir=1}render()}));
