const MON=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const DATA=window.DATA; // diisi assets/data.js (dibuat app.py dari Excel)
if(window.Chart&&Chart.defaults){Chart.defaults.font.family='Inter, system-ui, sans-serif';Chart.defaults.font.size=10
}let rows=DATA.rows,charts={},base25=DATA.base;
const $=id=>document.getElementById(id);
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const fmt=n=>Math.round(n).toLocaleString('id-ID');

function val(rs,act,met){
  const s=act?rs.filter(r=>r.act===act):rs;
  return met==='trx'?new Set(s.map(r=>r.trx)).size:s.reduce((a,r)=>a+r[met],0);
}
function by(rs,f){const m=new Map();rs.forEach(r=>{const k=f(r);(m.get(k)||m.set(k,[]).get(k)).push(r)});return m}
function yoy(a,b){if(!b)return '<span class="y">—</span>';const p=(a-b)/b*100;return `<span class="y ${p>=0?'up':'down'}">${p>=0?'▲':'▼'} ${Math.abs(p).toFixed(1)}%</span>`}

function fill(sel,items,all){sel.innerHTML=(all?`<option value="">${all}</option>`:'')+items.map(x=>`<option>${x}</option>`).join('')}
function setup(){
  const ys=[...new Set(rows.map(r=>r.y))].sort();
  fill($('year'),ys);$('year').value=ys[ys.length-1];
  fill($('wh'),[...new Set(rows.map(r=>r.wh))].sort(),'Semua warehouse');
  fill($('cust'),[...new Set(rows.map(r=>r.cust))].sort(),'Semua customer');
  if(base25)$('wh').value='LSB';
  render();
}

function render(){
  const met=$('met').value,y=+$('year').value,wh=$('wh').value,cu=$('cust').value,unit={trx:'transaksi',vol:'CBM',qty:'qty'}[met];
  const base=rows.filter(r=>(!wh||r.wh===wh)&&(!cu||r.cust===cu));
  const cur=base.filter(r=>r.y===y),prv=base.filter(r=>r.y===y-1);
  const last=cur.length?Math.max(...cur.map(r=>r.m)):0; 
  const prvYTD=prv.filter(r=>r.m<=last);
const bmap=base25&&(typeof base25.wh==='string'?{[base25.wh]:base25}:base25.wh);
const bw=base25&&base25.y===y-1&&wh&&bmap&&bmap[wh];
const useBase=!prv.length&&!!bw&&!cu&&met!=='qty';
const newWh=!!wh&&!prv.length&&!bw;
const bArr=useBase?bw[met].map(v=>v||0):null;  const bYTD=useBase?bArr.slice(0,last+1).reduce((a,v)=>a+v,0):0;
  const I=val(cur,'IN',met),O=val(cur,'OUT',met),T=val(cur,null,met),pI=val(prvYTD,'IN',met),pO=val(prvYTD,'OUT',met),pT=useBase?bYTD:val(prvYTD,null,met);
  const nc=new Set(cur.map(r=>r.cust)).size,pnc=new Set(prvYTD.map(r=>r.cust)).size;
  $('kpis').innerHTML=[
   ['in','Inbound',I,yoy(I,pI)],['out','Outbound',O,yoy(O,pO)],['tot','Total',T,yoy(T,pT)],['cus','Customer aktif',nc,yoy(nc,pnc)]
  ].map(([c,l,v,g])=>`<div class="card k ${c}"><div class="l">${l} ${c==='cus'?'':'('+unit+')'}</div><div class="v">${fmt(v)}</div>${g.includes('>—<')?'':g+' <span class="sub">vs '+(y-1)+' (Jan–'+MON[last]+')</span>'}</div>`).join('');

  const mi=MON.map((_,m)=>val(cur.filter(r=>r.m===m),'IN',met)),mo=MON.map((_,m)=>val(cur.filter(r=>r.m===m),'OUT',met));
  const pm=useBase?bArr:MON.map((_,m)=>val(prv.filter(r=>r.m===m),null,met)),tm=MON.map((_,m)=>val(cur.filter(r=>r.m===m),null,met));
  const tick=css('--mut'),grid=css('--line');
  const opt=(extra={})=>Object.assign({responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',align:'end',labels:{color:tick,usePointStyle:true,boxWidth:8,boxHeight:8}},
      tooltip:{callbacks:{label:c=>' '+c.dataset.label+': '+fmt(c.chart.options.indexAxis==='y'?c.parsed.x:c.parsed.y)}}},
    scales:{x:{ticks:{color:tick},grid:{display:false}},y:{ticks:{color:tick,callback:v=>v.toLocaleString('id-ID')},grid:{color:grid},beginAtZero:true}}},extra);
  const draw=(id,cfg)=>{charts[id]&&charts[id].destroy();charts[id]=new Chart($(id),cfg)};
  draw('c1',{type:'bar',data:{labels:MON,datasets:[{label:'Inbound',data:mi,backgroundColor:css('--in'),borderRadius:4,maxBarThickness:22},{label:'Outbound',data:mo,backgroundColor:css('--out'),borderRadius:4,maxBarThickness:22}]},options:opt()});
  draw('c2',{type:'line',data:{labels:MON,datasets:[{label:String(y),data:tm.map((v,i)=>cur.some(r=>r.m===i)?v:null),borderColor:css('--in'),backgroundColor:css('--in')+'22',fill:true,borderWidth:3,pointRadius:3.5,pointBackgroundColor:css('--in'),tension:.25},
    {label:String(y-1),data:pm,borderColor:css('--prev'),backgroundColor:css('--prev'),borderDash:[6,5],borderWidth:2,pointRadius:0,tension:.25}]},options:opt()});

  $('t1').innerHTML='<tr><th>Bulan</th><th>Inbound</th><th>Outbound</th><th>Total</th><th>'+(y-1)+'</th><th>YoY</th></tr>'+
    MON.map((n,i)=>cur.some(r=>r.m===i)?`<tr><td>${n}</td><td>${fmt(mi[i])}</td><td>${fmt(mo[i])}</td><td>${fmt(tm[i])}</td><td>${fmt(pm[i])}</td><td>${yoy(tm[i],pm[i])}</td></tr>`:'').join('');

  const rowsC=[...by(cur,r=>r.cust)].map(([k,rs])=>{const p=prvYTD.filter(r=>r.cust===k);const i=val(rs,'IN',met),o=val(rs,'OUT',met);return{k,i,o,t:i+o,p:val(p,null,met)}}).sort((a,b)=>b.t-a.t);
  const top=rowsC.slice(0,10);
  draw('c3',{type:'bar',data:{labels:top.map(x=>{const n=x.k.replace(/,\s*PT.*$/i,'');return n.length>26?n.slice(0,25)+'…':n}),datasets:[{label:'Inbound',data:top.map(x=>x.i),backgroundColor:css('--in')},{label:'Outbound',data:top.map(x=>x.o),backgroundColor:css('--out')}]},
    options:opt({indexAxis:'y',scales:{x:{stacked:true,ticks:{color:tick},grid:{color:grid}},y:{stacked:true,ticks:{color:tick},grid:{display:false}}}})});
  const tot=rowsC.reduce((a,x)=>a+x.t,0)||1;
  const anyP=rowsC.some(x=>x.p>0);
  $('t2').innerHTML='<tr><th>Customer</th><th>Inbound</th><th>Outbound</th><th>Total</th><th>Kontribusi</th>'+(anyP?'<th>YoY (YTD)</th>':'')+'</tr>'+
    rowsC.map(x=>`<tr><td>${x.k}</td><td>${fmt(x.i)}</td><td>${fmt(x.o)}</td><td>${fmt(x.t)}</td><td>${(x.t/tot*100).toFixed(1)}%</td>${anyP?'<td>'+yoy(x.t,x.p)+'</td>':''}</tr>`).join('');
  const noPrev=!prv.length&&!useBase?' YoY tidak tersedia untuk filter ini: data pembanding 2025 hanya ada untuk Warehouse LSB (total, tanpa pemisahan inbound/outbound dan customer).':'';
}

/* ================= Produktivitas per warehouse (sheet Produktivitas 2026) ================= */
const PROD=DATA.prod;
const nums=(arr,idx)=>idx.map(i=>arr[i]).filter(x=>typeof x==='number'&&x>0);
const mean=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;
function pstat(b){
  const idx=b.prod.map((v,i)=>typeof v==='number'?i:-1).filter(i=>i>=0);   // bulan yang sudah ada datanya
  const a26=mean(nums(b.prod,idx)),a25=mean(nums(b.prev,idx));             // pembanding: bulan yang sama tahun lalu
  const total=idx.reduce((s,i)=>s+(b.val[i]||0),0);
  const mp=idx.length?b.mp[idx[idx.length-1]]:null;
  return {idx,a26,a25,total,mp,last:idx.length?idx[idx.length-1]:0};
}
const f1=n=>n==null?'—':n.toLocaleString('id-ID',{maximumFractionDigits:n>=100?0:1});
function prodChart(id,b,label){
  const tick=css('--mut'),grid=css('--line');
  charts[id]&&charts[id].destroy();
  charts[id]=new Chart($(id),{type:'bar',data:{labels:MON,datasets:[
    {type:'bar',label:String(PROD.year),data:b.prod,backgroundColor:css('--in'),borderRadius:4,maxBarThickness:24,order:2},
    {type:'line',label:String(PROD.prev_year),data:b.prev,borderColor:css('--prev'),backgroundColor:css('--prev'),borderDash:[6,5],borderWidth:2,pointRadius:0,tension:.25,spanGaps:true,order:1}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',align:'end',labels:{color:tick,usePointStyle:true,boxWidth:8,boxHeight:8}},
      tooltip:{callbacks:{label:c=>' '+c.dataset.label+': '+f1(c.parsed.y)+' '+label}}},
      scales:{x:{ticks:{color:tick},grid:{display:false}},y:{beginAtZero:true,ticks:{color:tick,callback:v=>v.toLocaleString('id-ID')},grid:{color:grid}}}}});
}
function renderProd(){
  if(!PROD||!PROD.wh.length)return;
  $('prod').hidden=false;
  const w=PROD.wh.find(x=>x.name===$('pwh').value)||PROD.wh[0];
  const c=pstat(w.cbm),t=pstat(w.trx),lastM=MON[Math.max(c.last,t.last)];
  const tile=(l,v,g,s)=>`<div><div class="l">${l}</div><div class="v">${v}</div>${g.includes('>—<')?'<span class="sub">tidak ada data '+PROD.prev_year+'</span>':g+(s?' <span class="sub">'+s+'</span>':'')}</div>`;
  $('pk').innerHTML=
    tile('Throughput Jan–'+MON[c.last]+' (CBM)',f1(c.total),'','')+
    tile('Manpower (CBM)',f1(c.mp),'','orang')+
    tile('Produktivitas CBM/MP',f1(c.a26),yoy(c.a26,c.a25),'vs '+PROD.prev_year)+
    tile('Produktivitas transaksi/MP',f1(t.a26),yoy(t.a26,t.a25),'vs '+PROD.prev_year);
  prodChart('p1',w.cbm,'CBM/MP');prodChart('p2',w.trx,'transaksi/MP');
  $('pt').innerHTML='<tr><th>Warehouse</th><th>Throughput (CBM)</th><th>MP</th><th>CBM/MP '+PROD.year+'</th><th>CBM/MP '+PROD.prev_year+'</th><th>Selisih</th><th>Transaksi</th><th>MP</th><th>Transaksi/MP '+PROD.year+'</th><th>Transaksi/MP '+PROD.prev_year+'</th><th>Selisih</th></tr>'+
    PROD.wh.map(x=>{const a=pstat(x.cbm),b=pstat(x.trx),tt=x.trx.val.reduce((s,v)=>s+(v||0),0);
      return `<tr><td>${x.name}</td><td>${f1(a.total)}</td><td>${f1(a.mp)}</td><td>${f1(a.a26)}</td><td>${f1(a.a25)}</td><td>${yoy(a.a26,a.a25)}</td><td>${f1(tt)}</td><td>${f1(b.mp)}</td><td>${f1(b.a26)}</td><td>${f1(b.a25)}</td><td>${yoy(b.a26,b.a25)}</td></tr>`}).join('');
}
if(PROD&&PROD.wh.length){fill($('pwh'),PROD.wh.map(x=>x.name));$('pwh').onchange=renderProd;renderProd()}
['year','wh','cust','met'].forEach(id=>$(id).onchange=render);
setup();
