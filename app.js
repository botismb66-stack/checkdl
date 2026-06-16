let rawRows=[], groups=[], gameStats={}, agentTotals={};
const COL={account:0, regTime:1, fullName:2, agent:3, level:4, deposit:5, ip:6, lastDeposit:7, bank:8, branch:9};
const GAME_COL={agent:2, account:3, hall:4, validBet:7};
function val(r,i){return (r[i]??'').toString().trim()}
function num(v){if(v===null||v===undefined)return 0; return Number(String(v).replace(/,/g,'').replace(/[^0-9.-]/g,''))||0}
function parseDate(v){
  if(v instanceof Date) return v;
  if(typeof v==='number'){return new Date(Math.round((v-25569)*86400*1000));}
  let s=String(v||'').trim(); if(!s)return null;
  s=s.replace(/\./g,'/').replace(/-/g,'/');
  let m=s.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if(m)return new Date(+m[1],+m[2]-1,+m[3],+(m[4]||0),+(m[5]||0),+(m[6]||0));
  return new Date(s);
}
function fmtDate(v){let d=parseDate(v); if(!d||isNaN(d))return ''; return d.getFullYear()+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+String(d.getDate()).padStart(2,'0')+' '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');}
function fmtNum(n){return (Number(n)||0).toLocaleString('en-US',{maximumFractionDigits:2})}
function fmtPct(n){return (Number(n)||0).toFixed(2)+'%'}
function ip24(ip){let p=String(ip||'').trim().split('.'); return p.length>=3?p.slice(0,3).join('.')+'.xxx':''}
function normName(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');}
function keyAgentAccount(agent,account){return normName(agent)+'|'+normName(account)}
function addGroup(map,key,row,type,value){if(!key)return; if(!map[key])map[key]={type,value,rows:[]}; map[key].rows.push(row)}
function readWorkbook(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{resolve(XLSX.read(e.target.result,{type:'array',cellDates:true}));}catch(err){reject(err)}
    };
    reader.onerror=reject;
    reader.readAsArrayBuffer(file);
  });
}
async function handleFile(){
  const f=document.getElementById('fileInput').files[0];
  if(!f){alert('Vui lòng chọn file đại lý');return;}
  try{
    const wb=await readWorkbook(f);
    const ws=wb.Sheets[wb.SheetNames[0]];
    const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    rawRows=data.slice(1).map(r=>({
      account:val(r,COL.account), regTime:val(r,COL.regTime), fullName:val(r,COL.fullName), agent:val(r,COL.agent), level:val(r,COL.level),
      deposit:num(r[COL.deposit]), ip:val(r,COL.ip), lastDeposit:val(r,COL.lastDeposit), bank:val(r,COL.bank), branch:val(r,COL.branch)
    })).filter(x=>x.account && x.agent && x.deposit>0);

    gameStats={};
    const gf=document.getElementById('gameFileInput').files[0];
    if(gf){ await loadGameFile(gf); }
    analyze(); render();
  }catch(e){console.error(e); alert('Không đọc được file. Vui lòng kiểm tra lại định dạng Excel/CSV.');}
}
async function loadGameFile(file){
  const wb=await readWorkbook(file);
  const ws=wb.Sheets[wb.SheetNames[0]];
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  data.slice(1).forEach(r=>{
    const agent=val(r,GAME_COL.agent), account=val(r,GAME_COL.account), hall=val(r,GAME_COL.hall), bet=num(r[GAME_COL.validBet]);
    if(!agent || !account || !hall || bet<=0) return;
    const key=keyAgentAccount(agent,account);
    if(!gameStats[key]) gameStats[key]={total:0,halls:{}};
    gameStats[key].total += bet;
    gameStats[key].halls[hall] = (gameStats[key].halls[hall]||0) + bet;
  });
}
function getGameRows(agent,account){
  const st=gameStats[keyAgentAccount(agent,account)];
  if(!st || !st.total) return [];
  return Object.entries(st.halls).map(([hall,bet])=>({hall,bet,pct:bet/st.total*100,total:st.total})).sort((a,b)=>b.bet-a.bet);
}
function getTopGame(agent,account){
  const rows=getGameRows(agent,account);
  if(!rows.length) return {hall:'',bet:0,pct:0,total:0};
  return rows[0];
}
function analyze(){
  groups=[];
  agentTotals={};
  const byAgent={}; rawRows.forEach(r=>{(byAgent[r.agent] ||= []).push(r)});
  Object.keys(byAgent).forEach(agent=>{ agentTotals[agent]=new Set(byAgent[agent].map(r=>r.account)).size; });
  for(const agent in byAgent){
    const rows=byAgent[agent];
    const maps=[{}, {}, {}, {}];
    rows.forEach(r=>{
      addGroup(maps[0], normName(r.fullName), r, 'Họ tên trùng', r.fullName);
      addGroup(maps[1], r.ip, r, 'IP trùng hoàn toàn', r.ip);
      addGroup(maps[2], ip24(r.ip), r, 'IP cùng 3 dãy đầu', ip24(r.ip));
      addGroup(maps[3], normName(r.bank)+'|'+normName(r.branch), r, 'Ngân hàng + chi nhánh trùng', (r.bank||'')+' / '+(r.branch||''));
    });
    maps.forEach(mp=>Object.values(mp).forEach(g=>{ if(g.rows.length>=2) groups.push(makeGroup(agent,g)); }));
  }
  groups.sort((a,b)=>b.score-a.score || b.count-a.count || b.totalDeposit-a.totalDeposit);
}
function makeGroup(agent,g){
  const rows=g.rows.slice().sort((a,b)=>(parseDate(a.regTime)||0)-(parseDate(b.regTime)||0));
  const regDates=rows.map(r=>parseDate(r.regTime)).filter(d=>d&&!isNaN(d));
  const lastDepositDates=rows.map(r=>parseDate(r.lastDeposit)).filter(d=>d&&!isNaN(d));
  const firstReg=regDates.length ? regDates[0] : null;
  const lastDeposit=lastDepositDates.length ? lastDepositDates.sort((a,b)=>a-b)[lastDepositDates.length-1] : null;
  const days=(firstReg&&lastDeposit)?Math.max(0,Math.ceil((lastDeposit-firstReg)/86400000)):'';
  const totalDeposit=rows.reduce((s,r)=>s+r.deposit,0);
  const agentTotal=agentTotals[agent] || rows.length;
  const ratio=agentTotal ? rows.length / agentTotal * 100 : 0;
  let score=0;
  if(g.type.includes('IP trùng'))score+=60;
  if(g.type.includes('IP cùng'))score+=35;
  if(g.type.includes('Họ tên'))score+=45;
  if(g.type.includes('Ngân hàng +'))score+=35;
  if(rows.length>=3)score+=15; if(rows.length>=5)score+=20;
  if(ratio>=30)score+=20; else if(ratio>=10)score+=10;
  const risk=score>=70?'Cao':score>=40?'Theo dõi':'Thấp';
  return {agent,type:g.type,value:g.value,count:rows.length,agentTotal,ratio,first:firstReg?fmtDate(firstReg):'',last:lastDeposit?fmtDate(lastDeposit):'',days,totalDeposit,risk,score,rows};
}
function riskWeight(r){return r==='Cao'?3:r==='Theo dõi'?2:1}
function countRatioHtml(g){
  const cls=g.ratio>=30?'ratioHigh':g.ratio>=10?'ratioMid':'ratioLow';
  return `<span class="tag ${cls}">${g.count}/${g.agentTotal} (${fmtPct(g.ratio)})</span>`;
}
function setSortMode(mode){document.getElementById('sortMode').value=mode; render();}
function render(){
  const q=document.getElementById('searchBox').value.toLowerCase().trim();
  const rf=document.getElementById('riskFilter').value;
  const minCount=Number(document.getElementById('minCount').value||0);
  const maxRaw=document.getElementById('maxCount').value;
  const maxCount=maxRaw===''?999999:Number(maxRaw||999999);
  const sortMode=document.getElementById('sortMode').value;
  let filtered=groups.filter(g=>
    g.count>=minCount &&
    g.count<=maxCount &&
    (!rf||g.risk===rf) &&
    (!q || JSON.stringify(g).toLowerCase().includes(q))
  );
  filtered=filtered.slice().sort((a,b)=>{
    if(sortMode==='count_desc') return b.count-a.count || b.score-a.score || b.totalDeposit-a.totalDeposit;
    if(sortMode==='count_asc') return a.count-b.count || b.score-a.score || b.totalDeposit-a.totalDeposit;
    if(sortMode==='deposit_desc') return b.totalDeposit-a.totalDeposit || b.score-a.score;
    if(sortMode==='deposit_asc') return a.totalDeposit-b.totalDeposit || b.score-a.score;
    if(sortMode==='risk_desc') return riskWeight(b.risk)-riskWeight(a.risk) || b.score-a.score || b.count-a.count;
    return b.score-a.score || b.count-a.count || b.totalDeposit-a.totalDeposit;
  });
  document.getElementById('stAccounts').textContent=rawRows.length;
  document.getElementById('stAgents').textContent=new Set(rawRows.map(r=>r.agent)).size;
  document.getElementById('stGroups').textContent=groups.length;
  document.getElementById('stHigh').textContent=groups.filter(g=>g.risk==='Cao').length;
  const tb=document.querySelector('#resultTable tbody'); tb.innerHTML='';
  filtered.forEach((g,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${g.agent}</td><td>${g.type}</td><td>${g.value||''}</td><td class="right">${countRatioHtml(g)}</td><td class="right">${fmtNum(g.totalDeposit)}</td><td><span class="tag ${g.risk==='Cao'?'high':g.risk==='Theo dõi'?'mid':'low'}">${g.risk}</span></td><td><button onclick="toggleDetail(${i})">Xem</button></td>`;
    tb.appendChild(tr);
    const dt=document.createElement('tr'); dt.className='detail'; dt.id='detail_'+i;
    dt.innerHTML=`<td colspan="7"><div class="small"><b>Chi tiết nhóm:</b></div>${renderDetail(g)}</td>`;
    tb.appendChild(dt);
  });
}
function renderDetail(g){
  let html=`<div class="mini"><b>Thông tin thời gian của nhóm:</b> <span class="pill">Đăng ký sớm nhất: ${g.first || 'Không có'}</span> <span class="pill">Nạp cuối gần nhất: ${g.last || 'Không có'}</span> <span class="pill">Khoảng cách ngày: ${g.days === '' ? 'Không có' : g.days}</span></div>`;
  html += `<table><thead><tr><th>Hội viên</th><th>Đăng ký</th><th>Tổng nạp</th><th>IP</th><th>Nạp cuối</th><th>Họ tên</th><th>Cấp độ nhóm</th><th>Ngân hàng</th><th>Chi nhánh</th><th>Sảnh chính</th><th class="right">Tỷ lệ</th></tr></thead><tbody>`;
  html += g.rows.map(r=>{
    const top=getTopGame(g.agent,r.account);
    return `<tr><td>${r.account}</td><td>${fmtDate(r.regTime)}</td><td class="right">${fmtNum(r.deposit)}</td><td>${r.ip}</td><td>${fmtDate(r.lastDeposit)}</td><td>${r.fullName}</td><td>${r.level||''}</td><td>${r.bank}</td><td>${r.branch}</td><td>${top.hall||'<span class="muted">Không có</span>'}</td><td class="right">${top.hall?fmtPct(top.pct):''}</td></tr>`;
  }).join('');
  html += `</tbody></table>`;

  html += `<div class="mini"><b>Chi tiết sảnh trò chơi và cược hợp lệ:</b>`;
  g.rows.forEach(r=>{
    const rows=getGameRows(g.agent,r.account);
    html += `<div class="accountTitle">${r.account}</div>`;
    if(!rows.length){
      html += `<div class="noGame small">Không tìm thấy dữ liệu cược game cho hội viên này trong file cược.</div>`;
    }else{
      html += `<table><thead><tr><th>Sảnh trò chơi</th><th class="right">Cược hợp lệ</th><th class="right">Tỷ lệ %</th><th>Biểu đồ</th></tr></thead><tbody>`;
      html += rows.map(x=>`<tr><td>${x.hall}</td><td class="right">${fmtNum(x.bet)}</td><td class="right">${fmtPct(x.pct)}</td><td><div class="bar"><span style="width:${Math.min(100,x.pct)}%"></span></div></td></tr>`).join('');
      html += `<tr><td><b>Tổng</b></td><td class="right"><b>${fmtNum(rows[0].total)}</b></td><td class="right"><b>100.00%</b></td><td></td></tr>`;
      html += `</tbody></table>`;
    }
  });
  html += `</div>`;
  return html;
}
function toggleDetail(i){const e=document.getElementById('detail_'+i); e.style.display=e.style.display==='table-row'?'none':'table-row'}
function exportResult(){
  if(!groups.length){alert('Chưa có dữ liệu');return;}
  const data=groups.map(g=>({'Đại lý':g.agent,'Dấu hiệu':g.type,'Giá trị':g.value,'TK nghi vấn':g.count,'Tổng TK đại lý':g.agentTotal,'Tỷ lệ nghi vấn':fmtPct(g.ratio),'Đăng ký sớm nhất':g.first,'Nạp cuối gần nhất':g.last,'Khoảng cách ngày':g.days,'Tổng nạp':g.totalDeposit,'Mức rủi ro':g.risk,'Điểm':g.score}));
  const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Tong_hop'); XLSX.writeFile(wb,'kiem_tra_dai_ly_nghi_van.xlsx');
}
function exportGroups(){
  if(!groups.length){alert('Chưa có dữ liệu');return;}
  const data=[]; groups.forEach((g,idx)=>g.rows.forEach(r=>{
    const top=getTopGame(g.agent,r.account);
    data.push({'Nhóm':idx+1,'Đại lý':g.agent,'Dấu hiệu':g.type,'Mức rủi ro':g.risk,'TK nghi vấn':g.count,'Tổng TK đại lý':g.agentTotal,'Tỷ lệ nghi vấn':fmtPct(g.ratio),'Hội viên':r.account,'Đăng ký':fmtDate(r.regTime),'Tổng nạp':r.deposit,'IP':r.ip,'Nạp cuối':fmtDate(r.lastDeposit),'Họ tên':r.fullName,'Cấp độ nhóm':r.level,'Ngân hàng':r.bank,'Chi nhánh':r.branch,'Sảnh chính':top.hall,'Tổng cược sảnh chính':top.bet,'Tỷ lệ sảnh chính':top.hall?fmtPct(top.pct):''});
  }));
  const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Chi_tiet_nhom'); XLSX.writeFile(wb,'chi_tiet_nhom_nghi_van.xlsx');
}
