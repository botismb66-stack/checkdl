let rawRows = [];
let groups = [];
let gameStats = {};
let agentTotals = {};

const COL = {
  account:0,
  regTime:1,
  fullName:2,
  agent:3,
  deposit:5,
  ip:6,
  lastDeposit:7,
  bank:8,
  branch:9
};

const GAME_COL = {
  agent:2,
  account:3,
  hall:4,
  bet:7
};

// ===== UTIL =====
function val(r,i){return (r[i]??'').toString().trim()}
function num(v){return Number(String(v).replace(/,/g,'').replace(/[^0-9.-]/g,''))||0}

function parseDate(v){
  if(v instanceof Date) return v;
  if(!v) return null;
  return new Date(v);
}

function fmtDate(d){
  if(!d) return '';
  d=new Date(d);
  return d.toLocaleString();
}

function norm(s){
  return String(s||'')
  .toLowerCase()
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'')
  .replace(/[^a-z0-9]/g,'');
}

// ===== READ FILE =====
function readFile(file){
  return new Promise((res,rej)=>{
    const fr=new FileReader();
    fr.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:'array'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
        res(data);
      }catch(err){rej(err)}
    };
    fr.readAsArrayBuffer(file);
  });
}

// ===== HANDLE FILE =====
async function handleFile(){
  const f=document.getElementById('fileInput').files[0];
  if(!f) return alert("Chọn file");

  const gfile=document.getElementById('gameFileInput').files[0];

  const data = await readFile(f);

  rawRows = data.slice(1).map(r=>({
    account: val(r,COL.account),
    regTime: val(r,COL.regTime),
    fullName: val(r,COL.fullName),
    agent: val(r,COL.agent),
    deposit: num(r[COL.deposit]),
    ip: val(r,COL.ip),
    lastDeposit: val(r,COL.lastDeposit),
    bank: val(r,COL.bank),
    branch: val(r,COL.branch),
  })).filter(x=>x.account && x.agent && x.deposit>0);

  if(gfile) await loadGame(gfile);

  analyze();
  render();
}

// ===== GAME FILE =====
async function loadGame(file){
  const data = await readFile(file);

  gameStats={};

  data.slice(1).forEach(r=>{
    const agent=val(r,GAME_COL.agent);
    const acc=val(r,GAME_COL.account);
    const hall=val(r,GAME_COL.hall);
    const bet=num(r[GAME_COL.bet]);

    if(!agent||!acc||!hall||bet<=0) return;

    const k=agent+"|"+acc;
    if(!gameStats[k]) gameStats[k]={total:0,halls:{}};

    gameStats[k].total+=bet;
    gameStats[k].halls[hall]=(gameStats[k].halls[hall]||0)+bet;
  });
}

// ===== ANALYZE (FULL LOGIC) =====
function analyze(){
  groups=[];
  agentTotals={};

  const byAgent={};
  rawRows.forEach(r=>{
    (byAgent[r.agent] ||= []).push(r);
  });

  for(const agent in byAgent){

    const rows=byAgent[agent];
    agentTotals[agent]=new Set(rows.map(x=>x.account)).size;

    const maps=[{}, {}, {}, {}];

    rows.forEach(r=>{
      // name
      (maps[0][norm(r.fullName)] ||= []).push(r);

      // IP exact
      (maps[1][r.ip] ||= []).push(r);

      // IP 3 segment
      const ip3=r.ip.split('.').slice(0,3).join('.');
      (maps[2][ip3] ||= []).push(r);

      // bank
      (maps[3][norm(r.bank+"-"+r.branch)] ||= []).push(r);
    });

    maps.forEach(m=>{
      Object.values(m).forEach(g=>{
        if(g.length>=2){
          groups.push(buildGroup(agent,g));
        }
      });
    });
  }
}

// ===== BUILD GROUP =====
function buildGroup(agent,rows){

  const totalDeposit=rows.reduce((s,x)=>s+x.deposit,0);
  const agentTotal=agentTotals[agent]||rows.length;

  const ratio=(rows.length/agentTotal)*100;

  const risk =
    ratio>30 ? "Cao" :
    ratio>10 ? "Theo dõi" :
    "Thấp";

  const firstReg = rows.map(r=>new Date(r.regTime)).sort((a,b)=>a-b)[0];
  const lastDep = rows.map(r=>new Date(r.lastDeposit)).sort((a,b)=>b-a)[0];

  const days = firstReg&&lastDep
    ? Math.ceil((lastDep-firstReg)/86400000)
    : '';

  return {
    agent,
    type:"Trùng dữ liệu",
    value:"",
    rows,
    count:rows.length,
    agentTotal,
    ratio,
    totalDeposit,
    risk,
    first:firstReg,
    last:lastDep,
    days
  };
}

// ===== RENDER =====
function render(){

  const tb=document.querySelector("#resultTable tbody");
  tb.innerHTML="";

  groups.forEach((g,i)=>{

    const tr=document.createElement("tr");

    tr.innerHTML=`
      <td>${g.agent}</td>
      <td>${g.type}</td>
      <td>${g.value||''}</td>
      <td>${g.count}/${g.agentTotal} (${g.ratio.toFixed(1)}%)</td>
      <td>${g.totalDeposit.toLocaleString()}</td>
      <td>${g.risk}</td>
      <td><button onclick="toggle(${i})">Xem</button></td>
    `;

    tb.appendChild(tr);

    const detail=document.createElement("tr");
    detail.id="d"+i;
    detail.style.display="none";

    detail.innerHTML=`
      <td colspan="7">
        <b>Đăng ký sớm:</b> ${fmtDate(g.first)} |
        <b>Nạp cuối:</b> ${fmtDate(g.last)} |
        <b>Khoảng cách:</b> ${g.days}

        <table>
          <tr><th>TK</th><th>IP</th><th>Nạp</th><th>Bank</th></tr>
          ${g.rows.map(r=>`
            <tr>
              <td>${r.account}</td>
              <td>${r.ip}</td>
              <td>${r.deposit.toLocaleString()}</td>
              <td>${r.bank}</td>
            </tr>
          `).join('')}
        </table>
      </td>
    `;

    tb.appendChild(detail);
  });
}

function toggle(i){
  const el=document.getElementById("d"+i);
  el.style.display = el.style.display==="none"?"table-row":"none";
}

// ===== EXPORT (giữ đơn giản) =====
function exportResult(){
  alert("Export tổng (giữ logic cũ bạn có thể gắn lại XLSX)");
}

function exportGroups(){
  alert("Export chi tiết");
}
console.log("APP JS LOADED");