const WEEKS = Array.from({length:52},(_,i)=>i+1);

let _debounceTimer = null;
function debounce(fn, ms){ clearTimeout(_debounceTimer); _debounceTimer = setTimeout(fn, ms); }
function setFilter(key, val){ state[key]=val; debounce(render, 300); }


const TEAMS = ['Development','Platform','PMO'];
const TYPES = ['Project','Base Service','Charge On','Internal Initiative'];

const DEFAULT_SERVICES = [
  {name:'Integration Support',team:'Platform'},
  {name:'Infrastructure Support',team:'Platform'},
  {name:'Cloud Operations',team:'Platform'},
  {name:'Application Support',team:'Development'},
  {name:'Bug Fixing',team:'Development'},
  {name:'Maintenance',team:'Development'},
  {name:'PMO Governance',team:'PMO'},
  {name:'Reporting',team:'PMO'},
];

function loadSaved(){
  try { const raw=localStorage.getItem('rp_data'); if(raw) return JSON.parse(raw); } catch(e){}
  return null;
}
function saveData(){
  try {
    localStorage.setItem('rp_data', JSON.stringify({
      projects: state.projects,
      assignments: state.assignments,
      baseServices: state.baseServices,
      teamMembers: state.teamMembers,
      teamConfig: state.teamConfig,
      inboxItems: state.inboxItems,
      userName: document.getElementById('user-name')?.value||'Martina Vallgren',
      role: document.getElementById('role-sel')?.value||'Teamlead',
    }));
  } catch(e){}
}
function exportData(){
  const data = localStorage.getItem('rp_data') || JSON.stringify({
    projects: state.projects, assignments: state.assignments,
    baseServices: state.baseServices, teamMembers: state.teamMembers,
    teamConfig: state.teamConfig, inboxItems: state.inboxItems
  });
  const blob = new Blob([data], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'ResourcePlanner_backup_' + new Date().toISOString().split('T')[0] + '.json';
  a.click();
}

function importData(){
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.onchange = e => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        localStorage.setItem('rp_data', JSON.stringify(data));
        alert('Data imported! The page will now reload.');
        location.reload();
      } catch(err) {
        alert('Could not read file. Make sure it is a valid backup file.');
      }
    };
    reader.readAsText(file);
  };
  input.click();
}


function clearData(){
  if(!confirm('Clear all data? This cannot be undone.')) return;
  localStorage.removeItem('rp_data');
  location.reload();
}

const saved = loadSaved();

let state = {
  tab: 'dashboard',
  projects: saved?.projects || [],
  assignments: saved?.assignments || [],
  baseServices: saved?.baseServices || DEFAULT_SERVICES,
  addType: 'Project',
  aName:'', aTeam:'Development', aCountry:'Sweden', aSkill:'', aLevel:'Junior',
  aProjId:'', aService:'', aWork:'', aStart:1, aEnd:3, aPct:80,
  pName:'', pPm:'', pStart:'', pEnd:'', pDesc:'',
  sName:'', sTeam:'Development',
  msg: null,
  selectedProject: null,
  selectedTeam: null,
  prName:'', prTeam:'Development', prCountry:'Sweden', prSkill:'', prLevel:'Junior', prStart:1, prEnd:4, prPct:80,
  teamMembers: saved?.teamMembers || [],
  teamConfig: saved?.teamConfig || {
    Development: { teamlead: '', manager: '' },
    Platform:    { teamlead: '', manager: '' },
    PMO:         { teamlead: '', manager: '' },
  },
  tmName:'', tmCountry:'Sweden', tmSkill:'', tmLevel:'Junior', tmTeamlead:'', tmManager:'',
  inboxItems: saved?.inboxItems || [],
  iTitle:'', iDesc:'', iPriority:'Medium',
  editingMemberId: null,
  emName:'', emCountry:'Sweden', emSkill:'', emLevel:'Junior', emTeamlead:'', emManager:'',
  fTeam:'', fSkill:'', fLevel:'', fName:'', fAssignment:'', fStatus:'',
  dashAllocRange: 8,
  selectedPerson: null,
};

function role(){ return document.getElementById('role-sel').value; }
function userName(){ return document.getElementById('user-name').value; }
function canEdit(){ return role()==='Teamlead'||role()==='Manager'; }
function canPlan(){ return role()==='Teamlead'||role()==='Manager'||role()==='Project Manager'; }
function pmProjects(){ return state.projects.filter(p=>isPmProject(p)); }

function visibleAssignments(){
  const r = role();
  const un = userName().trim().toLowerCase();
  return state.assignments.filter(a => {
    if(r === 'Team Member') return a.committed && a.name.trim().toLowerCase() === un;
    return true; // Teamlead, Manager, Project Manager see all
  });
}
function getPeople(){
  const map=new Map();
  visibleAssignments().forEach(a=>{ if(!map.has(a.name.toLowerCase())) map.set(a.name.toLowerCase(),a); });
  return [...map.values()];
}
function getAlloc(a,w){
  return a.periods.filter(p=>w>=p.startWeek&&w<=p.endWeek).reduce((s,p)=>s+p.allocationPercent,0);
}
function getTotalAlloc(name,w){
  return visibleAssignments().filter(a=>a.name.toLowerCase()===name.toLowerCase()).reduce((s,a)=>s+getAlloc(a,w),0);
}
function wClass(t){ return t>100?'ao':t===100?'af':t>0?'ap':''; }
function cBg(c){ return c==='Sweden'?'#dbeafe':'#fef3c7'; }

function calcTeamAllocForWeek(teamMembers, team, w){
  const members = teamMembers[team] || [];
  if(!members.length) return {avg:0,fullyBooked:0,free:0,over:0,total:0};
  // personTotalAtWeek uses state directly — accessible from outer scope
  const allocs = members.map(function(name){
    return state.assignments.filter(function(a){ return a.name===name; }).reduce(function(s,a){ return s+getAlloc(a,w); },0);
  });
  return {
    avg: Math.round(allocs.reduce(function(s,v){ return s+v; },0)/members.length),
    fullyBooked: allocs.filter(function(v){ return v>=100; }).length,
    free: allocs.filter(function(v){ return v===0; }).length,
    over: allocs.filter(function(v){ return v>100; }).length,
    total: members.length,
  };
}

function buildTeamAllocCard(teamMembers, allocWeeks, cw, state){
  const rangeOpts = [4,8,12,26,52].map(n=>{
    const active = state.dashAllocRange===n;
    return '<button onclick="state.dashAllocRange='+n+';render()" style="padding:5px 14px;font-size:12px;font-weight:600;border-radius:20px;border:1px solid '+(active?'#1D9E75':'#e5e7eb')+';background:'+(active?'#1D9E75':'#fff')+';color:'+(active?'#fff':'#6b7280')+';cursor:pointer;font-family:inherit;transition:all .1s">'+n+'W</button>';
  }).join('');

  const wHeaders = allocWeeks.map(w=>{
    const isNow = w===cw;
    return '<th style="padding:8px 4px;text-align:center;min-width:58px;'+(isNow?'background:#f0fdf8;':'')+'">'+
      '<div style="font-size:10px;font-weight:700;color:'+(isNow?'#0f6e56':'#9ca3af')+';font-family:DM Mono,monospace">W'+w+'</div>'+
      (isNow?'<div style="width:4px;height:4px;background:#1D9E75;border-radius:50%;margin:2px auto 0"></div>':'')+
    '</th>';
  }).join('');

  const teamRows = ['Development','Platform','PMO'].map(function(team, ti){
    const icon = team==='Development'?'💻':team==='Platform'?'☁':'📊';
    const memberCount = teamMembers[team].length;
    const nowAlloc = calcTeamAllocForWeek(teamMembers, team, cw);
    const bigPct = nowAlloc.avg;
    const bigColor = memberCount===0?'#d1d5db':bigPct>100?'#b91c1c':bigPct>=80?'#0f6e56':bigPct>=50?'#b45309':'#185fa5';

    let weekCells = '';
    if(memberCount===0){
      weekCells = '<td colspan="'+allocWeeks.length+'" style="padding:16px 24px;vertical-align:middle">'
        +'<div style="display:flex;align-items:center;gap:10px">'
        +'<div style="flex:1;height:8px;background:#f3f4f6;border-radius:4px"></div>'
        +'<span style="font-size:12px;color:#d1d5db;font-style:italic">No members yet</span>'
        +'</div></td>';
    } else {
      weekCells = allocWeeks.map(function(w){
        const ta = calcTeamAllocForWeek(teamMembers, team, w);
        const isNow = w===cw;
        const barH = Math.min(Math.round((ta.avg/100)*52),52);
        const barColor = ta.avg>100?'#fca5a5':ta.avg>=80?'#6ee7b7':ta.avg>=50?'#fcd34d':ta.avg>0?'#93c5fd':'#e5e7eb';
        const textColor = ta.avg>100?'#b91c1c':ta.avg>=80?'#065f46':ta.avg>=50?'#92400e':ta.avg>0?'#1e40af':'#d1d5db';
        return '<td style="padding:8px 4px;text-align:center;vertical-align:bottom;'+(isNow?'background:#f0fdf8;':'')+'border-bottom:1px solid #f3f4f6">'
          +'<div style="display:flex;flex-direction:column;align-items:center;gap:3px">'
          +(ta.over?'<div style="font-size:9px;color:#b91c1c;font-weight:700">+'+ta.over+'</div>':'<div style="font-size:9px;color:transparent">·</div>')
          +'<div style="font-size:12px;font-weight:700;color:'+textColor+';font-family:DM Mono,monospace;line-height:1">'+(ta.avg>0?ta.avg+'%':'–')+'</div>'
          +'<div style="width:32px;height:52px;background:#f3f4f6;border-radius:4px;overflow:hidden;display:flex;align-items:flex-end">'
          +'<div style="width:100%;height:'+barH+'px;background:'+barColor+';border-radius:4px"></div>'
          +'</div></div></td>';
      }).join('');
    }

    const statusPills =
      (nowAlloc.fullyBooked?'<span style="font-size:10px;font-weight:700;color:#065f46;background:#d1fae5;padding:1px 6px;border-radius:20px">'+nowAlloc.fullyBooked+' full</span>':'')+
      (nowAlloc.free?'<span style="font-size:10px;font-weight:700;color:#185fa5;background:#dbeafe;padding:1px 6px;border-radius:20px;margin-left:4px">'+nowAlloc.free+' free</span>':'')+
      (nowAlloc.over?'<span style="font-size:10px;font-weight:700;color:#b91c1c;background:#fef2f2;padding:1px 6px;border-radius:20px;margin-left:4px">'+nowAlloc.over+' over</span>':'')+
      (memberCount===0?'<span style="font-size:10px;color:#d1d5db;font-style:italic">no members yet</span>':'');

    return '<tr style="border-bottom:'+(ti<2?'2px solid #e5e7eb':'1px solid #f3f4f6')+'">'
      +'<td style="padding:14px 24px;vertical-align:middle;min-width:220px">'
      +'<div style="display:flex;align-items:center;gap:12px">'
      +'<div style="width:44px;height:44px;border-radius:12px;background:#f0fdf8;display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0">'+icon+'</div>'
      +'<div style="flex:1;min-width:0">'
      +'<div style="font-size:14px;font-weight:700;color:#111827">'+team+'</div>'
      +'<div style="font-size:12px;color:#6b7280;margin-top:2px">'+memberCount+' member'+(memberCount!==1?'s':'')+'</div>'
      +'<div style="display:flex;gap:6px;margin-top:4px;flex-wrap:wrap">'+statusPills+'</div>'
      +'</div>'
      +'<div style="text-align:right;flex-shrink:0">'
      +'<div style="font-size:28px;font-weight:700;color:'+bigColor+';font-family:DM Mono,monospace;line-height:1">'+(memberCount===0?'—':bigPct+'%')+'</div>'
      +'<div style="font-size:10px;color:#9ca3af;margin-top:2px">now</div>'
      +'</div></div></td>'
      + weekCells
      +'</tr>';
  }).join('');

  return '<div class="card">'
    +'<div class="card-hdr" style="padding:16px 24px">'
    +'<div>'
    +'<span class="card-title" style="font-size:15px">📊 Team allocation</span>'
    +'<div style="font-size:12px;color:#9ca3af;margin-top:2px">Average allocation % per team · '+allocWeeks.length+' weeks shown</div>'
    +'</div>'
    +'<div style="display:flex;gap:6px;align-items:center">'+rangeOpts+'</div>'
    +'</div>'
    +'<div style="overflow-x:auto">'
    +'<table style="width:100%;border-collapse:collapse;min-width:'+(180+allocWeeks.length*62)+'px">'
    +'<thead><tr style="border-bottom:2px solid #e5e7eb">'
    +'<th style="padding:10px 24px;text-align:left;font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;min-width:180px">Team</th>'
    + wHeaders
    +'</tr></thead>'
    +'<tbody>'+teamRows+'</tbody>'
    +'</table></div>'
    +'<div style="padding:12px 24px;border-top:1px solid #f3f4f6;display:flex;gap:20px;font-size:11px;color:#6b7280;flex-wrap:wrap;align-items:center">'
    +'<span style="font-weight:600;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;font-size:10px">Legend</span>'
    +'<span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#93c5fd;display:inline-block"></span>Low &lt;50%</span>'
    +'<span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#fcd34d;display:inline-block"></span>Medium 50–79%</span>'
    +'<span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#6ee7b7;display:inline-block"></span>High 80–100%</span>'
    +'<span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#fca5a5;display:inline-block"></span>Overbooked</span>'
    +'<span style="display:flex;align-items:center;gap:5px"><span style="width:4px;height:4px;border-radius:50%;background:#1D9E75;display:inline-block"></span>Current week</span>'
    +'</div>'
    +'</div>';
}


function renderDashboard(){
  const r = role();
  const ce = canEdit();
  const now = new Date();
  // Approximate current week number
  const startOfYear = new Date(now.getFullYear(), 0, 1);
  const currentWeek = Math.ceil(((now - startOfYear) / 86400000 + startOfYear.getDay() + 1) / 7);
  const cw = Math.min(Math.max(currentWeek, 1), 52);
  const NEXT = 4; // look-ahead weeks

  // ── Data helpers ─────────────────────────────────────────────────────────
  function personTotalAtWeek(name, w){
    return state.assignments.filter(a=>a.name===name).reduce((s,a)=>s+getAlloc(a,w),0);
  }

  // People active now (has allocation in current week)
  const isTM = r==='Team Member';
  const un = userName().trim().toLowerCase();
  // Team Members only see their own data
  const allPeople = isTM
    ? [...new Set(state.assignments.filter(a=>a.committed&&a.name.trim().toLowerCase()===un).map(a=>a.name))]
    : [...new Set(state.assignments.map(a=>a.name))];
  const activeNow = allPeople.filter(name=>personTotalAtWeek(name,cw)>0);

  // Projects active now
  const activeProjects = state.projects.filter(p=>{
    const hasActiveResource = state.assignments.some(a=>a.workName===p.name && a.periods.some(per=>cw>=per.startWeek&&cw<=per.endWeek));
    return hasActiveResource;
  });

  // Starting soon — only assignments linked to active/ongoing projects
  const activeProjectNames = new Set(activeProjects.map(p=>p.name));
  const startingSoon = state.assignments.filter(a=>
    a.type==='Project' &&
    activeProjectNames.has(a.workName) &&
    a.periods.some(p=>p.startWeek>cw && p.startWeek<=cw+NEXT)
  );
  const startingSoonProjects = [...new Set(startingSoon.map(a=>a.workName))];

  // Ending soon — only assignments linked to active/ongoing projects (next NEXT weeks)
  const endingSoon = state.assignments.filter(a=>
    a.type==='Project' &&
    activeProjectNames.has(a.workName) &&
    a.periods.some(p=>p.endWeek>=cw && p.endWeek<=cw+NEXT)
  );

  // Overbooked now
  const overbookedNow = allPeople.filter(name=>personTotalAtWeek(name,cw)>100);

  // Free capacity now (registered but no allocation this week)
  const allRegistered = isTM
    ? [userName().trim()].filter(n=>n)
    : [...new Set([...state.teamMembers.map(m=>m.name),...state.assignments.map(a=>a.name)])];
  const freeNow = allRegistered.filter(name=>personTotalAtWeek(name,cw)===0);

  // Uncommitted planned
  const uncommitted = state.assignments.filter(a=>!a.committed);
  const uncommittedByProject = {};
  uncommitted.forEach(a=>{
    if(!uncommittedByProject[a.workName]) uncommittedByProject[a.workName]=[];
    uncommittedByProject[a.workName].push(a);
  });

  // Inbox needing attention
  const inboxPending = state.inboxItems.filter(i=>i.status==='new');

  // Upcoming project end dates
  const now_str = now.toISOString().split('T')[0];
  const upcomingProjectEnds = state.projects.filter(p=>p.endDate && p.endDate>=now_str).sort((a,b)=>a.endDate.localeCompare(b.endDate)).slice(0,5);

  // Next 8 weeks: who starts/ends each week
  const weeklyChanges = [];
  for(let w=cw;w<=Math.min(cw+NEXT-1,52);w++){
    const starts = state.assignments.filter(a=>a.periods.some(p=>p.startWeek===w));
    const ends   = state.assignments.filter(a=>a.periods.some(p=>p.endWeek===w));
    if(starts.length||ends.length) weeklyChanges.push({w,starts,ends});
  }

  function fmtDate(d){ if(!d)return'—'; return new Date(d).toLocaleDateString('en-SE',{day:'2-digit',month:'short'}); }

  // Team allocation across weeks
  const ALLOC_RANGE = state.dashAllocRange || 8;
  const allocWeeks = WEEKS.slice(cw-1, cw-1+ALLOC_RANGE).filter(w=>w<=52);

  // Build team member lists — normalize team names, include both registered members and assignment people
  const teamMembers = {'Development':[], 'Platform':[], 'PMO':[]};
  const teamNormalize = (t) => {
    if(!t) return null;
    const s = t.trim().toLowerCase();
    if(s==='development') return 'Development';
    if(s==='platform')    return 'Platform';
    if(s==='pmo')         return 'PMO';
    return null;
  };
  state.teamMembers.forEach(m => {
    const t = teamNormalize(m.team);
    if(t && !teamMembers[t].includes(m.name)) teamMembers[t].push(m.name);
  });
  state.assignments.forEach(a => {
    const t = teamNormalize(a.team);
    if(t && !teamMembers[t].includes(a.name)) teamMembers[t].push(a.name);
  });

  function teamAllocForWeek(team, w){
    return calcTeamAllocForWeek(teamMembers, team, w);
  }

  // Current week single values for KPIs
  const teamAllocNow = {};
  ['Development','Platform','PMO'].forEach(t=>{ teamAllocNow[t]=teamAllocForWeek(t,cw); });
  function pTag(label,color,bg){ return `<span class="dash-tag" style="background:${bg};color:${color}">${label}</span>`; }

  // ── Render ────────────────────────────────────────────────────────────────
  return `
    <!-- Top KPI strip -->
    <div class="metrics" style="grid-template-columns:repeat(6,1fr);margin-bottom:0">
      <div class="metric"><div class="metric-lbl">Current week</div><div class="metric-val">W${cw}</div></div>
      <div class="metric"><div class="metric-lbl">Active now</div><div class="metric-val" style="color:#0f6e56">${activeNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Free now</div><div class="metric-val" style="color:#185fa5">${freeNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Overbooked</div><div class="metric-val ${overbookedNow.length?'red':''}">${overbookedNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Uncommitted</div><div class="metric-val" style="color:#b45309">${uncommitted.length}</div></div>
      <div class="metric"><div class="metric-lbl">Inbox</div><div class="metric-val ${inboxPending.length?'red':''}">${inboxPending.length}</div></div>
    </div>

    <!-- Team allocation -->
    <div id='team-alloc-card'></div>

    <!-- Alerts & actions needed -->
    ${!isTM && (overbookedNow.length||inboxPending.length||uncommitted.length) ? `
    <div class="card">
      <div class="card-hdr"><span class="card-title">⚠ Needs attention</span></div>
      ${overbookedNow.map(name=>`
        <div class="alert-row">
          <span class="alert-icon">🔴</span>
          <div style="flex:1"><strong>${name}</strong> is overbooked in W${cw} (${personTotalAtWeek(name,cw)}%)</div>
          <button class="btn sm" onclick="openPersonDetail('${name.replace(/'/g,"\'")}')">View →</button>
        </div>`).join('')}
      ${inboxPending.map(i=>`
        <div class="alert-row">
          <span class="alert-icon">📥</span>
          <div style="flex:1"><strong>${i.title}</strong> — in inbox, needs classification</div>
          <button class="btn sm" onclick="setTab('inbox')">Go to Inbox →</button>
        </div>`).join('')}
      ${Object.entries(uncommittedByProject).slice(0,4).map(([proj,items])=>`
        <div class="alert-row">
          <span class="alert-icon">⏳</span>
          <div style="flex:1">${items.length} planned resource${items.length!==1?'s':''} not yet committed on <strong>${proj}</strong></div>
          ${state.projects.find(p=>p.name===proj)?`<button class="btn sm" onclick="openProject(${state.projects.find(p=>p.name===proj).id})">View →</button>`:''}
        </div>`).join('')}
    </div>` : `<div class="card"><div style="padding:14px 18px;font-size:13px;color:#0f6e56;font-weight:600">✅ Everything looks good — no immediate issues.</div></div>`}

    <div class="dash-grid">

      <!-- Active now -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">👥 Active this week (W${cw})</span></div>
        ${!activeNow.length
          ? `<div class="empty" style="padding:20px">No active resources this week.</div>`
          : activeNow.map(name=>{
              const alloc = personTotalAtWeek(name,cw);
              const info = state.assignments.find(a=>a.name===name)||{team:'',skillset:'',level:''};
              const works = [...new Set(state.assignments.filter(a=>a.name===name&&a.periods.some(p=>cw>=p.startWeek&&cw<=p.endWeek)).map(a=>a.workName))];
              return `<div class="dash-row" onclick="openPersonDetail('${name.replace(/'/g,"\'")}')">
                <div style="flex:1">
                  <div style="font-weight:600">${name}</div>
                  <div style="font-size:11px;color:#9ca3af">${info.team} · ${works.join(', ')}</div>
                </div>
                <span class="dash-tag" style="background:${alloc>100?'#fecaca;color:#b91c1c':alloc===100?'#d1fae5;color:#065f46':'#fef3c7;color:#92400e'}">${alloc}%</span>
              </div>`;
            }).join('')}
      </div>

      <!-- Active projects -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">💼 Active projects (W${cw})</span></div>
        ${!activeProjects.length
          ? `<div class="empty" style="padding:20px">No projects active this week.</div>`
          : activeProjects.map(p=>{
              const resources = state.assignments.filter(a=>a.workName===p.name&&a.periods.some(per=>cw>=per.startWeek&&cw<=per.endWeek));
              const comm = resources.filter(a=>a.committed).length;
              const plan = resources.filter(a=>!a.committed).length;
              return `<div class="dash-row" onclick="openProject(${p.id})">
                <div style="flex:1">
                  <div style="font-weight:600">${p.name}</div>
                  <div style="font-size:11px;color:#9ca3af">ends ${fmtDate(p.endDate)} · ${resources.length} resource${resources.length!==1?'s':''}</div>
                </div>
                <div style="display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end">
                  ${comm?pTag('✓ '+comm,'#065f46','#d1fae5'):''}
                  ${plan?pTag('⏳ '+plan,'#92400e','#fef3c7'):''}
                </div>
              </div>`;
            }).join('')}
      </div>

      <!-- Starting soon -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">🚀 Starting soon (next ${NEXT} weeks)</span></div>
        ${!startingSoonProjects.length
          ? `<div class="empty" style="padding:20px">Nothing starting in the next ${NEXT} weeks.</div>`
          : startingSoonProjects.map(workName=>{
              const items = startingSoon.filter(a=>a.workName===workName);
              const firstWeek = Math.min(...items.flatMap(a=>a.periods.filter(p=>p.startWeek>cw).map(p=>p.startWeek)));
              const proj = state.projects.find(p=>p.name===workName);
              return `<div class="dash-row" ${proj?`onclick="openProject(${proj.id})"`:''}>
                <div style="flex:1">
                  <div style="font-weight:600">${workName}</div>
                  <div style="font-size:11px;color:#9ca3af">${items.length} resource${items.length!==1?'s':''}</div>
                </div>
                <span class="week-badge">W${firstWeek}</span>
              </div>`;
            }).join('')}
      </div>

      <!-- Ending soon -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">🏁 Ending soon (next ${NEXT} weeks)</span></div>
        ${!endingSoon.length
          ? `<div class="empty" style="padding:20px">Nothing ending in the next ${NEXT} weeks.</div>`
          : [...new Map(endingSoon.map(a=>[a.name+a.workName,a])).values()].map(a=>{
              const lastWeek = Math.max(...a.periods.filter(p=>p.endWeek>=cw&&p.endWeek<=cw+NEXT).map(p=>p.endWeek));
              return `<div class="dash-row" onclick="openPersonDetail('${a.name.replace(/'/g,"\'")}')">
                <div style="flex:1">
                  <div style="font-weight:600">${a.name}</div>
                  <div style="font-size:11px;color:#9ca3af">${a.workName}</div>
                </div>
                <span class="week-badge">ends W${lastWeek}</span>
              </div>`;
            }).join('')}
      </div>

      <!-- Free capacity -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">🟢 Free this week (W${cw})</span></div>
        ${!freeNow.length
          ? `<div class="empty" style="padding:20px">Everyone is allocated this week.</div>`
          : freeNow.map(name=>{
              const info = state.teamMembers.find(m=>m.name===name) || state.assignments.find(a=>a.name===name) || {};
              // Find next week they have something
              let nextBusy = null;
              for(let w=cw+1;w<=52;w++){ if(personTotalAtWeek(name,w)>0){nextBusy=w;break;} }
              return `<div class="dash-row" onclick="openPersonDetail('${name.replace(/'/g,"\'")}')">
                <div style="flex:1">
                  <div style="font-weight:600">${name}</div>
                  <div style="font-size:11px;color:#9ca3af">${info.team||''} · ${info.skillset||''}</div>
                </div>
                <span style="font-size:11px;color:#9ca3af">${nextBusy?'busy W'+nextBusy:'free all year'}</span>
              </div>`;
            }).join('')}
      </div>

      <!-- Week-by-week next 8 weeks -->
      <div class="card">
        <div class="card-hdr"><span class="card-title">📆 Week-by-week (W${cw}–W${Math.min(cw+NEXT-1,52)})</span></div>
        ${!weeklyChanges.length
          ? `<div class="empty" style="padding:20px">No changes in the next ${NEXT} weeks.</div>`
          : weeklyChanges.map(({w,starts,ends})=>`
            <div style="padding:10px 14px;border-bottom:1px solid #f3f4f6">
              <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
                <span class="week-badge" style="background:#e0f2fe;color:#075985">W${w}</span>
                ${w===cw?'<span style="font-size:10px;font-weight:700;color:#0f6e56">← current</span>':''}
              </div>
              ${starts.length?`<div style="font-size:11px;color:#0f6e56;margin-bottom:3px">▶ Starts: ${[...new Set(starts.map(a=>a.name))].join(', ')}</div>`:''}
              ${ends.length?`<div style="font-size:11px;color:#b45309">■ Ends: ${[...new Set(ends.map(a=>a.name))].join(', ')}</div>`:''}
            </div>`).join('')}
      </div>

    </div>


    <!-- Org structure -->
    <div class="card">
      <div class="card-hdr"><span class="card-title">🏢 Organisation</span></div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:0;border-top:1px solid #f3f4f6">
        ${['Development','Platform','PMO'].map(team=>{
          const cfg = state.teamConfig[team]||{};
          const members = state.teamMembers.filter(m=>m.team===team);
          return `<div style="padding:14px 18px;border-right:1px solid #f3f4f6">
            <div style="font-size:13px;font-weight:700;color:#111827;margin-bottom:10px">${team==='Development'?'💻':team==='Platform'?'☁':'📊'} ${team}</div>
            <div style="margin-bottom:8px">
              <div class="org-label">Team Lead</div>
              <div class="org-person" style="color:${cfg.teamlead?'#0f6e56':'#9ca3af'}">${cfg.teamlead||'Not assigned'}</div>
            </div>
            <div style="margin-bottom:10px">
              <div class="org-label">Manager</div>
              <div class="org-person" style="color:${cfg.manager?'#185fa5':'#9ca3af'}">${cfg.manager||'Not assigned'}</div>
            </div>
            <div class="org-label">${members.length} member${members.length!==1?'s':''}</div>
            ${members.slice(0,4).map(m=>`<div style="font-size:12px;color:#6b7280;padding:2px 0;display:flex;align-items:center;gap:6px">
              <span style="width:6px;height:6px;border-radius:50%;background:${m.country==='Sweden'?'#3b82f6':'#f59e0b'};display:inline-block;flex-shrink:0"></span>
              ${m.name}
            </div>`).join('')}
            ${members.length>4?`<div style="font-size:11px;color:#9ca3af;margin-top:4px">+${members.length-4} more</div>`:''}
          </div>`;
        }).join('')}
      </div>
    </div>

    <!-- Project timeline -->
    ${upcomingProjectEnds.length ? `
    <div class="card">
      <div class="card-hdr"><span class="card-title">📅 Upcoming project end dates</span></div>
      ${upcomingProjectEnds.map(p=>{
        const resources = state.assignments.filter(a=>a.workName===p.name);
        const daysLeft = Math.ceil((new Date(p.endDate)-now)/86400000);
        const urgency = daysLeft<=14?'#b91c1c':daysLeft<=30?'#b45309':'#374151';
        return `<div class="dash-row" onclick="openProject(${p.id})">
          <div style="flex:1"><div style="font-weight:600">${p.name}</div>
          <div style="font-size:11px;color:#9ca3af">${resources.length} resource${resources.length!==1?'s':''}</div></div>
          <div style="text-align:right">
            <div style="font-size:12px;font-weight:700;color:${urgency}">${fmtDate(p.endDate)}</div>
            <div style="font-size:10px;color:#9ca3af">${daysLeft} day${daysLeft!==1?'s':''} left</div>
          </div>
        </div>`;
      }).join('')}
    </div>` : ''}
  `;
}


function openPersonDetail(name){
  // Team Members can only see their own detail
  if(role()==='Team Member' && name.trim().toLowerCase()!==userName().trim().toLowerCase()) return;
  state.selectedPerson = name;
  state.tab = 'person-detail';
  document.getElementById('tab-title').textContent = name;
  document.querySelectorAll('.nav-item').forEach(el=>el.classList.remove('active'));
  render();
}

function renderPersonDetail(){
  const name = state.selectedPerson;
  if(!name) return `<div class="card"><div class="empty">No person selected.</div></div>`;
  const r = role();
  const ce = canEdit();

  const personAssignments = visibleAssignments().filter(a=>a.name===name);
  const info = personAssignments[0] || state.teamMembers.find(m=>m.name===name);
  if(!info) return `<div class="card"><div class="empty">Person not found.</div></div>`;

  const committed = personAssignments.filter(a=>a.committed);
  const planned   = personAssignments.filter(a=>!a.committed);

  // Build week totals
  const weekTotals = WEEKS.map(w=>({w, t:personAssignments.reduce((s,a)=>s+getAlloc(a,w),0)}));
  const CHUNK = 13;
  const weekRows = [];
  for(let i=0;i<52;i+=CHUNK) weekRows.push(weekTotals.slice(i,i+CHUNK));

  function wStyle(t){
    if(t>100) return 'background:#fecaca;color:#b91c1c;font-weight:700';
    if(t===100) return 'background:#d1fae5;color:#065f46;font-weight:700';
    if(t>0) return 'background:#fef3c7;color:#92400e;font-weight:600';
    return 'background:#f9fafb;color:#d1d5db';
  }

  const calTable = `<div style="overflow-x:auto"><table style="border-collapse:separate;border-spacing:2px;min-width:100%">
    ${weekRows.map(chunk=>`<tr>${chunk.map(d=>`
      <td style="text-align:center;padding:4px 2px;font-size:10px;font-family:'DM Mono',monospace;border-radius:3px;min-width:34px;${wStyle(d.t)}">
        <div style="font-size:9px;opacity:.7;line-height:1">W${d.w}</div>
        <div style="line-height:1.3">${d.t>0?d.t+'%':'–'}</div>
      </td>`).join('')}</tr>`).join('')}
  </table></div>`;

  function assignmentCard(a){
    const idx = state.assignments.indexOf(a);
    const totalWks = a.periods.reduce((s,p)=>s+(p.endWeek-p.startWeek+1),0);
    return `<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:10px">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px">
        <div style="flex:1">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
            <span style="font-size:13px;font-weight:700">${a.workName}</span>
            <span class="badge b-type">${a.type}</span>
            ${a.committed?`<span class="status-committed">✓ Committed</span>`:`<span class="status-planned">⏳ Planned</span>`}
          </div>
          ${a.committed?`<div style="font-size:11px;color:#9ca3af;margin-bottom:6px">Committed by ${a.committedBy}</div>`:''}
          <div style="display:flex;flex-wrap:wrap;gap:4px">
            ${a.periods.map((p,pi)=>`<span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%${ce?` <span onclick="delPeriod(${idx},${pi})" style="cursor:pointer;opacity:.5;font-size:10px" title="Delete period">✕</span>`:''}</span>`).join('')}
          </div>
          <div style="font-size:11px;color:#9ca3af;margin-top:6px">${totalWks} week${totalWks!==1?'s':''} total</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;flex-shrink:0">
          ${!a.committed&&ce?`<button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button>`:''}
          ${a.committed&&ce?`<button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${idx})">↩ Uncommit</button>`:''}
          ${ce?`<button class="btn danger sm" onclick="delAssignment(${idx})">🗑 Delete</button>`:''}
        </div>
      </div>
    </div>`;
  }

  return `
    <div style="margin-bottom:12px">
      <button class="btn sm" onclick="setTab('overview')">← Back to Overview</button>
    </div>

    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr">
        <span class="card-title">👤 ${name}</span>
        <div style="display:flex;gap:14px;font-size:12px;color:#6b7280;align-items:center;flex-wrap:wrap">
          <span>${info.team || '—'}</span>
          <span>${info.skillset || '—'}</span>
          <span>${info.level || '—'}</span>
          <span>${info.country || '—'}</span>
          ${getTeamlead(name)?`<span>TL: <strong>${getTeamlead(name)}</strong></span>`:''}
          ${getManager(name)?`<span>Mgr: <strong>${getManager(name)}</strong></span>`:''}
          <span style="color:#0f6e56;font-weight:600">${committed.length} committed</span>
          <span style="color:#b45309;font-weight:600">${planned.length} planned</span>
        </div>
      </div>
    </div>

    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title">📅 Full year allocation</span></div>
      <div class="card-body">${calTable}</div>
    </div>

    ${committed.length?`
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title" style="color:#0f6e56">✓ Committed assignments (${committed.length})</span></div>
      <div class="card-body">${committed.map(assignmentCard).join('')}</div>
    </div>`:''}

    ${planned.length?`
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title" style="color:#b45309">⏳ Planned assignments (${planned.length})</span></div>
      <div class="card-body">${planned.map(assignmentCard).join('')}</div>
    </div>`:''}

    ${!personAssignments.length?`<div class="card"><div class="empty"><span class="empty-icon">📋</span>No assignments for ${name} yet.</div></div>`:''}
  `;
}


function openProject(projId){
  state.selectedProject = projId;
  state.tab = 'project-detail';
  document.getElementById('tab-title').textContent = 'Project Detail';
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  render();
}

function saveTeamConfig(team, field, value){
  if(!state.teamConfig[team]) state.teamConfig[team] = { teamlead:'', manager:'' };
  state.teamConfig[team][field] = value;
  // Don't render immediately — onblur on the input will trigger render
  // This prevents the input from losing focus while typing
}

function getTeamlead(name){
  // Per-person override first
  const m = state.teamMembers.find(m=>m.name===name);
  if(m && m.teamlead) return m.teamlead;
  // Fall back to team default
  const team = m?.team || state.assignments.find(a=>a.name===name)?.team;
  return team && state.teamConfig[team] ? state.teamConfig[team].teamlead : '';
}

function getManager(name){
  const m = state.teamMembers.find(m=>m.name===name);
  if(m && m.manager) return m.manager;
  const team = m?.team || state.assignments.find(a=>a.name===name)?.team;
  return team && state.teamConfig[team] ? state.teamConfig[team].manager : '';
}


function openTeam(teamName){
  if(role()==='Team Member') return; // Team Members cannot see team views
  state.selectedTeam = teamName;
  state.tab = 'team-detail';
  document.getElementById('tab-title').textContent = teamName + ' Team';
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const idMap = {'Development':'nav-dev','Platform':'nav-plat','PMO':'nav-pmo'};
  if(idMap[teamName]) document.getElementById(idMap[teamName])?.classList.add('active');
  render();
}

function isPmProject(proj){
  // A project belongs to current user if their name matches the PM field
  return proj.projectManager && proj.projectManager.trim().toLowerCase() === userName().trim().toLowerCase();
}


function toggleEditProject(id){
  const panel = document.getElementById('edit-proj-panel');
  if(panel) panel.style.display = panel.style.display==='none' ? 'block' : 'none';
}

function saveProjectEdits(id){
  const name  = document.getElementById('ep-name')?.value.trim();
  const pm    = document.getElementById('ep-pm')?.value||'';
  const start = document.getElementById('ep-start')?.value;
  const end   = document.getElementById('ep-end')?.value;
  const desc  = document.getElementById('ep-desc')?.value.trim();
  if(!name) return;
  // Update assignments that referenced the old name
  const oldProj = state.projects.find(p=>p.id===id);
  if(oldProj && oldProj.name !== name){
    state.assignments.forEach(a=>{ if(a.workName===oldProj.name && a.type==='Project') a.workName=name; });
  }
  state.projects = state.projects.map(p=>p.id===id ? {...p,name,projectManager:pm,startDate:start,endDate:end,description:desc} : p);
  flashMsg('Project updated!', true);
  render();
}


function addResourceToProject(){
  const proj = state.projects.find(p => p.id === state.selectedProject);
  if(!proj) return;
  if(!state.prName.trim() || !state.prSkill.trim()){ flashMsg('Please fill in name and skillset.',false); return; }
  const period = {id:Date.now(), startWeek:state.prStart, endWeek:state.prEnd, allocationPercent:state.prPct};
  // Check if assignment already exists for this person+project
  const ei = state.assignments.findIndex(a =>
    a.name.trim().toLowerCase()===state.prName.trim().toLowerCase() &&
    a.type==='Project' &&
    a.workName===proj.name
  );
  if(ei>=0){
    state.assignments[ei].periods.push(period);
    state.assignments[ei].committed=false;
    state.assignments[ei].committedBy=null;
  } else {
    state.assignments.push({
      id:Date.now(), name:state.prName.trim(), team:state.prTeam, country:state.prCountry,
      skillset:state.prSkill.trim(), level:state.prLevel, type:'Project',
      workName:proj.name, projectId:proj.id,
      periods:[period], confirmed:false, confirmedBy:null, committed:false, committedBy:null
    });
  }
  state.prName=''; state.prSkill=''; state.prStart=1; state.prEnd=4; state.prPct=80;
  flashMsg('Resource added!', true);
}

function renderProjectDetail(){
  const proj = state.projects.find(p => p.id === state.selectedProject);
  if(!proj) return `<div class="card"><div class="empty"><span class="empty-icon">💼</span>Project not found.</div></div>`;
  const r = role();
  const ce = canEdit();

  // PM can only see their own projects
  if(r === 'Project Manager' && !isPmProject(proj)){
    return `<div class="card"><div class="empty"><span class="empty-icon">🔒</span>You do not have access to this project.</div></div>`;
  }

  const allAssignments = state.assignments.filter(a => a.workName === proj.name);
  const visAssignments = allAssignments.filter(a => {
    if(r === 'Team Member') return a.committed && a.name.trim().toLowerCase() === userName().trim().toLowerCase();
    return true;
  });

  const planned   = visAssignments.filter(a => !a.committed);
  const committed = visAssignments.filter(a =>  a.committed);

  // Name display rules:
  // PM: planned → skillset only, committed → full name
  // Others: always full name
  function displayName(a){
    if(r === 'Project Manager'){
      return a.committed ? a.name : `${a.skillset} (${a.level})`;
    }
    return a.name;
  }

  function resourceCard(a){
    const realIdx = state.assignments.indexOf(a);
    const showName = displayName(a);
    const totalWeeks = a.periods.reduce((s,p) => s+(p.endWeek-p.startWeek+1), 0);
    const avgAlloc   = a.periods.length ? Math.round(a.periods.reduce((s,p)=>s+p.allocationPercent,0)/a.periods.length) : 0;
    return `<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:10px">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px">
        <div style="flex:1">
          <div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:3px">${showName}</div>
          <div style="font-size:12px;color:#6b7280;margin-bottom:8px">${a.team} · ${a.country} · ${a.skillset} · ${a.level}</div>
          <div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px">
            ${a.periods.map(p=>`<span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%</span>`).join('')}
          </div>
          <div style="font-size:11px;color:#9ca3af">${totalWeeks} week${totalWeeks!==1?'s':''} · avg ${avgAlloc}% allocation</div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px;flex-shrink:0">
          ${a.committed
            ? `<span class="badge b-committed">✓ Committed</span><div style="font-size:10px;color:#9ca3af;margin-top:2px">by ${a.committedBy}</div>${ce?`<button class="btn danger sm" style="font-size:10px;padding:2px 6px;margin-top:4px" onclick="uncommitA(${realIdx})">↩ Uncommit</button>`:''}`
            : `<span class="badge b-plan">Planned</span>${ce?`<button class="btn primary sm" style="margin-top:6px" onclick="commitA(${realIdx})">🔒 Commit</button>`:'<div style="font-size:11px;color:#9ca3af;margin-top:4px">Awaiting commit</div>'}`
          }
          ${ce?`<button class="btn danger sm" onclick="delAssignment(${realIdx})">🗑 Remove</button>`:''}
        </div>
      </div>
    </div>`;
  }

  return `
    <div style="margin-bottom:12px">
      <button class="btn sm" onclick="setTab('projects')">← Back to Projects</button>
    </div>

    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr">
        <span class="card-title">💼 ${proj.name}</span>
        <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
          <span style="font-size:12px;color:#6b7280">📅 <strong>${proj.startDate||'—'}</strong> → <strong>${proj.endDate||'—'}</strong></span>
          ${proj.projectManager?`<span style="font-size:12px;color:#6b7280">PM: <strong>${proj.projectManager}</strong></span>`:''}
          <span style="color:#b45309;font-weight:600;font-size:12px">${planned.length} planned</span>
          <span style="color:#0f6e56;font-weight:600;font-size:12px">${committed.length} committed</span>
          ${ce?`<button class="btn sm" onclick="toggleEditProject(${proj.id})">✏ Edit details</button>`:''}
          ${ce?`<button class="btn danger sm" onclick="if(confirm('Delete project and all its planning?')){state.projects=state.projects.filter(p=>p.id!==${proj.id});state.assignments=state.assignments.filter(a=>a.workName!=='${proj.name}');setTab('projects')}">🗑 Delete project</button>`:''}
        </div>
        ${proj.description ? `<div style="padding:10px 18px;font-size:13px;color:#374151;border-top:1px solid #f3f4f6;background:#fafafa;line-height:1.6">${proj.description}</div>` : ''}
      </div>
      ${ce?`<div id="edit-proj-panel" style="display:none;padding:16px;border-top:1px solid #e5e7eb;background:#f9fafb">
        <div class="fgrid">
          <div class="fg"><label class="lbl">Project name</label><input class="inp" id="ep-name" value="${proj.name}" /></div>
          <div class="fg"><label class="lbl">Project manager</label>${peopleSelectOptional('ep-pm', proj.projectManager||'', 'document._epPm=this.value', '')}</div>
          <div class="fg"><label class="lbl">Start date</label><input class="inp" type="date" id="ep-start" value="${proj.startDate||''}" /></div>
          <div class="fg"><label class="lbl">End date</label><input class="inp" type="date" id="ep-end" value="${proj.endDate||''}" /></div>
          <div class="fg" style="grid-column:1/-1"><label class="lbl">Description <span style="color:#9ca3af;font-weight:400">(optional)</span></label><textarea class="inp" id="ep-desc" rows="2" placeholder="Goals, scope, notes…" style="resize:vertical">${proj.description||''}</textarea></div>
          <div class="fg" style="justify-content:flex-end;padding-top:4px">
            <button class="btn primary" onclick="saveProjectEdits(${proj.id})">Save changes</button>
          </div>
        </div>
      </div>`:''}
    </div>

    ${(ce || (role()==='Project Manager' && isPmProject(proj))) ? `
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title">＋ Add resource to project</span></div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:16px">
        <div class="fgrid">
          ${role()!=='Project Manager' ? `<div class="fg"><label class="lbl">Full name *</label>
            ${peopleSelect('inp-prName', state.prName, "onPersonInput(this.value,'pr');clearTimeout(_debounceTimer);render()", '')}
          </div>` : ''}
          <div class="fg"><label class="lbl">Team</label>
            <select class="sel" onchange="state.prTeam=this.value">
              <option${state.prTeam==='Development'?' selected':''}>Development</option>
              <option${state.prTeam==='Platform'?' selected':''}>Platform</option>
              <option${state.prTeam==='PMO'?' selected':''}>PMO</option>
            </select>
          </div>
          <div class="fg"><label class="lbl">Country</label>
            <select class="sel" onchange="state.prCountry=this.value">
              <option value="Sweden"${state.prCountry==='Sweden'?' selected':''}>Sweden</option>
              <option value="Poland"${state.prCountry==='Poland'?' selected':''}>Poland</option>
            </select>
          </div>
          <div class="fg"><label class="lbl">Skillset *</label>
            <input class="inp" placeholder="e.g. React, DevOps" value="${state.prSkill}"
              oninput="state.prSkill=this.value" />
          </div>
          <div class="fg"><label class="lbl">Level</label>
            <select class="sel" onchange="state.prLevel=this.value">
              <option${state.prLevel==='Junior'?' selected':''}>Junior</option>
              <option${state.prLevel==='Mid'?' selected':''}>Mid</option>
              <option${state.prLevel==='Senior'?' selected':''}>Senior</option>
            </select>
          </div>
        </div>
        <div class="frow">
          <div class="fg"><label class="lbl">From week</label>
            <input class="inp narrow" type="number" min="1" max="52" value="${state.prStart}" oninput="state.prStart=+this.value" />
          </div>
          <div class="arrow">→</div>
          <div class="fg"><label class="lbl">To week</label>
            <input class="inp narrow" type="number" min="1" max="52" value="${state.prEnd}" oninput="state.prEnd=+this.value" />
          </div>
          <div class="fg"><label class="lbl">Allocation</label>
            <div style="display:flex;align-items:center;gap:4px">
              <input class="inp narrow" type="number" min="0" max="200" value="${state.prPct}" oninput="state.prPct=+this.value" />
              <span style="font-size:13px;color:#6b7280">%</span>
            </div>
          </div>
          <div class="fg" style="justify-content:flex-end;padding-top:18px">
            <button class="btn primary" onclick="addResourceToProject()">＋ Add resource</button>
          </div>
        </div>
      </div>
    </div>` : ''}

    <div class="card">
      <div class="card-hdr">
        <span class="card-title">📅 Weekly allocation — all 52 weeks</span>
        <span class="card-sub">Green row = person total · White rows = per-period breakdown</span>
      </div>
      ${visAssignments.length === 0
        ? `<div class="empty"><span class="empty-icon">👤</span>No resources assigned yet. Add one above!</div>`
        : `<div class="tbl-wrap"><table>
          <thead><tr>
            <th>Resource</th><th>Skill</th><th>Level</th><th>Country</th><th>Status</th>
            ${WEEKS.map(w=>`<th class="wk">W${w}</th>`).join('')}
            ${ce?'<th style="position:sticky;right:0;background:#f9fafb"></th>':''}
          </tr></thead>
          <tbody>
            ${visAssignments.map(a => {
              const realIdx = state.assignments.indexOf(a);
              const rawName = a.name&&a.name.startsWith('__pm_planned__')?'No name':a.name; const showName = r==='Project Manager' ? (a.committed ? rawName : 'No name') : rawName;
              const statusBadge = a.committed
                ? `<span class="badge b-committed" style="white-space:nowrap">✓ Committed</span><div style="font-size:10px;color:#9ca3af;margin-bottom:4px">${a.committedBy}</div>${ce?`<button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${realIdx})">↩ Uncommit</button>`:''}`
                : `<span class="badge b-plan">Planned</span>${ce?`<div style="margin-top:4px"><button class="btn primary sm" onclick="commitA(${realIdx})">🔒 Commit</button></div>`:''}`;
              // Total allocation per week across all periods
              const weekCells = WEEKS.map(w => {
                const al = getAlloc(a, w);
                return `<td class="wk ${al>100?'ao':al===100?'af':al>0?'ap':''}" style="font-weight:700">${al>0?al+'%':'–'}</td>`;
              }).join('');
              // Period breakdown rows
              const periodRows = a.periods.map((p,pi) => {
                const pWeeks = WEEKS.map(w => {
                  const inRange = w>=p.startWeek && w<=p.endWeek;
                  return `<td class="wk" style="font-size:10px;${inRange?'background:rgba(29,158,117,0.08);color:#0f6e56':''}">${inRange?p.allocationPercent+'%':''}</td>`;
                }).join('');
                return `<tr style="background:#fafafa">
                  <td colspan="4" style="padding:4px 12px 4px 28px;font-size:11px;color:#6b7280">
                    Period ${pi+1}: W${p.startWeek}–${p.endWeek} · ${p.allocationPercent}%
                  </td>
                  <td style="padding:4px 12px">
                    ${ce?`<button class="btn danger sm" style="padding:2px 6px;font-size:10px" onclick="delPeriod(${realIdx},${pi})">🗑</button>`:''}
                  </td>
                  ${pWeeks}
                  ${ce?'<td style="position:sticky;right:0;background:#fafafa"></td>':''}
                </tr>`;
              }).join('');
              return `
                <tr style="background:#f0fdf8;border-top:2px solid #e5e7eb">
                  <td style="padding:10px 12px;font-size:13px;font-weight:700;white-space:nowrap">${showName}</td>
                  <td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.skillset}</td>
                  <td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.level}</td>
                  <td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.country}</td>
                  <td style="padding:10px 12px">${statusBadge}</td>
                  ${weekCells}
                  ${ce?`<td style="padding:10px 8px;white-space:nowrap;position:sticky;right:0;background:#f0fdf8;box-shadow:-2px 0 4px rgba(0,0,0,0.06)">
                    <button class="btn danger sm" onclick="delAssignment(${realIdx})">🗑 Remove</button>
                  </td>`:''}
                </tr>
                ${periodRows}
              `;
            }).join('')}
          </tbody></table></div>`
      }
    </div>
  `;
}

function startEditMember(id){
  if(!canEdit()) return;
  // Toggle off if already editing this person
  if(state.editingMemberId === id){ state.editingMemberId=null; render(); return; }
  // Find the person from teamMembers or from assignments
  const member = state.teamMembers.find(m=>m.id===id);
  if(!member) return;
  state.editingMemberId = id;
  state.emName = member.name;
  state.emCountry = member.country;
  state.emSkill = member.skillset;
  state.emLevel = member.level;
  state.emTeamlead = member.teamlead || '';
  state.emManager = member.manager || '';
  render();
}

function saveMemberEdit(id){
  if(!state.emName.trim()) return;
  const oldMember = state.teamMembers.find(m=>m.id===id);
  const oldName = oldMember?.name;
  state.teamMembers = state.teamMembers.map(m => m.id===id
    ? {...m, name:state.emName.trim(), country:state.emCountry, skillset:state.emSkill.trim(), level:state.emLevel, teamlead:state.emTeamlead||'', manager:state.emManager||''}
    : m
  );
  // Always sync assignments — name, skill, country and level
  const matchName = (state.emName.trim() || oldName||'').trim().toLowerCase();
  const oldNameLower = (oldName||'').trim().toLowerCase();
  state.assignments = state.assignments.map(a => {
    const aName = a.name.trim().toLowerCase();
    if(aName === oldNameLower || aName === matchName){
      return {...a,
        name: state.emName.trim() || a.name,
        country: state.emCountry || a.country,
        skillset: state.emSkill.trim() || a.skillset,
        level: state.emLevel || a.level
      };
    }
    return a;
  });
  state.editingMemberId = null;
  flashMsg('Member updated!', true);
}

function saveMemberEditFromInputs(id){
  // Read current values directly from DOM inputs (uncontrolled)
  const nameEl  = document.getElementById('em-name-'+id);
  const skillEl = document.getElementById('em-skill-'+id);
  const tlEl    = document.getElementById('em-tl-'+id);
  const mgrEl   = document.getElementById('em-mgr-'+id);
  const nameVal  = nameEl?.value.trim()  || state.emName.trim();
  const skillVal = skillEl?.value.trim() || state.emSkill.trim();
  const tlVal    = tlEl?.value.trim()    || '';
  const mgrVal   = mgrEl?.value.trim()   || '';
  if(!nameVal) return;
  state.emName    = nameVal;
  state.emSkill   = skillVal;
  state.emTeamlead = tlVal;
  state.emManager  = mgrVal;
  saveMemberEdit(id);
}

// ── Inbox ──────────────────────────────────────────────────────────────────

function addInboxItem(){
  if(!state.iTitle.trim()) return;
  state.inboxItems.push({
    id: Date.now(),
    title: state.iTitle.trim(),
    description: state.iDesc.trim(),
    priority: state.iPriority,
    status: 'new',
    createdBy: userName(),
    createdAt: new Date().toLocaleDateString('en-SE'),
    convertedTo: null,
  });
  state.iTitle=''; state.iDesc=''; state.iPriority='Medium';
  render();
}

function convertInboxItem(id, to){
  const item = state.inboxItems.find(i=>i.id===id);
  if(!item) return;

  if(to==='revert'){
    // Revert back to unclassified — remove created project if it exists
    if(item.convertedTo==='Project'){
      const proj = state.projects.find(p=>p.name===item.title);
      if(proj && confirm('This will also delete the project "'+item.title+'" that was created from this item. Continue?')){
        state.projects = state.projects.filter(p=>p.name!==item.title);
      } else if(!proj){
        // project already renamed/deleted, just revert
      } else {
        return; // user cancelled
      }
    }
    item.status='new'; item.convertedTo=null;
    flashMsg('Reverted to inbox.', true);
    render(); return;
  }

  if(to==='project'){
    // If previously was initiative, just reclassify
    if(item.convertedTo==='Internal Initiative'){
      state.projects.push({id:Date.now(), name:item.title, projectManager:'', startDate:'', endDate:'', description:item.description});
      item.convertedTo='Project';
      flashMsg('Changed to project!', true);
    } else {
      state.projects.push({id:Date.now(), name:item.title, projectManager:'', startDate:'', endDate:'', description:item.description});
      item.status='converted'; item.convertedTo='Project';
      flashMsg('Converted to project!', true);
    }
  } else if(to==='initiative'){
    // If previously was project, remove the auto-created project
    if(item.convertedTo==='Project'){
      const proj = state.projects.find(p=>p.name===item.title);
      if(proj) state.projects = state.projects.filter(p=>p.name!==item.title);
    }
    item.status='converted'; item.convertedTo='Internal Initiative';
    flashMsg('Changed to Internal Initiative!', true);
  }
  render();
}

function deleteInboxItem(id){
  state.inboxItems = state.inboxItems.filter(i=>i.id!==id);
  render();
}

function renderInbox(){
  if(!canEdit()) return `<div class="card"><div class="locked">🔒<span>Only Teamlead and Manager can manage the inbox.</span></div></div>`;
  const active = state.inboxItems.filter(i=>i.status==='new');
  const converted = state.inboxItems.filter(i=>i.status==='converted');
  return `
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title">📥 Inbox — New incoming items</span><span class="card-sub">Items not yet classified as project or initiative</span></div>
      <div class="card-body">
        <div class="ibox">
          <div class="sec-title">Add new item</div>
          <div class="fgrid">
            <div class="fg" style="grid-column:1/-1"><label class="lbl">Title *</label><input class="inp" placeholder="What is this about?" value="${state.iTitle}" oninput="state.iTitle=this.value" /></div>
            <div class="fg" style="grid-column:1/-1"><label class="lbl">Description <span style="color:#9ca3af;font-weight:400">(optional)</span></label><textarea class="inp" rows="2" placeholder="Context, background, who requested it…" oninput="state.iDesc=this.value" style="resize:vertical">${state.iDesc}</textarea></div>
            <div class="fg"><label class="lbl">Priority</label>
              <select class="sel" onchange="state.iPriority=this.value">
                <option${state.iPriority==='High'?' selected':''}>High</option>
                <option${state.iPriority==='Medium'?' selected':''}>Medium</option>
                <option${state.iPriority==='Low'?' selected':''}>Low</option>
              </select>
            </div>
            <div style="padding-top:18px"><button class="btn primary" onclick="addInboxItem()">＋ Add to inbox</button></div>
          </div>
        </div>
      </div>
    </div>

    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title">🔍 Needs classification (${active.length})</span></div>
      <div class="card-body">
        ${!active.length ? `<div class="empty" style="padding:20px 0"><span class="empty-icon">✅</span>Inbox is empty.</div>` :
          active.map(item=>`
          <div class="inbox-card priority-${item.priority.toLowerCase()}">
            <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px">
              <div style="flex:1">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
                  <span style="font-size:13px;font-weight:700;color:#111827">${item.title}</span>
                  <span style="background:${item.priority==='High'?'#fef2f2;color:#dc2626':item.priority==='Medium'?'#fffbeb;color:#b45309':'#f0fdf8;color:#0f6e56'};padding:1px 8px;border-radius:20px;font-size:10px;font-weight:700">${item.priority}</span>
                </div>
                ${item.description?`<div style="font-size:12px;color:#6b7280;margin-bottom:6px">${item.description}</div>`:''}
                <div style="font-size:11px;color:#9ca3af">Added by ${item.createdBy} · ${item.createdAt}</div>
              </div>
              <div style="display:flex;flex-direction:column;gap:6px;flex-shrink:0">
                <button class="btn primary sm" onclick="convertInboxItem(${item.id},'project')">→ Make project</button>
                <button class="btn sm" style="border-color:#a78bfa;color:#7c3aed" onclick="convertInboxItem(${item.id},'initiative')">→ Mark as initiative</button>
                <button class="btn danger sm" onclick="deleteInboxItem(${item.id})">🗑 Dismiss</button>
              </div>
            </div>
          </div>`).join('')
        }
      </div>
    </div>

    ${converted.length ? `<div class="card">
      <div class="card-hdr"><span class="card-title" style="color:#9ca3af">✓ Processed (${converted.length})</span></div>
      <div class="card-body">
        ${converted.map(item=>`
          <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid #f3f4f6;gap:12px">
            <div style="flex:1">
              <span style="font-size:13px;font-weight:600;color:#374151">${item.title}</span>
              <span style="margin-left:8px;background:${item.convertedTo==='Project'?'#e0f2fe;color:#075985':'#ede9fe;color:#6d28d9'};padding:1px 8px;border-radius:20px;font-size:11px;font-weight:700">→ ${item.convertedTo}</span>
              ${item.description?`<div style="font-size:11px;color:#9ca3af;margin-top:2px">${item.description}</div>`:''}
            </div>
            <div style="display:flex;gap:6px;flex-shrink:0;flex-wrap:wrap;justify-content:flex-end">
              ${item.convertedTo==='Project'
                ? `<button class="btn sm" style="font-size:11px;border-color:#a78bfa;color:#7c3aed" onclick="convertInboxItem(${item.id},'initiative')">→ Change to initiative</button>`
                : `<button class="btn sm" style="font-size:11px;border-color:#0ea5e9;color:#0369a1" onclick="convertInboxItem(${item.id},'project')">→ Change to project</button>`
              }
              <button class="btn sm" style="font-size:11px" onclick="convertInboxItem(${item.id},'revert')">↩ Revert to inbox</button>
              <button class="btn danger sm" onclick="deleteInboxItem(${item.id})">🗑</button>
            </div>
          </div>`).join('')}
      </div>
    </div>` : ''}
  `;
}

// ── Pipeline ──────────────────────────────────────────────────────────────────

function renderPipeline(){
  const r = role();
  const planned = state.assignments.filter(a => !a.committed);
  if(!planned.length) return `<div class="card"><div class="empty"><span class="empty-icon">⏳</span>No planned (uncommitted) assignments yet.</div></div>`;

  // Group by project/work
  const grouped = new Map();
  planned.forEach(a => {
    if(!grouped.has(a.workName)) grouped.set(a.workName, []);
    grouped.get(a.workName).push(a);
  });

  const ce = canEdit();

  return `
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr">
        <span class="card-title">⏳ Pipeline — Planned but not committed</span>
        <span class="card-sub">${planned.length} assignment${planned.length!==1?'s':''} across ${grouped.size} work item${grouped.size!==1?'s':''}</span>
      </div>
    </div>

    ${[...grouped.entries()].map(([workName, items])=>{
      const proj = state.projects.find(p=>p.name===workName);
      return `<div class="card" style="margin-bottom:12px">
        <div class="card-hdr">
          <span class="card-title">${items[0].type==='Project'?'💼':'🔧'} ${workName}</span>
          <div style="display:flex;gap:10px;align-items:center">
            <span class="badge b-type">${items[0].type}</span>
            ${proj?`<span style="font-size:11px;color:#9ca3af">📅 ${proj.startDate||'—'} → ${proj.endDate||'—'}</span>`:''}
            <span style="font-size:11px;color:#b45309;font-weight:600">${items.length} resource${items.length!==1?'s':''} planned</span>
          </div>
        </div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Resource</th><th>Team</th><th>Skill</th><th>Level</th><th>Country</th><th>Periods</th><th>Total weeks</th>${ce?'<th>Action</th>':''}</tr></thead>
          <tbody>${items.map(a=>{
            const idx = state.assignments.indexOf(a);
            const _rn = a.name&&a.name.startsWith('__pm_planned__')?'No name':a.name; const displayName = r==='Project Manager' ? (a.committed ? _rn : 'No name') : _rn;
            const totalWks = a.periods.reduce((s,p)=>s+(p.endWeek-p.startWeek+1),0);
            const periods = a.periods.map(p=>`<span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%</span>`).join(' ');
            return `<tr>
              <td style="background:${a.country==='Sweden'?'#dbeafe':'#fef3c7'}"><strong>${displayName}</strong></td>
              <td>${a.team}</td><td>${a.skillset}</td><td>${a.level}</td><td>${a.country}</td>
              <td>${periods}</td>
              <td><strong style="font-family:'DM Mono',monospace">${totalWks}</strong> week${totalWks!==1?'s':''}</td>
              ${ce?`<td style="white-space:nowrap">
                <button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button>
                <button class="btn danger sm" style="margin-top:4px" onclick="delAssignment(${idx})">🗑 Delete</button>
              </td>`:''}
            </tr>`;
          }).join('')}</tbody>
        </table></div>
      </div>`;
    }).join('')}
  `;
}


function addTeamMember(){
  // Read uncontrolled inputs from DOM
  const nameEl  = document.getElementById('inp-tmName');
  const skillEl = document.getElementById('inp-tmSkill');
  const tlEl    = document.getElementById('inp-tmTl');
  const mgrEl   = document.getElementById('inp-tmMgr');
  if(nameEl?.value.trim())  state.tmName   = nameEl.value.trim();
  if(skillEl?.value.trim()) state.tmSkill  = skillEl.value.trim();
  if(tlEl?.value.trim())    state.tmTeamlead = tlEl.value.trim();
  if(mgrEl?.value.trim())   state.tmManager  = mgrEl.value.trim();
  if(!state.tmName.trim() || !state.tmSkill.trim()) return;
  const teamName = state.selectedTeam;
  // avoid duplicates
  const exists = state.teamMembers.find(m => m.name.trim().toLowerCase()===state.tmName.trim().toLowerCase() && m.team===teamName);
  if(exists){ flashMsg('This person is already in the team.',false); return; }
  state.teamMembers.push({
    id: Date.now(),
    name: state.tmName.trim(),
    team: teamName,
    country: state.tmCountry,
    skillset: state.tmSkill.trim(),
    level: state.tmLevel,
    teamlead: state.tmTeamlead.trim(),
    manager: state.tmManager.trim(),
  });
  state.tmName=''; state.tmSkill=''; state.tmTeamlead=''; state.tmManager='';
  render();
}

function removeTeamMember(id){
  state.teamMembers = state.teamMembers.filter(m => m.id !== id);
  render();
}

function renderTeamDetail(){
  const teamName = state.selectedTeam;
  const r = role();
  const ce = canEdit();

  // Auto-register anyone appearing in planning for this team but not yet a team member
  visibleAssignments().filter(a => a.team === teamName).forEach(a => {
    const key = a.name.trim().toLowerCase();
    const alreadyMember = state.teamMembers.some(m => m.name.trim().toLowerCase()===key && m.team===teamName);
    if(!alreadyMember){
      state.teamMembers.push({
        id: Date.now() + Math.random(),
        name: a.name.trim(), team: teamName,
        country: a.country, skillset: a.skillset, level: a.level
      });
    }
  });

  // Build people map: all registered members for this team
  const peopleMap = new Map();
  state.teamMembers.filter(m => m.team === teamName).forEach(m => {
    const key = m.name.trim().toLowerCase();
    if(!peopleMap.has(key)) peopleMap.set(key, {
      id: m.id, name: m.name, team: m.team, country: m.country,
      skillset: m.skillset, level: m.level, assignments: [], registered: true
    });
  });
  // Attach assignments
  visibleAssignments().filter(a => a.team === teamName).forEach(a => {
    const key = a.name.trim().toLowerCase();
    if(peopleMap.has(key)){
      if(!peopleMap.get(key).assignments.find(x=>x.id===a.id))
        peopleMap.get(key).assignments.push(a);
    }
  });
  const people = [...peopleMap.values()];

  const totalCommitted = people.reduce((s,p)=>s+p.assignments.filter(a=>a.committed).length,0);
  const totalPlanned   = people.reduce((s,p)=>s+p.assignments.filter(a=>!a.committed).length,0);

  // Week allocation helpers
  function getPersonTotalAlloc(person, w){
    return person.assignments.reduce((s,a)=>s+getAlloc(a,w),0);
  }
  function wCls(t){ return t>100?'ao':t===100?'af':t>0?'ap':''; }

  // Per-assignment breakdown row under each person
  function assignmentRows(person){
    if(!person.assignments.length) return `<tr><td colspan="${5+52}" style="padding:6px 12px;font-size:11px;color:#9ca3af;font-style:italic">No assignments yet</td></tr>`;
    return person.assignments.map(a=>{
      const label = r==='Project Manager' && !a.committed ? '— Planned —' : a.workName;
      const idx = state.assignments.indexOf(a);
      const statusDot = a.committed
        ? `<span style="background:#d1fae5;color:#065f46;padding:1px 7px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap">✓ Committed</span>
           ${ce?`<button class="btn danger sm" style="font-size:10px;padding:1px 6px;margin-left:4px" onclick="uncommitA(${idx})">↩</button>`:''}`
        : `<span style="background:#fef3c7;color:#92400e;padding:1px 7px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap">⏳ Planned</span>
           ${ce?`<button class="btn primary sm" style="font-size:10px;padding:1px 6px;margin-left:4px" onclick="commitA(${idx})">🔒 Commit</button>`:''}`;
      const weekCells = WEEKS.map(w=>{
        const al = getAlloc(a,w);
        return `<td class="wk ${wCls(al)}" style="font-size:10px">${al>0?al+'%':''}</td>`;
      }).join('');
      return `<tr style="background:#fafafa">
        <td style="padding:5px 12px 5px 28px;font-size:12px;color:#374151;white-space:nowrap">${label}</td>
        <td style="padding:5px 12px;font-size:11px;color:#9ca3af">${a.type}</td>
        <td colspan="2" style="padding:5px 12px;white-space:nowrap">${statusDot}</td>
        <td style="padding:5px 12px"></td>
        ${weekCells}
        ${ce?`<td style="padding:5px 8px"></td>`:''}
      </tr>`;
    }).join('');
  }

  const memberRows = people.map(person=>{
    const totalAllocs = WEEKS.map(w=>getPersonTotalAlloc(person,w));
    const weekCells = WEEKS.map((w,i)=>{
      const t=totalAllocs[i];
      return `<td class="wk ${wCls(t)}" style="font-weight:700">${t>0?t+'%':'–'}</td>`;
    }).join('');
    const badge = person.registered
      ? `<span style="background:#e0f2fe;color:#075985;font-size:10px;font-weight:700;padding:1px 6px;border-radius:20px;margin-left:6px">Member</span>`
      : `<span style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;padding:1px 6px;border-radius:20px;margin-left:6px">Planning</span>`;
    const isEditing = ce && person.registered && state.editingMemberId === person.id;
    const editRow = isEditing ? `
      <tr style="background:#f0fdf8"><td colspan="${6+52}" style="padding:0">
        <div class="edit-member-panel">
          <div style="font-size:11px;font-weight:700;color:#0f6e56;margin-bottom:12px;text-transform:uppercase;letter-spacing:.05em">✏ Editing: ${person.name}</div>
          <div class="fgrid">
            <div class="fg"><label class="lbl">Full name</label>
              <input class="inp" id="em-name-${person.id}" placeholder="${state.emName}"
                oninput="state.emName=this.value" />
            </div>
            <div class="fg"><label class="lbl">Country</label>
              <select class="sel" onchange="state.emCountry=this.value">
                <option value="Sweden"${state.emCountry==='Sweden'?' selected':''}>Sweden</option>
                <option value="Poland"${state.emCountry==='Poland'?' selected':''}>Poland</option>
              </select>
            </div>
            <div class="fg"><label class="lbl">Skillset</label>
              <input class="inp" id="em-skill-${person.id}" placeholder="${state.emSkill}"
                oninput="state.emSkill=this.value" />
            </div>
            <div class="fg"><label class="lbl">Level</label>
              <select class="sel" onchange="state.emLevel=this.value">
                <option${state.emLevel==='Junior'?' selected':''}>Junior</option>
                <option${state.emLevel==='Mid'?' selected':''}>Mid</option>
                <option${state.emLevel==='Senior'?' selected':''}>Senior</option>
              </select>
            </div>
            <div class="fg"><label class="lbl">Teamlead <span style="color:#9ca3af;font-weight:400">(blank = team default: ${getTeamlead(person.name)||'none'})</span></label>
              <input class="inp" id="em-tl-${person.id}" placeholder="${state.emTeamlead||'inherit from team'}"
                list="people-list-optional" autocomplete="off"
                oninput="state.emTeamlead=this.value" /></div>
            <div class="fg"><label class="lbl">Manager <span style="color:#9ca3af;font-weight:400">(blank = team default: ${getManager(person.name)||'none'})</span></label>
              <input class="inp" id="em-mgr-${person.id}" placeholder="${state.emManager||'inherit from team'}"
                list="people-list-optional" autocomplete="off"
                oninput="state.emManager=this.value" /></div>
            <div style="display:flex;gap:8px;padding-top:4px;align-items:flex-end;grid-column:1/-1">
              <button class="btn primary sm" onclick="saveMemberEditFromInputs(${person.id})">✓ Save</button>
              <button class="btn sm" onclick="state.editingMemberId=null;render()">✕ Cancel</button>
              <button class="btn danger sm" onclick="removeTeamMember(${person.id})">🗑 Remove</button>
            </div>
          </div>
        </div>
      </td></tr>` : '';
    return `
      <tr style="background:#f0fdf8;border-top:2px solid #e5e7eb;${ce?'cursor:pointer':''}">
        <td style="padding:10px 12px;font-size:13px;font-weight:700;white-space:nowrap;${ce&&person.registered?'cursor:pointer;':''}" ${ce&&person.registered?`title="Click to edit" onclick="startEditMember(${person.id})"`:''}>
          ${person.name}${badge}
        </td>
        <td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.skillset}</td>
        <td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.level}</td>
        <td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.country}</td>
        <td style="padding:10px 12px;font-size:12px;color:#6b7280">
          ${getTeamlead(person.name)?`<div style="font-size:10px;color:#0f6e56;font-weight:700">TL: ${getTeamlead(person.name)}</div>`:''}
          ${getManager(person.name)?`<div style="font-size:10px;color:#185fa5;font-weight:700">Mgr: ${getManager(person.name)}</div>`:''}
          ${!getTeamlead(person.name)&&!getManager(person.name)?`<span style="color:#d1d5db;font-size:11px">–</span>`:''}
        </td>
        <td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.assignments.length} asgmt${person.assignments.length!==1?'s':''}</td>
        ${weekCells}
        <td style="padding:10px 8px"></td>
      </tr>
      ${editRow}
      ${assignmentRows(person)}
    `;
  }).join('');

  return `
    <div style="margin-bottom:12px">
      <button class="btn sm" onclick="setTab('overview')">← Back to Overview</button>
    </div>

    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr">
        <span class="card-title">👥 ${teamName} Team</span>
        <div style="display:flex;gap:16px;font-size:12px;color:#6b7280">
          <span><strong style="color:#111827">${people.length}</strong> member${people.length!==1?'s':''}</span>
          <span><strong style="color:#0f6e56">${totalCommitted}</strong> committed</span>
          <span><strong style="color:#b45309">${totalPlanned}</strong> planned</span>
        </div>
      </div>
      <div style="padding:14px 18px;border-top:1px solid #f3f4f6;display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div>
          <div class="org-label">Team Lead</div>
          ${ce ? `<div style="display:flex;align-items:center;gap:8px">
            ${peopleSelectOptional('tc-tl-'+teamName, state.teamConfig[teamName]?.teamlead||'', 'saveTeamConfig(\''+teamName+'\',\'teamlead\',this.value)', 'flex:1;max-width:220px')}
            ${state.teamConfig[teamName]?.teamlead?`<span style="font-size:11px;color:#0f6e56;font-weight:600">✓ ${state.teamConfig[teamName].teamlead}</span>`:'<span style="font-size:11px;color:#9ca3af">Not assigned</span>'}
          </div>` : `<div class="org-person">${state.teamConfig[teamName]?.teamlead||'<span style="color:#9ca3af;font-weight:400">Not assigned</span>'}</div>`}
        </div>
        <div>
          <div class="org-label">Manager</div>
          ${ce ? `<div style="display:flex;align-items:center;gap:8px">
            ${peopleSelectOptional('tc-mgr-'+teamName, state.teamConfig[teamName]?.manager||'', 'saveTeamConfig(\''+teamName+'\',\'manager\',this.value)', 'flex:1;max-width:220px')}
            ${state.teamConfig[teamName]?.manager?`<span style="font-size:11px;color:#185fa5;font-weight:600">✓ ${state.teamConfig[teamName].manager}</span>`:'<span style="font-size:11px;color:#9ca3af">Not assigned</span>'}
          </div>` : `<div class="org-person">${state.teamConfig[teamName]?.manager||'<span style="color:#9ca3af;font-weight:400">Not assigned</span>'}</div>`}
        </div>
      </div>
    </div>

    ${ce ? `
    <div class="card" style="margin-bottom:16px">
      <div class="card-hdr"><span class="card-title">＋ Add team member</span></div>
      <div class="card-body" style="display:flex;flex-direction:column;gap:14px">
        <!-- Row 1: core fields -->
        <div style="display:grid;grid-template-columns:1.5fr 1fr 1.5fr 1fr auto;gap:12px;align-items:flex-end">
          <div class="fg"><label class="lbl">Full name *</label>
            <input class="inp" id="inp-tmName" list="people-list" placeholder="Type or pick a name…"
              autocomplete="off" oninput="onPersonInput(this.value,'tm')" />
          </div>
          <div class="fg"><label class="lbl">Country</label>
            <select class="sel" onchange="state.tmCountry=this.value">
              <option value="Sweden"${state.tmCountry==='Sweden'?' selected':''}>Sweden</option>
              <option value="Poland"${state.tmCountry==='Poland'?' selected':''}>Poland</option>
            </select>
          </div>
          <div class="fg"><label class="lbl">Skillset *</label>
            <input class="inp" id="inp-tmSkill" placeholder="e.g. React, DevOps" oninput="state.tmSkill=this.value" onkeydown="if(event.key==='Enter')addTeamMember()" />
          </div>
          <div class="fg"><label class="lbl">Level</label>
            <select class="sel" onchange="state.tmLevel=this.value">
              <option${state.tmLevel==='Junior'?' selected':''}>Junior</option>
              <option${state.tmLevel==='Mid'?' selected':''}>Mid</option>
              <option${state.tmLevel==='Senior'?' selected':''}>Senior</option>
            </select>
          </div>
          <button class="btn primary" onclick="addTeamMember()" style="white-space:nowrap">＋ Add member</button>
        </div>
        <!-- Row 2: optional reporting overrides -->
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:12px;background:#f9fafb;border-radius:8px;border:1px solid #f3f4f6">
          <div class="fg">
            <label class="lbl">Teamlead override <span style="color:#9ca3af;font-weight:400">— leave blank to use team default: <strong>${state.teamConfig[state.selectedTeam]?.teamlead||'none set'}</strong></span></label>
            <input class="inp" id="inp-tmTl" list="people-list-optional" placeholder="Blank = inherit from team"
              autocomplete="off" oninput="state.tmTeamlead=this.value" />
          </div>
          <div class="fg">
            <label class="lbl">Manager override <span style="color:#9ca3af;font-weight:400">— leave blank to use team default: <strong>${state.teamConfig[state.selectedTeam]?.manager||'none set'}</strong></span></label>
            <input class="inp" id="inp-tmMgr" list="people-list-optional" placeholder="Blank = inherit from team"
              autocomplete="off" oninput="state.tmManager=this.value" />
          </div>
        </div>
      </div>
    </div>` : ''}

    <div class="card">
      <div class="card-hdr">
        <span class="card-title">📅 Weekly allocation — all 52 weeks</span>
        <span class="card-sub">Green row = person total · White rows = per-assignment breakdown</span>
      </div>
      ${people.length === 0
        ? `<div class="empty"><span class="empty-icon">👥</span>No team members in ${teamName} yet.</div>`
        : `<div class="tbl-wrap"><table>
            <thead><tr>
              <th>Name</th><th>Skill</th><th>Level</th><th>Country</th><th>Reporting</th><th>Assignments</th>
              ${WEEKS.map(w=>`<th class="wk">W${w}</th>`).join('')}
              ${ce?'<th></th>':''}
            </tr></thead>
            <tbody>${memberRows}</tbody>
          </table></div>`
      }
    </div>`;
}



function getAllPeople(){
  const map = new Map();
  state.teamMembers.forEach(m => {
    map.set(m.name.trim().toLowerCase(), {name:m.name, team:m.team, country:m.country, skillset:m.skillset, level:m.level});
  });
  state.assignments.forEach(a => {
    const key = a.name.trim().toLowerCase();
    if(!map.has(key)) map.set(key, {name:a.name, team:a.team, country:a.country, skillset:a.skillset, level:a.level});
  });
  return [...map.values()].sort((a,b)=>a.name.localeCompare(b.name));
}

function onPersonInput(val, target){
  if(target==='add')      state.aName  = val;
  else if(target==='tm')  state.tmName = val;
  else if(target==='pr')  state.prName = val;

  // Check for exact match to auto-fill related fields
  const match = getAllPeople().find(p => p.name.trim().toLowerCase() === val.trim().toLowerCase());
  if(match){
    if(target==='add'){
      state.aName=match.name; state.aTeam=match.team; state.aCountry=match.country; state.aSkill=match.skillset; state.aLevel=match.level;
      clearTimeout(_debounceTimer); render();
    } else if(target==='tm'){
      state.tmName=match.name; state.tmCountry=match.country; state.tmSkill=match.skillset; state.tmLevel=match.level;
      // Don't render — preserve input focus. Fields update on Add click.
    } else if(target==='pr'){
      state.prName=match.name; state.prTeam=match.team; state.prCountry=match.country; state.prSkill=match.skillset; state.prLevel=match.level;
      clearTimeout(_debounceTimer); render();
    }
  } else if(target==='add'){
    debounce(render, 600);
  }
  // For 'tm' and 'pr' with no match: just store, no render
}


function peopleSelect(id, value, onchangeCode, extraStyle, emptyLabel){
  const ph = emptyLabel || 'Type or select person…';
  // Use defaultValue via JS after render so typing isn't interrupted
  setTimeout(()=>{ const el=document.getElementById(id); if(el&&document.activeElement!==el) el.value=value||''; },0);
  return `<input class="sel" id="${id}" list="people-list"
    placeholder="${ph}" style="${extraStyle||''}" autocomplete="off"
    oninput="${onchangeCode}"
    onblur="render()" />`;
}

function peopleSelectOptional(id, value, onchangeCode, extraStyle){
  setTimeout(()=>{ const el=document.getElementById(id); if(el&&document.activeElement!==el) el.value=value||''; },0);
  return `<input class="sel" id="${id}" list="people-list-optional"
    placeholder="Type name or leave blank…" style="${extraStyle||''}" autocomplete="off"
    oninput="${onchangeCode}"
    onblur="debounce(render,200)" />`;
}

function updateSidebarForRole(){
  const tm = role()==='Team Member';
  const el = document.getElementById('sidebar-teams-section');
  if(el) el.style.display = tm ? 'none' : '';
}


function toggleDark(){
  const isDark = document.body.classList.toggle('dark');
  localStorage.setItem('rp_dark', isDark ? '1' : '0');
  document.getElementById('dark-btn').textContent = isDark ? '☀ Light' : '🌙 Dark';
}

// Apply saved dark mode preference on load
(function(){
  if(localStorage.getItem('rp_dark')==='1'){
    document.body.classList.add('dark');
  }
})();

function buildDatalist(){
  const people = getAllPeople();
  const opts = people.map(p => `<option value="${p.name}">${p.skillset} · ${p.team}</option>`).join('');
  // Main list (required name fields)
  let dl = document.getElementById('people-list');
  if(!dl){ dl = document.createElement('datalist'); dl.id='people-list'; document.body.appendChild(dl); }
  dl.innerHTML = opts;
  // Optional list (manager/teamlead fields — same content)
  let dl2 = document.getElementById('people-list-optional');
  if(!dl2){ dl2 = document.createElement('datalist'); dl2.id='people-list-optional'; document.body.appendChild(dl2); }
  dl2.innerHTML = opts;
}


function setTab(t){
  state.tab=t;
  document.querySelectorAll('.nav-item').forEach(el=>{
    el.classList.toggle('active', el.textContent.trim().replace(/^./,'').trim()===({dashboard:'Dashboard',overview:'Overview','person-detail':'Person Detail',projects:'Projects','project-detail':'Project Detail',services:'Base Services','team-detail':'Team Detail',inbox:'Inbox',pipeline:'Pipeline',add:'Planning mode'}[t]));
  });
  document.getElementById('tab-title').textContent = {dashboard:'Dashboard',overview:'Overview','person-detail':'Person Detail',projects:'Projects','project-detail':'Project Detail',services:'Base Services','team-detail':'Team Detail',inbox:'Inbox',pipeline:'Pipeline',add:'Planning mode'}[t];
  render();
}

function flashMsg(text,ok){
  state.msg={text,ok};
  render();
  setTimeout(()=>{state.msg=null;render();},3000);
}

function render(){
  buildDatalist();
  updateSidebarForRole();
  const darkBtn = document.getElementById('dark-btn');
  if(darkBtn) darkBtn.textContent = document.body.classList.contains('dark') ? '☀ Light' : '🌙 Dark';
  document.getElementById('foot').textContent = `${state.assignments.length} assignment${state.assignments.length!==1?'s':''} · ${state.projects.length} project${state.projects.length!==1?'s':''}`;  saveData();
  const el = document.getElementById('content');
  const t = state.tab;
  if(t==='dashboard'){
    el.innerHTML = renderDashboard();
    // Inject team alloc card separately to avoid template literal issues
    const tac = document.getElementById('team-alloc-card');
    if(tac){
      const now2 = new Date();
      const soy = new Date(now2.getFullYear(),0,1);
      const cw2 = Math.min(Math.max(Math.ceil(((now2-soy)/86400000+soy.getDay()+1)/7),1),52);
      const AR = state.dashAllocRange||8;
      const aw = WEEKS.slice(cw2-1,cw2-1+AR).filter(w=>w<=52);
      const tm = {'Development':[],'Platform':[],'PMO':[]};
      const tn = t2=>{if(!t2)return null;const s=t2.trim().toLowerCase();return s==='development'?'Development':s==='platform'?'Platform':s==='pmo'?'PMO':null;};
      state.teamMembers.forEach(m=>{const t2=tn(m.team);if(t2&&!tm[t2].includes(m.name))tm[t2].push(m.name);});
      state.assignments.forEach(a=>{const t2=tn(a.team);if(t2&&!tm[t2].includes(a.name))tm[t2].push(a.name);});
      tac.innerHTML = buildTeamAllocCard(tm, aw, cw2, state);
    }
  }
  else if(t==='overview') el.innerHTML = renderOverview();
  else if(t==='person-detail') el.innerHTML = renderPersonDetail();
  else if(t==='planning') el.innerHTML = renderPlanning();
  else if(t==='projects') el.innerHTML = renderProjects();
  else if(t==='project-detail') el.innerHTML = renderProjectDetail();
  else if(t==='services') el.innerHTML = renderServices();
  else if(t==='team-detail') el.innerHTML = renderTeamDetail();
  else if(t==='inbox') el.innerHTML = renderInbox();
  else if(t==='pipeline') el.innerHTML = renderPipeline();
  else if(t==='add'){ if(role()==='Project Manager') state.addType='Project'; el.innerHTML = renderAdd(); }
}

function renderOverview(){
  const allPeople = getPeople();
  const over   = allPeople.filter(p=>WEEKS.some(w=>getTotalAlloc(p.name,w)>100)).length;
  const conf   = state.assignments.filter(a=>a.committed).length;
  const planned= state.assignments.filter(a=>!a.committed).length;
  const r = role();
  const ce = canEdit();

  // Collect unique skills, levels, teams, assignments for filter dropdowns
  const teams       = [...new Set(allPeople.map(p=>p.team))].sort();
  const skills      = [...new Set(allPeople.map(p=>p.skillset).filter(Boolean))].sort();
  const levels      = [...new Set(allPeople.map(p=>p.level).filter(Boolean))].sort();
  const assignments = [...new Set(state.assignments.map(a=>a.workName).filter(Boolean))].sort();

  // Apply filters
  let people = allPeople.filter(p => {
    const n = state.fName.trim().toLowerCase();
    const s = state.fSkill.trim().toLowerCase();
    if(state.fTeam && p.team !== state.fTeam) return false;
    if(n && !p.name.toLowerCase().includes(n)) return false;
    if(s && !p.skillset.toLowerCase().includes(s)) return false;
    if(state.fLevel && p.level !== state.fLevel) return false;
    if(state.fAssignment){
      const hasAssignment = visibleAssignments().some(a=>a.name===p.name && a.workName===state.fAssignment);
      if(!hasAssignment) return false;
    }
    if(state.fStatus){
      const personA = visibleAssignments().filter(a=>a.name===p.name);
      if(state.fStatus==='planned' && !personA.some(a=>!a.committed)) return false;
      if(state.fStatus==='committed' && !personA.some(a=>a.committed)) return false;
      if(state.fStatus==='overbooked' && !WEEKS.some(w=>getTotalAlloc(p.name,w)>100)) return false;
    }
    return true;
  });

  const activeFilters = [state.fTeam,state.fName,state.fSkill,state.fLevel,state.fAssignment,state.fStatus].filter(Boolean).length;

  // Filter bar
  const filterBar = `<div class="filter-bar">
    ${(function(){
      if(role()==='Team Member') return '';
      const opts = getAllPeople().map(p=>'<option value="'+p.name+'"'+(state.fName===p.name?' selected':'')+'>'+p.name+'</option>').join('');
      return '<div><label class="filter-lbl">Name</label><select class="filter-inp" style="width:150px" onchange="state.fName=this.value;render()"><option value="">All people</option>'+opts+'</select></div>';
    })()}
    <div><label class="filter-lbl">Team</label><select class="filter-inp" onchange="state.fTeam=this.value;render()">
      <option value="">All teams</option>${teams.map(t=>`<option value="${t}"${state.fTeam===t?' selected':''}>${t}</option>`).join('')}
    </select></div>
    <div><label class="filter-lbl">Skill</label><input class="filter-inp" style="width:120px" placeholder="Filter skill…" value="${state.fSkill}" oninput="setFilter('fSkill',this.value)" /></div>
    <div><label class="filter-lbl">Level</label><select class="filter-inp" onchange="state.fLevel=this.value;render()">
      <option value="">All levels</option>${levels.map(l=>`<option value="${l}"${state.fLevel===l?' selected':''}>${l}</option>`).join('')}
    </select></div>
    <div><label class="filter-lbl">Assignment</label><select class="filter-inp" style="max-width:180px" onchange="state.fAssignment=this.value;render()">
      <option value="">All assignments</option>${assignments.map(a=>`<option value="${a}"${state.fAssignment===a?' selected':''}>${a}</option>`).join('')}
    </select></div>
    <div><label class="filter-lbl">Status</label><select class="filter-inp" onchange="state.fStatus=this.value;render()">
      <option value="">All</option>
      <option value="planned"${state.fStatus==='planned'?' selected':''}>⏳ Planned only</option>
      <option value="committed"${state.fStatus==='committed'?' selected':''}>✓ Has committed</option>
      <option value="overbooked"${state.fStatus==='overbooked'?' selected':''}>🔴 Overbooked</option>
    </select></div>
    ${activeFilters>0 ? `<button class="btn sm" style="margin-top:16px" onclick="state.fTeam='';state.fName='';state.fSkill='';state.fLevel='';state.fAssignment='';state.fStatus='';render()">✕ Clear filters (${activeFilters})</button>` : ''}
  </div>`;

  let rows='';
  if(!people.length){
    rows=`<div class="empty"><span class="empty-icon">👥</span>${allPeople.length?'No results match your filters.':'No allocations yet. Go to Planning mode to get started.'}</div>`;
  } else {
    rows=`<div class="tbl-wrap"><table>
      <thead><tr>
        <th>Person</th><th>Team</th><th>Country</th><th>Skill</th><th>Level</th><th>Status</th>
        ${WEEKS.map(w=>`<th class="wk">W${w}</th>`).join('')}
      </tr></thead>
      <tbody>${people.map(p=>{
        const dn = (r==='Project Manager' && visibleAssignments().some(a=>a.name===p.name&&!a.committed)) ? '— Planned resource —' : p.name;
        const personA = visibleAssignments().filter(a=>a.name===p.name);
        const hasCom = personA.some(a=>a.committed);
        const hasPlan = personA.some(a=>!a.committed);
        const statusBadge = hasCom && hasPlan
          ? `<span class="status-mixed">Mixed</span>`
          : hasCom
            ? `<span class="status-committed">✓ Committed</span>`
            : hasPlan
              ? `<span class="status-planned">⏳ Planned</span>`
              : `<span style="font-size:10px;color:#d1d5db">–</span>`;
        const clickAttr = ce ? `onclick="openPersonDetail('${p.name.replace(/'/g,"\'")}')"`  : '';
        return `<tr class="${ce?'person-row-click':''}" ${clickAttr}>
          <td style="background:${cBg(p.country)}"><strong>${dn}</strong>${ce?`<span style="font-size:10px;color:#9ca3af;margin-left:6px">→</span>`:''}</td>
          <td><span style="cursor:pointer;color:#1D9E75;text-decoration:underline" onclick="event.stopPropagation();openTeam('${p.team}')">${p.team}</span></td>
          <td>${p.country}</td><td>${p.skillset}</td><td>${p.level}</td>
          <td>${statusBadge}</td>
          ${WEEKS.map(w=>{const t=getTotalAlloc(p.name,w);return`<td class="wk ${wClass(t)}">${t>0?t+'%':'–'}</td>`;}).join('')}
        </tr>`;
      }).join('')}</tbody></table></div>`;
  }

  return `
    <div class="metrics">
      <div class="metric"><div class="metric-lbl">People</div><div class="metric-val">${allPeople.length}</div></div>
      <div class="metric"><div class="metric-lbl">Assignments</div><div class="metric-val">${state.assignments.length}</div></div>
      <div class="metric"><div class="metric-lbl">Committed</div><div class="metric-val green">${conf}</div></div>
      <div class="metric"><div class="metric-lbl">Planned</div><div class="metric-val" style="color:#b45309">${planned}</div></div>
    </div>
    <div class="card">
      <div class="card-hdr">
        <span class="card-title">📅 Total allocation per person</span>
        <span class="card-sub">${activeFilters>0?`${people.length} of ${allPeople.length} shown`:'All 52 weeks · click a person for details'}</span>
      </div>
      ${filterBar}
      ${rows}
    </div>`;
}

function renderPlanning(){
  const vas=visibleAssignments();
  if(!vas.length) return `<div class="card"><div class="empty"><span class="empty-icon">📋</span>No planning entries yet.</div></div>`;
  const ce=canEdit();
  return `<div class="card">
    <div class="card-hdr"><span class="card-title">👥 Detailed planning</span><span class="card-sub">All 52 weeks · Set week/allocation in Planning mode before using "Add period"</span></div>
    <div class="tbl-wrap"><table>
      <thead><tr><th>Resource</th><th>Team</th><th>Country</th><th>Skill</th><th>Level</th><th>Type</th><th>Work</th>${WEEKS.map(w=>`<th class="wk">W${w}</th>`).join('')}<th>Status</th><th style="min-width:190px">Periods</th></tr></thead>
      <tbody>${vas.map((a,ai)=>{
        const dn = a.name&&a.name.startsWith('__pm_planned__') ? (a.committed?a.name:'No name') : (!a.committed&&role()==='Project Manager'?'— Planned resource —':a.name);
        const wks=WEEKS.map(w=>{const al=getAlloc(a,w);return`<td class="wk ${wClass(al)}">${al>0?al+'%':'–'}</td>`;}).join('');
        const periods=a.periods.map((p,pi)=>`
          <div style="display:flex;align-items:center;gap:4px;margin-bottom:4px">
            <span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%</span>
            ${ce?`<button class="btn danger sm" onclick="delPeriod(${ai},${pi})">🗑</button>`:''}
          </div>`).join('');
        const actions=ce?`
          <div style="display:flex;flex-direction:column;gap:4px;margin-top:6px">
            <button class="btn sm" onclick="addPeriod(${ai})">+ Add period</button>
            <button class="btn danger sm" onclick="delAssignment(${ai})">🗑 Delete</button>
          </div>`:'';
        const status = a.committed
          ? `<span class="badge b-committed">✓ Committed</span><div style="font-size:10px;color:#fff;background:#0f6e56;display:inline-block;padding:1px 6px;border-radius:4px;margin-top:3px;font-weight:600">${a.committedBy}</div>${ce?`<div style="margin-top:4px"><button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${ai})">↩ Uncommit</button></div>`:''}`
          : `<span class="badge b-plan">Planned</span>${ce ? `<div style="margin-top:6px"><button class="btn primary sm" onclick="commitA(${ai})">🔒 Commit resource</button></div>` : '<div style="font-size:11px;color:#9ca3af;margin-top:4px">Awaiting commit</div>'}`;
        return `<tr>
          <td style="background:${cBg(a.country)}"><strong>${dn}</strong><div style="font-size:11px;color:#9ca3af">${a.country}·${a.skillset}·${a.level}</div></td>
          <td>${a.team}</td><td>${a.country}</td><td>${a.skillset}</td><td>${a.level}</td>
          <td><span class="badge b-type">${a.type}</span></td><td>${a.workName}</td>
          ${wks}<td>${status}</td><td>${periods}${actions}</td></tr>`;
      }).join('')}</tbody></table></div></div>`;
}

function renderProjects(){
  const ce=canEdit();
  const r = role();

  // PM only sees their own projects
  const visibleProjects = r === 'Project Manager'
    ? state.projects.filter(p => isPmProject(p))
    : state.projects;

  function fmtDate(d){ if(!d) return '—'; const dt=new Date(d); return dt.toLocaleDateString('en-SE',{day:'2-digit',month:'short',year:'numeric'}); }

  const list = visibleProjects.length
    ? visibleProjects.map(p => {
        const allA = state.assignments.filter(a=>a.workName===p.name);
        const vis = r==='Team Member'
          ? allA.filter(a=>a.committed && a.name.trim().toLowerCase()===userName().trim().toLowerCase())
          : allA;
        const comm = vis.filter(a=>a.committed).length;
        const plan = r!=='Team Member' ? vis.filter(a=>!a.committed).length : 0;
        return `<div class="project-row" style="cursor:pointer;transition:box-shadow 0.15s" onmouseover="this.style.boxShadow='0 2px 8px rgba(0,0,0,0.08)'" onmouseout="this.style.boxShadow=''" onclick="openProject(${p.id})">
          <div style="flex:1">
            <div style="font-size:13px;font-weight:700;color:#111827;margin-bottom:2px">${p.name} <span style="font-size:11px;color:#9ca3af;font-weight:400">→ click to view</span></div>
            <div style="font-size:11px;color:#9ca3af;margin-top:2px">
              📅 ${fmtDate(p.startDate)} → ${fmtDate(p.endDate)}
              ${p.projectManager?` &nbsp;·&nbsp; PM: <strong style="color:#374151">${p.projectManager}</strong>`:''}
            </div>
            ${p.description?`<div style="font-size:11px;color:#6b7280;margin-top:3px;white-space:normal;max-width:400px;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical">${p.description}</div>`:''}
          </div>
          <div style="display:flex;gap:8px;align-items:center">
            ${comm>0?`<span style="background:#d1fae5;color:#065f46;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">✓ ${comm} committed</span>`:''}
            ${plan>0?`<span style="background:#fef3c7;color:#92400e;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">⏳ ${plan} planned</span>`:''}
            ${!comm&&!plan?`<span style="color:#d1d5db;font-size:11px">No resources yet</span>`:''}
            ${ce?`<button class="btn danger sm" style="margin-left:4px" onclick="event.stopPropagation();deleteProject(${p.id})">🗑</button>`:''}
          </div>
        </div>`;
      }).join('')
    : r === 'Project Manager'
      ? `<div class="empty" style="padding:24px 0"><span class="empty-icon">💼</span>No projects assigned to you yet.</div>`
      : `<div class="empty" style="padding:24px 0"><span class="empty-icon">💼</span>No projects yet.</div>`;
  return `<div class="card">
    <div class="card-hdr"><span class="card-title">💼 Projects</span></div>
    <div class="card-body">
      ${ce?`<div class="ibox">
        <div class="sec-title">Add project</div>
        <div class="frow">
          <div class="fg"><label class="lbl">Project name</label><input class="inp" placeholder="e.g. Platform Renewal" value="${state.pName}" oninput="state.pName=this.value" onkeydown="if(event.key==='Enter')addProject()" /></div>
          <div class="fg"><label class="lbl">Project manager</label>${peopleSelectOptional('add-pm', state.pPm, 'state.pPm=this.value', '')}</div>
          <div class="fg"><label class="lbl">Start date</label><input class="inp" type="date" value="${state.pStart}" oninput="state.pStart=this.value" /></div>
          <div class="fg"><label class="lbl">End date</label><input class="inp" type="date" value="${state.pEnd}" oninput="state.pEnd=this.value" /></div>
          <div class="fg" style="grid-column:1/-1"><label class="lbl">Description <span style="color:#9ca3af;font-weight:400">(optional)</span></label><textarea class="inp" rows="2" placeholder="Goals, scope, notes…" style="resize:vertical" oninput="state.pDesc=this.value">${state.pDesc}</textarea></div>
          <div style="padding-top:4px"><button class="btn primary" onclick="addProject()">＋ Add project</button></div>
        </div>
      </div>`:''}
      ${list}
    </div></div>`;
}

function renderServices(){
  const ce=canEdit();
  const groups=TEAMS.map(t=>{
    const svcs=state.baseServices.filter(s=>s.team===t);
    return `<div class="svc-grp-lbl">${t}</div>${svcs.length?svcs.map((s,i)=>`
      <div class="svc-row"><span>${s.name}</span>
        ${ce?`<div style="display:flex;gap:6px">
          <button class="btn sm" onclick="editSvc('${s.name}')">✏ Edit</button>
          <button class="btn danger sm" onclick="delSvc('${s.name}')">🗑</button>
        </div>`:''}
      </div>`).join(''):`<div style="font-size:12px;color:#9ca3af;padding-bottom:6px">No services</div>`}`;
  }).join('');
  return `<div class="card">
    <div class="card-hdr"><span class="card-title">🔧 Base Services</span></div>
    <div class="card-body">
      ${ce?`<div class="ibox">
        <div class="sec-title">Add base service</div>
        <div class="frow">
          <div class="fg"><label class="lbl">Team</label><select class="sel" style="width:140px" onchange="state.sTeam=this.value">${TEAMS.map(t=>`<option${state.sTeam===t?' selected':''}>${t}</option>`).join('')}</select></div>
          <div class="fg" style="flex:1"><label class="lbl">Service name</label><input class="inp" placeholder="e.g. API Support" value="${state.sName}" oninput="state.sName=this.value" onkeydown="if(event.key==='Enter')addSvc()" /></div>
          <button class="btn primary" onclick="addSvc()">＋ Add</button>
        </div>
      </div>`:''}
      <div style="max-height:400px;overflow-y:auto;border:1px solid #e5e7eb;border-radius:10px;padding:14px;background:#fafafa;">
        ${groups}
      </div>
    </div></div>`;
}

function buildPersonCalendar(name, selectedStart, selectedEnd){
  // Get all existing allocations for this person across all assignments
  const personAssignments = state.assignments.filter(a =>
    a.name.trim().toLowerCase() === name.trim().toLowerCase()
  );

  if(!personAssignments.length && !name){
    return '';
  }

  // Build week data: total existing alloc + highlight selected range
  const weekData = WEEKS.map(w => {
    const existing = personAssignments.reduce((s, a) => s + getAlloc(a, w), 0);
    const inNewRange = w >= selectedStart && w <= selectedEnd;
    const projectedTotal = existing + (inNewRange ? state.aPct : 0);
    return { w, existing, inNewRange, projectedTotal };
  });

  const hasAny = weekData.some(d => d.existing > 0 || d.inNewRange);

  // Build compact bar rows — group weeks into blocks of 13 (4 rows of 13 = 52)
  const CHUNK = 13;
  const rows = [];
  for(let i = 0; i < 52; i += CHUNK){
    rows.push(weekData.slice(i, i + CHUNK));
  }

  function cellStyle(d){
    if(d.inNewRange){
      if(d.projectedTotal > 100) return 'background:#fecaca;color:#b91c1c;font-weight:700;outline:2px solid #ef4444;outline-offset:-2px';
      if(d.projectedTotal === 100) return 'background:#6ee7b7;color:#065f46;font-weight:700;outline:2px solid #10b981;outline-offset:-2px';
      return 'background:#bfdbfe;color:#1e40af;font-weight:700;outline:2px solid #3b82f6;outline-offset:-2px';
    }
    if(d.existing > 100) return 'background:#fecaca;color:#b91c1c;font-weight:700';
    if(d.existing === 100) return 'background:#d1fae5;color:#065f46;font-weight:700';
    if(d.existing > 0) return 'background:#fef3c7;color:#92400e;font-weight:600';
    return 'background:#f9fafb;color:#d1d5db';
  }

  const tableRows = rows.map(chunk => `
    <tr>
      ${chunk.map(d => `<td style="text-align:center;padding:4px 2px;font-size:10px;font-family:'DM Mono',monospace;border-radius:3px;min-width:34px;${cellStyle(d)}">
        <div style="font-size:9px;color:inherit;opacity:.7;line-height:1">W${d.w}</div>
        <div style="line-height:1.3">${d.projectedTotal > 0 ? d.projectedTotal+'%' : '–'}</div>
      </td>`).join('')}
    </tr>`).join('');

  const overbooked = weekData.filter(d => d.inNewRange && d.projectedTotal > 100).length;
  const warn = overbooked ? `<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:6px;padding:6px 12px;font-size:12px;color:#b91c1c;margin-top:8px">⚠ ${overbooked} week${overbooked!==1?'s':''} would be overbooked with this new period</div>` : '';

  return `<div class="add-sec">
    <div class="sec-title" style="display:flex;align-items:center;justify-content:space-between">
      <span>📅 ${name} — current &amp; planned allocation</span>
      <div style="display:flex;gap:12px;font-size:10px;font-weight:600">
        <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#fef3c7;display:inline-block"></span>Partially booked</span>
        <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#d1fae5;display:inline-block"></span>Fully booked</span>
        <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#fecaca;display:inline-block"></span>Overbooked</span>
        <span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#bfdbfe;display:inline-block"></span>New period</span>
      </div>
    </div>
    ${!hasAny
      ? `<div style="padding:12px;background:#f9fafb;border-radius:8px;font-size:12px;color:#9ca3af;text-align:center">✓ No existing allocations — all weeks are free${selectedStart&&selectedEnd?'. New period highlighted in blue.':''}</div>`
      : `<div style="overflow-x:auto"><table style="border-collapse:separate;border-spacing:2px;min-width:100%">${tableRows}</table></div>`
    }
    ${warn}
    ${personAssignments.length ? `<div style="margin-top:8px;font-size:11px;color:#9ca3af">${personAssignments.length} existing assignment${personAssignments.length!==1?'s':''}: ${[...new Set(personAssignments.map(a=>a.workName))].join(', ')}</div>` : ''}
  </div>`;
}


function renderAdd(){
  const r = role();
  const isPM = r==='Project Manager';
  if(!canPlan()) return `<div class="card"><div class="locked">🔒<span>Only Teamlead, Manager and Project Manager can add planning.</span></div></div>`;
  const filtSvcs=state.baseServices.filter(s=>s.team===state.aTeam);
  // PM only sees their own projects
  const availableProjects = isPM ? pmProjects() : state.projects;

  const workField = state.addType==='Project'
    ?`<div class="fg" style="max-width:300px"><label class="lbl">Project *</label><select class="sel" onchange="state.aProjId=this.value">
        <option value="">Select project…</option>
        ${availableProjects.map(p=>`<option value="${p.id}"${state.aProjId==p.id?' selected':''}>${p.name}</option>`).join('')}
      </select>${!availableProjects.length?`<div class="hint">${isPM?'No projects assigned to you as PM yet.':'No projects yet — add one in the Projects tab.'}</div>`:''}</div>`
    :state.addType==='Base Service'
    ?`<div class="fg" style="max-width:300px"><label class="lbl">Base service *</label><select class="sel" onchange="state.aService=this.value">
        <option value="">Select service…</option>
        ${filtSvcs.map(s=>`<option value="${s.name}"${state.aService===s.name?' selected':''}>${s.name}</option>`).join('')}
      </select></div>`
    :`<div class="fg" style="max-width:300px"><label class="lbl">Work name *</label><input class="inp" placeholder="Name of work item" value="${state.aWork}" oninput="state.aWork=this.value" /></div>`;

  // Build person calendar if a name is entered
  const calendarSection = (!isPM && state.aName.trim()) ? buildPersonCalendar(state.aName.trim(), state.aStart, state.aEnd) : '';

  return `<div class="card">
    <div class="card-hdr"><span class="card-title">📋 Planning mode — add entry</span></div>
    <div class="card-body" style="display:flex;flex-direction:column;gap:24px">

      <div class="add-sec">
        <div class="sec-title">Assignment type</div>
        <div class="ttabs">${(isPM ? ['Project'] : TYPES).map(t=>`<div class="ttab${state.addType===t?' on':''}" onclick="state.addType='${t}';state.aProjId='';state.aService='';state.aWork='';render()">${t}</div>`).join('')}</div>
        ${isPM?`<div class="hint" style="margin-top:6px">As Project Manager you can plan resources for your own projects.</div>`:''}
      </div>

      <div class="add-sec">
        <div class="sec-title">Resource details</div>
        <div class="fgrid">
          ${!isPM ? `<div class="fg"><label class="lbl">Full name *</label>
            ${peopleSelect('inp-aName', state.aName, "onPersonInput(this.value,'add');clearTimeout(_debounceTimer);render()", '', 'Select person…')}
          </div>` : ''}
          <div class="fg"><label class="lbl">Team</label><select class="sel" onchange="state.aTeam=this.value;state.aService='';render()">${TEAMS.map(t=>`<option${state.aTeam===t?' selected':''}>${t}</option>`).join('')}</select></div>
          <div class="fg"><label class="lbl">Country</label><select class="sel" onchange="state.aCountry=this.value"><option value="Sweden"${state.aCountry==='Sweden'?' selected':''}>Sweden</option><option value="Poland"${state.aCountry==='Poland'?' selected':''}>Poland</option></select></div>
          <div class="fg"><label class="lbl">Skillset ${isPM?'*':'*'}</label><input class="inp" placeholder="e.g. React, DevOps" value="${state.aSkill}" oninput="state.aSkill=this.value" /></div>
          <div class="fg"><label class="lbl">Level</label><select class="sel" onchange="state.aLevel=this.value"><option${state.aLevel==='Junior'?' selected':''}>Junior</option><option${state.aLevel==='Mid'?' selected':''}>Mid</option><option${state.aLevel==='Senior'?' selected':''}>Senior</option></select></div>
        </div>
      </div>

      ${calendarSection}

      <div class="add-sec">
        <div class="sec-title">Assignment</div>
        ${workField}
      </div>

      <div class="add-sec">
        <div class="sec-title">Period &amp; allocation</div>
        <div class="frow">
          <div class="fg"><label class="lbl">From week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.aStart}" onchange="state.aStart=+this.value;render()" /></div>
          <div class="arrow">→</div>
          <div class="fg"><label class="lbl">To week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.aEnd}" onchange="state.aEnd=+this.value;render()" /></div>
          <div class="fg"><label class="lbl">Allocation</label><div style="display:flex;align-items:center;gap:4px"><input class="inp narrow" type="number" min="0" max="200" value="${state.aPct}" oninput="state.aPct=+this.value" /><span style="font-size:13px;color:#6b7280">%</span></div></div>
        </div>
        <div class="hint">💡 These fields are also used by "Add period" in Detailed Planning.</div>
      </div>

      <div>
        <button class="btn primary block" onclick="addAssignment()">✓ Add planning entry</button>
        ${state.msg?`<div class="msg ${state.msg.ok?'ok':'err'}">${state.msg.text}</div>`:'<div class="msg"></div>'}
      </div>
    </div></div>`;
}

// ── Actions ──────────────────────────────────────────────────────────────────

function deleteProject(id){
  if(!confirm('Delete this project? Planning entries linked to it will not be deleted.')) return;
  state.projects = state.projects.filter(p => p.id !== id);
  render();
}

function addProject(){
  if(!state.pName.trim()) return;
  state.projects.push({id:Date.now(),name:state.pName.trim(),projectManager:state.pPm.trim(),startDate:state.pStart,endDate:state.pEnd,description:state.pDesc.trim()});
  state.pName=''; state.pPm=''; state.pStart=''; state.pEnd=''; state.pDesc='';
  render();
}

function addSvc(){
  if(!state.sName.trim()) return;
  state.baseServices.push({name:state.sName.trim(),team:state.sTeam});
  state.sName='';
  render();
}

function editSvc(name){
  const n=prompt('New name:',name);
  if(!n?.trim()) return;
  state.baseServices=state.baseServices.map(s=>s.name===name?{...s,name:n.trim()}:s);
  render();
}

function delSvc(name){
  state.baseServices=state.baseServices.filter(s=>s.name!==name);
  render();
}

function getWorkName(){
  if(state.addType==='Project'){
    const p=state.projects.find(p=>p.id==state.aProjId);
    return p?p.name:'';
  }
  if(state.addType==='Base Service') return state.aService;
  return state.aWork;
}

function addAssignment(){
  if(!canPlan()){ flashMsg('You do not have permission to add planning.',false); return; }
  const isPM2 = role()==='Project Manager';
  if(!isPM2 && !state.aName.trim()){ flashMsg('Please fill in the name.',false); return; }
  if(!state.aSkill.trim()){ flashMsg('Please fill in the skillset.',false); return; }
  if(state.addType==='Project'&&!state.aProjId){ flashMsg('Please select a project.',false); return; }
  if(state.addType==='Base Service'&&!state.aService){ flashMsg('Please select a base service.',false); return; }
  if((state.addType==='Charge On'||state.addType==='Internal Initiative')&&!state.aWork.trim()){ flashMsg('Please enter a work name.',false); return; }
  const wn=getWorkName();
  const period={id:Date.now(),startWeek:state.aStart,endWeek:state.aEnd,allocationPercent:state.aPct};
  const ei=state.assignments.findIndex(a=>a.name.trim().toLowerCase()===state.aName.trim().toLowerCase()&&a.type===state.addType&&a.workName.trim().toLowerCase()===wn.trim().toLowerCase());
  if(ei>=0){
    state.assignments[ei].periods.push(period);
    state.assignments[ei].confirmed=false;
    state.assignments[ei].confirmedBy=null;
    state.assignments[ei].committed=false;
    state.assignments[ei].committedBy=null;
  } else {
    const assignName = (role()==='Project Manager') ? ('__pm_planned__'+Date.now()) : state.aName.trim();
    state.assignments.push({id:Date.now(),name:assignName,team:state.aTeam,country:state.aCountry,skillset:state.aSkill.trim(),level:state.aLevel,type:state.addType,workName:wn,projectId:state.addType==='Project'?+state.aProjId:null,periods:[period],confirmed:false,confirmedBy:null,committed:false,committedBy:null,pmPlanned:role()==='Project Manager'});
  }
  flashMsg('Entry added!',true);
  state.aName=''; state.aSkill=''; state.aProjId=''; state.aService=''; state.aWork='';
  state.aStart=1; state.aEnd=3; state.aPct=80;
}

function addPeriod(ai){
  const period={id:Date.now(),startWeek:state.aStart,endWeek:state.aEnd,allocationPercent:state.aPct};
  state.assignments[ai].periods.push(period);
  state.assignments[ai].confirmed=false;
  state.assignments[ai].confirmedBy=null;
  state.assignments[ai].committed=false;
  state.assignments[ai].committedBy=null;
  render();
}

function uncommitA(ai){
  state.assignments[ai].committed = false;
  state.assignments[ai].committedBy = null;
  state.assignments[ai].confirmed = false;
  state.assignments[ai].confirmedBy = null;
  render();
}

function commitA(ai){
  state.assignments[ai].committed = true;
  state.assignments[ai].committedBy = userName();
  // also mark confirmed for backward compat
  state.assignments[ai].confirmed = true;
  state.assignments[ai].confirmedBy = userName();
  render();
}

function delAssignment(ai){
  state.assignments.splice(ai,1);
  render();
}

function delPeriod(ai,pi){
  // pi can be either a period id (from planning table) or index (from project detail)
  // If pi is a number index (project detail uses index directly)
  if(typeof pi === 'number' && pi < state.assignments[ai].periods.length){
    state.assignments[ai].periods.splice(pi,1);
  } else {
    state.assignments[ai].periods = state.assignments[ai].periods.filter(p=>p.id!==pi);
  }
  state.assignments[ai].confirmed=false;
  state.assignments[ai].confirmedBy=null;
  state.assignments[ai].committed=false;
  state.assignments[ai].committedBy=null;
  render();
}


function autoRegisterAllTeamMembers(){
  // Ensure everyone with an assignment is registered as a team member
  state.assignments.forEach(a => {
    if(!a.name || !a.team) return;
    const key = a.name.trim().toLowerCase();
    const team = a.team.trim();
    const alreadyMember = state.teamMembers.some(
      m => m.name.trim().toLowerCase()===key && m.team===team
    );
    if(!alreadyMember){
      state.teamMembers.push({
        id: Date.now() + Math.random(),
        name: a.name.trim(), team: team,
        country: a.country||'Sweden',
        skillset: a.skillset||'',
        level: a.level||'Junior'
      });
    }
  });
}

// Auto-register all team members from assignments on startup
autoRegisterAllTeamMembers();

// Restore saved name & role
if(saved?.userName) document.getElementById('user-name').value = saved.userName;
if(saved?.role) document.getElementById('role-sel').value = saved.role;
render();
