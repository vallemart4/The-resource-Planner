// ─────────────────────────────────────────────────────────────────────────────
// Resource Planner — script.js (clean)
// ─────────────────────────────────────────────────────────────────────────────

const WEEKS = Array.from({length:52}, (_,i) => i+1);
const TEAMS = ['Development','Platform','PMO'];
const TYPES = ['Project','Base Service','Charge On','Internal Initiative'];

const _now = new Date();
const _soy = new Date(_now.getFullYear(), 0, 1);
const CURRENT_WEEK = Math.min(Math.max(Math.ceil(
  (((_now - _soy) / 86400000) + _soy.getDay() + 1) / 7
), 1), 52);

const DEFAULT_SERVICES = [
  {name:'Integration Support',   team:'Platform'},
  {name:'Infrastructure Support', team:'Platform'},
  {name:'Cloud Operations',       team:'Platform'},
  {name:'Application Support',    team:'Development'},
  {name:'Bug Fixing',             team:'Development'},
  {name:'Maintenance',            team:'Development'},
  {name:'PMO Governance',         team:'PMO'},
  {name:'Reporting',              team:'PMO'},
];

const API_URL = 'https://resource-planner-api-sgit-theresourceplanner.apps.openshift-dev.stenacloud.com';

async function loadFromApi(){
  try {
    const res = await fetch(`${API_URL}/api/data`);
    if(!res.ok) throw new Error('API error');
    return await res.json();
  } catch(e) {
    console.warn('API unavailable, using localStorage:', e.message);
    return null;
  }
}

async function saveData(){
  const data = {
    projects:    state.projects,
    assignments: state.assignments,
    baseServices:state.baseServices,
    teamMembers: state.teamMembers,
    teamConfig:  state.teamConfig,
    inboxItems:  state.inboxItems,
    userName: document.getElementById('user-name')?.value || '',
    role:     document.getElementById('role-sel')?.value  || 'Teamlead',
  };
  try { localStorage.setItem('rp_data', JSON.stringify(data)); } catch(e){}
  try {
    await fetch(`${API_URL}/api/data`, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(data),
    });
  } catch(e) { console.warn('Could not save to API:', e.message); }
}

function exportData(){
  const data = localStorage.getItem('rp_data') || JSON.stringify({
    projects:state.projects, assignments:state.assignments,
    baseServices:state.baseServices, teamMembers:state.teamMembers,
    teamConfig:state.teamConfig, inboxItems:state.inboxItems,
  });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([data], {type:'application/json'}));
  a.download = 'ResourcePlanner_backup_' + new Date().toISOString().split('T')[0] + '.json';
  a.click();
}

function importData(){
  const input = document.createElement('input');
  input.type = 'file'; input.accept = '.json';
  input.onchange = e => {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        localStorage.setItem('rp_data', JSON.stringify(data));
        alert('Data imported! The page will now reload.');
        location.reload();
      } catch(err) { alert('Could not read file. Make sure it is a valid backup.'); }
    };
    reader.readAsText(file);
  };
  input.click();
}

function clearData(){
  if(!confirm('Clear all data? This cannot be undone.')) return;
  localStorage.removeItem('rp_data'); location.reload();
}

function importExcel(){
  const input = document.createElement('input');
  input.type = 'file'; input.accept = '.xlsx,.xls';
  input.onchange = e => {
    const file = e.target.files[0]; if(!file) return;
    if(!window.XLSX){ alert('Excel library not loaded. Please refresh the page and try again.'); return; }
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, {type:'binary'});
        const sheetName = wb.SheetNames.includes('Resource Allocation') ? 'Resource Allocation' : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, raw:true});
        const header = rows[2] || [];
        const weekMap = {};
        for(let c = 9; c <= 60; c++){
          const v = header[c];
          if(v !== null && v !== undefined){ const n = parseInt(v); if(n >= 1 && n <= 52) weekMap[c] = n; }
        }
        const dataRows = rows.slice(3);
        const newProjects = [], newAssignments = [], newMembers = [];
        const seenProjects = new Set(state.projects.map(p => p.name.toLowerCase()));
        const seenMembers  = new Set(state.teamMembers.map(m => m.name.toLowerCase()));
        let importedCount = 0;
        dataRows.forEach(row => {
          if(!row || !row[5]) return;
          const name       = ((row[0]||'').toString().trim()).replace(/^"+|"+$/g,'').trim();
          const role       = (row[1]||'').toString().trim();
          const area       = (row[2]||'').toString().trim();
          const tl         = (row[3]||'').toString().trim();
          const status     = (row[4]||'').toString().trim().toLowerCase();
          const assignment = (row[5]||'').toString().trim();
          const typeRaw    = (row[6]||'').toString().trim().toLowerCase();
          if(!assignment) return;
          let appType = 'Project';
          if(typeRaw.includes('base service')) appType = 'Base Service';
          else if(typeRaw.includes('charge on')) appType = 'Charge On';
          else if(typeRaw.includes('initiative') && !typeRaw.includes('project')) appType = 'Internal Initiative';
          if(appType === 'Project' && !seenProjects.has(assignment.toLowerCase())){
            seenProjects.add(assignment.toLowerCase());
            newProjects.push({id:Date.now()+Math.random(), name:assignment, projectManager:'', startDate:'', endDate:'', description:''});
          }
          const nameKey = name.toLowerCase();
          if(name && !nameKey.startsWith('nn') && !seenMembers.has(nameKey)){
            seenMembers.add(nameKey);
            let team = 'Development';
            const a = area.toLowerCase();
            if(a.includes('pmo')||a==='pmo') team='PMO'; else if(a.includes('platform')) team='Platform';
            newMembers.push({id:Date.now()+Math.random(), name, team, country:'Sweden', skillset:role||'', level:'Mid', teamlead:tl, manager:tl});
          }
          const weekAllocs = [];
          for(const [colStr, week] of Object.entries(weekMap)){
            const col = parseInt(colStr); if(col >= row.length) continue;
            const raw = row[col]; if(raw === null || raw === undefined || raw === '') continue;
            let pct = 0;
            if(typeof raw === 'number') pct = Math.round(raw * 100);
            else { const s = raw.toString().replace('%','').trim(); const n = parseFloat(s); if(!isNaN(n)) pct = n > 1 ? Math.round(n) : Math.round(n*100); }
            if(pct > 0) weekAllocs.push({week, pct});
          }
          if(!weekAllocs.length) return;
          weekAllocs.sort((a,b) => a.week - b.week);
          const periods = [];
          let pStart = weekAllocs[0].week, pPct = weekAllocs[0].pct, pPrev = weekAllocs[0].week;
          for(let i = 1; i < weekAllocs.length; i++){
            const {week, pct} = weekAllocs[i];
            if(week <= pPrev+2 && pct === pPct){ pPrev = week; }
            else { periods.push({id:Date.now()+Math.random(), startWeek:pStart, endWeek:pPrev, allocationPercent:pPct}); pStart=week; pPct=pct; pPrev=week; }
          }
          periods.push({id:Date.now()+Math.random(), startWeek:pStart, endWeek:pPrev, allocationPercent:pPct});
          const committed = status==='commited'||status==='committed';
          let team = 'Development';
          const a = area.toLowerCase();
          if(a.includes('pmo')) team='PMO'; else if(a.includes('platform')) team='Platform';
          newAssignments.push({id:Date.now()+Math.random(), name:name||'NN', team, country:'Sweden', skillset:role||'', level:'Mid',
            type:appType, workName:assignment, projectId:null, periods, committed,
            committedBy:committed?(tl||'Excel import'):null, confirmed:committed, confirmedBy:committed?(tl||'Excel import'):null});
          importedCount++;
        });
        newProjects.forEach(p => state.projects.push(p));
        newMembers.forEach(m => state.teamMembers.push(m));
        newAssignments.forEach(a => state.assignments.push(a));
        saveData();
        flashMsg(`✓ Imported ${importedCount} assignments, ${newProjects.length} projects, ${newMembers.length} team members`, true);
        render();
      } catch(err) { console.error('Excel import error:', err); alert('Could not read Excel file: ' + err.message); }
    };
    reader.readAsBinaryString(file);
  };
  input.click();
}

function loadSaved(){
  try { const r=localStorage.getItem('rp_data'); if(r) return JSON.parse(r); } catch(e){}
  return null;
}
const saved = loadSaved();

let state = {
  tab: 'dashboard', selectedProject: null, selectedTeam: null, selectedPerson: null,
  projects:     saved?.projects     || [],
  assignments:  saved?.assignments  || [],
  baseServices: saved?.baseServices || DEFAULT_SERVICES,
  teamMembers:  saved?.teamMembers  || [],
  teamConfig:   saved?.teamConfig   || {
    Development: {teamlead:'', manager:''}, Platform: {teamlead:'', manager:''}, PMO: {teamlead:'', manager:''},
  },
  inboxItems: saved?.inboxItems || [],
  addType: 'Project', aName:'', aTeam:'Development', aCountry:'Sweden', aSkill:'', aLevel:'Junior',
  aProjId:'', aService:'', aWork:'', aStart:1, aEnd:3, aPct:80,
  pName:'', pPm:'', pStart:'', pEnd:'', pDesc:'',
  sName:'', sTeam:'Development', sTargetPct:0,
  prName:'', prTeam:'Development', prCountry:'Sweden', prSkill:'', prLevel:'Junior', prStart:1, prEnd:4, prPct:80,
  tmName:'', tmCountry:'Sweden', tmSkill:'', tmLevel:'Junior', tmTeamlead:'', tmManager:'',
  editingMemberId: null, emName:'', emCountry:'Sweden', emSkill:'', emLevel:'Junior', emTeamlead:'', emManager:'', emTeam:'Development',
  editingAssignmentId: null, eaStart: 1, eaEnd: 52, eaPct: 50,
  iTitle:'', iDesc:'', iPriority:'Medium',
  fTeam:'', fSkill:'', fLevel:'', fName:'', fAssignment:'', fStatus:'',
  teamFilterNames: new Set(),
  dashAllocRange: 8, showAllWeeks: false, weekOffset: 0, debtSectionOpen: false,
  msg: null,
};

function role()    { return document.getElementById('role-sel').value; }
function userName(){ return document.getElementById('user-name').value; }
function canEdit() { return role()==='Teamlead' || role()==='Manager'; }
function canPlan() { return role()==='Teamlead' || role()==='Manager' || role()==='Project Manager'; }
function pmProjects(){ return state.projects.filter(p => isPmProject(p)); }
function isPmProject(proj){ return proj.projectManager && proj.projectManager.trim().toLowerCase() === userName().trim().toLowerCase(); }

function visibleAssignments(){
  const r = role(), un = userName().trim().toLowerCase();
  return state.assignments.filter(a => { if(r === 'Team Member') return a.committed && a.name.trim().toLowerCase() === un; return true; });
}
function getPeople(){
  const map = new Map();
  visibleAssignments().forEach(a => { if(!map.has(a.name.toLowerCase())) map.set(a.name.toLowerCase(), a); });
  return [...map.values()];
}
function getAlloc(a, w){ return a.periods.filter(p => w>=p.startWeek && w<=p.endWeek).reduce((s,p) => s+p.allocationPercent, 0); }
function getEffectiveAlloc(a, w){
  const raw = getAlloc(a, w); if(!raw || a.type !== 'Base Service') return raw;
  const otherAlloc = visibleAssignments().filter(x => x!==a && x.name.toLowerCase()===a.name.toLowerCase() && x.committed).reduce((s,x) => s+getAlloc(x,w), 0);
  return Math.max(0, Math.min(raw, 100 - otherAlloc));
}
function getTotalAlloc(name, w){ return visibleAssignments().filter(a => a.name.toLowerCase()===name.toLowerCase() && a.committed).reduce((s,a) => s+getEffectiveAlloc(a,w), 0); }
function getTeamlead(name){
  const m = state.teamMembers.find(m => m.name===name); if(m && m.teamlead) return m.teamlead;
  const team = m?.team || state.assignments.find(a => a.name===name)?.team;
  return team && state.teamConfig[team] ? state.teamConfig[team].teamlead : '';
}
function getManager(name){
  const m = state.teamMembers.find(m => m.name===name); if(m && m.manager) return m.manager;
  const team = m?.team || state.assignments.find(a => a.name===name)?.team;
  return team && state.teamConfig[team] ? state.teamConfig[team].manager : '';
}
function getAllPeople(){
  const map = new Map();
  state.teamMembers.forEach(m => map.set(m.name.trim().toLowerCase(), {name:m.name, team:m.team, country:m.country, skillset:m.skillset, level:m.level}));
  state.assignments.forEach(a => { const key = a.name.trim().toLowerCase(); if(!map.has(key)) map.set(key, {name:a.name, team:a.team, country:a.country, skillset:a.skillset, level:a.level}); });
  return [...map.values()].sort((a,b) => a.name.localeCompare(b.name));
}
function getSvcAlloc(svcName, w){ return state.assignments.filter(a => a.type==='Base Service' && a.workName===svcName).reduce((s,a) => s+getEffectiveAlloc(a,w), 0); }
function calcSvcDebt(svc){
  const target = svc.targetPct || 0; if(!target) return {debtPct:0, debtWeeks:0};
  let debtPct = 0; for(let w=1; w<CURRENT_WEEK; w++) debtPct += Math.max(0, target - getSvcAlloc(svc.name, w));
  return {debtPct: Math.round(debtPct), debtWeeks: Math.round((debtPct/100)*10)/10};
}
function calcTeamDebt(team){
  const svcs = state.baseServices.filter(s => s.team===team && s.targetPct>0);
  if(!svcs.length) return {debtPct:0, debtWeeks:0, svcs:[]};
  let total = 0;
  const details = svcs.map(s => { const d=calcSvcDebt(s); total+=d.debtPct; return {name:s.name, target:s.targetPct, ...d}; });
  return {debtPct:Math.round(total), debtWeeks:Math.round((total/100)*10)/10, svcs:details};
}
function setSvcTarget(name, val){
  state.baseServices = state.baseServices.map(s => s.name===name ? {...s, targetPct: Math.max(0, Math.min(200, +val||0))} : s);
  saveData(); render();
}
function calcTeamAllocForWeek(teamMemberMap, team, w){
  const members = teamMemberMap[team] || []; if(!members.length) return {avg:0, fullyBooked:0, free:0, over:0, total:0};
  const allocs = members.map(name => state.assignments.filter(a => a.name===name && a.committed).reduce((s,a) => s+getAlloc(a,w), 0));
  return {avg: Math.round(allocs.reduce((s,v)=>s+v,0)/members.length), fullyBooked: allocs.filter(v=>v>=100).length, free: allocs.filter(v=>v===0).length, over: allocs.filter(v=>v>100).length, total: members.length};
}
function buildTeamMemberMap(){
  const map = {Development:[], Platform:[], PMO:[]};
  const norm = t => { if(!t) return null; const s=t.trim().toLowerCase(); return s==='development'?'Development':s==='platform'?'Platform':s==='pmo'?'PMO':null; };
  state.teamMembers.forEach(m => { const t=norm(m.team); if(t && !map[t].includes(m.name)) map[t].push(m.name); });
  state.assignments.forEach(a => { const t=norm(a.team); if(t && !map[t].includes(a.name)) map[t].push(a.name); });
  return map;
}

let _debounceTimer = null;
function debounce(fn, ms){ clearTimeout(_debounceTimer); _debounceTimer = setTimeout(fn, ms); }
function setFilter(key, val){ state[key] = val; debounce(render, 300); }
function wClass(t){ return t>100?'ao': t===100?'af': t>0?'ap': ''; }
function cBg(c)   { return c==='Sweden'?'#dbeafe':'#fef3c7'; }

function visibleWeeks(){
  if(state.showAllWeeks) return WEEKS;
  const start = Math.max(0, Math.min(CURRENT_WEEK - 1 + state.weekOffset, 40));
  return WEEKS.slice(start, start + 12);
}

function wkHdr(w){
  const isCurrent = w === CURRENT_WEEK, isPast = w < CURRENT_WEEK;
  return `<th class="wk" style="${isCurrent?'background:var(--green-bg);':''}${isPast?'color:#b0b8c8;':''}">${isCurrent?'▼':''}W${w}</th>`;
}
function wkCell(w, alloc){
  const isPast = w < CURRENT_WEEK;
  return `<td class="wk ${wClass(alloc)}" style="${isPast&&!alloc?'opacity:.5;':''}${isPast&&alloc?'filter:saturate(.6);':''}">${alloc>0 ? alloc+'%' : '–'}</td>`;
}
function wkCellA(a, w){
  const eff = getEffectiveAlloc(a, w), isPast = w < CURRENT_WEEK;
  return `<td class="wk ${wClass(eff)}" style="${isPast&&!eff?'opacity:.5;':''}${isPast&&eff?'filter:saturate(.6);':''}">${eff>0 ? eff+'%' : '–'}</td>`;
}

// ── Week navigation ───────────────────────────────────────────────────────────
function weekNav(delta){
  const minOffset = -(CURRENT_WEEK - 1), maxOffset = 52 - CURRENT_WEEK - 11;
  state.weekOffset = Math.max(minOffset, Math.min(maxOffset, state.weekOffset + delta));
  render();
}
function weekNavReset(){ state.weekOffset = 0; render(); }
function toggleAddMember(){ state.addMemberOpen=!state.addMemberOpen; render(); }
function weekNavToggleAll(){ state.showAllWeeks = !state.showAllWeeks; state.weekOffset = 0; render(); }

// ── FIXED: weekRangeToggle — buttons call weekNav() directly, pointer-events:none disables ──
function weekRangeToggle(){
  const minOffset = -(CURRENT_WEEK - 1), maxOffset = 52 - CURRENT_WEEK - 11;
  const atStart = state.weekOffset <= minOffset, atEnd = state.weekOffset >= maxOffset;
  const wks = visibleWeeks(), firstW = wks[0], lastW = wks[wks.length - 1];
  const dim = 'opacity:.35;pointer-events:none;';
  return `<div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap">
    <button class="btn sm" onclick="weekNavToggleAll()" style="font-size:11px">${state.showAllWeeks ? '📅 Aktuell vecka' : '⊞ Alla veckor'}</button>
    ${!state.showAllWeeks ? `
    <button class="btn sm" onclick="weekNav(-12)" style="font-size:11px;padding:3px 8px;${atStart?dim:''}" title="12 veckor bakåt">◀ Bakåt</button>
    <button class="btn sm" onclick="weekNav(-1)"  style="font-size:11px;padding:3px 7px;${atStart?dim:''}" title="1 vecka bakåt">‹</button>
    <span style="font-size:11px;color:#6b7280;min-width:68px;text-align:center;font-family:'DM Mono',monospace;user-select:none">W${firstW}–W${lastW}</span>
    <button class="btn sm" onclick="weekNav(1)"   style="font-size:11px;padding:3px 7px;${atEnd?dim:''}" title="1 vecka framåt">›</button>
    <button class="btn sm" onclick="weekNav(12)"  style="font-size:11px;padding:3px 8px;${atEnd?dim:''}" title="12 veckor framåt">Framåt ▶</button>
    <button class="btn sm" onclick="weekNavReset()" style="font-size:11px;padding:3px 8px;${state.weekOffset===0?dim:'color:#0f6e56'}" title="Gå till aktuell vecka">↩ Nu</button>
    ` : ''}
  </div>`;
}

function flashMsg(text, ok){ state.msg = {text, ok}; render(); setTimeout(() => { state.msg = null; render(); }, 3000); }
function fmtDate(d){ if(!d) return '—'; return new Date(d).toLocaleDateString('en-SE', {day:'2-digit', month:'short'}); }
function fmtDateLong(d){ if(!d) return '—'; return new Date(d).toLocaleDateString('en-SE', {day:'2-digit', month:'short', year:'numeric'}); }
function badge(label, bg, color){ return `<span style="background:${bg};color:${color};padding:1px 8px;border-radius:20px;font-size:11px;font-weight:700">${label}</span>`; }

function peopleSelect(id, value, onchangeCode, extraStyle, placeholder){
  const ph = placeholder || 'Type or select person…';
  setTimeout(() => { const el=document.getElementById(id); if(el && document.activeElement!==el) el.value=value||''; }, 0);
  return `<input class="sel" id="${id}" list="people-list" placeholder="${ph}" style="${extraStyle||''}" autocomplete="off" oninput="${onchangeCode}" onblur="render()" />`;
}
function peopleSelectOptional(id, value, onchangeCode, extraStyle){
  setTimeout(() => { const el=document.getElementById(id); if(el && document.activeElement!==el) el.value=value||''; }, 0);
  return `<input class="sel" id="${id}" list="people-list-optional" placeholder="Type name or leave blank…" style="${extraStyle||''}" autocomplete="off" oninput="${onchangeCode}" onblur="debounce(render,200)" />`;
}
function buildDatalist(){
  const people = getAllPeople();
  const opts = people.map(p => `<option value="${p.name}">${p.skillset} · ${p.team}</option>`).join('');
  ['people-list','people-list-optional'].forEach(id => {
    let dl = document.getElementById(id);
    if(!dl){ dl = document.createElement('datalist'); dl.id = id; document.body.appendChild(dl); }
    dl.innerHTML = opts;
  });
}
function updateSidebarForRole(){ const el = document.getElementById('sidebar-teams-section'); if(el) el.style.display = role()==='Team Member' ? 'none' : ''; }
function toggleDark(){ const isDark = document.body.classList.toggle('dark'); localStorage.setItem('rp_dark', isDark ? '1' : '0'); document.getElementById('dark-btn').textContent = isDark ? '☀ Light' : '🌙 Dark'; }
(function(){ if(localStorage.getItem('rp_dark')==='1') document.body.classList.add('dark'); })();

function onPersonInput(val, target){
  if(target==='add') state.aName=val; else if(target==='tm') state.tmName=val; else if(target==='pr') state.prName=val;
  const match = getAllPeople().find(p => p.name.trim().toLowerCase() === val.trim().toLowerCase());
  if(match){
    if(target==='add'){ Object.assign(state, {aName:match.name, aTeam:match.team, aCountry:match.country, aSkill:match.skillset, aLevel:match.level}); clearTimeout(_debounceTimer); render(); }
    else if(target==='tm'){ Object.assign(state, {tmName:match.name, tmCountry:match.country, tmSkill:match.skillset, tmLevel:match.level}); }
    else if(target==='pr'){ Object.assign(state, {prName:match.name, prTeam:match.team, prCountry:match.country, prSkill:match.skillset, prLevel:match.level}); clearTimeout(_debounceTimer); render(); }
  } else if(target==='add'){ debounce(render, 600); }
}

const TAB_NAMES = {dashboard:'Dashboard', overview:'Overview', 'person-detail':'Person Detail', projects:'Projects', 'project-detail':'Project Detail', services:'Base Services', 'team-detail':'Team Detail', inbox:'Inbox', pipeline:'Pipeline', add:'Planning mode'};
function setTab(t){ state.tab=t; document.querySelectorAll('.nav-item').forEach(el => el.classList.toggle('active', el.textContent.trim().replace(/^./,'').trim()===TAB_NAMES[t])); document.getElementById('tab-title').textContent=TAB_NAMES[t]; render(); }
function openPersonDetail(name){ if(role()==='Team Member' && name.trim().toLowerCase()!==userName().trim().toLowerCase()) return; state.selectedPerson=name; state.tab='person-detail'; document.getElementById('tab-title').textContent=name; document.querySelectorAll('.nav-item').forEach(el=>el.classList.remove('active')); render(); }
function openProject(projId){ state.selectedProject=projId; state.tab='project-detail'; document.getElementById('tab-title').textContent='Project Detail'; document.querySelectorAll('.nav-item').forEach(el=>el.classList.remove('active')); render(); }
function openTeam(teamName){ if(role()==='Team Member') return; state.selectedTeam=teamName; state.tab='team-detail'; state.teamFilterNames=new Set(); document.getElementById('tab-title').textContent=teamName+' Team'; document.querySelectorAll('.nav-item').forEach(el=>el.classList.remove('active')); const idMap={Development:'nav-dev',Platform:'nav-plat',PMO:'nav-pmo'}; if(idMap[teamName]) document.getElementById(idMap[teamName])?.classList.add('active'); render(); }

function commitA(ai){ const a=state.assignments[ai]; a.committed=true; a.committedBy=userName(); a.confirmed=true; a.confirmedBy=userName(); render(); }
function uncommitA(ai){ const a=state.assignments[ai]; a.committed=false; a.committedBy=null; a.confirmed=false; a.confirmedBy=null; render(); }
function delAssignment(ai){ state.assignments.splice(ai,1); render(); }
function delPeriod(ai,pi){ state.assignments[ai].periods.splice(pi,1); state.assignments[ai].committed=false; state.assignments[ai].committedBy=null; state.assignments[ai].confirmed=false; state.assignments[ai].confirmedBy=null; render(); }
function addPeriod(ai){ state.assignments[ai].periods.push({id:Date.now(), startWeek:state.aStart, endWeek:state.aEnd, allocationPercent:state.aPct}); state.assignments[ai].committed=false; state.assignments[ai].committedBy=null; render(); }
function addProject(){ if(!state.pName.trim()) return; state.projects.push({id:Date.now(), name:state.pName.trim(), projectManager:state.pPm.trim(), startDate:state.pStart, endDate:state.pEnd, description:state.pDesc.trim()}); state.pName=''; state.pPm=''; state.pStart=''; state.pEnd=''; state.pDesc=''; render(); }
function deleteProject(id){ if(!confirm('Delete this project? Planning entries linked to it will not be deleted.')) return; state.projects=state.projects.filter(p=>p.id!==id); render(); }
function saveProjectEdits(id){
  const name=document.getElementById('ep-name')?.value.trim(), pm=document.getElementById('ep-pm')?.value||'', start=document.getElementById('ep-start')?.value, end=document.getElementById('ep-end')?.value, desc=document.getElementById('ep-desc')?.value.trim();
  if(!name) return;
  const old=state.projects.find(p=>p.id===id);
  if(old && old.name!==name) state.assignments.forEach(a => { if(a.workName===old.name && a.type==='Project') a.workName=name; });
  state.projects=state.projects.map(p => p.id===id ? {...p,name,projectManager:pm,startDate:start,endDate:end,description:desc} : p);
  flashMsg('Project updated!', true);
}
function toggleEditProject(){ const el=document.getElementById('edit-proj-panel'); if(el) el.style.display=el.style.display==='none'?'block':'none'; }
function addResourceToProject(){
  const proj=state.projects.find(p=>p.id===state.selectedProject); if(!proj) return;
  if(!state.prName.trim()||!state.prSkill.trim()){ flashMsg('Please fill in name and skillset.',false); return; }
  const period={id:Date.now(), startWeek:state.prStart, endWeek:state.prEnd, allocationPercent:state.prPct};
  const ei=state.assignments.findIndex(a => a.name.trim().toLowerCase()===state.prName.trim().toLowerCase() && a.type==='Project' && a.workName===proj.name);
  if(ei>=0){ state.assignments[ei].periods.push(period); state.assignments[ei].committed=false; state.assignments[ei].committedBy=null; }
  else state.assignments.push({id:Date.now(), name:state.prName.trim(), team:state.prTeam, country:state.prCountry, skillset:state.prSkill.trim(), level:state.prLevel, type:'Project', workName:proj.name, projectId:proj.id, periods:[period], confirmed:false, confirmedBy:null, committed:false, committedBy:null});
  state.prName=''; state.prSkill=''; state.prStart=1; state.prEnd=4; state.prPct=80;
  flashMsg('Resource added!', true);
}
function addSvc(){ if(!state.sName.trim()) return; state.baseServices.push({name:state.sName.trim(), team:state.sTeam, targetPct:state.sTargetPct||0}); state.sName=''; state.sTargetPct=0; render(); }
function editSvc(name){ const n=prompt('New name:',name); if(!n?.trim()) return; state.baseServices=state.baseServices.map(s=>s.name===name?{...s,name:n.trim()}:s); render(); }
function delSvc(name){ state.baseServices=state.baseServices.filter(s=>s.name!==name); render(); }
function addTeamMember(){
  const nameEl=document.getElementById('inp-tmName'), skillEl=document.getElementById('inp-tmSkill'), tlEl=document.getElementById('inp-tmTl'), mgrEl=document.getElementById('inp-tmMgr');
  if(nameEl?.value.trim()) state.tmName=nameEl.value.trim(); if(skillEl?.value.trim()) state.tmSkill=skillEl.value.trim();
  if(tlEl?.value.trim()) state.tmTeamlead=tlEl.value.trim(); if(mgrEl?.value.trim()) state.tmManager=mgrEl.value.trim();
  if(!state.tmName.trim()||!state.tmSkill.trim()) return;
  const teamName=state.selectedTeam;
  if(state.teamMembers.find(m=>m.name.trim().toLowerCase()===state.tmName.trim().toLowerCase()&&m.team===teamName)){ flashMsg('This person is already in the team.',false); return; }
  state.teamMembers.push({id:Date.now(), name:state.tmName.trim(), team:teamName, country:state.tmCountry, skillset:state.tmSkill.trim(), level:state.tmLevel, teamlead:state.tmTeamlead.trim(), manager:state.tmManager.trim()});
  state.tmName=''; state.tmSkill=''; state.tmTeamlead=''; state.tmManager=''; render();
}
function removeTeamMember(id){ state.teamMembers=state.teamMembers.filter(m=>m.id!==id); render(); }
function toggleTeamFilter(name){ if(state.teamFilterNames.has(name)) state.teamFilterNames.delete(name); else state.teamFilterNames.add(name); render(); }
function startEditAssignment(idx){
  if(!canEdit()) return;
  if(state.editingAssignmentId===idx){ state.editingAssignmentId=null; render(); return; }
  const a=state.assignments[idx]; if(!a) return;
  state.editingAssignmentId=idx;
  if(a.periods.length){ state.eaStart=a.periods[0].startWeek; state.eaEnd=a.periods[a.periods.length-1].endWeek; state.eaPct=a.periods[0].allocationPercent; }
  state.editingMemberId=null; render();
  setTimeout(() => {
    const el = document.getElementById('ea-start-'+idx);
    if(el) {
      el.closest('tr')?.scrollIntoView({behavior:'smooth', block:'nearest'});
    }
  }, 50);
}
function addPeriodToAssignment(idx){
  const start=parseInt(document.getElementById('ea-start-'+idx)?.value)||state.eaStart;
  const end=parseInt(document.getElementById('ea-end-'+idx)?.value)||state.eaEnd;
  const pctRaw=document.getElementById('ea-pct-'+idx)?.value;
  const pct=pctRaw===''||pctRaw===null||pctRaw===undefined ? state.eaPct : parseInt(pctRaw);
  if(start>end){ flashMsg('Start week must be before end week.',false); return; }

  const a = state.assignments[idx];
  const newPeriods = [];

  // Split existing periods around the new one
  a.periods.forEach(p => {
    const hasOverlap = p.startWeek <= end && p.endWeek >= start;
    if(!hasOverlap) {
      newPeriods.push(p);
    } else {
      if(p.startWeek < start)
        newPeriods.push({id:Date.now()+Math.random(), startWeek:p.startWeek, endWeek:start-1, allocationPercent:p.allocationPercent});
      if(p.endWeek > end)
        newPeriods.push({id:Date.now()+Math.random(), startWeek:end+1, endWeek:p.endWeek, allocationPercent:p.allocationPercent});
    }
  });

  // Add new period (only skip if pct is 0 AND user explicitly wants to remove — we keep 0% for e.g. vacation)
  if(pct > 0) {
    newPeriods.push({id:Date.now(), startWeek:start, endWeek:end, allocationPercent:pct});
  }
  // If pct === 0, the gap just means no allocation those weeks (no period needed)

  // Sort by start week
  newPeriods.sort((a,b) => a.startWeek - b.startWeek);

  // Merge adjacent periods with same allocation %
  const merged = [];
  newPeriods.forEach(p => {
    const last = merged[merged.length-1];
    if(last && last.allocationPercent === p.allocationPercent && last.endWeek + 1 === p.startWeek) {
      last.endWeek = p.endWeek; // extend
    } else {
      merged.push({...p});
    }
  });

  a.periods = merged;
  a.committed = false;
  a.committedBy = null;
  state.eaStart = Math.min(end+1, 52);
  state.eaEnd   = Math.min(end+4, 52);
  flashMsg('Period added!', true);
}
function saveAssignmentEdit(idx){
  const start=parseInt(document.getElementById('ea-start-'+idx)?.value)||state.eaStart;
  const end=parseInt(document.getElementById('ea-end-'+idx)?.value)||state.eaEnd;
  const pct=parseInt(document.getElementById('ea-pct-'+idx)?.value)||state.eaPct;
  if(start>end){ flashMsg('Start week must be before end week.',false); return; }
  state.assignments[idx].periods=[{id:Date.now(), startWeek:start, endWeek:end, allocationPercent:pct}];
  state.assignments[idx].committed=false; state.assignments[idx].committedBy=null;
  state.editingAssignmentId=null; flashMsg('Assignment updated!',true);
}
function startEditMember(id){
  if(!canEdit()) return;
  if(state.editingMemberId===id){ state.editingMemberId=null; render(); return; }
  const m=state.teamMembers.find(m=>m.id===id); if(!m) return;
  state.editingMemberId=id; state.emName=m.name; state.emCountry=m.country; state.emSkill=m.skillset; state.emLevel=m.level; state.emTeamlead=m.teamlead||''; state.emManager=m.manager||''; state.emTeam=m.team||teamName; render();
}
function saveMemberEditFromInputs(id){
  const nv=document.getElementById('em-name-'+id)?.value.trim()||state.emName.trim();
  const sv=document.getElementById('em-skill-'+id)?.value.trim()||state.emSkill.trim();
  const tv=document.getElementById('em-tl-'+id)?.value.trim()||'';
  const mv=document.getElementById('em-mgr-'+id)?.value.trim()||'';
  if(!nv) return;
  const old=state.teamMembers.find(m=>m.id===id), oldName=(old?.name||'').trim().toLowerCase();
  state.teamMembers=state.teamMembers.map(m => m.id===id ? {...m,name:nv,country:state.emCountry,skillset:sv,level:state.emLevel,teamlead:tv,manager:mv,team:state.emTeam} : m);
  state.assignments=state.assignments.map(a => { if(a.name.trim().toLowerCase()===oldName) return {...a,name:nv,country:state.emCountry||a.country,skillset:sv||a.skillset,level:state.emLevel||a.level,team:state.emTeam||a.team}; return a; });
  state.editingMemberId=null; flashMsg('Member updated!',true);
}
function saveTeamConfig(team,field,value){ if(!state.teamConfig[team]) state.teamConfig[team]={teamlead:'',manager:''}; state.teamConfig[team][field]=value; }
function addAssignment(){
  if(!canPlan()){ flashMsg('You do not have permission to add planning.',false); return; }
  const isPM=role()==='Project Manager';
  if(!isPM&&!state.aName.trim()){ flashMsg('Please fill in the name.',false); return; }
  if(!state.aSkill.trim()){ flashMsg('Please fill in the skillset.',false); return; }
  if(state.addType==='Project'&&!state.aProjId){ flashMsg('Please select a project.',false); return; }
  if(state.addType==='Base Service'&&!state.aService){ flashMsg('Please select a base service.',false); return; }
  if((state.addType==='Charge On'||state.addType==='Internal Initiative')&&!state.aWork.trim()){ flashMsg('Please enter a work name.',false); return; }
  let wn='';
  if(state.addType==='Project'){ const p=state.projects.find(p=>p.id==state.aProjId); wn=p?p.name:''; }
  else if(state.addType==='Base Service') wn=state.aService; else wn=state.aWork;
  const period={id:Date.now(), startWeek:state.aStart, endWeek:state.aEnd, allocationPercent:state.aPct};
  const ei=state.assignments.findIndex(a => a.name.trim().toLowerCase()===state.aName.trim().toLowerCase() && a.type===state.addType && a.workName.trim().toLowerCase()===wn.trim().toLowerCase());
  if(ei>=0){ state.assignments[ei].periods.push(period); state.assignments[ei].committed=false; state.assignments[ei].committedBy=null; }
  else { const assignName=isPM?('__pm_planned__'+Date.now()):state.aName.trim(); state.assignments.push({id:Date.now(), name:assignName, team:state.aTeam, country:state.aCountry, skillset:state.aSkill.trim(), level:state.aLevel, type:state.addType, workName:wn, projectId:state.addType==='Project'?+state.aProjId:null, periods:[period], confirmed:false, confirmedBy:null, committed:false, committedBy:null, pmPlanned:isPM}); }
  flashMsg('Entry added!',true);
  state.aName=''; state.aSkill=''; state.aProjId=''; state.aService=''; state.aWork=''; state.aStart=1; state.aEnd=3; state.aPct=80;
}
function addInboxItem(){ if(!state.iTitle.trim()) return; state.inboxItems.push({id:Date.now(), title:state.iTitle.trim(), description:state.iDesc.trim(), priority:state.iPriority, status:'new', createdBy:userName(), createdAt:new Date().toLocaleDateString('en-SE'), convertedTo:null}); state.iTitle=''; state.iDesc=''; state.iPriority='Medium'; render(); }
function deleteInboxItem(id){ state.inboxItems=state.inboxItems.filter(i=>i.id!==id); render(); }
function convertInboxItem(id,to){
  const item=state.inboxItems.find(i=>i.id===id); if(!item) return;
  if(to==='revert'){ if(item.convertedTo==='Project'){ const proj=state.projects.find(p=>p.name===item.title); if(proj&&confirm('This will also delete the project "'+item.title+'". Continue?')) state.projects=state.projects.filter(p=>p.name!==item.title); else if(!proj){} else return; } item.status='new'; item.convertedTo=null; flashMsg('Reverted to inbox.',true); render(); return; }
  if(to==='project'){ state.projects.push({id:Date.now(), name:item.title, projectManager:'', startDate:'', endDate:'', description:item.description}); item.status='converted'; item.convertedTo='Project'; flashMsg('Converted to project!',true); }
  else if(to==='initiative'){ if(item.convertedTo==='Project'){ const proj=state.projects.find(p=>p.name===item.title); if(proj) state.projects=state.projects.filter(p=>p.name!==item.title); } item.status='converted'; item.convertedTo='Internal Initiative'; flashMsg('Changed to Internal Initiative!',true); }
  render();
}
function autoRegisterTeamMembers(){ state.assignments.forEach(a => { if(!a.name||!a.team) return; const key=a.name.trim().toLowerCase(), team=a.team.trim(); if(!state.teamMembers.some(m=>m.name.trim().toLowerCase()===key&&m.team===team)) state.teamMembers.push({id:Date.now()+Math.random(), name:a.name.trim(), team, country:a.country||'Sweden', skillset:a.skillset||'', level:a.level||'Junior'}); }); }

function renderAddMemberCard(teamName, cfg, state){
  if(!canEdit()) return '';
  const open = state.addMemberOpen;
  const body = open ? `
    <div class="card-body" style="display:flex;flex-direction:column;gap:14px">
      <div style="display:grid;grid-template-columns:1.5fr 1fr 1.5fr 1fr auto;gap:12px;align-items:flex-end">
        <div class="fg"><label class="lbl">Full name *</label><input class="inp" id="inp-tmName" list="people-list" placeholder="Type or pick a name…" autocomplete="off" oninput="onPersonInput(this.value,'tm')" /></div>
        <div class="fg"><label class="lbl">Country</label><select class="sel" onchange="state.tmCountry=this.value"><option value="Sweden"${state.tmCountry==='Sweden'?' selected':''}>Sweden</option><option value="Poland"${state.tmCountry==='Poland'?' selected':''}>Poland</option></select></div>
        <div class="fg"><label class="lbl">Skillset *</label><input class="inp" id="inp-tmSkill" placeholder="e.g. React, DevOps" oninput="state.tmSkill=this.value" onkeydown="if(event.key==='Enter')addTeamMember()" /></div>
        <div class="fg"><label class="lbl">Level</label><select class="sel" onchange="state.tmLevel=this.value">${['Junior','Mid','Senior'].map(l=>`<option${state.tmLevel===l?' selected':''}>${l}</option>`).join('')}</select></div>
        <button class="btn primary" onclick="addTeamMember()" style="white-space:nowrap">＋ Add member</button>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:12px;background:#f9fafb;border-radius:8px;border:1px solid #f3f4f6">
        <div class="fg"><label class="lbl">Teamlead override <span style="color:#9ca3af;font-weight:400">— blank = team default: <strong>${cfg.teamlead||'none set'}</strong></span></label><input class="inp" id="inp-tmTl" list="people-list-optional" placeholder="Blank = inherit from team" autocomplete="off" oninput="state.tmTeamlead=this.value" /></div>
        <div class="fg"><label class="lbl">Manager override <span style="color:#9ca3af;font-weight:400">— blank = team default: <strong>${cfg.manager||'none set'}</strong></span></label><input class="inp" id="inp-tmMgr" list="people-list-optional" placeholder="Blank = inherit from team" autocomplete="off" oninput="state.tmManager=this.value" /></div>
      </div>
    </div>` : '';
  return `<div class="card" style="margin-bottom:16px">
    <div class="card-hdr" onclick="toggleAddMember()" style="cursor:pointer;user-select:none">
      <div style="display:flex;align-items:center;gap:8px">
        <span style="font-size:13px">${open?'▾':'▸'}</span>
        <span class="card-title">＋ Add team member</span>
      </div>
    </div>
    ${body}
  </div>`;
}

function render(){
  buildDatalist(); updateSidebarForRole();
  const darkBtn=document.getElementById('dark-btn'); if(darkBtn) darkBtn.textContent=document.body.classList.contains('dark')?'☀ Light':'🌙 Dark';
  document.getElementById('foot').textContent=`${state.assignments.length} assignment${state.assignments.length!==1?'s':''} · ${state.projects.length} project${state.projects.length!==1?'s':''}`;
  saveData();
  const el=document.getElementById('content');
  switch(state.tab){
    case 'dashboard':      el.innerHTML=renderDashboard();    injectTeamAllocCard(); break;
    case 'overview':       el.innerHTML=renderOverview();     break;
    case 'person-detail':  el.innerHTML=renderPersonDetail(); break;
    case 'planning':       el.innerHTML=renderPlanning();     break;
    case 'projects':       el.innerHTML=renderProjects();     break;
    case 'project-detail': el.innerHTML=renderProjectDetail();break;
    case 'services':       el.innerHTML=renderServices();     break;
    case 'team-detail':    el.innerHTML=renderTeamDetail();   break;
    case 'inbox':          el.innerHTML=renderInbox();        break;
    case 'pipeline':       el.innerHTML=renderPipeline();     break;
    case 'add': if(role()==='Project Manager') state.addType='Project'; el.innerHTML=renderAdd(); break;
  }
}
function injectTeamAllocCard(){ const el=document.getElementById('team-alloc-card'); if(!el) return; const cw=CURRENT_WEEK,AR=state.dashAllocRange||8; const allocWeeks=WEEKS.slice(cw-1,cw-1+AR).filter(w=>w<=52); el.innerHTML=buildTeamAllocCard(buildTeamMemberMap(),allocWeeks,cw,state); }

function buildTeamAllocCard(tmMap, allocWeeks, cw, state){
  const rangeOpts = [4,8,12,26,52].map(n => { const on=state.dashAllocRange===n; return `<button onclick="state.dashAllocRange=${n};render()" style="padding:5px 14px;font-size:12px;font-weight:600;border-radius:20px;border:1px solid ${on?'#1D9E75':'#e5e7eb'};background:${on?'#1D9E75':'#fff'};color:${on?'#fff':'#6b7280'};cursor:pointer;font-family:inherit">${n}W</button>`; }).join('');
  const wHeaders = allocWeeks.map(w => { const isNow=w===cw; return `<th style="padding:8px 4px;text-align:center;min-width:58px;${isNow?'background:var(--green-bg)':''}"><div style="font-size:10px;font-weight:700;color:${isNow?'#0f6e56':'#9ca3af'};font-family:DM Mono,monospace">W${w}</div>${isNow?'<div style="width:4px;height:4px;background:#1D9E75;border-radius:50%;margin:2px auto 0"></div>':''}</th>`; }).join('');
  const teamRows = TEAMS.map((team,ti) => {
    const icon=team==='Development'?'💻':team==='Platform'?'☁':'📊', mc=tmMap[team].length, now=calcTeamAllocForWeek(tmMap,team,cw);
    const bigColor=mc===0?'#d1d5db':now.avg>100?'#b91c1c':now.avg>=80?'#0f6e56':now.avg>=50?'#b45309':'#185fa5';
    const pills=(now.fullyBooked?`<span style="font-size:10px;font-weight:700;color:#065f46;background:#d1fae5;padding:1px 6px;border-radius:20px">${now.fullyBooked} full</span>`:'')+(now.free?`<span style="font-size:10px;font-weight:700;color:#185fa5;background:#dbeafe;padding:1px 6px;border-radius:20px;margin-left:4px">${now.free} free</span>`:'')+(now.over?`<span style="font-size:10px;font-weight:700;color:#b91c1c;background:#fef2f2;padding:1px 6px;border-radius:20px;margin-left:4px">${now.over} over</span>`:'')+(mc===0?'<span style="font-size:10px;color:#d1d5db;font-style:italic">no members yet</span>':'');
    const wCells=mc===0?`<td colspan="${allocWeeks.length}" style="padding:16px 24px"><div style="display:flex;align-items:center;gap:10px"><div style="flex:1;height:8px;background:#f3f4f6;border-radius:4px"></div><span style="font-size:12px;color:#d1d5db;font-style:italic">No members yet</span></div></td>`:allocWeeks.map(w=>{const ta=calcTeamAllocForWeek(tmMap,team,w),isNow=w===cw,barH=Math.min(Math.round((ta.avg/100)*52),52),barColor=ta.avg>100?'#fca5a5':ta.avg>=80?'#6ee7b7':ta.avg>=50?'#fcd34d':ta.avg>0?'#93c5fd':'#e5e7eb',textColor=ta.avg>100?'#b91c1c':ta.avg>=80?'#065f46':ta.avg>=50?'#92400e':ta.avg>0?'#1e40af':'#d1d5db';return `<td style="padding:8px 4px;text-align:center;vertical-align:bottom;${isNow?'background:var(--green-bg);':''}border-bottom:1px solid #f3f4f6"><div style="display:flex;flex-direction:column;align-items:center;gap:3px">${ta.over?`<div style="font-size:9px;color:#b91c1c;font-weight:700">+${ta.over}</div>`:'<div style="font-size:9px;color:transparent">·</div>'}<div style="font-size:12px;font-weight:700;color:${textColor};font-family:DM Mono,monospace;line-height:1">${ta.avg>0?ta.avg+'%':'–'}</div><div style="width:32px;height:52px;background:#f3f4f6;border-radius:4px;overflow:hidden;display:flex;align-items:flex-end"><div style="width:100%;height:${barH}px;background:${barColor};border-radius:4px"></div></div></div></td>`;}).join('');
    return `<tr style="border-bottom:${ti<2?'2px solid #e5e7eb':'1px solid #f3f4f6'}"><td style="padding:14px 24px;vertical-align:middle;min-width:220px"><div style="display:flex;align-items:center;gap:12px"><div style="width:44px;height:44px;border-radius:12px;background:var(--green-bg);display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0">${icon}</div><div style="flex:1;min-width:0"><div style="font-size:14px;font-weight:700;color:#111827">${team}</div><div style="font-size:12px;color:#6b7280;margin-top:2px">${mc} member${mc!==1?'s':''}</div><div style="display:flex;gap:6px;margin-top:4px;flex-wrap:wrap">${pills}</div></div><div style="text-align:right;flex-shrink:0"><div style="font-size:28px;font-weight:700;color:${bigColor};font-family:DM Mono,monospace;line-height:1">${mc===0?'—':now.avg+'%'}</div><div style="font-size:10px;color:#9ca3af;margin-top:2px">now</div></div></div></td>${wCells}</tr>`;
  }).join('');
  return `<div class="card"><div class="card-hdr" style="padding:16px 24px"><div><span class="card-title" style="font-size:15px">📊 Team allocation</span><div style="font-size:12px;color:#9ca3af;margin-top:2px">Average allocation % per team · ${allocWeeks.length} weeks shown</div></div><div style="display:flex;gap:6px;align-items:center">${rangeOpts}</div></div><div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;min-width:${180+allocWeeks.length*62}px"><thead><tr style="border-bottom:2px solid #e5e7eb"><th style="padding:10px 24px;text-align:left;font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;min-width:180px">Team</th>${wHeaders}</tr></thead><tbody>${teamRows}</tbody></table></div><div style="padding:12px 24px;border-top:1px solid #f3f4f6;display:flex;gap:20px;font-size:11px;color:#6b7280;flex-wrap:wrap;align-items:center"><span style="font-weight:600;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;font-size:10px">Legend</span><span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#93c5fd;display:inline-block"></span>Low &lt;50%</span><span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#fcd34d;display:inline-block"></span>Medium 50–79%</span><span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#6ee7b7;display:inline-block"></span>High 80–100%</span><span style="display:flex;align-items:center;gap:5px"><span style="width:12px;height:12px;border-radius:3px;background:#fca5a5;display:inline-block"></span>Overbooked</span><span style="display:flex;align-items:center;gap:5px"><span style="width:4px;height:4px;border-radius:50%;background:#1D9E75;display:inline-block"></span>Current week</span></div></div>`;
}

function renderDashboard(){
  const cw=CURRENT_WEEK,NEXT=4,isTM=role()==='Team Member',un=userName().trim().toLowerCase();
  function ptw(name,w){return state.assignments.filter(a=>a.name===name&&a.committed).reduce((s,a)=>s+getAlloc(a,w),0);}
  const allPeople=isTM?[...new Set(state.assignments.filter(a=>a.committed&&a.name.trim().toLowerCase()===un).map(a=>a.name))]:[...new Set(state.assignments.map(a=>a.name))];
  const activeNow=allPeople.filter(n=>ptw(n,cw)>0);
  const overbookedNow=allPeople.filter(n=>ptw(n,cw)>100&&!n.trim().toLowerCase().startsWith('nn'));
  const allRegistered=isTM?[userName().trim()].filter(Boolean):[...new Set([...state.teamMembers.map(m=>m.name),...state.assignments.map(a=>a.name)])];
  const freeNow=allRegistered.filter(n=>ptw(n,cw)===0);
  const uncommitted=state.assignments.filter(a=>!a.committed);
  const inboxPending=state.inboxItems.filter(i=>i.status==='new');
  const activeProjects=state.projects.filter(p=>state.assignments.some(a=>a.workName===p.name&&a.periods.some(per=>cw>=per.startWeek&&cw<=per.endWeek)));
  const activeProjectNames=new Set(activeProjects.map(p=>p.name));
  const startingSoon=state.assignments.filter(a=>a.type==='Project'&&activeProjectNames.has(a.workName)&&a.periods.some(p=>p.startWeek>cw&&p.startWeek<=cw+NEXT));
  const startingSoonProjects=[...new Set(startingSoon.map(a=>a.workName))];
  const endingSoon=state.assignments.filter(a=>a.type==='Project'&&activeProjectNames.has(a.workName)&&a.periods.some(p=>p.endWeek>=cw&&p.endWeek<=cw+NEXT));
  const uncommittedByProject={};
  uncommitted.forEach(a=>{if(!uncommittedByProject[a.workName])uncommittedByProject[a.workName]=[];uncommittedByProject[a.workName].push(a);});
  const todayStr=new Date().toISOString().split('T')[0];
  const upcomingEnds=state.projects.filter(p=>p.endDate&&p.endDate>=todayStr).sort((a,b)=>a.endDate.localeCompare(b.endDate)).slice(0,5);
  const weeklyChanges=[];
  for(let w=cw;w<=Math.min(cw+NEXT-1,52);w++){const starts=state.assignments.filter(a=>a.periods.some(p=>p.startWeek===w)),ends=state.assignments.filter(a=>a.periods.some(p=>p.endWeek===w));if(starts.length||ends.length)weeklyChanges.push({w,starts,ends});}
  function pTag(label,color,bg){return `<span class="dash-tag" style="background:${bg};color:${color}">${label}</span>`;}
  const needsAttention=!isTM&&(overbookedNow.length||inboxPending.length||uncommitted.length);
  const attentionHtml=needsAttention?`<div class="card"><div class="card-hdr"><span class="card-title">⚠ Needs attention</span></div>${overbookedNow.map(n=>`<div class="alert-row"><span class="alert-icon">🔴</span><div style="flex:1"><strong>${n}</strong> is overbooked in W${cw} (${ptw(n,cw)}%)</div><button class="btn sm" onclick="openPersonDetail('${n.replace(/'/g,"\\'")}')">View →</button></div>`).join('')}${inboxPending.map(i=>`<div class="alert-row"><span class="alert-icon">📥</span><div style="flex:1"><strong>${i.title}</strong> — in inbox, needs classification</div><button class="btn sm" onclick="setTab('inbox')">Go to Inbox →</button></div>`).join('')}${Object.entries(uncommittedByProject).slice(0,4).map(([proj,items])=>{const p=state.projects.find(p=>p.name===proj);return `<div class="alert-row"><span class="alert-icon">⏳</span><div style="flex:1">${items.length} planned resource${items.length!==1?'s':''} not yet committed on <strong>${proj}</strong></div>${p?`<button class="btn sm" onclick="openProject(${p.id})">View →</button>`:''}</div>`;}).join('')}</div>`:`<div class="card"><div style="padding:14px 18px;font-size:13px;color:#0f6e56;font-weight:600">✅ Everything looks good — no immediate issues.</div></div>`;
  const hasTargets=state.baseServices.some(s=>s.targetPct>0);
  const debtHtml=hasTargets?(()=>{const rows=TEAMS.map(team=>{const d=calcTeamDebt(team),svcsWT=state.baseServices.filter(s=>s.team===team&&s.targetPct>0);if(!svcsWT.length)return '';const icon=team==='Development'?'💻':team==='Platform'?'☁':'📊',color=d.debtPct===0?'#0f6e56':d.debtPct<200?'#b45309':'#b91c1c',maxDebt=svcsWT.reduce((s,sv)=>s+sv.targetPct*(CURRENT_WEEK-1),0),barPct=maxDebt>0?Math.min(100,Math.round((d.debtPct/maxDebt)*100)):0;return `<div class="dash-row" style="cursor:pointer" onclick="openTeam('${team}')"><div style="display:flex;align-items:center;gap:10px;flex:1;min-width:0"><span style="font-size:18px">${icon}</span><div style="flex:1;min-width:0"><div style="font-weight:600;font-size:13px">${team}</div><div style="margin-top:4px;height:5px;background:#f3f4f6;border-radius:3px;overflow:hidden"><div style="height:100%;width:${barPct}%;background:${color};border-radius:3px"></div></div><div style="font-size:10px;color:#9ca3af;margin-top:2px">${svcsWT.length} service${svcsWT.length!==1?'s':''} tracked</div></div></div><div style="text-align:right;flex-shrink:0;margin-left:16px"><div style="font-size:20px;font-weight:700;color:${color};font-family:'DM Mono',monospace;line-height:1">${d.debtPct===0?'✓':d.debtPct+'%'}</div>${d.debtPct>0?`<div style="font-size:10px;color:#9ca3af;margin-top:2px">${d.debtWeeks}w fulltime</div>`:'<div style="font-size:10px;color:#0f6e56;margin-top:2px">on track</div>'}</div><span style="font-size:12px;color:#9ca3af;margin-left:8px">→</span></div>`;}).filter(Boolean).join('');return rows?`<div class="card"><div class="card-hdr"><span class="card-title">🔧 Technical debt — Base Services</span><span class="card-sub">Accumulated W1–W${cw-1} · click team for details</span></div>${rows}</div>`:'';})():'';
  return `
    <div class="metrics" style="grid-template-columns:repeat(6,1fr);margin-bottom:0">
      <div class="metric"><div class="metric-lbl">Current week</div><div class="metric-val">W${cw}</div></div>
      <div class="metric"><div class="metric-lbl">Active now</div><div class="metric-val" style="color:#0f6e56">${activeNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Free now</div><div class="metric-val" style="color:#185fa5">${freeNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Overbooked</div><div class="metric-val ${overbookedNow.length?'red':''}">${overbookedNow.length}</div></div>
      <div class="metric"><div class="metric-lbl">Uncommitted</div><div class="metric-val" style="color:#b45309">${uncommitted.length}</div></div>
      <div class="metric"><div class="metric-lbl">Inbox</div><div class="metric-val ${inboxPending.length?'red':''}">${inboxPending.length}</div></div>
    </div>
    <div id="team-alloc-card"></div>
    ${attentionHtml}
    <div class="dash-grid">
      <div class="card"><div class="card-hdr"><span class="card-title">👥 Active this week (W${cw})</span></div>${!activeNow.length?`<div class="empty" style="padding:20px">No active resources this week.</div>`:activeNow.map(name=>{const alloc=ptw(name,cw),info=state.assignments.find(a=>a.name===name)||{team:'',skillset:''},works=[...new Set(state.assignments.filter(a=>a.name===name&&a.periods.some(p=>cw>=p.startWeek&&cw<=p.endWeek)).map(a=>a.workName))];return `<div class="dash-row" onclick="openPersonDetail('${name.replace(/'/g,"\\'")}')"><div style="flex:1"><div style="font-weight:600">${name}</div><div style="font-size:11px;color:#9ca3af">${info.team} · ${works.join(', ')}</div></div><span class="dash-tag" style="background:${alloc>100?'#fecaca;color:#b91c1c':alloc===100?'#d1fae5;color:#065f46':'#fef3c7;color:#92400e'}">${alloc}%</span></div>`;}).join('')}</div>
      <div class="card"><div class="card-hdr"><span class="card-title">💼 Active projects (W${cw})</span></div>${!activeProjects.length?`<div class="empty" style="padding:20px">No projects active this week.</div>`:activeProjects.map(p=>{const res=state.assignments.filter(a=>a.workName===p.name&&a.periods.some(per=>cw>=per.startWeek&&cw<=per.endWeek)),comm=res.filter(a=>a.committed).length,plan=res.filter(a=>!a.committed).length;return `<div class="dash-row" onclick="openProject(${p.id})"><div style="flex:1"><div style="font-weight:600">${p.name}</div><div style="font-size:11px;color:#9ca3af">ends ${fmtDate(p.endDate)} · ${res.length} resource${res.length!==1?'s':''}</div></div><div style="display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end">${comm?pTag('✓ '+comm,'#065f46','#d1fae5'):''}${plan?pTag('⏳ '+plan,'#92400e','#fef3c7'):''}</div></div>`;}).join('')}</div>
      <div class="card"><div class="card-hdr"><span class="card-title">🚀 Starting soon (next ${NEXT} weeks)</span></div>${!startingSoonProjects.length?`<div class="empty" style="padding:20px">Nothing starting in the next ${NEXT} weeks.</div>`:startingSoonProjects.map(wn=>{const items=startingSoon.filter(a=>a.workName===wn),firstW=Math.min(...items.flatMap(a=>a.periods.filter(p=>p.startWeek>cw).map(p=>p.startWeek))),proj=state.projects.find(p=>p.name===wn);return `<div class="dash-row" ${proj?`onclick="openProject(${proj.id})"`:''}><div style="flex:1"><div style="font-weight:600">${wn}</div><div style="font-size:11px;color:#9ca3af">${items.length} resource${items.length!==1?'s':''}</div></div><span class="week-badge">W${firstW}</span></div>`;}).join('')}</div>
      <div class="card"><div class="card-hdr"><span class="card-title">🏁 Ending soon (next ${NEXT} weeks)</span></div>${!endingSoon.length?`<div class="empty" style="padding:20px">Nothing ending in the next ${NEXT} weeks.</div>`:[...new Map(endingSoon.map(a=>[a.name+a.workName,a])).values()].map(a=>{const lastW=Math.max(...a.periods.filter(p=>p.endWeek>=cw&&p.endWeek<=cw+NEXT).map(p=>p.endWeek));return `<div class="dash-row" onclick="openPersonDetail('${a.name.replace(/'/g,"\\'")}')"><div style="flex:1"><div style="font-weight:600">${a.name}</div><div style="font-size:11px;color:#9ca3af">${a.workName}</div></div><span class="week-badge">ends W${lastW}</span></div>`;}).join('')}</div>
      <div class="card"><div class="card-hdr"><span class="card-title">🟢 Free this week (W${cw})</span></div>${!freeNow.length?`<div class="empty" style="padding:20px">Everyone is allocated this week.</div>`:freeNow.map(name=>{const info=state.teamMembers.find(m=>m.name===name)||state.assignments.find(a=>a.name===name)||{};let nextBusy=null;for(let w=cw+1;w<=52;w++){if(ptw(name,w)>0){nextBusy=w;break;}}return `<div class="dash-row" onclick="openPersonDetail('${name.replace(/'/g,"\\'")}')"><div style="flex:1"><div style="font-weight:600">${name}</div><div style="font-size:11px;color:#9ca3af">${info.team||''} · ${info.skillset||''}</div></div><span style="font-size:11px;color:#9ca3af">${nextBusy?'busy W'+nextBusy:'free all year'}</span></div>`;}).join('')}</div>
      <div class="card"><div class="card-hdr"><span class="card-title">📆 Week-by-week (W${cw}–W${Math.min(cw+NEXT-1,52)})</span></div>${!weeklyChanges.length?`<div class="empty" style="padding:20px">No changes in the next ${NEXT} weeks.</div>`:weeklyChanges.map(({w,starts,ends})=>`<div style="padding:10px 14px;border-bottom:1px solid #f3f4f6"><div style="display:flex;align-items:center;gap:8px;margin-bottom:6px"><span class="week-badge" style="background:#e0f2fe;color:#075985">W${w}</span>${w===cw?'<span style="font-size:10px;font-weight:700;color:#0f6e56">← current</span>':''}</div>${starts.length?`<div style="font-size:11px;color:#0f6e56;margin-bottom:3px">▶ Starts: ${[...new Set(starts.map(a=>a.name))].join(', ')}</div>`:''}${ends.length?`<div style="font-size:11px;color:#b45309">■ Ends: ${[...new Set(ends.map(a=>a.name))].join(', ')}</div>`:''}</div>`).join('')}</div>
    </div>
    <div class="card"><div class="card-hdr"><span class="card-title">🏢 Organisation</span></div><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:0;border-top:1px solid #f3f4f6">${TEAMS.map(team=>{const cfg=state.teamConfig[team]||{},members=state.teamMembers.filter(m=>m.team===team),icon=team==='Development'?'💻':team==='Platform'?'☁':'📊';return `<div style="padding:14px 18px;border-right:1px solid #f3f4f6"><div style="font-size:13px;font-weight:700;color:#111827;margin-bottom:10px">${icon} ${team}</div><div style="margin-bottom:8px"><div class="org-label">Team Lead</div><div class="org-person" style="color:${cfg.teamlead?'#0f6e56':'#9ca3af'}">${cfg.teamlead||'Not assigned'}</div></div><div style="margin-bottom:10px"><div class="org-label">Manager</div><div class="org-person" style="color:${cfg.manager?'#185fa5':'#9ca3af'}">${cfg.manager||'Not assigned'}</div></div><div class="org-label">${members.length} member${members.length!==1?'s':''}</div>${members.slice(0,4).map(m=>`<div style="font-size:12px;color:#6b7280;padding:2px 0;display:flex;align-items:center;gap:6px"><span style="width:6px;height:6px;border-radius:50%;background:${m.country==='Sweden'?'#3b82f6':'#f59e0b'};display:inline-block;flex-shrink:0"></span>${m.name}</div>`).join('')}${members.length>4?`<div style="font-size:11px;color:#9ca3af;margin-top:4px">+${members.length-4} more</div>`:''}</div>`;}).join('')}</div></div>
    ${upcomingEnds.length?`<div class="card"><div class="card-hdr"><span class="card-title">📅 Upcoming project end dates</span></div>${upcomingEnds.map(p=>{const res=state.assignments.filter(a=>a.workName===p.name),days=Math.ceil((new Date(p.endDate)-new Date())/86400000),urgency=days<=14?'#b91c1c':days<=30?'#b45309':'#374151';return `<div class="dash-row" onclick="openProject(${p.id})"><div style="flex:1"><div style="font-weight:600">${p.name}</div><div style="font-size:11px;color:#9ca3af">${res.length} resource${res.length!==1?'s':''}</div></div><div style="text-align:right"><div style="font-size:12px;font-weight:700;color:${urgency}">${fmtDate(p.endDate)}</div><div style="font-size:10px;color:#9ca3af">${days} day${days!==1?'s':''} left</div></div></div>`;}).join('')}</div>`:''}
    ${debtHtml}`;
}

function renderPersonDetail(){
  const name=state.selectedPerson; if(!name) return `<div class="card"><div class="empty">No person selected.</div></div>`;
  const ce=canEdit(), personAssignments=visibleAssignments().filter(a=>a.name===name), info=personAssignments[0]||state.teamMembers.find(m=>m.name===name);
  if(!info) return `<div class="card"><div class="empty">Person not found.</div></div>`;
  const committed=personAssignments.filter(a=>a.committed), planned=personAssignments.filter(a=>!a.committed);
  const weekTotals=WEEKS.map(w=>({w,t:personAssignments.reduce((s,a)=>s+getAlloc(a,w),0)}));
  const calRows=[]; for(let i=0;i<52;i+=13) calRows.push(weekTotals.slice(i,i+13));
  function cellStyle(d){const cur=d.w===CURRENT_WEEK,outline=cur?'outline:2px solid #1D9E75;outline-offset:-2px;':'';if(d.t>100)return `background:#fecaca;color:#b91c1c;font-weight:700;${outline}`;if(d.t===100)return `background:#d1fae5;color:#065f46;font-weight:700;${outline}`;if(d.t>0)return `background:#fef3c7;color:#92400e;font-weight:600;${outline}`;return `background:${cur?'rgba(29,158,117,0.07)':'#f9fafb'};color:#d1d5db;${outline}`;}
  const calTable=`<div style="overflow-x:auto"><table style="border-collapse:separate;border-spacing:2px;min-width:100%">${calRows.map(chunk=>`<tr>${chunk.map(d=>`<td style="text-align:center;padding:4px 2px;font-size:10px;font-family:'DM Mono',monospace;border-radius:3px;min-width:34px;${cellStyle(d)}"><div style="font-size:9px;opacity:.7;line-height:1">${d.w===CURRENT_WEEK?'▼':''}W${d.w}</div><div style="line-height:1.3">${d.t>0?d.t+'%':'–'}</div></td>`).join('')}</tr>`).join('')}</table></div>`;
  function assignmentCard(a){const idx=state.assignments.indexOf(a),totalWks=a.periods.reduce((s,p)=>s+(p.endWeek-p.startWeek+1),0);return `<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;margin-bottom:10px"><div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px"><div style="flex:1"><div style="display:flex;align-items:center;gap:8px;margin-bottom:4px"><span style="font-size:13px;font-weight:700">${a.workName}</span><span class="badge b-type">${a.type}</span>${a.committed?`<span class="status-committed">✓ Committed</span>`:`<span class="status-planned">⏳ Planned</span>`}</div>${a.committed?`<div style="font-size:11px;color:#9ca3af;margin-bottom:6px">Committed by ${a.committedBy}</div>`:''}<div style="display:flex;flex-wrap:wrap;gap:4px">${a.periods.map((p,pi)=>`<span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%${ce?` <span onclick="delPeriod(${idx},${pi})" style="cursor:pointer;opacity:.5;font-size:10px">✕</span>`:''}</span>`).join('')}</div><div style="font-size:11px;color:#9ca3af;margin-top:6px">${totalWks} week${totalWks!==1?'s':''} total</div></div><div style="display:flex;flex-direction:column;gap:6px;flex-shrink:0">${!a.committed&&ce?`<button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button>`:''}${a.committed&&ce?`<button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${idx})">↩ Uncommit</button>`:''}${ce?`<button class="btn danger sm" onclick="delAssignment(${idx})">🗑 Delete</button>`:''}</div></div></div>`;}
  return `<div style="margin-bottom:12px"><button class="btn sm" onclick="setTab('overview')">← Back to Overview</button></div>
    <div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">👤 ${name}</span><div style="display:flex;gap:14px;font-size:12px;color:#6b7280;align-items:center;flex-wrap:wrap"><span>${info.team||'—'}</span><span>${info.skillset||'—'}</span><span>${info.level||'—'}</span><span>${info.country||'—'}</span>${getTeamlead(name)?`<span>TL: <strong>${getTeamlead(name)}</strong></span>`:''}${getManager(name)?`<span>Mgr: <strong>${getManager(name)}</strong></span>`:''}<span style="color:#0f6e56;font-weight:600">${committed.length} committed</span><span style="color:#b45309;font-weight:600">${planned.length} planned</span></div></div></div>
    <div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">📅 Full year allocation</span><span class="card-sub">▼ = current week</span></div><div class="card-body">${calTable}</div></div>
    ${committed.length?`<div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title" style="color:#0f6e56">✓ Committed (${committed.length})</span></div><div class="card-body">${committed.map(assignmentCard).join('')}</div></div>`:''}
    ${planned.length?`<div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title" style="color:#b45309">⏳ Planned (${planned.length})</span></div><div class="card-body">${planned.map(assignmentCard).join('')}</div></div>`:''}
    ${!personAssignments.length?`<div class="card"><div class="empty"><span class="empty-icon">📋</span>No assignments for ${name} yet.</div></div>`:''}`;
}

function renderOverview(){
  const allPeople=getPeople(),conf=state.assignments.filter(a=>a.committed).length,plan=state.assignments.filter(a=>!a.committed).length,r=role(),ce=canEdit();
  const teams=[...new Set(['Development','Platform','PMO',...allPeople.map(p=>p.team)].filter(Boolean))].sort();
  const levels=[...new Set(allPeople.map(p=>p.level).filter(Boolean))].sort();
  const assignmentNames=[...new Set(state.assignments.map(a=>a.workName).filter(Boolean))].sort();
  let people=allPeople.filter(p=>{const n=state.fName.trim().toLowerCase(),s=state.fSkill.trim().toLowerCase();if(state.fTeam&&p.team!==state.fTeam)return false;if(n&&!p.name.toLowerCase().includes(n))return false;if(s&&!p.skillset.toLowerCase().includes(s))return false;if(state.fLevel&&p.level!==state.fLevel)return false;if(state.fAssignment&&!visibleAssignments().some(a=>a.name===p.name&&a.workName===state.fAssignment))return false;if(state.fStatus){const pa=visibleAssignments().filter(a=>a.name===p.name);if(state.fStatus==='planned'&&!pa.some(a=>!a.committed))return false;if(state.fStatus==='committed'&&!pa.some(a=>a.committed))return false;if(state.fStatus==='overbooked'&&!WEEKS.some(w=>getTotalAlloc(p.name,w)>100))return false;}return true;});
  const activeFilters=[state.fTeam,state.fName,state.fSkill,state.fLevel,state.fAssignment,state.fStatus].filter(Boolean).length;
  const filterBar=`<div class="filter-bar">${r!=='Team Member'?`<div><label class="filter-lbl">Name</label><select class="filter-inp" style="width:150px" onchange="state.fName=this.value;render()"><option value="">All people</option>${getAllPeople().map(p=>`<option value="${p.name}"${state.fName===p.name?' selected':''}>${p.name}</option>`).join('')}</select></div>`:''}<div><label class="filter-lbl">Team</label><select class="filter-inp" onchange="state.fTeam=this.value;render()"><option value="">All teams</option>${teams.map(t=>`<option value="${t}"${state.fTeam===t?' selected':''}>${t}</option>`).join('')}</select></div><div><label class="filter-lbl">Skill</label><input class="filter-inp" style="width:120px" placeholder="Filter skill…" value="${state.fSkill}" oninput="setFilter('fSkill',this.value)" /></div><div><label class="filter-lbl">Level</label><select class="filter-inp" onchange="state.fLevel=this.value;render()"><option value="">All levels</option>${levels.map(l=>`<option value="${l}"${state.fLevel===l?' selected':''}>${l}</option>`).join('')}</select></div><div><label class="filter-lbl">Assignment</label><select class="filter-inp" style="max-width:180px" onchange="state.fAssignment=this.value;render()"><option value="">All assignments</option>${assignmentNames.map(a=>`<option value="${a}"${state.fAssignment===a?' selected':''}>${a}</option>`).join('')}</select></div><div><label class="filter-lbl">Status</label><select class="filter-inp" onchange="state.fStatus=this.value;render()"><option value="">All</option><option value="planned"${state.fStatus==='planned'?' selected':''}>⏳ Planned only</option><option value="committed"${state.fStatus==='committed'?' selected':''}>✓ Has committed</option><option value="overbooked"${state.fStatus==='overbooked'?' selected':''}>🔴 Overbooked</option></select></div>${activeFilters>0?`<button class="btn sm" style="margin-top:16px" onclick="state.fTeam='';state.fName='';state.fSkill='';state.fLevel='';state.fAssignment='';state.fStatus='';render()">✕ Clear (${activeFilters})</button>`:''}</div>`;
  const wks=visibleWeeks();
  const rows=!people.length?`<div class="empty"><span class="empty-icon">👥</span>${allPeople.length?'No results match your filters.':'No allocations yet. Go to Planning mode to get started.'}</div>`:`<div class="tbl-wrap"><table><thead><tr><th>Name</th><th>Skill</th><th>Level</th><th>Country</th><th>Team</th><th>Status</th>${wks.map(w=>wkHdr(w)).join('')}</tr></thead><tbody>${people.map(p=>{const dn=(r==='Project Manager'&&visibleAssignments().some(a=>a.name===p.name&&!a.committed))?'— Planned resource —':p.name,pa=visibleAssignments().filter(a=>a.name===p.name),hasCom=pa.some(a=>a.committed),hasPlan=pa.some(a=>!a.committed),statusBadge=hasCom&&hasPlan?`<span class="status-mixed">Mixed</span>`:hasCom?`<span class="status-committed">✓ Committed</span>`:hasPlan?`<span class="status-planned">⏳ Planned</span>`:`<span style="font-size:10px;color:#d1d5db">–</span>`;return `<tr class="${ce?'person-row-click':''}" ${ce?`onclick="openPersonDetail('${p.name.replace(/'/g,"\\'")}')"`:''}><td style="background:${cBg(p.country)}"><strong>${dn}</strong>${ce?`<span style="font-size:10px;color:#9ca3af;margin-left:6px">→</span>`:''}</td><td>${p.skillset}</td><td>${p.level}</td><td>${p.country}</td><td><span style="cursor:pointer;color:#1D9E75;text-decoration:underline" onclick="event.stopPropagation();openTeam('${p.team}')">${p.team}</span></td><td>${statusBadge}</td>${wks.map(w=>wkCell(w,getTotalAlloc(p.name,w))).join('')}</tr>`;}).join('')}</tbody></table></div>`;
  return `<div class="metrics"><div class="metric"><div class="metric-lbl">People</div><div class="metric-val">${allPeople.length}</div></div><div class="metric"><div class="metric-lbl">Assignments</div><div class="metric-val">${state.assignments.length}</div></div><div class="metric"><div class="metric-lbl">Committed</div><div class="metric-val green">${conf}</div></div><div class="metric"><div class="metric-lbl">Planned</div><div class="metric-val" style="color:#b45309">${plan}</div></div></div>
    <div class="card"><div class="card-hdr"><span class="card-title">📅 Total allocation per person</span><div style="display:flex;align-items:center;gap:12px">${weekRangeToggle()}<span class="card-sub">${activeFilters>0?`${people.length} of ${allPeople.length} shown`:''}</span></div></div>${filterBar}${rows}</div>`;
}

function renderProjects(){
  const ce=canEdit(),r=role(),visible=r==='Project Manager'?state.projects.filter(p=>isPmProject(p)):state.projects;
  const list=visible.length?(()=>{
    const ongoing=[],planned=[],noRes=[];
    visible.forEach(p=>{const allA=state.assignments.filter(a=>a.workName===p.name),hasAsgOngoing=allA.some(a=>a.periods.some(per=>CURRENT_WEEK>=per.startWeek&&CURRENT_WEEK<=per.endWeek)),startWeek=p.startDate?Math.ceil((new Date(p.startDate)-new Date(new Date().getFullYear(),0,1))/604800000):null,endWeek=p.endDate?Math.ceil((new Date(p.endDate)-new Date(new Date().getFullYear(),0,1))/604800000):null,hasDateOngoing=startWeek!==null&&endWeek!==null&&CURRENT_WEEK>=startWeek&&CURRENT_WEEK<=endWeek,isOngoing=hasDateOngoing||hasAsgOngoing,hasAny=allA.length>0;if(isOngoing)ongoing.push(p);else if(hasAny)planned.push(p);else noRes.push(p);});
    function renderProject(p){const allA=state.assignments.filter(a=>a.workName===p.name),vis=r==='Team Member'?allA.filter(a=>a.committed&&a.name.trim().toLowerCase()===userName().trim().toLowerCase()):allA,comm=vis.filter(a=>a.committed).length,pl=r!=='Team Member'?vis.filter(a=>!a.committed).length:0,hasAsgOngoing=allA.some(a=>a.periods.some(per=>CURRENT_WEEK>=per.startWeek&&CURRENT_WEEK<=per.endWeek)),startWeek=p.startDate?Math.ceil((new Date(p.startDate)-new Date(new Date().getFullYear(),0,1))/604800000):null,endWeek=p.endDate?Math.ceil((new Date(p.endDate)-new Date(new Date().getFullYear(),0,1))/604800000):null,hasDateOngoing=startWeek!==null&&endWeek!==null&&CURRENT_WEEK>=startWeek&&CURRENT_WEEK<=endWeek,isOngoing=hasDateOngoing||hasAsgOngoing,rowBg=isOngoing?'#f0fdf4':'#fff',ongoingBadge=isOngoing?`<span style="background:#d1fae5;color:#065f46;padding:2px 10px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:.05em;text-transform:uppercase">● Ongoing</span>`:'';return `<div class="project-row" style="cursor:pointer;transition:box-shadow .15s;background:${rowBg};border-left:${isOngoing?'3px solid #1D9E75':'3px solid transparent'}" onmouseover="this.style.boxShadow='0 2px 8px rgba(0,0,0,0.08)'" onmouseout="this.style.boxShadow=''" onclick="openProject(${p.id})"><div style="flex:1"><div style="font-size:13px;font-weight:700;color:#111827;margin-bottom:2px;display:flex;align-items:center;gap:8px">${p.name} ${ongoingBadge}<span style="font-size:11px;color:#9ca3af;font-weight:400">→ click to view</span></div><div style="font-size:11px;color:#9ca3af;margin-top:2px">📅 ${fmtDateLong(p.startDate)} → ${fmtDateLong(p.endDate)}${p.projectManager?` &nbsp;·&nbsp; PM: <strong style="color:#374151">${p.projectManager}</strong>`:''}</div>${p.description?`<div style="font-size:11px;color:#6b7280;margin-top:3px;max-width:400px;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical">${p.description}</div>`:''}</div><div style="display:flex;gap:8px;align-items:center">${comm>0?`<span style="background:#d1fae5;color:#065f46;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">✓ ${comm} committed</span>`:''}${pl>0?`<span style="background:#fef3c7;color:#92400e;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">⏳ ${pl} planned</span>`:''}${!comm&&!pl?`<span style="color:#d1d5db;font-size:11px">No resources yet</span>`:''}${ce?`<button class="btn danger sm" style="margin-left:4px" onclick="event.stopPropagation();deleteProject(${p.id})">🗑</button>`:''}</div></div>`;}
    const section=(label,items)=>items.length?`<div style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.08em;padding:14px 18px 6px">${label} (${items.length})</div>${items.map(renderProject).join('')}`:'';
    return section('🟢 Ongoing',ongoing)+section('⏳ Planned',planned)+section('📭 No resources',noRes);
  })():r==='Project Manager'?`<div class="empty" style="padding:24px 0"><span class="empty-icon">💼</span>No projects assigned to you yet.</div>`:`<div class="empty" style="padding:24px 0"><span class="empty-icon">💼</span>No projects yet.</div>`;
  return `<div class="card"><div class="card-hdr"><span class="card-title">💼 Projects</span></div><div class="card-body">${ce?`<div class="ibox"><div class="sec-title">Add project</div><div class="frow"><div class="fg"><label class="lbl">Project name</label><input class="inp" placeholder="e.g. Platform Renewal" value="${state.pName}" oninput="state.pName=this.value" onkeydown="if(event.key==='Enter')addProject()" /></div><div class="fg"><label class="lbl">Project manager</label>${peopleSelectOptional('add-pm',state.pPm,'state.pPm=this.value','')}</div><div class="fg"><label class="lbl">Start date</label><input class="inp" type="date" value="${state.pStart}" oninput="state.pStart=this.value" /></div><div class="fg"><label class="lbl">End date</label><input class="inp" type="date" value="${state.pEnd}" oninput="state.pEnd=this.value" /></div><div class="fg" style="grid-column:1/-1"><label class="lbl">Description <span style="color:#9ca3af;font-weight:400">(optional)</span></label><textarea class="inp" rows="2" style="resize:vertical" oninput="state.pDesc=this.value">${state.pDesc}</textarea></div><div style="padding-top:4px"><button class="btn primary" onclick="addProject()">＋ Add project</button></div></div></div>`:''}${list}</div></div>`;
}

function renderProjectDetail(){
  const proj=state.projects.find(p=>p.id===state.selectedProject);
  if(!proj) return `<div class="card"><div class="empty"><span class="empty-icon">💼</span>Project not found.</div></div>`;
  const r=role(),ce=canEdit();
  if(r==='Project Manager'&&!isPmProject(proj)) return `<div class="card"><div class="empty"><span class="empty-icon">🔒</span>You do not have access to this project.</div></div>`;
  const allA=state.assignments.filter(a=>a.workName===proj.name);
  const visA=allA.filter(a=>{if(r==='Team Member') return a.committed&&a.name.trim().toLowerCase()===userName().trim().toLowerCase(); return true;});
  const planned=visA.filter(a=>!a.committed),committed=visA.filter(a=>a.committed);
  const wks=visibleWeeks();
  return `<div style="margin-bottom:12px"><button class="btn sm" onclick="setTab('projects')">← Back to Projects</button></div>
    <div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">💼 ${proj.name}</span><div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap"><span style="font-size:12px;color:#6b7280">📅 <strong>${proj.startDate||'—'}</strong> → <strong>${proj.endDate||'—'}</strong></span>${proj.projectManager?`<span style="font-size:12px;color:#6b7280">PM: <strong>${proj.projectManager}</strong></span>`:''}<span style="color:#b45309;font-weight:600;font-size:12px">${planned.length} planned</span><span style="color:#0f6e56;font-weight:600;font-size:12px">${committed.length} committed</span>${ce?`<button class="btn sm" onclick="toggleEditProject()">✏ Edit</button>`:''}${ce?`<button class="btn danger sm" onclick="if(confirm('Delete project and all its planning?')){state.projects=state.projects.filter(p=>p.id!==${proj.id});state.assignments=state.assignments.filter(a=>a.workName!=='${proj.name.replace(/'/g,"\\'")}');setTab('projects')}">🗑 Delete</button>`:''}</div>${proj.description?`<div style="padding:10px 18px;font-size:13px;color:#374151;border-top:1px solid #f3f4f6;background:#fafafa;line-height:1.6">${proj.description}</div>`:''}</div>${ce?`<div id="edit-proj-panel" style="display:none;padding:16px;border-top:1px solid #e5e7eb;background:#f9fafb"><div class="fgrid"><div class="fg"><label class="lbl">Project name</label><input class="inp" id="ep-name" value="${proj.name}" /></div><div class="fg"><label class="lbl">Project manager</label>${peopleSelectOptional('ep-pm',proj.projectManager||'','document._epPm=this.value','')}</div><div class="fg"><label class="lbl">Start date</label><input class="inp" type="date" id="ep-start" value="${proj.startDate||''}" /></div><div class="fg"><label class="lbl">End date</label><input class="inp" type="date" id="ep-end" value="${proj.endDate||''}" /></div><div class="fg" style="grid-column:1/-1"><label class="lbl">Description</label><textarea class="inp" id="ep-desc" rows="2" style="resize:vertical">${proj.description||''}</textarea></div><div class="fg" style="justify-content:flex-end;padding-top:4px"><button class="btn primary" onclick="saveProjectEdits(${proj.id})">Save changes</button></div></div></div>`:''}</div>
    ${(ce||(r==='Project Manager'&&isPmProject(proj)))?`<div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">＋ Add resource to project</span></div><div class="card-body" style="display:flex;flex-direction:column;gap:16px"><div class="fgrid">${r!=='Project Manager'?`<div class="fg"><label class="lbl">Full name *</label>${peopleSelect('inp-prName',state.prName,"onPersonInput(this.value,'pr');clearTimeout(_debounceTimer);render()",'')}</div>`:''}<div class="fg"><label class="lbl">Team</label><select class="sel" onchange="state.prTeam=this.value">${TEAMS.map(t=>`<option${state.prTeam===t?' selected':''}>${t}</option>`).join('')}</select></div><div class="fg"><label class="lbl">Country</label><select class="sel" onchange="state.prCountry=this.value"><option value="Sweden"${state.prCountry==='Sweden'?' selected':''}>Sweden</option><option value="Poland"${state.prCountry==='Poland'?' selected':''}>Poland</option></select></div><div class="fg"><label class="lbl">Skillset *</label><input class="inp" placeholder="e.g. React, DevOps" value="${state.prSkill}" oninput="state.prSkill=this.value" /></div><div class="fg"><label class="lbl">Level</label><select class="sel" onchange="state.prLevel=this.value">${['Junior','Mid','Senior'].map(l=>`<option${state.prLevel===l?' selected':''}>${l}</option>`).join('')}</select></div></div><div class="frow"><div class="fg"><label class="lbl">From week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.prStart}" oninput="state.prStart=+this.value" /></div><div class="arrow">→</div><div class="fg"><label class="lbl">To week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.prEnd}" oninput="state.prEnd=+this.value" /></div><div class="fg"><label class="lbl">Allocation</label><div style="display:flex;align-items:center;gap:4px"><input class="inp narrow" type="number" min="0" max="200" value="${state.prPct}" oninput="state.prPct=+this.value" /><span style="font-size:13px;color:#6b7280">%</span></div></div><div class="fg" style="justify-content:flex-end;padding-top:18px"><button class="btn primary" onclick="addResourceToProject()">＋ Add resource</button></div></div></div></div>`:''}
    <div class="card"><div class="card-hdr"><span class="card-title">📅 Weekly allocation</span><div style="display:flex;align-items:center;gap:12px">${weekRangeToggle()}<span class="card-sub">Green = person total · White = per-period</span></div></div>
    ${!visA.length?`<div class="empty"><span class="empty-icon">👤</span>No resources assigned yet.</div>`:`<div class="tbl-wrap"><table><thead><tr><th>Resource</th><th>Skill</th><th>Level</th><th>Country</th><th>Status</th>${wks.map(w=>wkHdr(w)).join('')}${ce?'<th></th>':''}</tr></thead><tbody>${visA.map(a=>{const idx=state.assignments.indexOf(a),rawN=a.name&&a.name.startsWith('__pm_planned__')?'No name':a.name,showN=r==='Project Manager'?(a.committed?rawN:'No name'):rawN,statusBadge=a.committed?`<span class="badge b-committed" style="white-space:nowrap">✓ Committed</span><div style="font-size:10px;color:#9ca3af;margin-bottom:4px">${a.committedBy}</div>${ce?`<button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${idx})">↩ Uncommit</button>`:''}`:`<span class="badge b-plan">Planned</span>${ce?`<div style="margin-top:4px"><button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button></div>`:''}`,weekCells=wks.map(w=>wkCell(w,getEffectiveAlloc(a,w))).join(''),periodRows=a.periods.map((p,pi)=>{const pW=wks.map(w=>{const inR=w>=p.startWeek&&w<=p.endWeek;return `<td class="wk" style="font-size:10px;${inR?'background:rgba(29,158,117,0.08);color:#0f6e56':''}">${inR?p.allocationPercent+'%':''}</td>`;}).join('');return `<tr style="background:#fafafa"><td colspan="5" style="padding:4px 12px 4px 28px;font-size:11px;color:#6b7280">Period ${pi+1}: W${p.startWeek}–${p.endWeek} · ${p.allocationPercent}%  ${ce?`<button class="btn danger sm" style="padding:2px 6px;font-size:10px" onclick="delPeriod(${idx},${pi})">🗑</button>`:''}</td>${pW}${ce?'<td></td>':''}</tr>`;}).join('');return `<tr style="background:var(--green-bg);border-top:2px solid #e5e7eb"><td style="padding:10px 12px;font-size:13px;font-weight:700">${showN}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.skillset}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.level}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${a.country}</td><td style="padding:10px 12px">${statusBadge}</td>${weekCells}${ce?`<td style="padding:10px 8px"><button class="btn danger sm" onclick="delAssignment(${idx})">🗑 Remove</button></td>`:''}</tr>${periodRows}`;}).join('')}</tbody></table></div>`}</div>`;
}

function renderTeamDetail(){
  const teamName=state.selectedTeam,r=role(),ce=canEdit();
  visibleAssignments().filter(a=>a.team===teamName).forEach(a=>{const key=a.name.trim().toLowerCase();if(!state.teamMembers.some(m=>m.name.trim().toLowerCase()===key&&m.team===teamName))state.teamMembers.push({id:Date.now()+Math.random(),name:a.name.trim(),team:teamName,country:a.country,skillset:a.skillset,level:a.level});});
  const peopleMap=new Map();
  state.teamMembers.filter(m=>m.team===teamName).forEach(m=>peopleMap.set(m.name.trim().toLowerCase(),{id:m.id,name:m.name,team:m.team,country:m.country,skillset:m.skillset,level:m.level,assignments:[],registered:true}));
  visibleAssignments().filter(a=>a.team===teamName).forEach(a=>{const key=a.name.trim().toLowerCase();if(peopleMap.has(key)&&!peopleMap.get(key).assignments.find(x=>x.id===a.id))peopleMap.get(key).assignments.push(a);});
  // Sort: Base Services last
  peopleMap.forEach(p => p.assignments.sort((a,b) => (a.type==='Base Service'?1:0)-(b.type==='Base Service'?1:0)));
  const people=[...peopleMap.values()];
  const filteredPeople=state.teamFilterNames.size>0?people.filter(p=>state.teamFilterNames.has(p.name)):people;
  const totalCommitted=people.reduce((s,p)=>s+p.assignments.filter(a=>a.committed).length,0);
  const totalPlanned=people.reduce((s,p)=>s+p.assignments.filter(a=>!a.committed).length,0);
  const filterBar=`<div style="padding:12px 18px;border-bottom:1px solid #f3f4f6;display:flex;flex-wrap:wrap;gap:8px;align-items:center"><span style="font-size:11px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap">Filter people:</span>${people.map(p=>{const active=state.teamFilterNames.has(p.name);return `<button onclick="toggleTeamFilter('${p.name.replace(/'/g,"\\'")}');event.stopPropagation()" style="padding:3px 12px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;border:1px solid ${active?'#1D9E75':'#e5e7eb'};background:${active?'#1D9E75':'#fff'};color:${active?'#fff':'#6b7280'};transition:all .15s">${p.name}</button>`;}).join('')}${state.teamFilterNames.size>0?`<button onclick="state.teamFilterNames=new Set();render()" style="padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;cursor:pointer;font-family:inherit;border:1px solid #fca5a5;background:#fef2f2;color:#b91c1c">✕ Clear filter (${state.teamFilterNames.size})</button>`:''}</div>`;
  function ptw(person,w){return person.assignments.filter(a=>a.committed).reduce((s,a)=>s+getAlloc(a,w),0);}
  function wCls(t){return t>100?'ao':t===100?'af':t>0?'ap':'';}
  const wks=visibleWeeks();
  function assignmentRows(person){
    if(!person.assignments.length) return `<tr><td colspan="${6+wks.length}" style="padding:6px 12px;font-size:11px;color:#9ca3af;font-style:italic">No assignments yet</td></tr>`;
    return person.assignments.map(a=>{
      const label=r==='Project Manager'&&!a.committed?'— Planned —':a.workName,idx=state.assignments.indexOf(a),isEditingA=ce&&state.editingAssignmentId===idx;
      const dot=a.committed?`<span style="background:#d1fae5;color:#065f46;padding:1px 7px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap">✓ Committed</span>${ce?`<button class="btn danger sm" style="font-size:10px;padding:1px 6px;margin-left:4px" onclick="uncommitA(${idx})">↩</button>`:''}`:`<span style="background:#fef3c7;color:#92400e;padding:1px 7px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap">⏳ Planned</span>${ce?`<button class="btn primary sm" style="font-size:10px;padding:1px 6px;margin-left:4px" onclick="commitA(${idx})">🔒 Commit</button>`:''}`;
      const wCells=wks.map(w=>{const al=getEffectiveAlloc(a,w);return `<td class="wk ${a.committed?wCls(al):''}" style="font-size:10px;${!a.committed?'color:#c4c9d4':''}">${al>0?al+'%':''}</td>`;}).join('');
      const editPanel=isEditingA?`<tr style="background:#f0fdf8;border-bottom:2px solid #e5e7eb"><td colspan="${6+wks.length}" style="padding:0"><div style="padding:12px 16px"><div style="font-size:11px;font-weight:700;color:#0f6e56;margin-bottom:10px">✏ ${a.workName}</div>${a.periods.length?`<div style="margin-bottom:10px"><div style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">Current periods</div><div style="display:flex;flex-wrap:wrap;gap:6px">${a.periods.map((p,pi)=>`<span style="background:#e0f2fe;color:#075985;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}% <span onclick="delPeriod(${idx},${pi});event.stopPropagation()" style="cursor:pointer;opacity:.6;margin-left:2px">✕</span></span>`).join('')}</div></div>`:''}<div style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Add period</div><div style="display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap"><div class="fg"><label class="lbl">From week</label><input class="inp narrow" type="number" min="1" max="52" id="ea-start-${idx}" value="${state.eaStart}" oninput="state.eaStart=+this.value" /></div><div style="padding-bottom:8px;color:#9ca3af">→</div><div class="fg"><label class="lbl">To week</label><input class="inp narrow" type="number" min="1" max="52" id="ea-end-${idx}" value="${state.eaEnd}" oninput="state.eaEnd=+this.value" /></div><div class="fg"><label class="lbl">Allocation %</label><input class="inp narrow" type="number" min="0" max="200" id="ea-pct-${idx}" value="${state.eaPct}" oninput="state.eaPct=+this.value" /></div><div style="display:flex;gap:6px;padding-bottom:4px"><button class="btn primary sm" onclick="addPeriodToAssignment(${idx})">＋ Add period</button><button class="btn sm" onclick="state.editingAssignmentId=null;render()">✕ Close</button>${a.committed?`<button class="btn danger sm" onclick="uncommitA(${idx})">↩ Uncommit</button>`:`<button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button>`}<button class="btn danger sm" onclick="if(confirm('Delete this assignment?')){delAssignment(${idx})}">🗑 Delete</button></div></div></div></td></tr>`:'';
      const rowBg=a.committed?'#fafafa':'#fefefe',rowColor=a.committed?'#374151':'#c4c9d4';
      return `<tr style="background:${rowBg};${ce?'cursor:pointer':''}" ${ce?`onclick="startEditAssignment(${idx})"`:''}}><td style="padding:5px 12px 5px 28px;font-size:12px;color:${rowColor};white-space:nowrap">${label}${ce?` <span style="font-size:10px;color:#9ca3af">✏</span>`:''}</td><td style="padding:5px 12px;font-size:11px;color:#9ca3af">${a.type}</td><td colspan="3" style="padding:5px 12px;white-space:nowrap">${dot}</td><td style="padding:5px 12px"></td>${wCells}${ce?`<td style="padding:5px 8px"></td>`:''}</tr>${editPanel}`;
    }).join('');
  }
  const memberRows=filteredPeople.map(person=>{
    const allT=WEEKS.map(w=>ptw(person,w)),wCells=wks.map(w=>{const t=allT[WEEKS.indexOf(w)];return `<td class="wk ${wCls(t)}" style="font-weight:700">${t>0?t+'%':'–'}</td>`;}).join(''),mbadge=person.registered?`<span style="background:#e0f2fe;color:#075985;font-size:10px;font-weight:700;padding:1px 6px;border-radius:20px;margin-left:6px">Member</span>`:`<span style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;padding:1px 6px;border-radius:20px;margin-left:6px">Planning</span>`,isEditing=ce&&person.registered&&state.editingMemberId===person.id;
    const editRow=isEditing?`<tr style="background:var(--green-bg)"><td colspan="${6+wks.length}" style="padding:0"><div class="edit-member-panel"><div style="font-size:11px;font-weight:700;color:#0f6e56;margin-bottom:12px;text-transform:uppercase;letter-spacing:.05em">✏ Editing: ${person.name}</div><div class="fgrid"><div class="fg"><label class="lbl">Full name</label><input class="inp" id="em-name-${person.id}" placeholder="${state.emName}" oninput="state.emName=this.value" /></div><div class="fg"><label class="lbl">Country</label><select class="sel" onchange="state.emCountry=this.value"><option value="Sweden"${state.emCountry==='Sweden'?' selected':''}>Sweden</option><option value="Poland"${state.emCountry==='Poland'?' selected':''}>Poland</option></select></div><div class="fg"><label class="lbl">Team</label><select class="sel" onchange="state.emTeam=this.value">${TEAMS.map(t=>`<option${state.emTeam===t?' selected':''} value="${t}">${t}</option>`).join('')}</select></div><div class="fg"><label class="lbl">Skillset</label><input class="inp" id="em-skill-${person.id}" placeholder="${state.emSkill}" oninput="state.emSkill=this.value" /></div><div class="fg"><label class="lbl">Level</label><select class="sel" onchange="state.emLevel=this.value">${['Junior','Mid','Senior'].map(l=>`<option${state.emLevel===l?' selected':''}>${l}</option>`).join('')}</select></div><div class="fg"><label class="lbl">Teamlead</label><input class="inp" id="em-tl-${person.id}" placeholder="${state.emTeamlead||'inherit from team'}" list="people-list-optional" autocomplete="off" oninput="state.emTeamlead=this.value" /></div><div class="fg"><label class="lbl">Manager</label><input class="inp" id="em-mgr-${person.id}" placeholder="${state.emManager||'inherit from team'}" list="people-list-optional" autocomplete="off" oninput="state.emManager=this.value" /></div><div style="display:flex;gap:8px;padding-top:4px;align-items:flex-end;grid-column:1/-1"><button class="btn primary sm" onclick="saveMemberEditFromInputs(${person.id})">✓ Save</button><button class="btn sm" onclick="state.editingMemberId=null;render()">✕ Cancel</button><button class="btn danger sm" onclick="removeTeamMember(${person.id})">🗑 Remove</button></div></div></div></td></tr>`:'';
    return `<tr style="background:var(--green-bg);border-top:2px solid #e5e7eb"><td style="padding:10px 12px;font-size:13px;font-weight:700;white-space:nowrap;${ce&&person.registered?'cursor:pointer':''}" ${ce&&person.registered?`onclick="startEditMember(${person.id})"`:''}}>${person.name}${mbadge}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.skillset}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.level}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.country}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${getTeamlead(person.name)?`<div style="font-size:10px;color:#0f6e56;font-weight:700">TL: ${getTeamlead(person.name)}</div>`:''}${getManager(person.name)?`<div style="font-size:10px;color:#185fa5;font-weight:700">Mgr: ${getManager(person.name)}</div>`:''}${!getTeamlead(person.name)&&!getManager(person.name)?`<span style="color:#d1d5db;font-size:11px">–</span>`:''}</td><td style="padding:10px 12px;font-size:12px;color:#6b7280">${person.assignments.length} asgmt${person.assignments.length!==1?'s':''}</td>${wCells}<td style="padding:10px 8px"></td></tr>${editRow}${assignmentRows(person)}`;
  }).join('');
  const svcsWithTarget=state.baseServices.filter(s=>s.team===teamName&&s.targetPct>0);
  const debtSection=svcsWithTarget.length?(()=>{
    const td=calcTeamDebt(teamName);
    const open=state.debtSectionOpen;
    const svcRows=svcsWithTarget.map(s=>{const debt=calcSvcDebt(s),curAlloc=getSvcAlloc(s.name,CURRENT_WEEK),onTrack=curAlloc>=s.targetPct,dc=debt.debtPct===0?'#0f6e56':debt.debtPct<s.targetPct*4?'#b45309':'#b91c1c',db=debt.debtPct===0?'#d1fae5':debt.debtPct<s.targetPct*4?'#fef3c7':'#fef2f2',barPct=s.targetPct*(CURRENT_WEEK-1)>0?Math.min(100,Math.round((debt.debtPct/(s.targetPct*(CURRENT_WEEK-1)))*100)):0;return `<div style="padding:12px 18px;border-bottom:1px solid #f3f4f6"><div style="display:flex;align-items:center;gap:12px"><div style="flex:1"><div style="font-size:13px;font-weight:600;color:#111827">${s.name}</div><div style="font-size:11px;color:#6b7280;margin-top:2px">Target <strong>${s.targetPct}%</strong> · Now <strong style="color:${onTrack?'#0f6e56':'#b91c1c'}">${curAlloc}%</strong> ${onTrack?'✓':('↓ '+(s.targetPct-curAlloc)+'% below target')}</div><div style="margin-top:6px;height:5px;background:#f3f4f6;border-radius:3px;overflow:hidden;max-width:300px"><div style="height:100%;width:${barPct}%;background:${dc};border-radius:3px"></div></div></div><div style="text-align:right;flex-shrink:0"><span style="background:${db};color:${dc};padding:2px 10px;border-radius:20px;font-size:12px;font-weight:700">${debt.debtPct===0?'✓ No debt':debt.debtPct+'%'}</span>${debt.debtPct>0?`<div style="font-size:10px;color:#9ca3af;margin-top:4px">= ${debt.debtWeeks} weeks fulltime</div>`:''}</div></div></div>`;}).join('');
    return `<div class="card" style="margin-bottom:16px">
      <div class="card-hdr" onclick="state.debtSectionOpen=!state.debtSectionOpen;render()" style="cursor:pointer;user-select:none">
        <div style="display:flex;align-items:center;gap:8px">
          <span style="font-size:13px">${open?'▾':'▸'}</span>
          <span class="card-title">🔧 Technical debt — Base Services</span>
          <span class="card-sub">Accumulated W1–W${CURRENT_WEEK-1} · ${svcsWithTarget.length} services</span>
        </div>
        <span style="font-size:13px;font-weight:700;color:${td.debtPct===0?'#0f6e56':td.debtPct<400?'#b45309':'#b91c1c'}">${td.debtPct===0?'✓ No debt':td.debtPct+'% · '+td.debtWeeks+'w'}</span>
      </div>
      ${open?`${svcRows}<div style="padding:10px 18px;background:#f9fafb;border-top:1px solid #f3f4f6;display:flex;align-items:center;justify-content:space-between"><span style="font-size:12px;color:#6b7280">${svcsWithTarget.length} services tracked</span><span style="font-size:13px;font-weight:700;color:${td.debtPct===0?'#0f6e56':td.debtPct<400?'#b45309':'#b91c1c'}">Total: ${td.debtPct===0?'✓ No debt':td.debtPct+'% · '+td.debtWeeks+'w'}</span></div>`:''}
    </div>`;
  })():'';
  const cfg=state.teamConfig[teamName]||{};
  return `<div style="margin-bottom:12px"><button class="btn sm" onclick="setTab('overview')">← Back to Overview</button></div>
    <div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">👥 ${teamName} Team</span><div style="display:flex;gap:16px;font-size:12px;color:#6b7280"><span><strong style="color:#111827">${people.length}</strong> member${people.length!==1?'s':''}</span><span><strong style="color:#0f6e56">${totalCommitted}</strong> committed</span><span><strong style="color:#b45309">${totalPlanned}</strong> planned</span></div></div><div style="padding:14px 18px;border-top:1px solid #f3f4f6;display:grid;grid-template-columns:1fr 1fr;gap:12px"><div><div class="org-label">Team Lead</div>${ce?`<div style="display:flex;align-items:center;gap:8px">${peopleSelectOptional('tc-tl-'+teamName,cfg.teamlead||'','saveTeamConfig(\''+teamName+'\',\'teamlead\',this.value)','flex:1;max-width:220px')}${cfg.teamlead?`<span style="font-size:11px;color:#0f6e56;font-weight:600">✓ ${cfg.teamlead}</span>`:'<span style="font-size:11px;color:#9ca3af">Not assigned</span>'}</div>`:`<div class="org-person">${cfg.teamlead||'<span style="color:#9ca3af;font-weight:400">Not assigned</span>'}</div>`}</div><div><div class="org-label">Manager</div>${ce?`<div style="display:flex;align-items:center;gap:8px">${peopleSelectOptional('tc-mgr-'+teamName,cfg.manager||'','saveTeamConfig(\''+teamName+'\',\'manager\',this.value)','flex:1;max-width:220px')}${cfg.manager?`<span style="font-size:11px;color:#185fa5;font-weight:600">✓ ${cfg.manager}</span>`:'<span style="font-size:11px;color:#9ca3af">Not assigned</span>'}</div>`:`<div class="org-person">${cfg.manager||'<span style="color:#9ca3af;font-weight:400">Not assigned</span>'}</div>`}</div></div></div>
    ${debtSection}
    ${renderAddMemberCard(teamName, cfg, state)}
    <div class="card"><div class="card-hdr"><span class="card-title">📅 Weekly allocation</span><div style="display:flex;align-items:center;gap:12px">${weekRangeToggle()}<span class="card-sub">Green = person total · White = per-assignment</span></div></div>${!people.length?`<div class="empty"><span class="empty-icon">👥</span>No team members in ${teamName} yet.</div>`:`${filterBar}<div class="tbl-wrap"><table><thead><tr><th>Name</th><th>Skill</th><th>Level</th><th>Country</th><th>Reporting</th><th>Assignments</th>${wks.map(w=>wkHdr(w)).join('')}${ce?'<th></th>':''}</tr></thead><tbody>${memberRows}</tbody></table></div>`}</div>`;
}

function renderServices(){
  const ce=canEdit();
  const groups=TEAMS.map(t=>{const svcs=state.baseServices.filter(s=>s.team===t),teamDebt=calcTeamDebt(t),hasTargets=svcs.some(s=>s.targetPct>0);
    const svcRows=svcs.length?svcs.map(s=>{const debt=calcSvcDebt(s),target=s.targetPct||0,curAlloc=getSvcAlloc(s.name,CURRENT_WEEK),onTrack=!target||curAlloc>=target,dc=debt.debtPct===0?'#0f6e56':debt.debtPct<target*4?'#b45309':'#b91c1c',db=debt.debtPct===0?'#d1fae5':debt.debtPct<target*4?'#fef3c7':'#fef2f2';return `<div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:12px 14px;margin-bottom:8px"><div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap"><div style="flex:1;min-width:140px"><div style="font-size:13px;font-weight:600;color:#111827">${s.name}</div>${target?`<div style="font-size:11px;color:#6b7280;margin-top:2px">Target: <strong>${target}%</strong> · Now: <strong style="color:${onTrack?'#0f6e56':'#b91c1c'}">${curAlloc}%</strong></div>`:`<div style="font-size:11px;color:#9ca3af;margin-top:2px">No target set</div>`}</div>${target?`<div style="text-align:right;flex-shrink:0"><span style="background:${db};color:${dc};padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700">${debt.debtPct===0?'✓ No debt':'⚠ '+debt.debtPct+'% debt'}</span>${debt.debtPct>0?`<div style="font-size:10px;color:#9ca3af;margin-top:3px">= ${debt.debtWeeks} weeks fulltime</div>`:''}</div>`:''} ${ce?`<div style="display:flex;align-items:center;gap:6px;flex-shrink:0"><label style="font-size:10px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.05em">Target %</label><input type="number" min="0" max="200" value="${target||''}" placeholder="0" style="width:60px;padding:4px 6px;font-size:12px;border:1px solid #d1d5db;border-radius:6px;font-family:DM Mono,monospace;text-align:center" oninput="setSvcTarget('${s.name.replace(/'/g,"\\'")}',this.value)" /><button class="btn sm" onclick="editSvc('${s.name.replace(/'/g,"\\'")}')">✏</button><button class="btn danger sm" onclick="delSvc('${s.name.replace(/'/g,"\\'")}')">🗑</button></div>`:''}</div>${target&&debt.debtPct>0?`<div style="margin-top:10px"><div style="display:flex;justify-content:space-between;font-size:10px;color:#9ca3af;margin-bottom:3px"><span>Accumulated debt W1–W${CURRENT_WEEK-1}</span><span>${debt.debtPct}% of one FTE</span></div><div style="height:6px;background:#f3f4f6;border-radius:3px;overflow:hidden"><div style="height:100%;width:${Math.min(100,Math.round((debt.debtPct/(target*CURRENT_WEEK))*100))}%;background:${dc};border-radius:3px;transition:width .3s"></div></div></div>`:''}</div>`;}).join(''):`<div style="font-size:12px;color:#9ca3af;padding-bottom:6px">No services</div>`;
    const teamBadge=hasTargets?(teamDebt.debtPct===0?`<span style="background:#d1fae5;color:#065f46;font-size:11px;font-weight:700;padding:1px 8px;border-radius:20px">✓ No debt</span>`:`<span style="background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;padding:1px 8px;border-radius:20px">⚠ ${teamDebt.debtPct}% · ${teamDebt.debtWeeks}w</span>`):'';
    return `<div style="margin-bottom:20px"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px"><div class="svc-grp-lbl" style="margin:0">${t}</div>${teamBadge}</div>${svcRows}</div>`;
  }).join('');
  return `<div class="card"><div class="card-hdr"><span class="card-title">🔧 Base Services</span><span class="card-sub">Set a target % to track technical debt</span></div><div class="card-body">${ce?`<div class="ibox"><div class="sec-title">Add base service</div><div class="frow"><div class="fg"><label class="lbl">Team</label><select class="sel" style="width:140px" onchange="state.sTeam=this.value">${TEAMS.map(t=>`<option${state.sTeam===t?' selected':''}>${t}</option>`).join('')}</select></div><div class="fg" style="flex:1"><label class="lbl">Service name</label><input class="inp" placeholder="e.g. API Support" value="${state.sName}" oninput="state.sName=this.value" onkeydown="if(event.key==='Enter')addSvc()" /></div><div class="fg"><label class="lbl">Target %</label><input class="inp" type="number" min="0" max="200" placeholder="e.g. 10" value="${state.sTargetPct||''}" style="width:80px" oninput="state.sTargetPct=+this.value" /></div><button class="btn primary" onclick="addSvc()">＋ Add</button></div></div>`:''}${groups}</div></div>`;
}

function renderInbox(){
  if(!canEdit()) return `<div class="card"><div class="locked">🔒<span>Only Teamlead and Manager can manage the inbox.</span></div></div>`;
  const active=state.inboxItems.filter(i=>i.status==='new'),converted=state.inboxItems.filter(i=>i.status==='converted');
  return `<div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">📥 Inbox</span><span class="card-sub">Items not yet classified as project or initiative</span></div><div class="card-body"><div class="ibox"><div class="sec-title">Add new item</div><div class="fgrid"><div class="fg" style="grid-column:1/-1"><label class="lbl">Title *</label><input class="inp" placeholder="What is this about?" value="${state.iTitle}" oninput="state.iTitle=this.value" /></div><div class="fg" style="grid-column:1/-1"><label class="lbl">Description <span style="color:#9ca3af;font-weight:400">(optional)</span></label><textarea class="inp" rows="2" oninput="state.iDesc=this.value" style="resize:vertical">${state.iDesc}</textarea></div><div class="fg"><label class="lbl">Priority</label><select class="sel" onchange="state.iPriority=this.value">${['High','Medium','Low'].map(p=>`<option${state.iPriority===p?' selected':''}>${p}</option>`).join('')}</select></div><div style="padding-top:18px"><button class="btn primary" onclick="addInboxItem()">＋ Add to inbox</button></div></div></div></div></div>
    <div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">🔍 Needs classification (${active.length})</span></div><div class="card-body">${!active.length?`<div class="empty" style="padding:20px 0"><span class="empty-icon">✅</span>Inbox is empty.</div>`:active.map(item=>`<div class="inbox-card priority-${item.priority.toLowerCase()}"><div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px"><div style="flex:1"><div style="display:flex;align-items:center;gap:8px;margin-bottom:4px"><span style="font-size:13px;font-weight:700;color:#111827">${item.title}</span><span style="background:${item.priority==='High'?'#fef2f2;color:#dc2626':item.priority==='Medium'?'#fffbeb;color:#b45309':'var(--green-bg);color:#0f6e56'};padding:1px 8px;border-radius:20px;font-size:10px;font-weight:700">${item.priority}</span></div>${item.description?`<div style="font-size:12px;color:#6b7280;margin-bottom:6px">${item.description}</div>`:''}<div style="font-size:11px;color:#9ca3af">Added by ${item.createdBy} · ${item.createdAt}</div></div><div style="display:flex;flex-direction:column;gap:6px;flex-shrink:0"><button class="btn primary sm" onclick="convertInboxItem(${item.id},'project')">→ Make project</button><button class="btn sm" style="border-color:#a78bfa;color:#7c3aed" onclick="convertInboxItem(${item.id},'initiative')">→ Mark as initiative</button><button class="btn danger sm" onclick="deleteInboxItem(${item.id})">🗑 Dismiss</button></div></div></div>`).join('')}</div></div>
    ${converted.length?`<div class="card"><div class="card-hdr"><span class="card-title" style="color:#9ca3af">✓ Processed (${converted.length})</span></div><div class="card-body">${converted.map(item=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:10px 0;border-bottom:1px solid #f3f4f6;gap:12px"><div style="flex:1"><span style="font-size:13px;font-weight:600;color:#374151">${item.title}</span><span style="margin-left:8px;background:${item.convertedTo==='Project'?'#e0f2fe;color:#075985':'#ede9fe;color:#6d28d9'};padding:1px 8px;border-radius:20px;font-size:11px;font-weight:700">→ ${item.convertedTo}</span>${item.description?`<div style="font-size:11px;color:#9ca3af;margin-top:2px">${item.description}</div>`:''}</div><div style="display:flex;gap:6px;flex-shrink:0;flex-wrap:wrap;justify-content:flex-end">${item.convertedTo==='Project'?`<button class="btn sm" style="font-size:11px;border-color:#a78bfa;color:#7c3aed" onclick="convertInboxItem(${item.id},'initiative')">→ Change to initiative</button>`:`<button class="btn sm" style="font-size:11px;border-color:#0ea5e9;color:#0369a1" onclick="convertInboxItem(${item.id},'project')">→ Change to project</button>`}<button class="btn sm" style="font-size:11px" onclick="convertInboxItem(${item.id},'revert')">↩ Revert</button><button class="btn danger sm" onclick="deleteInboxItem(${item.id})">🗑</button></div></div>`).join('')}</div></div>`:''}`;
}

function renderPipeline(){
  const r=role(),ce=canEdit(),planned=state.assignments.filter(a=>!a.committed);
  if(!planned.length) return `<div class="card"><div class="empty"><span class="empty-icon">⏳</span>No planned (uncommitted) assignments yet.</div></div>`;
  const grouped=new Map(); planned.forEach(a=>{if(!grouped.has(a.workName))grouped.set(a.workName,[]);grouped.get(a.workName).push(a);});
  return `<div class="card" style="margin-bottom:16px"><div class="card-hdr"><span class="card-title">⏳ Pipeline — Planned but not committed</span><span class="card-sub">${planned.length} assignment${planned.length!==1?'s':''} across ${grouped.size} work item${grouped.size!==1?'s':''}</span></div></div>${[...grouped.entries()].map(([workName,items])=>{const proj=state.projects.find(p=>p.name===workName);return `<div class="card" style="margin-bottom:12px"><div class="card-hdr"><span class="card-title">${items[0].type==='Project'?'💼':'🔧'} ${workName}</span><div style="display:flex;gap:10px;align-items:center"><span class="badge b-type">${items[0].type}</span>${proj?`<span style="font-size:11px;color:#9ca3af">📅 ${proj.startDate||'—'} → ${proj.endDate||'—'}</span>`:''}<span style="font-size:11px;color:#b45309;font-weight:600">${items.length} resource${items.length!==1?'s':''} planned</span></div></div><div class="tbl-wrap"><table><thead><tr><th>Resource</th><th>Team</th><th>Skill</th><th>Level</th><th>Country</th><th>Periods</th><th>Total weeks</th>${ce?'<th>Action</th>':''}</tr></thead><tbody>${items.map(a=>{const idx=state.assignments.indexOf(a),rawN=a.name&&a.name.startsWith('__pm_planned__')?'No name':a.name,dn=r==='Project Manager'?(a.committed?rawN:'No name'):rawN,totalW=a.periods.reduce((s,p)=>s+(p.endWeek-p.startWeek+1),0),periods=a.periods.map(p=>`<span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%</span>`).join(' ');return `<tr><td style="background:${cBg(a.country)}"><strong>${dn}</strong></td><td>${a.team}</td><td>${a.skillset}</td><td>${a.level}</td><td>${a.country}</td><td>${periods}</td><td><strong style="font-family:'DM Mono',monospace">${totalW}</strong> week${totalW!==1?'s':''}</td>${ce?`<td style="white-space:nowrap"><button class="btn primary sm" onclick="commitA(${idx})">🔒 Commit</button><button class="btn danger sm" style="margin-top:4px" onclick="delAssignment(${idx})">🗑 Delete</button></td>`:''}</tr>`;}).join('')}</tbody></table></div></div>`;}).join('')}`;
}

function renderPlanning(){
  const vas=visibleAssignments(),ce=canEdit();
  if(!vas.length) return `<div class="card"><div class="empty"><span class="empty-icon">📋</span>No planning entries yet.</div></div>`;
  return `<div class="card"><div class="card-hdr"><span class="card-title">👥 Detailed planning</span><span class="card-sub">All 52 weeks</span></div><div class="tbl-wrap"><table><thead><tr><th>Resource</th><th>Team</th><th>Country</th><th>Skill</th><th>Level</th><th>Type</th><th>Work</th>${WEEKS.map(w=>wkHdr(w)).join('')}<th>Status</th><th style="min-width:190px">Periods</th></tr></thead><tbody>${vas.map((a,ai)=>{const dn=a.name&&a.name.startsWith('__pm_planned__')?(a.committed?a.name:'No name'):(!a.committed&&role()==='Project Manager'?'— Planned resource —':a.name),wks=WEEKS.map(w=>wkCell(w,getEffectiveAlloc(a,w))).join(''),periods=a.periods.map((p,pi)=>`<div style="display:flex;align-items:center;gap:4px;margin-bottom:4px"><span class="ptag">W${p.startWeek}–${p.endWeek}: ${p.allocationPercent}%</span>${ce?`<button class="btn danger sm" onclick="delPeriod(${ai},${pi})">🗑</button>`:''}</div>`).join(''),actions=ce?`<div style="display:flex;flex-direction:column;gap:4px;margin-top:6px"><button class="btn sm" onclick="addPeriod(${ai})">+ Add period</button><button class="btn danger sm" onclick="delAssignment(${ai})">🗑 Delete</button></div>`:'',status=a.committed?`<span class="badge b-committed">✓ Committed</span><div style="font-size:10px;color:#fff;background:#0f6e56;display:inline-block;padding:1px 6px;border-radius:4px;margin-top:3px;font-weight:600">${a.committedBy}</div>${ce?`<div style="margin-top:4px"><button class="btn danger sm" style="font-size:10px;padding:2px 6px" onclick="uncommitA(${ai})">↩ Uncommit</button></div>`:''}`:`<span class="badge b-plan">Planned</span>${ce?`<div style="margin-top:6px"><button class="btn primary sm" onclick="commitA(${ai})">🔒 Commit resource</button></div>`:'<div style="font-size:11px;color:#9ca3af;margin-top:4px">Awaiting commit</div>'}`;return `<tr><td style="background:${cBg(a.country)}"><strong>${dn}</strong><div style="font-size:11px;color:#9ca3af">${a.country}·${a.skillset}·${a.level}</div></td><td>${a.team}</td><td>${a.country}</td><td>${a.skillset}</td><td>${a.level}</td><td><span class="badge b-type">${a.type}</span></td><td>${a.workName}</td>${wks}<td>${status}</td><td>${periods}${actions}</td></tr>`;}).join('')}</tbody></table></div></div>`;
}

function buildPersonCalendar(name, selStart, selEnd){
  const pa=state.assignments.filter(a=>a.name.trim().toLowerCase()===name.trim().toLowerCase());
  if(!pa.length&&!name) return '';
  const weekData=WEEKS.map(w=>{const existing=pa.reduce((s,a)=>s+getAlloc(a,w),0),inNew=w>=selStart&&w<=selEnd;return {w,existing,inNew,proj:existing+(inNew?state.aPct:0)};});
  const hasAny=weekData.some(d=>d.existing>0||d.inNew);
  const rows=[]; for(let i=0;i<52;i+=13) rows.push(weekData.slice(i,i+13));
  function cellStyle(d){const cur=d.w===CURRENT_WEEK,outline=cur?'outline:2px solid #1D9E75;outline-offset:-2px;':'';if(d.inNew){if(d.proj>100)return `background:#fecaca;color:#b91c1c;font-weight:700;${outline}`;if(d.proj===100)return `background:#6ee7b7;color:#065f46;font-weight:700;${outline}`;return `background:#bfdbfe;color:#1e40af;font-weight:700;${outline}`;}if(d.existing>100)return `background:#fecaca;color:#b91c1c;font-weight:700;${outline}`;if(d.existing===100)return `background:#d1fae5;color:#065f46;font-weight:700;${outline}`;if(d.existing>0)return `background:#fef3c7;color:#92400e;font-weight:600;${outline}`;return `background:${cur?'rgba(29,158,117,0.07)':'#f9fafb'};color:#d1d5db;${outline}`;}
  const tableRows=rows.map(chunk=>`<tr>${chunk.map(d=>`<td style="text-align:center;padding:4px 2px;font-size:10px;font-family:'DM Mono',monospace;border-radius:3px;min-width:34px;${cellStyle(d)}"><div style="font-size:9px;opacity:.7;line-height:1">${d.w===CURRENT_WEEK?'▼':''}W${d.w}</div><div style="line-height:1.3">${d.proj>0?d.proj+'%':'–'}</div></td>`).join('')}</tr>`).join('');
  const over=weekData.filter(d=>d.inNew&&d.proj>100).length;
  return `<div class="add-sec"><div class="sec-title" style="display:flex;align-items:center;justify-content:space-between"><span>📅 ${name} — current &amp; planned allocation</span><div style="display:flex;gap:12px;font-size:10px;font-weight:600"><span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#fef3c7;display:inline-block"></span>Partial</span><span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#d1fae5;display:inline-block"></span>Full</span><span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#fecaca;display:inline-block"></span>Over</span><span style="display:inline-flex;align-items:center;gap:4px"><span style="width:10px;height:10px;border-radius:2px;background:#bfdbfe;display:inline-block"></span>New</span></div></div>${!hasAny?`<div style="padding:12px;background:#f9fafb;border-radius:8px;font-size:12px;color:#9ca3af;text-align:center">✓ No existing allocations — all weeks free.</div>`:`<div style="overflow-x:auto"><table style="border-collapse:separate;border-spacing:2px;min-width:100%">${tableRows}</table></div>`}${over?`<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:6px;padding:6px 12px;font-size:12px;color:#b91c1c;margin-top:8px">⚠ ${over} week${over!==1?'s':''} would be overbooked with this new period</div>`:''}${pa.length?`<div style="margin-top:8px;font-size:11px;color:#9ca3af">${pa.length} existing assignment${pa.length!==1?'s':''}: ${[...new Set(pa.map(a=>a.workName))].join(', ')}</div>`:''}</div>`;
}

function renderAdd(){
  const isPM=role()==='Project Manager';
  if(!canPlan()) return `<div class="card"><div class="locked">🔒<span>Only Teamlead, Manager and Project Manager can add planning.</span></div></div>`;
  const avail=isPM?pmProjects():state.projects, filtSvcs=state.baseServices.filter(s=>s.team===state.aTeam);
  const workField=state.addType==='Project'?`<div class="fg" style="max-width:300px"><label class="lbl">Project *</label><select class="sel" onchange="state.aProjId=this.value"><option value="">Select project…</option>${avail.map(p=>`<option value="${p.id}"${state.aProjId==p.id?' selected':''}>${p.name}</option>`).join('')}</select>${!avail.length?`<div class="hint">${isPM?'No projects assigned to you as PM yet.':'No projects yet — add one in the Projects tab.'}</div>`:''}</div>`:state.addType==='Base Service'?`<div class="fg" style="max-width:300px"><label class="lbl">Base service *</label><select class="sel" onchange="state.aService=this.value"><option value="">Select service…</option>${filtSvcs.map(s=>`<option value="${s.name}"${state.aService===s.name?' selected':''}>${s.name}</option>`).join('')}</select></div>`:`<div class="fg" style="max-width:300px"><label class="lbl">Work name *</label><input class="inp" placeholder="Name of work item" value="${state.aWork}" oninput="state.aWork=this.value" /></div>`;
  const calSection=(!isPM&&state.aName.trim())?buildPersonCalendar(state.aName.trim(),state.aStart,state.aEnd):'';
  return `<div class="card"><div class="card-hdr"><span class="card-title">📋 Planning mode — add entry</span></div><div class="card-body" style="display:flex;flex-direction:column;gap:24px">
    <div class="add-sec"><div class="sec-title">Assignment type</div><div class="ttabs">${(isPM?['Project']:TYPES).map(t=>`<div class="ttab${state.addType===t?' on':''}" onclick="state.addType='${t}';state.aProjId='';state.aService='';state.aWork='';render()">${t}</div>`).join('')}</div>${isPM?`<div class="hint" style="margin-top:6px">As Project Manager you can plan resources for your own projects.</div>`:''}</div>
    <div class="add-sec"><div class="sec-title">Resource details</div><div class="fgrid">${!isPM?`<div class="fg"><label class="lbl">Full name *</label>${peopleSelect('inp-aName',state.aName,"onPersonInput(this.value,'add');clearTimeout(_debounceTimer);render()",'','Select person…')}</div>`:''}<div class="fg"><label class="lbl">Team</label><select class="sel" onchange="state.aTeam=this.value;state.aService='';render()">${TEAMS.map(t=>`<option${state.aTeam===t?' selected':''}>${t}</option>`).join('')}</select></div><div class="fg"><label class="lbl">Country</label><select class="sel" onchange="state.aCountry=this.value"><option value="Sweden"${state.aCountry==='Sweden'?' selected':''}>Sweden</option><option value="Poland"${state.aCountry==='Poland'?' selected':''}>Poland</option></select></div><div class="fg"><label class="lbl">Skillset *</label><input class="inp" placeholder="e.g. React, DevOps" value="${state.aSkill}" oninput="state.aSkill=this.value" /></div><div class="fg"><label class="lbl">Level</label><select class="sel" onchange="state.aLevel=this.value">${['Junior','Mid','Senior'].map(l=>`<option${state.aLevel===l?' selected':''}>${l}</option>`).join('')}</select></div></div></div>
    ${calSection}
    <div class="add-sec"><div class="sec-title">Assignment</div>${workField}</div>
    <div class="add-sec"><div class="sec-title">Period &amp; allocation</div><div class="frow"><div class="fg"><label class="lbl">From week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.aStart}" onchange="state.aStart=+this.value;render()" /></div><div class="arrow">→</div><div class="fg"><label class="lbl">To week</label><input class="inp narrow" type="number" min="1" max="52" value="${state.aEnd}" onchange="state.aEnd=+this.value;render()" /></div><div class="fg"><label class="lbl">Allocation</label><div style="display:flex;align-items:center;gap:4px"><input class="inp narrow" type="number" min="0" max="200" value="${state.aPct}" oninput="state.aPct=+this.value" /><span style="font-size:13px;color:#6b7280">%</span></div></div></div><div class="hint">💡 Du kan ange vilken vecka som helst (1–52) för att lägga in uppgifter bakåt i tid.</div></div>
    <div><button class="btn primary block" onclick="addAssignment()">✓ Add planning entry</button>${state.msg?`<div class="msg ${state.msg.ok?'ok':'err'}">${state.msg.text}</div>`:'<div class="msg"></div>'}</div>
  </div></div>`;
}

async function init(){
  document.getElementById('content').innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:#9ca3af;font-size:14px">Loading data…</div>';
  const apiData=await loadFromApi();
  if(apiData&&Object.keys(apiData).length>0){
    if(apiData.projects)    state.projects    =apiData.projects;
    if(apiData.assignments) state.assignments =apiData.assignments;
    if(apiData.baseServices)state.baseServices=apiData.baseServices;
    if(apiData.teamMembers) state.teamMembers =apiData.teamMembers;
    if(apiData.teamConfig)  state.teamConfig  =apiData.teamConfig;
    if(apiData.inboxItems)  state.inboxItems  =apiData.inboxItems;
    if(apiData.userName) document.getElementById('user-name').value=apiData.userName;
    if(apiData.role)     document.getElementById('role-sel').value=apiData.role;
  } else if(saved){
    if(saved.userName) document.getElementById('user-name').value=saved.userName;
    if(saved.role)     document.getElementById('role-sel').value=saved.role;
  }
  autoRegisterTeamMembers();
  render();
}
init();
