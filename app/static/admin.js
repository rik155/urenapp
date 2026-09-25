let data=[],projects=[],pending=null;
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
function toast(msg,bad=false){const t=$('toast');t.textContent=msg;t.className='toast show'+(bad?' bad':'');setTimeout(()=>t.className='toast',3000)}
async function api(url,opt={}){
  const r=await fetch(url,opt);
  if(!r.ok){let d={};try{d=await r.json()}catch{}throw new Error(d.detail||'Er ging iets mis.')}
  return r;
}
async function load(){
  try{
    const [statusResponse,projectsResponse]=await Promise.all([api('/api/status'),api('/api/projects')]);
    data=await statusResponse.json();projects=await projectsResponse.json();
    fillWeeks();render();renderProjects();
  }catch(e){toast(e.message,true)}
}
function fillWeeks(){
  const current=$('weekFilter').value,weeks=[...new Set(data.map(x=>x.week))].sort((a,b)=>b-a);
  $('weekFilter').innerHTML='<option value="">Alle weken</option>'+weeks.map(w=>`<option value="${w}">Week ${w}</option>`).join('');
  $('weekFilter').value=current;
}
function filtered(){
  const q=$('search').value.trim().toLowerCase(),w=$('weekFilter').value;
  return data.filter(x=>(!q||x.name.toLowerCase().includes(q)||String(x.work_addresses||'').toLowerCase().includes(q)||String(x.temporary_names||'').toLowerCase().includes(q)||String(x.absence_types||'').toLowerCase().includes(q))&&(!w||String(x.week)===w));
}
function render(){
  const d=filtered();
  $('rows').innerHTML=d.length?d.map(x=>`<tr>
    <td><div class="person"><div class="avatar">${esc(x.name.charAt(0).toUpperCase())}</div><span><b>${esc(x.name)}</b><small class="worker-type">${x.temporary_names?`Uitzendkracht: ${esc(x.temporary_names)} (${Number(x.temporary_hours).toFixed(2)} u)`:'Medewerker'}</small></span></div></td>
    <td class="address">${esc(x.work_addresses||'—')}</td>
    <td><span class="weekpill">W${x.week}</span></td><td>${x.period}</td>
    <td>${Number(x.work_hours).toFixed(2)}</td>
    <td>${Number(x.absence_hours||0).toFixed(2)}<small class="worker-type">${esc(x.absence_types||'')}</small></td>
    <td><b>${Number(x.total_hours).toFixed(2)}</b></td><td class="muted">${esc(x.submitted_at)}</td>
    <td><div class="rowactions"><button class="iconbtn" title="Deze week verwijderen" onclick="askDeleteWeek(${x.id},'${encodeURIComponent(x.name)}',${x.week})">×</button><button class="textdanger" onclick="askDeleteEmployee('${encodeURIComponent(x.name_key)}','${encodeURIComponent(x.name)}')">Alles verwijderen</button></div></td>
  </tr>`).join(''):'<tr><td colspan="9" class="empty">Geen inzendingen gevonden.</td></tr>';
  $('statSub').textContent=data.length;
  $('statEmp').textContent=new Set(data.map(x=>x.name_key)).size;
  $('statWork').textContent=data.reduce((a,x)=>a+Number(x.work_hours),0).toFixed(2);
  $('statAbsence').textContent=data.reduce((a,x)=>a+Number(x.absence_hours||0),0).toFixed(2);
}
function renderProjects(){
  $('projectList').innerHTML=projects.length?projects.map(p=>`<span class="project-pill"><b>${esc(p.name)}</b><button type="button" onclick="removeProject(${p.id},'${encodeURIComponent(p.name)}')" aria-label="Project verwijderen">×</button></span>`).join(''):'<p class="project-empty">Nog geen projecten toegevoegd.</p>';
}
$('projectForm').addEventListener('submit',async e=>{
  e.preventDefault();const input=$('projectName'),name=input.value.trim();if(!name)return;
  try{
    await api('/api/projects',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
    input.value='';toast('Project toegevoegd.');await load();
  }catch(err){toast(err.message,true)}
});
window.removeProject=async(id,encodedName)=>{
  const name=decodeURIComponent(encodedName);
  if(!window.confirm(`Project "${name}" verwijderen uit de keuzelijst? Eerder geregistreerde uren blijven bewaard.`))return;
  try{const r=await api('/api/projects/'+id,{method:'DELETE'});const d=await r.json();toast(d.message);await load()}catch(e){toast(e.message,true)}
};
window.askDeleteWeek=(id,name,week)=>{
  name=decodeURIComponent(name);pending={type:'week',id};
  $('modalTitle').textContent='Week verwijderen?';$('modalText').textContent=`Week ${week} van ${name} wordt definitief verwijderd en komt niet meer in de Excel-export.`;$('modal').hidden=false;
};
window.askDeleteEmployee=(key,name)=>{
  name=decodeURIComponent(name);pending={type:'employee',key:decodeURIComponent(key)};
  $('modalTitle').textContent='Alle gegevens verwijderen?';$('modalText').textContent=`Alle ingediende weken van ${name} worden definitief verwijderd.`;$('modal').hidden=false;
};
$('cancelDelete').onclick=()=>{$('modal').hidden=true;pending=null};
$('confirmDelete').onclick=async()=>{
  if(!pending)return;
  try{
    const r=pending.type==='week'?await api('/api/submission/'+pending.id,{method:'DELETE'}):await api('/api/employee?name_key='+encodeURIComponent(pending.key),{method:'DELETE'});
    const d=await r.json();$('modal').hidden=true;pending=null;toast(d.message);await load();
  }catch(e){toast(e.message,true)}
};
$('refreshBtn').onclick=load;$('search').oninput=render;$('weekFilter').onchange=render;$('excelBtn').onclick=()=>{window.location.href='/api/excel'};
load();setInterval(load,30000);
