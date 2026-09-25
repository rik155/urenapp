const days=['Maandag','Dinsdag','Woensdag','Donderdag','Vrijdag','Zaterdag','Zondag'];
const absenceTypes=['Vakantie','Tandarts','Ziekenhuisafspraak','Doktersafspraak','Bijzonder verlof','Overig'];
const daysEl=document.getElementById('days');
const employeeSelect=document.getElementById('employee');
const weekInput=document.getElementById('week');
let projects=[];
const esc=s=>String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));

function isoWeek(d){
  d=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  d.setUTCDate(d.getUTCDate()+4-(d.getUTCDay()||7));
  const y=new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d-y)/86400000)+1)/7);
}
weekInput.value=isoWeek(new Date());
document.getElementById('weekLabel').textContent=weekInput.value;
weekInput.addEventListener('input',()=>document.getElementById('weekLabel').textContent=weekInput.value||'–');

function projectOptions(selected=''){
  if(!projects.length)return '<option value="">Kantoor moet eerst een project toevoegen</option>';
  return '<option value="">Kies project</option>'+projects.map(p=>`<option value="${esc(p.name)}"${p.name===selected?' selected':''}>${esc(p.name)}</option>`).join('');
}
function absenceOptions(){
  return '<option value="">Geen afwezigheid</option>'+absenceTypes.map(x=>`<option value="${x}">${x}</option>`).join('');
}
function dayCard(day,i){
  return `<section class="day-card">
    <div class="day-head">
      <div><span class="day-index">${i+1}</span><h3>${day}</h3></div>
      <div class="day-actions"><span class="day-total" id="total-${i}">0 u</span><button class="add-btn" onclick="addLine(${i})" aria-label="Extra projectregel toevoegen">＋</button></div>
    </div>
    <div id="lines-${i}" class="lines"></div>
    <div id="temporary-${i}" class="temporary-list"></div>
    <button class="temporary-add" type="button" onclick="addTemporary(${i})">＋ Uitzendkracht toevoegen</button>
    <div class="absence-box">
      <div class="absence-title">Verlof of afspraak <small>optioneel</small></div>
      <div class="absence-fields">
        <label>Soort<select class="absence-type">${absenceOptions()}</select></label>
        <label class="hours-label">Uren<input class="absence-hours" type="number" min="0.25" max="24" step="0.25" inputmode="decimal" placeholder="0"></label>
      </div>
      <label>Opmerking<input class="absence-note" placeholder="Bijv. afspraak om 10:30" autocomplete="off"></label>
    </div>
  </section>`;
}
daysEl.innerHTML=days.map(dayCard).join('');

function lineHtml(i){
  return `<div class="work-line">
    <div class="work-main">
      <label>Project<select class="job">${projectOptions()}</select></label>
      <label class="hours-label">Uren<input class="hours" type="number" min="0" max="24" step="0.25" inputmode="decimal" placeholder="0"></label>
    </div>
    <div class="work-extra">
      <label class="free-wrap">Vrije uren<input class="free" type="number" min="0" max="24" step="0.25" inputmode="decimal" placeholder="0"></label>
      <label>Opmerking<input class="note" placeholder="Optioneel" autocomplete="off"></label>
      <button class="remove" type="button" onclick="removeLine(this,${i})" aria-label="Regel verwijderen">×</button>
    </div>
  </div>`;
}
function temporaryHtml(i){
  return `<div class="temporary-line">
    <div class="temporary-title"><strong>Uitzendkracht</strong><button type="button" onclick="removeTemporary(this,${i})" aria-label="Uitzendkracht verwijderen">×</button></div>
    <div class="temporary-main">
      <label>Naam<input class="temp-name" placeholder="Voor- en achternaam" autocomplete="off"></label>
      <label class="hours-label">Uren<input class="temp-hours" type="number" min="0.25" max="24" step="0.25" inputmode="decimal" placeholder="0"></label>
    </div>
    <div class="temporary-extra">
      <label>Project<select class="temp-job">${projectOptions()}</select></label>
      <label>Opmerking<input class="temp-note" placeholder="Optioneel" autocomplete="off"></label>
    </div>
  </div>`;
}
function addLine(i){
  document.getElementById(`lines-${i}`).insertAdjacentHTML('beforeend',lineHtml(i));
  syncZero();bindCalculations();
}
function addTemporary(i){
  const wrap=document.getElementById(`temporary-${i}`);
  wrap.insertAdjacentHTML('beforeend',temporaryHtml(i));
  bindCalculations();wrap.lastElementChild.querySelector('.temp-name').focus();
}
window.addLine=addLine;window.addTemporary=addTemporary;
window.removeLine=(btn,i)=>{
  const wrap=document.getElementById(`lines-${i}`);
  if(wrap.children.length>1)btn.closest('.work-line').remove();
  else{const line=btn.closest('.work-line');line.querySelectorAll('input').forEach(x=>x.value='');line.querySelectorAll('select').forEach(x=>x.value='')}
  updateTotals();
};
window.removeTemporary=btn=>{btn.closest('.temporary-line').remove();updateTotals()};

days.forEach((_,i)=>{addLine(i);addLine(i)});
async function loadProjects(){
  try{
    const r=await fetch('/api/projects');if(!r.ok)throw new Error();
    projects=await r.json();
    document.querySelectorAll('.job,.temp-job').forEach(select=>{const current=select.value;select.innerHTML=projectOptions(current);select.value=current});
  }catch{show('Projecten konden niet worden geladen. Probeer de pagina opnieuw.',false)}
}

function syncZero(){
  const z=document.getElementById('zero').checked;
  document.querySelectorAll('.free-wrap').forEach(e=>e.style.display=z?'flex':'none');
  updateTotals();
}
document.getElementById('zero').addEventListener('change',syncZero);
function bindCalculations(){
  document.querySelectorAll('.hours,.free,.temp-hours,.absence-hours').forEach(el=>{
    if(el.dataset.bound)return;el.dataset.bound='1';el.addEventListener('input',updateTotals);
  });
}
function updateTotals(){
  let total=0,free=0;
  days.forEach((_,i)=>{
    let d=0;
    document.querySelectorAll(`#lines-${i} .hours`).forEach(x=>d+=Number(x.value||0));
    document.querySelectorAll(`#temporary-${i} .temp-hours`).forEach(x=>d+=Number(x.value||0));
    const absence=document.querySelectorAll('.absence-hours')[i];d+=Number(absence.value||0);
    document.querySelectorAll(`#lines-${i} .free`).forEach(x=>free+=document.getElementById('zero').checked?Number(x.value||0):0);
    total+=d;document.getElementById(`total-${i}`).textContent=`${fmt(d)} u`;
  });
  document.getElementById('weekTotal').textContent=`${fmt(total)} uur`;
  document.getElementById('freeTotal').textContent=`${fmt(free)} vrij`;
  document.getElementById('freeStat').style.display=document.getElementById('zero').checked?'flex':'none';
}
function fmt(n){return Number.isInteger(n)?String(n):n.toFixed(2).replace(/0$/,'').replace('.',',')}
function show(t,ok){
  const m=document.getElementById('msg');m.textContent=t;m.className='msg '+(ok?'ok':'badmsg');m.hidden=false;
  m.scrollIntoView({behavior:'smooth',block:'center'});
}

async function submitWeek(){
  const name=employeeSelect.value.trim(),week=Number(weekInput.value),zero=document.getElementById('zero').checked;
  if(!name){show('Kies eerst je naam.',false);employeeSelect.focus();return}
  if(!week||week<1||week>53){show('Vul een geldig weeknummer in.',false);return}
  document.querySelectorAll('.field-error').forEach(x=>x.classList.remove('field-error'));
  const lines=[],temporary_workers=[],absences=[];
  days.forEach((day,i)=>{
    document.querySelectorAll(`#lines-${i} .work-line`).forEach(l=>{
      const job=l.querySelector('.job').value,hours=Number(l.querySelector('.hours').value||0),free=Number(l.querySelector('.free').value||0),note=l.querySelector('.note').value.trim();
      if((hours>0||free>0)&&!job)l.querySelector('.job').classList.add('field-error');
      if(job&&(hours>0||free>0))lines.push({day,job,hours,free_hours:zero?free:0,note});
    });
    document.querySelectorAll(`#temporary-${i} .temporary-line`).forEach(l=>{
      const tempName=l.querySelector('.temp-name').value.trim(),hours=Number(l.querySelector('.temp-hours').value||0),job=l.querySelector('.temp-job').value,note=l.querySelector('.temp-note').value.trim();
      if(!tempName)l.querySelector('.temp-name').classList.add('field-error');
      if(hours<=0)l.querySelector('.temp-hours').classList.add('field-error');
      if(!job)l.querySelector('.temp-job').classList.add('field-error');
      if(tempName&&hours>0&&job)temporary_workers.push({day,name:tempName,hours,job,note});
    });
    const box=document.querySelectorAll('.absence-box')[i],category=box.querySelector('.absence-type').value,hours=Number(box.querySelector('.absence-hours').value||0),note=box.querySelector('.absence-note').value.trim();
    if(category||hours>0||note){
      if(!category)box.querySelector('.absence-type').classList.add('field-error');
      if(hours<=0)box.querySelector('.absence-hours').classList.add('field-error');
      if(category&&hours>0)absences.push({day,category,hours,note});
    }
  });
  if(document.querySelector('.field-error')){show('Controleer de rood gemarkeerde velden.',false);return}
  if(!lines.length&&!temporary_workers.length&&!absences.length){show('Vul minimaal één project, uitzendkracht of afwezigheid in.',false);return}
  const btn=document.getElementById('submitBtn');btn.disabled=true;btn.innerHTML='<span class="spinner"></span> Bezig met indienen…';
  try{
    const r=await fetch('/api/submit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,week,zero_hours_contract:zero,lines,temporary_workers,absences})});
    const d=await r.json();if(!r.ok)throw new Error(d.detail||'Opslaan mislukt');
    const temp=d.temporary_total>0?`, ${fmt(d.temporary_total)} uitzenduren`:'';
    const absent=d.absence_total>0?`, ${fmt(d.absence_total)} afwezigheidsuren`:'';
    show(`Week ${week} is opgeslagen: ${fmt(d.work_total)} projecturen${temp}${absent}.`,true);
  }catch(e){show(e.message,false)}
  finally{btn.disabled=false;btn.innerHTML='Week indienen <span>→</span>'}
}
window.submitWeek=submitWeek;
syncZero();bindCalculations();updateTotals();loadProjects();
