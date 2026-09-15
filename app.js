
const STORE='raporSmpV2';
const MAPELS=['Pendidikan Agama Islam','Bahasa Indonesia','Matematika','Ilmu Pengetahuan Alam','Ilmu Pengetahuan Sosial','Bahasa Inggris','Pendidikan Pancasila','Bahasa Arab','PJOK','Informatika'];
let state={students:[],mapel:[],tahfizh:[],ekskul:[],walas:[]};

function esc(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function uid(p){return p+'-'+Date.now()+'-'+Math.floor(Math.random()*1e6)}
function save(){localStorage.setItem(STORE,JSON.stringify(state));refreshAll()}
function load(){try{const x=JSON.parse(localStorage.getItem(STORE)||'null');if(x)state=x}catch(e){} refreshAll()}
function classes(){return [...new Set(state.students.map(s=>s.cls).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id',{numeric:true}))}
function sortedStudents(cls=''){return state.students.filter(s=>!cls||s.cls===cls).sort((a,b)=>(+a.order||999)-(+b.order||999)||a.name.localeCompare(b.name,'id'))}
function recordBy(arr, sid, extraKey, extraVal){return arr.find(x=>x.sid===sid && (!extraKey||x[extraKey]===extraVal))}
function go(id){document.querySelectorAll('.section').forEach(x=>x.classList.remove('active'));document.getElementById(id).classList.add('active');document.querySelectorAll('.nav button').forEach(x=>x.classList.toggle('active',x.dataset.target===id))}
document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>go(b.dataset.target));

document.getElementById('roleSelect').onchange=applyRole;
function applyRole(){
  const role=document.getElementById('roleSelect').value;
  document.querySelectorAll('.nav button[data-role]').forEach(b=>{b.style.display=b.dataset.role.split(',').includes(role)?'block':'none'});
  const active=document.querySelector('.nav button.active');
  if(active && active.style.display==='none') go('dashboard');
}

function optionsFor(select, vals, placeholder){
  const cur=select.value;
  select.innerHTML=(placeholder?`<option value="">${placeholder}</option>`:'')+vals.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if(vals.includes(cur))select.value=cur;
}

function refreshAll(){
  const cs=classes();
  ['masterClass','mapelClass','tahfizhClass','ekskulClass','walasClass','previewClass'].forEach(id=>optionsFor(document.getElementById(id),cs,id==='masterClass'?'Semua kelas':'Pilih kelas'));
  optionsFor(document.getElementById('mapelName'),MAPELS,'Pilih mapel');
  document.getElementById('kStudents').textContent=state.students.length;
  document.getElementById('kClasses').textContent=cs.length;
  document.getElementById('kMapel').textContent=state.mapel.length;
  const ready=state.students.filter(s=>state.mapel.some(x=>x.sid===s.id)&&state.tahfizh.some(x=>x.sid===s.id)&&state.walas.some(x=>x.sid===s.id)).length;
  document.getElementById('kReady').textContent=ready;
  renderMaster();renderMapelInput();renderTahfizhInput();renderEkskulInput();renderWalasInput();refreshPreviewStudents();
}

document.getElementById('dbFile').addEventListener('change',e=>{const f=e.target.files[0];if(f)readDb(f)});
function rows(ws){return XLSX.utils.sheet_to_json(ws,{defval:''})}
function readDb(file){
  const r=new FileReader();r.onload=e=>{try{
    const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'}), s=wb.Sheets['Siswa'];
    if(!s)throw new Error("Sheet Siswa tidak ditemukan");
    state.students=rows(s).filter(x=>x.Nama).map((x,i)=>({
      id:String(x.StudentID||uid('S')),nisn:String(x.NISN||''),nis:String(x.NIS||''),name:String(x.Nama||'').trim(),
      cls:String(x.Kelas||'').trim(),gender:String(x.JenisKelamin||''),order:Number(x.NoUrut||i+1),homeroom:String(x.WaliKelas||'').trim()
    }));
    save();document.getElementById('fileStatus').innerHTML='<span class="ok">✓ Database siswa berhasil dibaca.</span>';
  }catch(err){alert(err.message)}};
  r.readAsArrayBuffer(file)
}

function addStudent(){
  const name=prompt('Nama siswa');if(!name)return;
  const cls=prompt('Kelas, contoh 7A');if(!cls)return;
  const nisn=prompt('NISN (opsional)')||'';
  const homeroom=prompt('Wali kelas (opsional)')||'';
  state.students.push({id:uid('S'),nisn,nis:'',name:name.trim(),cls:cls.trim(),gender:'',order:sortedStudents(cls).length+1,homeroom});
  save();
}
function editStudent(id){
  const s=state.students.find(x=>x.id===id);if(!s)return;
  const name=prompt('Nama siswa',s.name);if(name===null)return;
  const cls=prompt('Kelas',s.cls);if(cls===null)return;
  const order=prompt('No urut',s.order);if(order===null)return;
  s.name=name.trim();s.cls=cls.trim();s.order=Number(order)||s.order;save();
}
function delStudent(id){if(confirm('Hapus siswa ini?')){state.students=state.students.filter(x=>x.id!==id);save()}}
function renderMaster(){
  const cls=document.getElementById('masterClass').value;
  document.getElementById('masterBody').innerHTML=sortedStudents(cls).map(s=>`<tr><td>${s.order}</td><td>${esc(s.nisn)}</td><td><b>${esc(s.name)}</b></td><td>${esc(s.cls)}</td><td>${esc(s.homeroom)}</td><td><button class="btn" onclick="editStudent('${s.id}')">Edit</button> <button class="btn danger" onclick="delStudent('${s.id}')">Hapus</button></td></tr>`).join('')||'<tr><td colspan="6" class="muted">Belum ada siswa.</td></tr>';
}

function renderMapelInput(){
  const cls=document.getElementById('mapelClass').value, mapel=document.getElementById('mapelName').value;
  document.getElementById('mapelBody').innerHTML=sortedStudents(cls).map(s=>{const r=recordBy(state.mapel,s.id,'subject',mapel)||{};return `<tr data-sid="${s.id}"><td>${s.order}</td><td><b>${esc(s.name)}</b></td><td><input class="score" type="number" min="0" max="100" value="${r.score??''}"></td><td><textarea class="desc">${esc(r.desc||'')}</textarea></td></tr>`}).join('')||'<tr><td colspan="4" class="muted">Pilih kelas.</td></tr>';
}
function saveMapel(){
  const cls=mapelClass.value, subject=mapelName.value;if(!cls||!subject)return alert('Pilih kelas dan mapel.');
  document.querySelectorAll('#mapelBody tr[data-sid]').forEach(tr=>{const sid=tr.dataset.sid,score=tr.querySelector('.score').value,desc=tr.querySelector('.desc').value;let r=recordBy(state.mapel,sid,'subject',subject);if(!r){r={id:uid('M'),sid,cls,subject};state.mapel.push(r)}r.score=score===''?null:Number(score);r.desc=desc});
  save();alert('Nilai mapel tersimpan.')
}

function renderTahfizhInput(){
  const cls=tahfizhClass.value;
  tahfizhBody.innerHTML=sortedStudents(cls).map(s=>{const r=recordBy(state.tahfizh,s.id)||{};return `<tr data-sid="${s.id}"><td>${s.order}</td><td><b>${esc(s.name)}</b></td><td><input class="score" type="number" min="0" max="100" value="${r.score??''}"></td><td><input class="haf" value="${esc(r.haf||'')}"></td><td><input class="taj" value="${esc(r.taj||'')}"></td><td><input class="mur" value="${esc(r.mur||'')}"></td><td><textarea class="desc">${esc(r.desc||'')}</textarea></td></tr>`}).join('')||'<tr><td colspan="7" class="muted">Pilih kelas.</td></tr>';
}
function saveTahfizh(){
  const cls=tahfizhClass.value;if(!cls)return alert('Pilih kelas.');
  document.querySelectorAll('#tahfizhBody tr[data-sid]').forEach(tr=>{const sid=tr.dataset.sid;let r=recordBy(state.tahfizh,sid);if(!r){r={id:uid('T'),sid,cls};state.tahfizh.push(r)}r.score=tr.querySelector('.score').value===''?null:Number(tr.querySelector('.score').value);r.haf=tr.querySelector('.haf').value;r.taj=tr.querySelector('.taj').value;r.mur=tr.querySelector('.mur').value;r.desc=tr.querySelector('.desc').value});
  save();alert('Nilai Tahfizh tersimpan.')
}

function renderEkskulInput(){
  const cls=ekskulClass.value, ex=ekskulName.value.trim();
  ekskulBody.innerHTML=sortedStudents(cls).map(s=>{const r=recordBy(state.ekskul,s.id,'name',ex)||{};return `<tr data-sid="${s.id}"><td>${s.order}</td><td><b>${esc(s.name)}</b></td><td><input class="join" type="checkbox" ${r.join?'checked':''}></td><td><select class="pred"><option></option>${['A','B','C','D'].map(x=>`<option ${r.pred===x?'selected':''}>${x}</option>`).join('')}</select></td><td><textarea class="desc">${esc(r.desc||'')}</textarea></td></tr>`}).join('')||'<tr><td colspan="5" class="muted">Pilih kelas.</td></tr>';
}
function saveEkskul(){
  const cls=ekskulClass.value, name=ekskulName.value.trim();if(!cls||!name)return alert('Pilih kelas dan isi nama ekskul.');
  document.querySelectorAll('#ekskulBody tr[data-sid]').forEach(tr=>{const sid=tr.dataset.sid, join=tr.querySelector('.join').checked;let r=recordBy(state.ekskul,sid,'name',name);if(join){if(!r){r={id:uid('E'),sid,cls,name};state.ekskul.push(r)}r.join=true;r.pred=tr.querySelector('.pred').value;r.desc=tr.querySelector('.desc').value}else if(r){state.ekskul=state.ekskul.filter(x=>x!==r)}});
  save();alert('Data ekskul tersimpan.')
}

function renderWalasInput(){
  const cls=walasClass.value;
  walasBody.innerHTML=sortedStudents(cls).map(s=>{const r=recordBy(state.walas,s.id)||{};return `<tr data-sid="${s.id}"><td>${s.order}</td><td><b>${esc(s.name)}</b></td><td><select class="disc"><option></option>${['Sangat Baik','Baik','Cukup','Perlu Bimbingan'].map(x=>`<option ${r.disc===x?'selected':''}>${x}</option>`).join('')}</select></td><td><input class="sick" type="number" min="0" value="${r.sick??0}"></td><td><input class="permit" type="number" min="0" value="${r.permit??0}"></td><td><input class="absent" type="number" min="0" value="${r.absent??0}"></td><td><textarea class="reward">${esc(r.reward||'')}</textarea></td><td><textarea class="violation">${esc(r.violation||'')}</textarea></td><td><textarea class="note">${esc(r.note||'')}</textarea></td></tr>`}).join('')||'<tr><td colspan="9" class="muted">Pilih kelas.</td></tr>';
}
function saveWalas(){
  const cls=walasClass.value;if(!cls)return alert('Pilih kelas.');
  document.querySelectorAll('#walasBody tr[data-sid]').forEach(tr=>{const sid=tr.dataset.sid;let r=recordBy(state.walas,sid);if(!r){r={id:uid('W'),sid,cls};state.walas.push(r)}r.disc=tr.querySelector('.disc').value;r.sick=Number(tr.querySelector('.sick').value||0);r.permit=Number(tr.querySelector('.permit').value||0);r.absent=Number(tr.querySelector('.absent').value||0);r.reward=tr.querySelector('.reward').value;r.violation=tr.querySelector('.violation').value;r.note=tr.querySelector('.note').value});
  save();alert('Bagian wali kelas tersimpan.')
}

function refreshPreviewStudents(){
  const cls=previewClass.value, list=sortedStudents(cls), cur=previewStudent.value;
  previewStudent.innerHTML='<option value="">Pilih siswa</option>'+list.map(s=>`<option value="${s.id}">${esc(s.name)}</option>`).join('');
  if(list.some(x=>x.id===cur))previewStudent.value=cur;renderPreview();
}
function renderPreview(){
  const sid=previewStudent.value,s=state.students.find(x=>x.id===sid);if(!s){reportShell.innerHTML='<div class="muted">Pilih siswa untuk preview.</div>';previewStatus.textContent='';return}
  const ms=state.mapel.filter(x=>x.sid===sid).sort((a,b)=>MAPELS.indexOf(a.subject)-MAPELS.indexOf(b.subject)), t=recordBy(state.tahfizh,sid), es=state.ekskul.filter(x=>x.sid===sid), w=recordBy(state.walas,sid)||{};
  const missing=[];if(!ms.length)missing.push('nilai mapel');if(!t)missing.push('Tahfizh');if(!state.walas.some(x=>x.sid===sid))missing.push('bagian wali kelas');
  previewStatus.innerHTML=missing.length?`<span class="warning">Belum lengkap: ${missing.join(', ')}.</span>`:'<span class="ok">✓ Data utama rapor sudah terisi.</span>';
  reportShell.innerHTML=`<div class="report">
    <div class="report-header"><h2>SMP Tahfizhpreneur Cahaya Qur'an</h2><h3>RAPOR PENILAIAN TENGAH SEMESTER</h3><p>Tahun Pelajaran 2026/2027 • Semester Ganjil</p></div>
    <div class="identity"><div>Nama</div><div>: <b>${esc(s.name)}</b></div><div>Kelas</div><div>: ${esc(s.cls)}</div><div>NIS / NISN</div><div>: ${esc(s.nis||'-')} / ${esc(s.nisn||'-')}</div><div>Wali Kelas</div><div>: ${esc(s.homeroom||'-')}</div></div>
    <div class="block-title">A. Capaian Mata Pelajaran</div>
    <table><thead><tr><th>No</th><th>Mata Pelajaran</th><th>Nilai</th><th>Deskripsi</th></tr></thead><tbody>
    ${ms.map((r,i)=>`<tr><td>${i+1}</td><td>${esc(r.subject)}</td><td style="text-align:center"><b>${r.score??'-'}</b></td><td>${esc(r.desc||'-')}</td></tr>`).join('')||'<tr><td colspan="4">Belum ada nilai mapel.</td></tr>'}
    </tbody></table>
    <div class="block-title">B. Tahfizh Al-Qur'an</div>
    <table><thead><tr><th>Nilai</th><th>Hafalan</th><th>Tajwid</th><th>Murajaah</th><th>Deskripsi</th></tr></thead><tbody><tr><td>${t?.score??'-'}</td><td>${esc(t?.haf||'-')}</td><td>${esc(t?.taj||'-')}</td><td>${esc(t?.mur||'-')}</td><td>${esc(t?.desc||'-')}</td></tr></tbody></table>
    <div class="block-title">C. Ekstrakurikuler</div>
    <table><thead><tr><th>No</th><th>Ekskul</th><th>Predikat</th><th>Keterangan</th></tr></thead><tbody>${es.map((r,i)=>`<tr><td>${i+1}</td><td>${esc(r.name)}</td><td>${esc(r.pred||'-')}</td><td>${esc(r.desc||'-')}</td></tr>`).join('')||'<tr><td colspan="4">-</td></tr>'}</tbody></table>
    <div class="block-title">D. Kedisiplinan & Kehadiran</div>
    <div class="two-col"><table><tr><td>Kedisiplinan</td><td><b>${esc(w.disc||'-')}</b></td></tr><tr><td>Reward</td><td>${esc(w.reward||'-')}</td></tr><tr><td>Pelanggaran</td><td>${esc(w.violation||'-')}</td></tr></table><table><tr><td>Sakit</td><td>${w.sick??0} hari</td></tr><tr><td>Izin</td><td>${w.permit??0} hari</td></tr><tr><td>Alpa</td><td>${w.absent??0} hari</td></tr></table></div>
    <div class="block-title">E. Catatan Wali Kelas</div><div style="border:1px solid #333;min-height:70px;padding:8px">${esc(w.note||'-')}</div>
    <div class="sign"><div>Orang Tua/Wali<div class="sign-space"></div>(______________________)</div><div>Wali Kelas<div class="sign-space"></div><b>${esc(s.homeroom||'______________________')}</b></div></div>
  </div>`;
}
function printReport(){if(!previewStudent.value)return alert('Pilih siswa dulu.');window.print()}
function clearAll(){if(confirm('Hapus semua data di browser ini?')){localStorage.removeItem(STORE);state={students:[],mapel:[],tahfizh:[],ekskul:[],walas:[]};refreshAll()}}
applyRole();load();
