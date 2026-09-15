
const state = { students:[], grades:[], settings:{}, selectedNisn:null };
const DEFAULT = { schoolName:"SMP Tahfizhpreneur Cahaya Qur'an", reportName:"Rapor Penilaian Tengah Semester", academicYear:"2026/2027", semester:"Ganjil", kkm:75 };

function norm(v){ return (v ?? "").toString().trim(); }
function esc(v){ return norm(v).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }
function num(v){ const n=Number(v); return Number.isFinite(n)?n:null; }

document.querySelectorAll(".nav button").forEach(btn=>{
  btn.addEventListener("click",()=>{
    document.querySelectorAll(".nav button").forEach(x=>x.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".section").forEach(x=>x.classList.remove("active"));
    document.getElementById(btn.dataset.target).classList.add("active");
  });
});

document.getElementById("excelFile").addEventListener("change", e=>{
  const file=e.target.files[0]; if(file) readWorkbook(file);
});
const drop=document.getElementById("drop");
drop.addEventListener("dragover",e=>{e.preventDefault();drop.style.borderColor="#2E73B9"});
drop.addEventListener("dragleave",()=>drop.style.borderColor="");
drop.addEventListener("drop",e=>{e.preventDefault();drop.style.borderColor="";const f=e.dataTransfer.files[0];if(f)readWorkbook(f)});

function rowsFrom(ws){ return XLSX.utils.sheet_to_json(ws,{defval:""}); }
function readWorkbook(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const data=new Uint8Array(e.target.result), wb=XLSX.read(data,{type:"array"});
      const siswa=wb.Sheets["Data Siswa"], nilai=wb.Sheets["Nilai"], peng=wb.Sheets["Pengaturan"];
      if(!siswa || !nilai) throw new Error("Sheet 'Data Siswa' dan 'Nilai' wajib tersedia.");
      state.students=rowsFrom(siswa).map(r=>({
        nisn:norm(r["NISN"]), nis:norm(r["NIS"]), name:norm(r["Nama"]), cls:norm(r["Kelas"]),
        gender:norm(r["Jenis Kelamin"]), homeroom:norm(r["Wali Kelas"]), note:norm(r["Catatan Wali Kelas"]),
        sick:num(r["Sakit"])||0, permit:num(r["Izin"])||0, absent:num(r["Alpa"])||0
      })).filter(x=>x.nisn || x.name);
      state.grades=rowsFrom(nilai).map(r=>({
        nisn:norm(r["NISN"]), name:norm(r["Nama"]), cls:norm(r["Kelas"]),
        subject:norm(r["Mata Pelajaran"]), score:num(r["Nilai"]), desc:norm(r["Deskripsi"])
      })).filter(x=>x.nisn && x.subject);
      if(peng){
        const p=rowsFrom(peng); const map={}; p.forEach(r=>map[norm(r["Kunci"])]=r["Nilai"]);
        state.settings=map;
      }
      localStorage.setItem("raporSMP",JSON.stringify(state));
      refreshAll();
      document.getElementById("fileStatus").innerHTML=`<span class="status-ok">✓ ${esc(file.name)}</span>`;
    }catch(err){ alert("File tidak dapat dibaca: "+err.message); }
  };
  reader.readAsArrayBuffer(file);
}
function loadSaved(){
  try{
    const s=JSON.parse(localStorage.getItem("raporSMP")||"null");
    if(s){ Object.assign(state,s); refreshAll(); document.getElementById("fileStatus").innerHTML='<span class="status-ok">✓ Data tersimpan di browser</span>'; }
  }catch(e){}
}
function resetData(){
  if(!confirm("Hapus data rapor dari browser ini?")) return;
  localStorage.removeItem("raporSMP"); state.students=[];state.grades=[];state.settings={};state.selectedNisn=null;refreshAll();
  document.getElementById("fileStatus").textContent="Belum ada file";
}
function getSetting(key, fallback){ return state.settings[key] ?? fallback; }
function gradeRows(nisn){ return state.grades.filter(g=>g.nisn===nisn); }

function refreshAll(){
  const classes=[...new Set(state.students.map(s=>s.cls).filter(Boolean))].sort();
  const filled=state.students.filter(s=>gradeRows(s.nisn).length>0).length;
  document.getElementById("kStudents").textContent=state.students.length;
  document.getElementById("kClasses").textContent=classes.length;
  document.getElementById("kFilled").textContent=filled;
  document.getElementById("kIncomplete").textContent=Math.max(0,state.students.length-filled);

  const body=document.getElementById("studentBody");
  body.innerHTML=state.students.map((s,i)=>{
    const gs=gradeRows(s.nisn), avg=gs.length?gs.reduce((a,g)=>a+(g.score??0),0)/gs.length:null;
    return `<tr><td>${i+1}</td><td>${esc(s.name)}</td><td>${esc(s.cls)}</td><td>${gs.length}</td><td>${avg===null?"-":avg.toFixed(1)}</td><td><button class="btn" onclick="openReport('${esc(s.nisn)}')">Buka Rapor</button></td></tr>`;
  }).join("") || '<tr><td colspan="6" class="muted">Belum ada data siswa.</td></tr>';

  const sel=document.getElementById("studentSelect");
  sel.innerHTML='<option value="">Pilih siswa...</option>'+state.students.map(s=>`<option value="${esc(s.nisn)}">${esc(s.name)} — ${esc(s.cls)}</option>`).join("");
}

function openReport(nisn){
  state.selectedNisn=nisn;
  document.querySelector('[data-target="printSection"]').click();
  document.getElementById("studentSelect").value=nisn;
  renderReport();
}
document.getElementById("studentSelect").addEventListener("change",e=>{state.selectedNisn=e.target.value;renderReport()});

function renderReport(){
  const s=state.students.find(x=>x.nisn===state.selectedNisn);
  const box=document.getElementById("reportBox");
  if(!s){ box.innerHTML='<div class="muted">Pilih siswa untuk menampilkan rapor.</div>'; return; }
  const gs=gradeRows(s.nisn);
  const kkm=Number(getSetting("KKM",DEFAULT.kkm))||75;
  const avg=gs.length?gs.reduce((a,g)=>a+(g.score??0),0)/gs.length:0;
  box.innerHTML=`
  <div class="report">
    <div class="report-header">
      <h2>${esc(getSetting("Nama Sekolah",DEFAULT.schoolName))}</h2>
      <h3>${esc(getSetting("Jenis Rapor","Penilaian Tengah Semester (PTS)"))}</h3>
      <p>Tahun Pelajaran ${esc(getSetting("Tahun Pelajaran",DEFAULT.academicYear))} • Semester ${esc(getSetting("Semester",DEFAULT.semester))}</p>
    </div>
    <div class="identity">
      <div>Nama</div><div>: <b>${esc(s.name)}</b></div><div>Kelas</div><div>: ${esc(s.cls)}</div>
      <div>NIS / NISN</div><div>: ${esc(s.nis)} / ${esc(s.nisn)}</div><div>Wali Kelas</div><div>: ${esc(s.homeroom)}</div>
    </div>
    <table>
      <thead><tr><th style="width:36px">No.</th><th>Mata Pelajaran</th><th style="width:75px;text-align:center">Nilai</th><th>Deskripsi Capaian</th></tr></thead>
      <tbody>
        ${gs.map((g,i)=>`<tr><td>${i+1}</td><td>${esc(g.subject)}</td><td style="text-align:center;font-weight:700">${g.score??"-"}</td><td>${esc(g.desc)}</td></tr>`).join("") || '<tr><td colspan="4">Belum ada nilai.</td></tr>'}
        ${gs.length?`<tr><td colspan="2"><b>Rata-rata</b></td><td style="text-align:center"><b>${avg.toFixed(1)}</b></td><td>${avg>=kkm?"Capaian umum memenuhi KKM.":"Perlu tindak lanjut pada beberapa capaian."}</td></tr>`:""}
      </tbody>
    </table>
    <div style="margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:18px">
      <div><b>Catatan Wali Kelas</b><div style="border:1px solid #333;min-height:70px;padding:8px;margin-top:6px">${esc(s.note)||"-"}</div></div>
      <div><b>Ketidakhadiran</b><table style="margin-top:6px"><tr><td>Sakit</td><td>${s.sick} hari</td></tr><tr><td>Izin</td><td>${s.permit} hari</td></tr><tr><td>Alpa</td><td>${s.absent} hari</td></tr></table></div>
    </div>
    <div class="sign">
      <div>Orang Tua/Wali<div class="sign-space"></div><div>(________________________)</div></div>
      <div>Wali Kelas<div class="sign-space"></div><div><b>${esc(s.homeroom)||"(________________________)"}</b></div></div>
    </div>
  </div>`;
}
function printReport(){ if(!state.selectedNisn){alert("Pilih siswa terlebih dahulu.");return;} window.print(); }
function printAll(){
  if(!state.students.length){alert("Belum ada data.");return;}
  alert("Versi MVP mencetak per siswa. Fitur gabungan satu kelas siap dikembangkan saat format final rapor dikunci.");
}
loadSaved();
