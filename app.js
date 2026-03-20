let latest20=[];
document.getElementById("fileInput").addEventListener("change",e=>{
  const r=new FileReader();
  r.onload=ev=>{
    const wb=XLSX.read(ev.target.result,{type:"binary"});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const data=XLSX.utils.sheet_to_json(ws);
    latest20=data.slice(-20);
    render();
  };
  r.readAsBinaryString(e.target.files[0]);
});
function render(){
  const freq={};
  latest20.forEach(r=>{
    const nums=[r.N1,r.N2,r.N3,r.N4,r.N5,r.N6,r['特別號']];
    nums.forEach(n=>freq[n]=(freq[n]||0)+1);
  });
  const tbody=document.getElementById('freq');
  tbody.innerHTML='';
  Object.entries(freq).sort((a,b)=>b[1]-a[1]).forEach(([n,c])=>{
    tbody.innerHTML+=`<tr><td>${n}</td><td>${c}</td></tr>`;
  });
}
function downloadExcel(){
  alert('請用 Python 版本匯出 Excel');
}