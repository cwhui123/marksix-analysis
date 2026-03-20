async function load(){
  const res = await fetch('data.xlsx');
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws);
  const latest20 = data.slice(-20);

  // render frequency
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

  window._latest20 = latest20;
}

function downloadExcel(){
  const latest20 = window._latest20 || [];
  const wb=XLSX.utils.book_new();

  const freq={};
  latest20.forEach(r=>{
    const nums=[r.N1,r.N2,r.N3,r.N4,r.N5,r.N6,r['特別號']];
    nums.forEach(n=>freq[n]=(freq[n]||0)+1);
  });
  const freqArr=[["號碼","出現次數"],...Object.entries(freq).sort((a,b)=>b[1]-a[1])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(freqArr),'號碼統計（含特別號）');

  const pairCount={};
  latest20.forEach(r=>{
    const nums=[r.N1,r.N2,r.N3,r.N4,r.N5,r.N6,r['特別號']].sort((a,b)=>a-b);
    for(let i=0;i<nums.length;i++)for(let j=i+1;j<nums.length;j++){
      const k=`${nums[i]}-${nums[j]}`;
      pairCount[k]=(pairCount[k]||0)+1;
    }
  });
  const pairArr=[["同期兩個號碼","出現次數"],...Object.entries(pairCount).sort((a,b)=>b[1]-a[1])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pairArr),'同期號碼組合（含特別號）');

  XLSX.writeFile(wb,'marksix_last20_with_special.xlsx');
}

load();