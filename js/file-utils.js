/* ========== FILE UTILITIES ========== */

/* Read Excel/CSV via XLSX library — returns Promise<Array<Object>> */
function readFileXLSX(file){
  return new Promise(resolve=>{
    const fr=new FileReader();
    fr.onload=()=>{
      const data=new Uint8Array(fr.result);
      const wb=XLSX.read(data,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const json=XLSX.utils.sheet_to_json(ws,{defval:'',raw:false});
      resolve(json);
    };
    fr.readAsArrayBuffer(file);
  });
}

/* Read file as plain text — returns Promise<string> */
function readFileText(file){
  return new Promise(resolve=>{
    const r=new FileReader();
    r.onload=()=>resolve(r.result);
    r.readAsText(file);
  });
}

/* Parse CSV text into array of arrays */
function parseCSV(text){
  const lines=text.split(/\r?\n/);
  const rows=[];
  for(const line of lines){
    const row=[];
    let inQuote=false,cell='';
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(ch==='"'){inQuote=!inQuote;}
      else if(ch===','&&!inQuote){row.push(cell.trim());cell='';}
      else{cell+=ch;}
    }
    row.push(cell.trim());
    rows.push(row);
  }
  return rows;
}

/* Extract dimension sections from CSV rows (looks for #accounts header pattern) */
function extractSections(rows){
  const sections={};
  let i=0;
  while(i<rows.length){
    const row=rows[i];
    const headerIdx=row.findIndex(c=>/^#\s*accounts$/i.test(c));
    if(headerIdx>=0){
      let dimName='';
      for(let j=headerIdx-1;j>=0;j--){if(row[j]){dimName=row[j];break;}}
      const dimLower=dimName.toLowerCase();
      let key='';
      if(dimLower.includes('industr'))key='industry';
      else if(dimLower.includes('country')||dimLower.includes('geo'))key='geography';
      else if(dimLower.includes('employee')||dimLower.includes('size'))key='employee_size';
      else key=dimLower.replace(/[^a-z0-9]/g,'_');

      const data=[];
      let actualLabelCol=0;
      for(let j=0;j<headerIdx;j++){if(row[j])actualLabelCol=j;}

      i++;
      while(i<rows.length){
        const r=rows[i];
        const label=r[actualLabelCol]||'';
        const countStr=r[headerIdx]||'';
        if(!label&&!countStr)break;
        if(/grand\s*total/i.test(label)){i++;break;}
        const count=parseInt(countStr.replace(/,/g,''))||0;
        if(label||count){data.push({label:label||'(blank)',count});}
        i++;
      }
      const total=data.reduce((s,d)=>s+d.count,0);
      data.forEach(d=>d.pct=total?d.count/total:0);
      sections[key]={name:dimName,data,total};
    }else{i++;}
  }
  return sections;
}

/* Set up drag-and-drop on an upload box */
function setupDragDrop(boxId,onDrop){
  const b=document.getElementById(boxId);
  b.addEventListener('dragover',e=>{e.preventDefault();b.style.borderColor='var(--bg-brand)';});
  b.addEventListener('dragleave',()=>{b.style.borderColor='';});
  b.addEventListener('drop',e=>{
    e.preventDefault();b.style.borderColor='';
    const f=e.dataTransfer.files[0];
    if(!f)return;
    const inp=b.querySelector('input');
    const dt=new DataTransfer();dt.items.add(f);inp.files=dt.files;
    inp.dispatchEvent(new Event('change'));
  });
}
