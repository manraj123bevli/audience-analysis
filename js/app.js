/* ========== APP SHELL — NAV, STEPS, FILE UPLOAD ========== */

/* ---- Top-level navigation ---- */
function switchSection(key,btn){
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.app-section').forEach(s=>s.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('section-'+key).classList.add('active');
}

/* ---- Step wizard (Contact Analysis) ---- */
let CA_highestStep=1;

function goToStep(n){
  if(n>CA_highestStep)CA_highestStep=n;
  document.querySelectorAll('.step').forEach(s=>s.classList.remove('active'));
  document.getElementById('step'+n).classList.add('active');
  ['sd1','sd2','sd3'].forEach((id,i)=>{
    const el=document.getElementById(id);
    el.className='step-dot';
    if(i+1<n)el.classList.add('done');
    else if(i+1===n)el.classList.add('current');
    if(i+1<=CA_highestStep&&i+1!==n){
      el.style.cursor='pointer';
    }else{
      el.style.cursor='default';
    }
  });
  window.scrollTo({top:0,behavior:'smooth'});
}

function caNavToStep(n){
  if(n<=CA_highestStep) goToStep(n);
}

function goToStep2(){
  const cMap=autoDetect(CA.contactCols,CONTACT_FIELDS,CA.contactSamples);
  const aMap=autoDetect(CA.accountCols,ACCOUNT_FIELDS,CA.accountSamples);
  renderMappingCard('contact-mappings','Contacts File Columns',CONTACT_FIELDS,CA.contactCols,cMap,CA.contactSamples,'c');
  renderMappingCard('account-mappings','Accounts File Columns',ACCOUNT_FIELDS,CA.accountCols,aMap,CA.accountSamples,'a');
  goToStep(2);
}

function checkStep1(){
  document.getElementById('btn-next1').disabled=!(CA.contactsRaw.length&&CA.accountsRaw.length);
}

/* ---- Contact Analysis file uploads ---- */
document.getElementById('file-contacts').addEventListener('change',async e=>{
  const f=e.target.files[0];if(!f)return;
  document.getElementById('name-contacts').textContent=f.name;
  document.getElementById('box-contacts').classList.add('loaded');
  CA.contactsRaw=await readFileXLSX(f);
  CA.contactCols=Object.keys(CA.contactsRaw[0]||{});
  CA.contactSamples={};
  CA.contactCols.forEach(col=>{
    CA.contactSamples[col]=CA.contactsRaw.slice(0,50).map(r=>r[col]).filter(v=>v&&String(v).trim()).slice(0,3).map(v=>String(v).substring(0,60));
  });
  document.getElementById('meta-contacts').textContent=`${CA.contactsRaw.length.toLocaleString()} rows, ${CA.contactCols.length} columns`;
  checkStep1();
});

document.getElementById('file-accounts').addEventListener('change',async e=>{
  const f=e.target.files[0];if(!f)return;
  document.getElementById('name-accounts').textContent=f.name;
  document.getElementById('box-accounts').classList.add('loaded');
  CA.accountsRaw=await readFileXLSX(f);
  CA.accountCols=Object.keys(CA.accountsRaw[0]||{});
  CA.accountSamples={};
  CA.accountCols.forEach(col=>{
    CA.accountSamples[col]=CA.accountsRaw.slice(0,50).map(r=>r[col]).filter(v=>v&&String(v).trim()).slice(0,3).map(v=>String(v).substring(0,60));
  });
  document.getElementById('meta-accounts').textContent=`${CA.accountsRaw.length.toLocaleString()} rows, ${CA.accountCols.length} columns`;
  checkStep1();
});

/* ---- DC Step wizard (4 steps) ---- */
DC._highestStep=1; // track the furthest step reached

function dcGoToStep(n){
  if(n>DC._highestStep)DC._highestStep=n;
  document.querySelectorAll('.dc-step').forEach(s=>s.classList.remove('active'));
  document.getElementById('dc-step'+n).classList.add('active');
  ['dc-sd1','dc-sd2','dc-sd3','dc-sd4'].forEach((id,i)=>{
    const el=document.getElementById(id);
    el.className='step-dot';
    if(i+1<n)el.classList.add('done');
    else if(i+1===n)el.classList.add('current');
    // Visually indicate clickable completed steps
    if(i+1<=DC._highestStep&&i+1!==n){
      el.style.cursor='pointer';
    }else{
      el.style.cursor='default';
    }
  });
  window.scrollTo({top:0,behavior:'smooth'});
}

/* Navigate to a step via step indicator click — only if already visited */
function dcNavToStep(n){
  if(n<=DC._highestStep) dcGoToStep(n);
}

function dcGoToStep2(){
  // Auto-detect columns for both files
  const tMap=autoDetect(DC.targetCols,DC_FIELDS,DC.targetSamples);
  const cMap=autoDetect(DC.currentCols,DC_FIELDS,DC.currentSamples);
  dcRenderMappingCard('dc-target-mappings','Target Accounts Columns',DC.targetCols,tMap,DC.targetSamples,'t');
  dcRenderMappingCard('dc-current-mappings','Current Customers Columns',DC.currentCols,cMap,DC.currentSamples,'c');
  // Reset custom dimensions
  DC.customDimCounter=0;
  document.getElementById('dc-custom-dims').innerHTML='';
  // Build standard dimension checkboxes
  dcUpdateDimensions();
  dcGoToStep(2);
}

function dcCheckStep1(){
  document.getElementById('dc-btn-next1').disabled=!(DC.targetRaw.length&&DC.currentRaw.length);
}

function buildSamples(rows,cols){
  const samples={};
  cols.forEach(col=>{
    samples[col]=rows.slice(0,50).map(r=>r[col]).filter(v=>v&&String(v).trim()).slice(0,3).map(v=>String(v).substring(0,60));
  });
  return samples;
}

/* ---- Distribution Comparison file uploads ---- */
document.getElementById('dc-file-target').addEventListener('change',async e=>{
  const f=e.target.files[0];if(!f)return;
  document.getElementById('dc-name-target').textContent=f.name;
  document.getElementById('dc-box-target').classList.add('loaded');
  DC.targetRaw=await readFileXLSX(f);
  DC.targetCols=Object.keys(DC.targetRaw[0]||{});
  DC.targetSamples=buildSamples(DC.targetRaw,DC.targetCols);
  document.getElementById('dc-meta-target').textContent=`${DC.targetRaw.length.toLocaleString()} rows, ${DC.targetCols.length} columns`;
  dcCheckStep1();
});

document.getElementById('dc-file-current').addEventListener('change',async e=>{
  const f=e.target.files[0];if(!f)return;
  document.getElementById('dc-name-current').textContent=f.name;
  document.getElementById('dc-box-current').classList.add('loaded');
  DC.currentRaw=await readFileXLSX(f);
  DC.currentCols=Object.keys(DC.currentRaw[0]||{});
  DC.currentSamples=buildSamples(DC.currentRaw,DC.currentCols);
  document.getElementById('dc-meta-current').textContent=`${DC.currentRaw.length.toLocaleString()} rows, ${DC.currentCols.length} columns`;
  dcCheckStep1();
});

/* ---- Drag and drop for all upload boxes ---- */
setupDragDrop('box-contacts');
setupDragDrop('box-accounts');
setupDragDrop('dc-box-current');
setupDragDrop('dc-box-target');
