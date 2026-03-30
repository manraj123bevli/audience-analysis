/* ========== DISTRIBUTION COMPARISON ========== */
window.DC={
  targetRaw:[],currentRaw:[],
  targetCols:[],currentCols:[],
  targetSamples:{},currentSamples:{},
  customDimCounter:0,
  pendingDims:[],       // dimensions gathered from step 2, used through step 3→4
  reconciliations:{},   // dimKey → [{value, source, masterLabel, confidence}]
};

/* ========== STANDARD DIMENSION FIELDS ========== */
const DC_FIELDS=[
  {key:'industry',label:'Industry',required:false,
   hints:['industry','sector','vertical','industry group','industry_group','industry category']},
  {key:'country',label:'Country / Geography',required:false,
   hints:['country','country/region','country_region','geography','geo','region','nation','location country','hq country','headquarters country']},
  {key:'employeeSize',label:'Employee Size / Range',required:false,
   hints:['employee','employees','employee range','employee_range','headcount','company size','size','num employees','employee count','number of employees','revised: employee range','employee size']},
];

/* ========== NORMALIZATION & FUZZY MATCHING ========== */
function normalizeForMatch(s){
  return s.toLowerCase().replace(/&/g,'and').replace(/[^a-z0-9\s]/g,'').replace(/\s+/g,' ').trim();
}

/* Tokenize into meaningful words, splitting compound words like "healthcare" → ["health","care"] */
function tokenize(s){
  const norm=normalizeForMatch(s);
  const words=norm.split(/\s+/).filter(w=>w.length>1);
  // Also split camelCase/compound: "healthcare" → ["health","care"]
  const expanded=[];
  for(const w of words){
    expanded.push(w);
    // Try common compound splits
    const compounds=[
      [/^(health)(care)$/],
      [/^(bio)(tech|technology)$/],
      [/^(fin)(tech)$/],
      [/^(ed)(tech)$/],
      [/^(info)(sec|security)$/],
      [/^(cyber)(security)$/],
    ];
    for(const[re]of compounds){
      const m=w.match(re);
      if(m){expanded.push(m[1],m[2]);}
    }
  }
  return[...new Set(expanded)];
}

/* Score how similar two label strings are (0-1) */
function similarityScore(a,b){
  const tokA=tokenize(a);
  const tokB=tokenize(b);
  if(!tokA.length||!tokB.length)return 0;

  // Exact normalized match
  if(normalizeForMatch(a)===normalizeForMatch(b))return 1;

  // Word overlap (both directions)
  const setA=new Set(tokA);
  const setB=new Set(tokB);
  const common=[...setA].filter(w=>setB.has(w)&&w.length>2).length;
  const overlapScore=common/Math.max(setA.size,setB.size);

  // Substring containment boost: "banking" contains "bank", "financial" contains "financ"
  let substringBonus=0;
  for(const wa of setA){
    for(const wb of setB){
      if(wa.length>=4&&wb.length>=4){
        if(wa.includes(wb)||wb.includes(wa))substringBonus+=0.15;
      }
    }
  }

  return Math.min(overlapScore+substringBonus,1);
}

/* ========== EMPLOYEE SIZE NORMALIZATION ========== */
const EMP_BUCKETS=[
  {min:1,max:10,label:'1-10'},
  {min:11,max:50,label:'11-50'},
  {min:51,max:200,label:'51-200'},
  {min:201,max:500,label:'201-500'},
  {min:501,max:1000,label:'501-1,000'},
  {min:1001,max:5000,label:'1,001-5,000'},
  {min:5001,max:10000,label:'5,001-10,000'},
  {min:10001,max:Infinity,label:'10,001+'},
];

/* Expand k/K/m/M shorthand: "10k" → 10000, "1.5k" → 1500, "1m" → 1000000 */
function expandShorthand(s){
  return s.replace(/(\d+(?:\.\d+)?)\s*([kKmM])\b/g,(_,num,unit)=>{
    const n=parseFloat(num);
    if(unit==='k'||unit==='K')return String(Math.round(n*1000));
    if(unit==='m'||unit==='M')return String(Math.round(n*1000000));
    return _;
  });
}

function parseNumber(s){
  const cleaned=expandShorthand(String(s).replace(/[$~,]/g,'').trim());
  const m=cleaned.match(/^(\d+)/);
  return m?parseInt(m[1],10):null;
}

function extractRange(s){
  const cleaned=expandShorthand(String(s).replace(/,/g,'').replace(/employees?/gi,'').trim());
  // "10001+" or "10k+"
  const plusMatch=cleaned.match(/^(\d+)\s*\+/);
  if(plusMatch)return{min:parseInt(plusMatch[1],10),max:Infinity};
  // "201-500" or "201 - 500" or "201 to 500" or "1k-5k"
  const rangeMatch=cleaned.match(/^(\d+)\s*[-–—to]+\s*(\d+)/i);
  if(rangeMatch)return{min:parseInt(rangeMatch[1],10),max:parseInt(rangeMatch[2],10)};
  return null;
}

function numberToBucket(n){
  for(const b of EMP_BUCKETS){
    if(n>=b.min&&n<=b.max)return b.label;
  }
  return EMP_BUCKETS[EMP_BUCKETS.length-1].label;
}

/* Check which bucket a range's min and max fall into.
   Returns bucket label if both ends land in the SAME bucket, null if they span multiple. */
function rangeToBucket(range){
  if(range.max===Infinity){
    // "X+" — open-ended ranges: find the bucket where X starts
    // If X is at/near the top of a bucket, bump to next (e.g., "10k+" → "10,001+")
    for(let i=0;i<EMP_BUCKETS.length;i++){
      const b=EMP_BUCKETS[i];
      if(range.min>=b.min&&range.min<=b.max){
        const bucketSpan=b.max-b.min;
        const posInBucket=(range.min-b.min)/bucketSpan;
        if(posInBucket>=0.8&&i<EMP_BUCKETS.length-1)return EMP_BUCKETS[i+1].label;
        return b.label;
      }
    }
    return EMP_BUCKETS[EMP_BUCKETS.length-1].label;
  }
  // Bounded range: check if min and max fall in the same bucket
  let minBucket=null,maxBucket=null;
  for(let i=0;i<EMP_BUCKETS.length;i++){
    const b=EMP_BUCKETS[i];
    if(range.min>=b.min&&range.min<=b.max)minBucket=i;
    if(range.max>=b.min&&range.max<=b.max)maxBucket=i;
  }
  if(minBucket!==null&&maxBucket!==null&&minBucket===maxBucket){
    return EMP_BUCKETS[minBucket].label;
  }
  // Spans multiple buckets — return null to let it pass through to reconciliation
  return null;
}

/* Fix Excel date-mangled employee ranges: "Nov-50" → "11-50", "10-Feb" → "2-10", etc. */
const MONTH_TO_NUM={
  'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
  'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
  'january':1,'february':2,'march':3,'april':4,'june':6,
  'july':7,'august':8,'september':9,'october':10,'november':11,'december':12,
};

function fixExcelDateMangling(s){
  const trimmed=s.trim();
  // Pattern 1: "Nov-50", "Dec-10", "Aug-200" → month-number (Excel turned "11-50" into "Nov-50")
  const m1=trimmed.match(/^([A-Za-z]+)-(\d+)$/);
  if(m1){
    const monthNum=MONTH_TO_NUM[m1[1].toLowerCase()];
    if(monthNum!==undefined) return monthNum+'-'+m1[2];
  }
  // Pattern 2: "10-Feb", "50-Nov" → number-month (Excel turned "2-10" into "10-Feb" with day-month swap)
  const m2=trimmed.match(/^(\d+)-([A-Za-z]+)$/);
  if(m2){
    const monthNum=MONTH_TO_NUM[m2[2].toLowerCase()];
    if(monthNum!==undefined) return monthNum+'-'+m2[1];
  }
  return null;
}

function normalizeEmployeeValue(val){
  let s=String(val).trim();
  if(!s||s==='(blank)'||/^not\s*found$/i.test(s)||s==='-'||s==='N/A'||s==='n/a'||s==='NA')return'(blank)';

  // Fix XLSX date serial numbers: XLSX.js may convert "2-10" to 46063 (a date serial).
  // Employee counts should never be 5-digit numbers in the 40000-50000 range — those are date serials.
  // Convert them back: serial → Date → "month-day" → treat as employee range.
  if(/^\d{5}$/.test(s)){
    const n=parseInt(s,10);
    if(n>=30000&&n<=60000){
      // Almost certainly an Excel date serial, not an employee count
      const d=new Date((n-25569)*86400000); // Excel epoch offset
      const month=d.getUTCMonth()+1;
      const day=d.getUTCDate();
      s=month+'-'+day; // reconstruct the original range like "2-10", "11-50"
    }
  }

  // Fix Excel date mangling in text: "Nov-50" → "11-50", "10-Feb" → "2-10"
  const fixed=fixExcelDateMangling(s);
  if(fixed) s=fixed;

  // Check if it's already a known bucket label (exact match after normalization)
  const normS=s.replace(/,/g,'').replace(/employees?/gi,'').trim().toLowerCase();
  for(const b of EMP_BUCKETS){
    const normB=b.label.replace(/,/g,'').toLowerCase();
    if(normS===normB)return b.label;
  }

  // Try as a range string: "201-500", "1001 - 5000", "501 to 1,000"
  const range=extractRange(s);
  if(range){
    const bucket=rangeToBucket(range);
    if(bucket)return bucket;
    // Range spans multiple buckets — format it cleanly but don't auto-bucket.
    // It will flow into reconciliation for manual assignment.
    if(range.max===Infinity)return range.min.toLocaleString()+'+';
    return range.min.toLocaleString()+'-'+range.max.toLocaleString();
  }

  // Try as a plain number: "3500", "150 employees"
  const num=parseNumber(s);
  if(num!==null&&num>0)return numberToBucket(num);

  // Can't parse — return as-is
  return s;
}

/* ========== BUILD DISTRIBUTION FROM RAW DATA ========== */
function buildDistribution(rows,colName,dimLabel,dimKey){
  const counts=new Map();
  const isEmpSize=dimKey==='employeeSize';
  rows.forEach(r=>{
    let val=String(r[colName]||'').trim()||'(blank)';
    if(isEmpSize)val=normalizeEmployeeValue(val);
    counts.set(val,(counts.get(val)||0)+1);
  });
  const data=[...counts.entries()].map(([label,count])=>({label,count})).sort((a,b)=>b.count-a.count);
  const total=data.reduce((s,d)=>s+d.count,0);
  data.forEach(d=>d.pct=total?d.count/total:0);
  return{name:dimLabel,data,total};
}

/* ========== BUILD DISTRIBUTION WITH RECONCILIATION MAP ========== */
function buildDistributionWithMap(rows,colName,dimLabel,valueMap,dimKey){
  // valueMap: original value → master label
  const isEmpSize=dimKey==='employeeSize';
  const counts=new Map();
  rows.forEach(r=>{
    let raw=String(r[colName]||'').trim()||'(blank)';
    if(isEmpSize)raw=normalizeEmployeeValue(raw);
    const mapped=valueMap[raw]||raw;
    counts.set(mapped,(counts.get(mapped)||0)+1);
  });
  const data=[...counts.entries()].map(([label,count])=>({label,count})).sort((a,b)=>b.count-a.count);
  const total=data.reduce((s,d)=>s+d.count,0);
  data.forEach(d=>d.pct=total?d.count/total:0);
  return{name:dimLabel,data,total};
}

/* ========== RECONCILIATION ALGORITHM ========== */
const STANDARD_EMP_LABELS=new Set(EMP_BUCKETS.map(b=>b.label));
function isStandardBucket(label){
  return STANDARD_EMP_LABELS.has(label)||label==='(blank)';
}

function dcBuildReconciliation(targetValues,currentValues,dimKey){
  // targetValues/currentValues: [{label, count}]
  // Returns: {masterLabels: string[], mappings: [{value, source, masterLabel, confidence, count}]}

  const masterSet=new Map(); // normalized → display label
  const mappings=[];

  // Step 1: Target values become master labels — BUT for employee size,
  // non-standard ranges (like "2-50") are flagged as needing manual assignment
  for(const tv of targetValues){
    const norm=normalizeForMatch(tv.label);
    const isNonStandard=dimKey==='employeeSize'&&!isStandardBucket(tv.label);
    if(!isNonStandard){
      if(!masterSet.has(norm)) masterSet.set(norm,tv.label);
      mappings.push({value:tv.label,source:'target',masterLabel:tv.label,confidence:'exact',count:tv.count});
    }else{
      // Non-standard range from target — needs manual assignment
      mappings.push({value:tv.label,source:'target',masterLabel:tv.label,confidence:'manual',count:tv.count});
    }
  }

  // Step 2: Try to map each current value to a master label
  for(const cv of currentValues){
    const norm=normalizeForMatch(cv.label);

    // 2a: Exact normalized match to a target value
    if(masterSet.has(norm)){
      mappings.push({value:cv.label,source:'current',masterLabel:masterSet.get(norm),confidence:'exact',count:cv.count});
      continue;
    }

    // 2b: Similarity scoring against all master labels
    let bestMaster=null,bestScore=0;
    for(const[mNorm,mLabel]of masterSet){
      const score=similarityScore(cv.label,mLabel);
      if(score>bestScore){bestScore=score;bestMaster=mLabel;}
    }

    if(bestScore>=0.6){
      // High confidence — auto match
      mappings.push({value:cv.label,source:'current',masterLabel:bestMaster,confidence:'auto',count:cv.count});
    }else if(bestScore>=0.3){
      // Medium confidence — suggest but flag for review
      mappings.push({value:cv.label,source:'current',masterLabel:bestMaster,confidence:'suggest',count:cv.count});
    }else{
      // No match — unmapped, needs manual assignment
      masterSet.set(norm,cv.label);
      mappings.push({value:cv.label,source:'current',masterLabel:cv.label,confidence:'manual',count:cv.count});
    }
  }

  const masterLabels=[...new Set(mappings.filter(m=>m.source==='target').map(m=>m.masterLabel))];
  // Add any current-only master labels
  const currentOnly=mappings.filter(m=>m.source==='current'&&m.confidence==='manual').map(m=>m.masterLabel);
  currentOnly.forEach(l=>{if(!masterLabels.includes(l))masterLabels.push(l);});

  return{masterLabels,mappings};
}

/* ========== STEP 2: MAPPING & DIMENSION UI ========== */
function dcRenderMappingCard(containerId,title,cols,mapping,samples,prefix){
  const el=document.getElementById(containerId);
  let html=`<h3>${title}</h3>`;
  DC_FIELDS.forEach(field=>{
    const m=mapping[field.key];
    const selId='dc-'+prefix+'_'+field.key;
    html+=`<div class="map-row">
      <span class="map-label">${field.label}</span>
      <select class="map-select" id="${selId}" onchange="dcUpdatePreview('${selId}','${prefix}');dcUpdateDimensions()">
        <option value="">— not mapped —</option>
        ${cols.map(c=>`<option value="${esc(c)}" ${m.col===c?'selected':''}>${esc(c)}</option>`).join('')}
      </select>
      ${m.col&&m.auto?'<span class="map-badge map-auto">auto</span>':m.col?'<span class="map-badge map-manual">suggest</span>':''}
    </div>
    <div class="preview-row" id="prev_${selId}">${m.col&&samples[m.col]?'e.g. '+samples[m.col].join(' | '):''}</div>`;
  });
  el.innerHTML=html;
}

function dcUpdatePreview(selId,prefix){
  const sel=document.getElementById(selId);
  const col=sel.value;
  const samples=prefix==='t'?DC.targetSamples:DC.currentSamples;
  const prev=document.getElementById('prev_'+selId);
  prev.textContent=col&&samples[col]?'e.g. '+samples[col].join(' | '):'';
}

function dcGetMapping(prefix,key){
  const el=document.getElementById('dc-'+prefix+'_'+key);
  return el?el.value:'';
}

function dcUpdateDimensions(){
  const container=document.getElementById('dc-standard-dims');
  let html='';
  DC_FIELDS.forEach(field=>{
    const tCol=dcGetMapping('t',field.key);
    const cCol=dcGetMapping('c',field.key);
    const bothMapped=!!(tCol&&cCol);
    html+=`<div class="dim-row ${bothMapped?'':'disabled'}">
      <input type="checkbox" id="dc-dim-${field.key}" ${bothMapped?'checked':''} ${bothMapped?'':'disabled'}>
      <span class="dim-label">${field.label}</span>
      <span class="dim-cols">
        ${tCol?`Target: <span>${esc(tCol)}</span>`:'<span class="dim-unavailable">Target: not mapped</span>'}
        &nbsp;&nbsp;|&nbsp;&nbsp;
        ${cCol?`Current: <span>${esc(cCol)}</span>`:'<span class="dim-unavailable">Current: not mapped</span>'}
      </span>
    </div>`;
  });
  container.innerHTML=html;
}

function dcAddCustomDimension(){
  DC.customDimCounter++;
  const id=DC.customDimCounter;
  const container=document.getElementById('dc-custom-dims');
  const row=document.createElement('div');
  row.className='custom-dim-row';
  row.id='dc-custom-'+id;
  row.innerHTML=`
    <input type="text" placeholder="Label" id="dc-cdlabel-${id}" value="">
    <select id="dc-cdtarget-${id}" onchange="dcAutoFillCustomLabel(${id})">
      <option value="">Target column...</option>
      ${DC.targetCols.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('')}
    </select>
    <span class="dim-arrow">&#8596;</span>
    <select id="dc-cdcurrent-${id}">
      <option value="">Current column...</option>
      ${DC.currentCols.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('')}
    </select>
    <button class="custom-dim-remove" onclick="this.parentElement.remove()" title="Remove">&times;</button>`;
  container.appendChild(row);
}

function dcAutoFillCustomLabel(id){
  const labelEl=document.getElementById('dc-cdlabel-'+id);
  const targetEl=document.getElementById('dc-cdtarget-'+id);
  if(!labelEl.value&&targetEl.value) labelEl.value=targetEl.value;
}

function dcGatherDimensions(){
  const dims=[];
  DC_FIELDS.forEach(field=>{
    const cb=document.getElementById('dc-dim-'+field.key);
    if(cb&&cb.checked){
      const tCol=dcGetMapping('t',field.key);
      const cCol=dcGetMapping('c',field.key);
      if(tCol&&cCol) dims.push({key:field.key,label:field.label,targetCol:tCol,currentCol:cCol});
    }
  });
  const customRows=document.querySelectorAll('[id^="dc-custom-"]');
  customRows.forEach(row=>{
    const id=row.id.replace('dc-custom-','');
    const label=document.getElementById('dc-cdlabel-'+id)?.value||'';
    const tCol=document.getElementById('dc-cdtarget-'+id)?.value||'';
    const cCol=document.getElementById('dc-cdcurrent-'+id)?.value||'';
    if(label&&tCol&&cCol) dims.push({key:'custom_'+id,label,targetCol:tCol,currentCol:cCol});
  });
  return dims;
}

/* ========== STEP 3: RECONCILIATION UI (one dimension at a time) ========== */
function dcPrepareReconciliation(){
  const dims=dcGatherDimensions();
  if(dims.length===0){alert('Please select at least one dimension to compare.');return;}
  DC.pendingDims=dims;
  DC.reconciliations={};
  DC._reconPanelsHtml={};
  DC._mergeState={};
  DC._recsCache={};

  // Pre-compute reconciliations and render all panels
  for(const dim of dims){
    const targetDist=buildDistribution(DC.targetRaw,dim.targetCol,dim.label,dim.key);
    const currentDist=buildDistribution(DC.currentRaw,dim.currentCol,dim.label,dim.key);
    const recon=dcBuildReconciliation(targetDist.data,currentDist.data,dim.key);
    DC.reconciliations[dim.key]=recon;
    DC._reconPanelsHtml[dim.key]=dcRenderReconPanel(dim,recon);
  }

  dcRenderReconStep();
  dcGoToStep(3);
}

/* Navigate back to step 3 from step 4, preserving reconciliation state */
function dcBackToReconcile(){
  if(!DC.pendingDims||!DC.pendingDims.length){dcGoToStep(2);return;}
  // Re-render panels from existing reconciliations
  for(const dim of DC.pendingDims){
    DC._reconPanelsHtml[dim.key]=dcRenderReconPanel(dim,DC.reconciliations[dim.key]);
  }
  dcRenderReconStep();
  dcGoToStep(3);
}

function dcRenderReconStep(){
  const dims=DC.pendingDims;
  const tabsEl=document.getElementById('dc-recon-tabs');
  tabsEl.innerHTML=dims.map((dim,i)=>{
    const unmatchedCount=DC.reconciliations[dim.key].mappings.filter(m=>m.confidence==='manual').length;
    const badge=unmatchedCount?` <span class="badge b-red" style="margin-left:4px">${unmatchedCount}</span>`:'';
    return `<button class="tab-btn${i===0?' active':''}" onclick="dcShowReconDim('${dim.key}',this)">${esc(dim.label)}${badge}</button>`;
  }).join('');

  const container=document.getElementById('dc-reconciliation-panels');
  container.innerHTML=dims.map((dim,i)=>
    `<div class="dc-recon-dim${i===0?' active':''}" id="dc-recon-dim-${dim.key}">${DC._reconPanelsHtml[dim.key]}</div>`
  ).join('');
}

function dcShowReconDim(dimKey,btn){
  document.querySelectorAll('#dc-recon-tabs .tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.dc-recon-dim').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('dc-recon-dim-'+dimKey).classList.add('active');
}

function dcRenderReconPanel(dim,recon){
  // Build groups from reconciliation mappings
  const groups=new Map();
  for(const m of recon.mappings){
    const key=m.masterLabel;
    if(!groups.has(key))groups.set(key,{targetCount:0,currentCount:0,members:[]});
    const g=groups.get(key);
    if(m.source==='target')g.targetCount+=m.count;
    else g.currentCount+=m.count;
    g.members.push(m);
  }

  // Merge state
  if(!DC._mergeState)DC._mergeState={};
  if(!DC._mergeState[dim.key])DC._mergeState[dim.key]={};
  const ms=DC._mergeState[dim.key];

  const allLabelsUnsorted=[...groups.keys()];
  const independentLabels=allLabelsUnsorted.filter(l=>!ms[l]);

  // Reverse map: which labels are absorbed into each label
  const absorbedBy={};
  for(const l of allLabelsUnsorted)absorbedBy[l]=[];
  for(const[src,tgt]of Object.entries(ms)){if(tgt&&absorbedBy[tgt])absorbedBy[tgt].push(src);}

  // Sort
  if(!DC._reconSort)DC._reconSort={};
  const sort=DC._reconSort[dim.key]||{col:'total',dir:'desc'};
  const entries=[...groups.entries()];
  const sortFns={
    value:(a,b)=>a[0].localeCompare(b[0]),
    target:(a,b)=>a[1].targetCount-b[1].targetCount,
    current:(a,b)=>a[1].currentCount-b[1].currentCount,
    total:(a,b)=>(a[1].targetCount+a[1].currentCount)-(b[1].targetCount+b[1].currentCount),
  };
  const fn=sortFns[sort.col]||sortFns.total;
  entries.sort((a,b)=>sort.dir==='asc'?fn(a,b):-fn(a,b));
  const sorted=entries;
  const allLabels=sorted.map(([l])=>l);

  // Precompute similarity recommendations (cached per dimension)
  if(!DC._recsCache)DC._recsCache={};
  if(!DC._recsCache[dim.key]){
    DC._recsCache[dim.key]={};
    for(const label of allLabels){
      DC._recsCache[dim.key][label]=allLabels
        .filter(l=>l!==label)
        .map(l=>({label:l,score:similarityScore(label,l)}))
        .filter(s=>s.score>=0.15)
        .sort((a,b)=>b.score-a.score)
        .slice(0,5);
    }
  }

  const mergedCount=allLabels.length-independentLabels.length;

  let html=`<div style="margin-bottom:var(--space-400);font-size:13px;color:var(--text-subtle)">${allLabels.length} unique values &rarr; <strong>${independentLabels.length} final groups</strong>${mergedCount?` <span style="color:var(--text-warning)">(${mergedCount} merged)</span>`:''}</div>`;

  // Available to absorb: independent labels not already absorbed by someone else
  const availableForAbsorb=allLabels.filter(l=>!ms[l]);

  function sortHdr(label,colKey,width,cls){
    const arrow=sort.col===colKey?(sort.dir==='asc'?' &#9650;':' &#9660;'):'';
    const style=width?`style="width:${width};cursor:pointer"`:'style="cursor:pointer"';
    return`<th class="${cls||''}" ${style} onclick="dcReconSort('${dim.key}','${colKey}')">${label}${arrow}</th>`;
  }
  html+=`<div class="table-wrap scroll-table"><table class="recon-table">
    <thead><tr>
      ${sortHdr('Value','value','')}
      ${sortHdr('Target #','target','80px','num')}
      ${sortHdr('Current #','current','80px','num')}
      <th>Merge Into This</th>
      <th>Suggested Matches</th>
    </tr></thead><tbody>`;

  for(const[label,data]of sorted){
    const mergedInto=ms[label]||'';
    const isMerged=!!mergedInto;
    const absorbed=absorbedBy[label]||[];
    const hasManual=data.members.some(m=>m.confidence==='manual');

    // Show raw member values if group has multiple
    const rawVals=[...new Set(data.members.map(m=>m.value))];
    const memberHint=rawVals.length>1
      ?`<div style="font-size:11px;color:var(--text-subtle);margin-top:2px">Includes: ${rawVals.map(v=>esc(v)).join(', ')}</div>`
      :'';

    // Absorb dropdown: show available values (independent, not self)
    const absorbOpts=availableForAbsorb
      .filter(l=>l!==label)
      .map(l=>{const g=groups.get(l);return`<option value="${esc(l)}">${esc(l)} (${g.targetCount}T / ${g.currentCount}C)</option>`})
      .join('');

    // Absorbed chips (with remove X)
    const absorbedChips=absorbed.map(src=>{
      const g=groups.get(src);
      const cnt=g?(g.targetCount+g.currentCount):0;
      return`<span class="recon-absorbed-chip">${esc(src)} <span class="recon-absorbed-count">(${cnt})</span> <span class="recon-absorbed-x" data-dim="${dim.key}" data-source="${esc(src)}" onclick="event.stopPropagation();dcReconUnmerge(this.dataset.dim,this.dataset.source)">&times;</span></span>`;
    }).join(' ');

    // Recommendations: only show values that are still available to absorb
    const recLabels=new Set((DC._recsCache[dim.key][label]||[]).filter(r=>!ms[r.label]&&r.label!==label).slice(0,5).map(r=>r.label));
    const recChips=[...recLabels].map(l=>`<span class="recon-rec-chip" data-dim="${dim.key}" data-source="${esc(l)}" data-target="${esc(label)}" onclick="dcReconAbsorb(this.dataset.dim,this.dataset.target,this.dataset.source)">${esc(l)}</span>`).join(' ');

    // All other available values not in top recommendations
    const restLabels=availableForAbsorb.filter(l=>l!==label&&!recLabels.has(l));
    const toggleId=`dc-rest-${dim.key}-${gi}`;
    const restChips=restLabels.map(l=>`<span class="recon-rec-chip recon-rest-chip" data-dim="${dim.key}" data-source="${esc(l)}" data-target="${esc(label)}" onclick="dcReconAbsorb(this.dataset.dim,this.dataset.target,this.dataset.source)">${esc(l)}</span>`).join(' ');
    const restToggle=restLabels.length
      ?`<span class="recon-show-all" onclick="this.nextElementSibling.style.display='inline';this.style.display='none'">+${restLabels.length} more</span><span style="display:none">${restChips}</span>`
      :'';

    const recHtml=recChips||restToggle
      ?`${recChips}${recChips&&restToggle?' ':''}${restToggle}`
      :'<span style="color:var(--text-disabled);font-size:11px">&mdash;</span>';

    html+=`<tr class="${isMerged?'recon-row-merged':''}">
      <td><strong>${esc(label)}</strong>${hasManual?' <span class="badge b-red" style="font-size:10px">Unmatched</span>':''}${isMerged?` <span style="font-size:11px;color:var(--text-subtle)">&rarr; merged into ${esc(mergedInto)}</span>`:''}${memberHint}</td>
      <td class="num">${data.targetCount.toLocaleString()}</td>
      <td class="num">${data.currentCount.toLocaleString()}</td>
      <td>${isMerged
        ?'<span style="font-size:12px;color:var(--text-disabled)">Absorbed</span>'
        :`<select class="recon-select" data-dim="${dim.key}" data-label="${esc(label)}" onchange="dcReconAbsorb(this.dataset.dim,this.dataset.label,this.value);this.selectedIndex=0;">
          <option value="">+ Add value&hellip;</option>
          ${absorbOpts}
        </select>${absorbedChips?`<div style="margin-top:4px">${absorbedChips}</div>`:''}`
      }</td>
      <td>${isMerged?'':recHtml}</td>
    </tr>`;
  }

  html+=`</tbody></table></div>`;
  return html;
}

/* ===== Reconciliation table event handlers ===== */
function dcReconAbsorb(dimKey,targetLabel,sourceLabel){
  if(!sourceLabel||!targetLabel||sourceLabel===targetLabel)return;
  if(!DC._mergeState)DC._mergeState={};
  if(!DC._mergeState[dimKey])DC._mergeState[dimKey]={};
  DC._mergeState[dimKey][sourceLabel]=targetLabel;
  dcRefreshReconPanel(dimKey);
}

function dcReconUnmerge(dimKey,sourceLabel){
  if(DC._mergeState&&DC._mergeState[dimKey]){
    delete DC._mergeState[dimKey][sourceLabel];
  }
  dcRefreshReconPanel(dimKey);
}

function dcReconSort(dimKey,colKey){
  if(!DC._reconSort)DC._reconSort={};
  const cur=DC._reconSort[dimKey]||{col:'total',dir:'desc'};
  if(cur.col===colKey){
    cur.dir=cur.dir==='asc'?'desc':'asc';
  }else{
    cur.col=colKey;
    cur.dir=colKey==='value'?'asc':'desc';
  }
  DC._reconSort[dimKey]=cur;
  dcRefreshReconPanel(dimKey);
}

function dcRefreshReconPanel(dimKey){
  const dim=DC.pendingDims.find(d=>d.key===dimKey);
  const recon=DC.reconciliations[dimKey];
  const container=document.getElementById('dc-recon-dim-'+dimKey);
  container.innerHTML=dcRenderReconPanel(dim,recon);
  // Update tab badge
  const ms=DC._mergeState[dimKey]||{};
  const mergedCount=Object.keys(ms).filter(k=>ms[k]).length;
  const tabs=document.querySelectorAll('#dc-recon-tabs .tab-btn');
  tabs.forEach(btn=>{
    if(btn.textContent.includes(dim.label.split(' ')[0])){
      const existing=btn.querySelector('.badge');
      if(existing)existing.remove();
      if(mergedCount){
        btn.insertAdjacentHTML('beforeend',` <span class="badge b-grn" style="margin-left:4px">${mergedCount} merged</span>`);
      }
    }
  });
}

/* ========== STEP 3→4: CONFIRM AND RUN ========== */
function dcConfirmAndRun(){
  const dims=DC.pendingDims;
  const dimensions=[];

  for(const dim of dims){
    const valueMap={};
    const recon=DC.reconciliations[dim.key];
    const ms=DC._mergeState&&DC._mergeState[dim.key]?DC._mergeState[dim.key]:{};

    // Resolve merge chains: if A→B and B→C, A should end up in C
    function resolveTarget(label){
      const seen=new Set();
      let current=label;
      while(ms[current]&&!seen.has(current)){
        seen.add(current);
        current=ms[current];
      }
      return current;
    }

    if(recon){
      for(const m of recon.mappings){
        const resolved=resolveTarget(m.masterLabel);
        valueMap[m.value]=resolved;
      }
    }

    const targetDist=buildDistributionWithMap(DC.targetRaw,dim.targetCol,dim.label,valueMap,dim.key);
    const currentDist=buildDistributionWithMap(DC.currentRaw,dim.currentCol,dim.label,valueMap,dim.key);

    const comp=compareSections(targetDist,currentDist,null);
    dimensions.push({key:dim.key,name:dim.label,comparison:comp,target:targetDist,current:currentDist,showMapped:false,
      targetCol:dim.targetCol,currentCol:dim.currentCol,valueMap});
  }

  dcGoToStep(4);
  dcRenderResults(dimensions);
}

/* ========== COMPARISON LOGIC ========== */
function compareSections(targetSection,currentSection,mapFn){
  const results=[];
  const usedCurrent=new Set();
  const targetNorm=targetSection.data.map(d=>({...d,norm:normalizeForMatch(d.label)}));
  const currentMapped=currentSection.data.map(d=>{
    const mapped=mapFn?mapFn(d.label):normalizeForMatch(d.label);
    return{...d,mappedTo:mapped};
  });

  for(const t of targetNorm){
    const matching=currentMapped.filter(c=>{
      if(mapFn)return c.mappedTo===t.norm||c.mappedTo===normalizeForMatch(t.label);
      return normalizeForMatch(c.label)===t.norm;
    });
    matching.forEach(m=>usedCurrent.add(m.label));
    const curCount=matching.reduce((s,m)=>s+m.count,0);
    const curPct=currentSection.total?curCount/currentSection.total:0;
    results.push({
      targetLabel:t.label,targetCount:t.count,targetPct:t.pct,
      currentLabels:matching.map(m=>m.label),currentCount:curCount,currentPct:curPct,
      delta:curPct-t.pct,
    });
  }

  for(const c of currentMapped){
    if(!usedCurrent.has(c.label)){
      results.push({
        targetLabel:'(no target match) '+c.label,targetCount:0,targetPct:0,
        currentLabels:[c.label],currentCount:c.count,
        currentPct:currentSection.total?c.count/currentSection.total:0,
        delta:currentSection.total?c.count/currentSection.total:0,
      });
    }
  }
  return results;
}

function getAlignment(delta){
  const abs=Math.abs(delta);
  if(abs<=0.03)return{label:'Aligned',cls:'align-good'};
  if(abs<=0.07)return{label:'Minor gap',cls:'align-warn'};
  return{label:'Significant gap',cls:'align-bad'};
}

/* ========== EXECUTIVE SUMMARY ========== */
function dcBuildExecSummary(dimensions){
  const el=document.getElementById('dc-exec-summary');
  if(!el)return;

  const tTotal=DC.targetRaw.length;
  const cTotal=DC.currentRaw.length;

  // Analyze each dimension relative to target
  const dimSummaries=dimensions.map(dim=>{
    const comp=dim.comparison;
    const avgAbsDelta=comp.reduce((s,r)=>s+Math.abs(r.delta),0)/comp.length;

    // Gaps: where current customer mix doesn't match target profile
    // Negative delta = under-penetrated vs target (we need MORE here)
    // Positive delta = over-concentrated vs target (we're heavy here relative to where we should be)
    const underPenetrated=[...comp].filter(r=>r.delta<-0.03&&r.targetPct>0).sort((a,b)=>a.delta-b.delta);
    const overConcentrated=[...comp].filter(r=>r.delta>0.03).sort((a,b)=>b.delta-a.delta);
    const onTrack=comp.filter(r=>Math.abs(r.delta)<=0.03&&r.targetPct>0);

    // Whitespace: categories in target with zero or near-zero current presence
    const whitespace=[...comp].filter(r=>r.targetPct>=0.02&&r.currentCount===0).sort((a,b)=>b.targetPct-a.targetPct);

    // Biggest single gap
    const biggestGap=underPenetrated.length>0?underPenetrated[0]:null;

    // Total target share represented by under-penetrated categories
    const underPenTargetShare=underPenetrated.reduce((s,r)=>s+r.targetPct,0);

    const health=avgAbsDelta<=0.03?'well':avgAbsDelta<=0.08?'moderate':'poor';

    return{name:dim.name,avgAbsDelta,health,comp,
      underPenetrated,overConcentrated,onTrack,whitespace,biggestGap,underPenTargetShare,
      totalCategories:comp.filter(r=>r.targetPct>0).length};
  });

  // Overall
  const overallAvg=dimSummaries.reduce((s,d)=>s+d.avgAbsDelta,0)/dimSummaries.length;
  const overallCls=overallAvg<=0.03?'health-good':overallAvg<=0.08?'health-warn':'health-bad';

  // Collect top action items across all dimensions
  const actionItems=[];
  for(const ds of dimSummaries){
    for(const r of ds.whitespace.slice(0,2)){
      actionItems.push({priority:'high',text:`${(r.targetPct*100).toFixed(1)}% of target list is in <strong>${esc(r.targetLabel)}</strong> (${ds.name}), but there are no current customers there`});
    }
    for(const r of ds.underPenetrated.slice(0,2)){
      if(r.currentCount>0){
        actionItems.push({priority:'medium',text:`${(r.targetPct*100).toFixed(1)}% of target list is in <strong>${esc(r.targetLabel)}</strong> (${ds.name}), but only ${(r.currentPct*100).toFixed(1)}% of current customers are — gap of ${(Math.abs(r.delta)*100).toFixed(1)}pp`});
      }
    }
  }

  let html=`<h2>Executive Summary</h2>`;

  // Top-line narrative
  html+=`<div class="exec-overview">`;
  html+=`<p>Your target list of <strong>${tTotal.toLocaleString()} accounts</strong> defines where you want to be. Your current <strong>${cTotal.toLocaleString()} customers</strong> are compared below to identify gaps and opportunities across <strong>${dimensions.length} dimension${dimensions.length>1?'s':''}</strong>.</p>`;

  // One-line verdict
  const wellCount=dimSummaries.filter(d=>d.health==='well').length;
  const poorCount=dimSummaries.filter(d=>d.health==='poor').length;
  if(poorCount>0){
    html+=`<p><span class="exec-health health-bad">Action needed</span> — ${poorCount} of ${dimSummaries.length} dimension${dimSummaries.length>1?'s':''} show${poorCount===1?'s':''} significant misalignment with your target profile.</p>`;
  }else if(wellCount===dimSummaries.length){
    html+=`<p><span class="exec-health health-good">On track</span> — your current customer mix closely matches your target profile across all dimensions.</p>`;
  }else{
    html+=`<p><span class="exec-health health-warn">Partially aligned</span> — some dimensions match your target profile, others have gaps to close.</p>`;
  }
  html+=`</div>`;

  // Priority actions
  if(actionItems.length>0){
    html+=`<div class="exec-actions">`;
    html+=`<h3>Key Gaps to Close</h3>`;
    html+=`<ul>`;
    for(const item of actionItems.slice(0,6)){
      const cls=item.priority==='high'?'action-high':'action-medium';
      html+=`<li class="${cls}">${item.text}</li>`;
    }
    html+=`</ul></div>`;
  }


  el.innerHTML=html;
  el.style.display='';
}

/* ========== RENDER RESULTS (Step 4) ========== */

// Store results data for sorting
DC._resultDimensions=[];

function dcRenderResults(dimensions){
  DC._resultDimensions=dimensions;
  const tTotal=DC.targetRaw.length;
  const cTotal=DC.currentRaw.length;

  dcBuildExecSummary(dimensions);

  document.getElementById('dc-total-target').textContent=tTotal.toLocaleString();
  document.getElementById('dc-total-current').textContent=cTotal.toLocaleString();

  const summaryGrid=document.getElementById('dc-summary-grid');
  summaryGrid.innerHTML='';
  for(const dim of dimensions){
    const avgAbsDelta=dim.comparison.reduce((s,r)=>s+Math.abs(r.delta),0)/dim.comparison.length;
    const maxDelta=dim.comparison.reduce((m,r)=>Math.abs(r.delta)>Math.abs(m.delta)?r:m,dim.comparison[0]);
    const align=avgAbsDelta<=0.03?{label:'Well Aligned',cls:'b-grn'}:avgAbsDelta<=0.08?{label:'Moderately Aligned',cls:'b-ylw'}:{label:'Poorly Aligned',cls:'b-red'};
    const card=document.createElement('div');
    card.className='summary-card';
    card.innerHTML=`
      <h3>${esc(dim.name)}</h3>
      <div class="score"><span class="badge ${align.cls}">${align.label}</span></div>
      <div class="detail">Biggest gap: <strong>${esc(maxDelta.targetLabel)}</strong> (${maxDelta.delta>=0?'+':''}${(maxDelta.delta*100).toFixed(1)}% delta)</div>`;
    summaryGrid.appendChild(card);
  }

  const tabsEl=document.getElementById('dc-tabs');
  const panelsEl=document.getElementById('dc-tab-panels');
  tabsEl.innerHTML='';
  panelsEl.innerHTML='';

  dimensions.forEach((dim,idx)=>{
    const btn=document.createElement('button');
    btn.className='tab-btn'+(idx===0?' active':'');
    btn.textContent=dim.name;
    btn.onclick=()=>switchTab(dim.key,btn,'dc-results-area');
    tabsEl.appendChild(btn);

    const panel=document.createElement('div');
    panel.className='tab-panel'+(idx===0?' active':'');
    panel.id='panel-'+dim.key;
    panelsEl.appendChild(panel);

    // Initial render (no sort)
    dcRenderSortableTable(dim.key,dim.comparison,dim.target.total,dim.current.total,null,null);
  });
}

/* ========== SORTABLE COMPARISON TABLE ========== */
const DC_SORT_COLS=[
  {key:'category',label:'Category',numeric:false,getter:r=>r.targetLabel},
  {key:'targetCount',label:'Target #',numeric:true,getter:r=>r.targetCount},
  {key:'targetPct',label:'Target %',numeric:true,getter:r=>r.targetPct},
  {key:'currentCount',label:'Current #',numeric:true,getter:r=>r.currentCount},
  {key:'currentPct',label:'Current %',numeric:true,getter:r=>r.currentPct},
  {key:'delta',label:'Delta',numeric:true,getter:r=>r.delta},
  {key:'alignment',label:'Alignment',numeric:true,getter:r=>Math.abs(r.delta)},
];

// Track current sort state per dimension
DC._sortState={};

function dcSortTable(dimKey,colKey){
  const state=DC._sortState[dimKey]||{col:null,dir:null};
  if(state.col===colKey){
    state.dir=state.dir==='asc'?'desc':'asc';
  }else{
    state.col=colKey;
    // Default: numeric cols start descending, text cols start ascending
    const colDef=DC_SORT_COLS.find(c=>c.key===colKey);
    state.dir=colDef&&colDef.numeric?'desc':'asc';
  }
  DC._sortState[dimKey]=state;

  const dim=DC._resultDimensions.find(d=>d.key===dimKey);
  if(dim) dcRenderSortableTable(dimKey,dim.comparison,dim.target.total,dim.current.total,state.col,state.dir);
}

/* ========== DRILLDOWN: show companies for a category ========== */
function dcGuessNameCol(cols){
  const hints=['company name','company','account name','account','name','organization','org name','organisation','company_name','account_name'];
  for(const h of hints){
    const match=cols.find(c=>c.toLowerCase().trim()===h);
    if(match)return match;
  }
  // Partial match
  for(const h of hints){
    const match=cols.find(c=>c.toLowerCase().includes(h));
    if(match)return match;
  }
  return cols[0]; // fallback to first column
}

function dcToggleDrilldown(dimKey,categoryLabel,rowEl){
  const existing=rowEl.nextElementSibling;
  if(existing&&existing.classList.contains('drilldown-row')){
    existing.remove();
    rowEl.classList.remove('drilldown-open');
    return;
  }

  // Close any other open drilldown in this table
  const table=rowEl.closest('table');
  table.querySelectorAll('.drilldown-row').forEach(r=>r.remove());
  table.querySelectorAll('.drilldown-open').forEach(r=>r.classList.remove('drilldown-open'));

  const dim=DC._resultDimensions.find(d=>d.key===dimKey);
  if(!dim)return;

  const isEmpSize=dimKey==='employeeSize';
  const targetNameCol=dcGuessNameCol(DC.targetCols);
  const currentNameCol=dcGuessNameCol(DC.currentCols);

  // Find target accounts matching this category
  const targetAccounts=[];
  DC.targetRaw.forEach(r=>{
    let raw=String(r[dim.targetCol]||'').trim()||'(blank)';
    if(isEmpSize)raw=normalizeEmployeeValue(raw);
    const mapped=dim.valueMap[raw]||raw;
    if(mapped===categoryLabel){
      targetAccounts.push({
        name:String(r[targetNameCol]||'').trim()||'(unnamed)',
        rawValue:String(r[dim.targetCol]||'').trim()||'(blank)'
      });
    }
  });

  // Find current accounts matching this category
  const currentAccounts=[];
  DC.currentRaw.forEach(r=>{
    let raw=String(r[dim.currentCol]||'').trim()||'(blank)';
    if(isEmpSize)raw=normalizeEmployeeValue(raw);
    const mapped=dim.valueMap[raw]||raw;
    if(mapped===categoryLabel){
      currentAccounts.push({
        name:String(r[currentNameCol]||'').trim()||'(unnamed)',
        rawValue:String(r[dim.currentCol]||'').trim()||'(blank)'
      });
    }
  });

  targetAccounts.sort((a,b)=>a.name.localeCompare(b.name));
  currentAccounts.sort((a,b)=>a.name.localeCompare(b.name));

  const colSpan=7;
  const tr=document.createElement('tr');
  tr.className='drilldown-row';
  const td=document.createElement('td');
  td.colSpan=colSpan;

  const dimLabel=dim.name; // e.g. "Industry", "Country / Geography", "Employee Size / Range"

  let html='<div class="drilldown-content"><div class="drilldown-columns">';
  html+=`<div class="drilldown-col"><div class="drilldown-col-header">Target Accounts <span class="drilldown-count">(${targetAccounts.length})</span></div>`;
  if(targetAccounts.length){
    html+='<ul class="drilldown-list">'+targetAccounts.map(a=>
      `<li><span class="drilldown-name">${esc(a.name)}</span><span class="drilldown-field">${esc(a.rawValue)}</span></li>`
    ).join('')+'</ul>';
  }else{
    html+='<div class="drilldown-empty">No target accounts</div>';
  }
  html+='</div>';
  html+=`<div class="drilldown-col"><div class="drilldown-col-header">Current Customers <span class="drilldown-count">(${currentAccounts.length})</span></div>`;
  if(currentAccounts.length){
    html+='<ul class="drilldown-list">'+currentAccounts.map(a=>
      `<li><span class="drilldown-name">${esc(a.name)}</span><span class="drilldown-field">${esc(a.rawValue)}</span></li>`
    ).join('')+'</ul>';
  }else{
    html+='<div class="drilldown-empty">No current customers</div>';
  }
  html+='</div></div></div>';

  td.innerHTML=html;
  tr.appendChild(td);
  rowEl.after(tr);
  rowEl.classList.add('drilldown-open');
}

function dcRenderSortableTable(dimKey,comparison,targetTotal,currentTotal,sortCol,sortDir){
  const panel=document.getElementById('panel-'+dimKey);
  if(!panel)return;

  // Sort rows (excluding total)
  let rows=[...comparison];
  if(sortCol&&sortDir){
    const colDef=DC_SORT_COLS.find(c=>c.key===sortCol);
    if(colDef){
      const getter=colDef.getter;
      const mult=sortDir==='asc'?1:-1;
      rows.sort((a,b)=>{
        const va=getter(a),vb=getter(b);
        if(colDef.numeric)return(va-vb)*mult;
        return String(va).localeCompare(String(vb))*mult;
      });
    }
  }

  // Build header
  let html='<div class="table-wrap"><table><thead><tr>';
  for(const col of DC_SORT_COLS){
    const isNum=col.numeric;
    const isSorted=sortCol===col.key;
    const dirCls=isSorted?(sortDir==='asc'?'sort-asc':'sort-desc'):'';
    html+=`<th class="${isNum?'num ':' '}sortable ${dirCls}" onclick="dcSortTable('${dimKey}','${col.key}')">${col.label}</th>`;
  }
  html+='</tr></thead><tbody>';

  // Body rows
  for(const row of rows){
    const align=getAlignment(row.delta);
    const deltaAbs=Math.abs(row.delta);
    const barWidth=Math.min(deltaAbs/0.25,1)*100;
    const barColor=align.cls==='align-good'?'var(--text-success)':align.cls==='align-warn'?'var(--text-warning)':'var(--text-error)';
    html+=`<tr class="category-row" data-dim="${dimKey}" data-category="${esc(row.targetLabel)}" onclick="dcToggleDrilldown(this.dataset.dim,this.dataset.category,this)" style="cursor:pointer">`;
    html+=`<td class="category-cell">${esc(row.targetLabel)} <span class="drilldown-arrow">&#9654;</span></td>`;
    html+=`<td class="num">${row.targetCount.toLocaleString()}</td>`;
    html+=`<td class="num">${(row.targetPct*100).toFixed(1)}%</td>`;
    html+=`<td class="num">${row.currentCount.toLocaleString()}</td>`;
    html+=`<td class="num">${(row.currentPct*100).toFixed(1)}%</td>`;
    html+=`<td><div class="delta-bar">
      <span class="delta-val" style="color:${barColor}">${row.delta>=0?'+':''}${(row.delta*100).toFixed(1)}%</span>
      <div class="bar-track"><div class="bar-fill" style="width:${barWidth}%;background:${barColor};left:0"></div></div>
    </div></td>`;
    html+=`<td><span class="align-tag ${align.cls}">${align.label}</span></td>`;
    html+='</tr>';
  }

  // Total row (always last, not sortable)
  html+=`<tr class="total-row"><td>Grand Total</td>`;
  html+=`<td class="num">${targetTotal.toLocaleString()}</td><td class="num">100%</td>`;
  html+=`<td class="num">${currentTotal.toLocaleString()}</td><td class="num">100%</td>`;
  html+='<td></td><td></td></tr></tbody></table></div>';

  panel.innerHTML=html;
}

/* ========== EXPORT RESULTS ========== */
function dcExportResults(){
  const wb=XLSX.utils.book_new();
  const dims=DC._resultDimensions;
  if(!dims||!dims.length){alert('No results to export.');return;}

  // Sheet 1: Config
  const configRows=[
    {Setting:'Export Date',Value:new Date().toLocaleString()},
    {Setting:'Target Accounts',Value:DC.targetRaw.length},
    {Setting:'Current Customers',Value:DC.currentRaw.length},
    {Setting:'',Value:''},
    {Setting:'--- Dimensions ---',Value:''},
  ];
  dims.forEach(d=>{configRows.push({Setting:d.name,Value:`Target col: ${d.targetCol}, Current col: ${d.currentCol}`})});

  // Merge decisions
  if(DC._mergeState){
    configRows.push({Setting:'',Value:''},{Setting:'--- Merge Decisions ---',Value:''});
    for(const dimKey of Object.keys(DC._mergeState)){
      const ms=DC._mergeState[dimKey];
      const dimLabel=dims.find(d=>d.key===dimKey)?.name||dimKey;
      for(const[src,tgt]of Object.entries(ms)){
        if(tgt)configRows.push({Setting:`${dimLabel}: "${src}"`,Value:`merged into "${tgt}"`});
      }
    }
  }
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(configRows),'Config');

  // Sheet per dimension: comparison table
  for(const dim of dims){
    const rows=dim.comparison.map(r=>({
      Category:r.targetLabel,
      'Target Count':r.targetCount,
      'Target %':(r.targetPct*100).toFixed(1)+'%',
      'Current Count':r.currentCount,
      'Current %':(r.currentPct*100).toFixed(1)+'%',
      'Delta (pp)':(r.delta*100).toFixed(1)+'%',
      Alignment:Math.abs(r.delta)<=0.03?'Aligned':Math.abs(r.delta)<=0.08?'Slight gap':'Significant gap',
    }));
    // Add totals
    rows.push({
      Category:'Grand Total',
      'Target Count':dim.target.total,
      'Target %':'100%',
      'Current Count':dim.current.total,
      'Current %':'100%',
      'Delta (pp)':'',
      Alignment:'',
    });
    const name=dim.name.length>31?dim.name.slice(0,31):dim.name;
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),name);
  }

  // Sheet: Value Mappings (what raw values mapped to what final labels)
  const mapRows=[];
  for(const dim of dims){
    if(!dim.valueMap)continue;
    for(const[raw,final]of Object.entries(dim.valueMap)){
      if(raw!==final)mapRows.push({Dimension:dim.name,'Original Value':raw,'Mapped To':final});
    }
  }
  if(mapRows.length){
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(mapRows),'Value Mappings');
  }

  XLSX.writeFile(wb,'Distribution_Comparison_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
