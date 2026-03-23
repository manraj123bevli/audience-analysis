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

  // Pre-compute reconciliations and render all panels
  for(const dim of dims){
    const targetDist=buildDistribution(DC.targetRaw,dim.targetCol,dim.label,dim.key);
    const currentDist=buildDistribution(DC.currentRaw,dim.currentCol,dim.label,dim.key);
    const recon=dcBuildReconciliation(targetDist.data,currentDist.data,dim.key);
    DC.reconciliations[dim.key]=recon;
    DC._reconPanelsHtml[dim.key]=dcRenderReconPanel(dim,recon);
  }

  // Build dimension tabs
  const tabsEl=document.getElementById('dc-recon-tabs');
  tabsEl.innerHTML=dims.map((dim,i)=>{
    const unmatchedCount=DC.reconciliations[dim.key].mappings.filter(m=>m.confidence==='manual').length;
    const badge=unmatchedCount?` <span class="badge b-red" style="margin-left:4px">${unmatchedCount}</span>`:'';
    return `<button class="tab-btn${i===0?' active':''}" onclick="dcShowReconDim('${dim.key}',this)">${esc(dim.label)}${badge}</button>`;
  }).join('');

  // Render all panels (hidden), show first
  const container=document.getElementById('dc-reconciliation-panels');
  container.innerHTML=dims.map((dim,i)=>
    `<div class="dc-recon-dim${i===0?' active':''}" id="dc-recon-dim-${dim.key}">${DC._reconPanelsHtml[dim.key]}</div>`
  ).join('');

  dcGoToStep(3);
}

function dcShowReconDim(dimKey,btn){
  document.querySelectorAll('#dc-recon-tabs .tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.dc-recon-dim').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('dc-recon-dim-'+dimKey).classList.add('active');
}

function dcRenderReconPanel(dim,recon){
  const groups=new Map();
  const needsReview=[];  // items needing user action

  for(const m of recon.mappings){
    if(m.confidence==='manual'){
      needsReview.push(m);
    }else{
      const key=m.masterLabel;
      if(!groups.has(key))groups.set(key,[]);
      groups.get(key).push(m);
    }
  }

  // Also find groups that look similar to each other (potential merges)
  const groupLabels=[...groups.keys()];
  const similarPairs=[];
  for(let i=0;i<groupLabels.length;i++){
    for(let j=i+1;j<groupLabels.length;j++){
      const score=similarityScore(groupLabels[i],groupLabels[j]);
      if(score>=0.25)similarPairs.push({a:groupLabels[i],b:groupLabels[j],score});
    }
  }
  similarPairs.sort((a,b)=>b.score-a.score);

  // Sort groups by count
  const sortedGroups=[...groups.entries()].sort((a,b)=>{
    const countA=a[1].reduce((s,m)=>s+m.count,0);
    const countB=b[1].reduce((s,m)=>s+m.count,0);
    return countB-countA;
  });

  let html='';

  // ===== SECTION 1: NEEDS REVIEW (prominent) =====
  // Combine unmatched items + similar pair suggestions
  const reviewItems=[...needsReview.map(m=>{
    let bestGroup='',bestScore=0;
    for(const gl of groupLabels){
      const sc=similarityScore(m.value,gl);
      if(sc>bestScore){bestScore=sc;bestGroup=gl;}
    }
    const bestCount=bestGroup?groups.get(bestGroup)?.reduce((s,x)=>s+x.count,0)||0:0;
    return{type:'unmatched',value:m.value,source:m.source,count:m.count,bestGroup,bestScore,bestCount};
  }),...similarPairs.map(p=>{
    const countA=groups.get(p.a)?.reduce((s,x)=>s+x.count,0)||0;
    const countB=groups.get(p.b)?.reduce((s,x)=>s+x.count,0)||0;
    return{type:'similar',a:p.a,b:p.b,score:p.score,countA,countB};
  })];

  if(reviewItems.length){
    for(const item of reviewItems){
      if(item.type==='unmatched'){
        const selId=`dc-recon-${dim.key}-${normalizeForMatch(item.value).replace(/\s/g,'_')}`;
        html+=`<div class="review-card">
          <span class="review-card-flag flag-unmatched">Needs assignment</span>
          <div class="review-card-value">${esc(item.value)}</div>
          <div class="review-card-meta">${item.count.toLocaleString()} accounts from ${item.source==='target'?'Target':'Current'} file</div>
          ${item.bestGroup&&item.bestScore>=0.2
            ?`<div class="review-card-suggestion">Looks similar to <strong>${esc(item.bestGroup)}</strong> (${item.bestCount.toLocaleString()} accounts)</div>`
            :''}
          <div class="review-card-actions">
            <select id="${selId}">
              <option value="${esc(item.value)}">Keep as separate category</option>
              ${groupLabels.map(gl=>`<option value="${esc(gl)}"${gl===item.bestGroup&&item.bestScore>=0.2?' selected':''}>${esc(gl)} (${(groups.get(gl)?.reduce((s,x)=>s+x.count,0)||0).toLocaleString()})</option>`).join('')}
            </select>
          </div>
        </div>`;
      }else{
        // Similar pair — suggest merging two existing groups
        const mergeId=`dc-simpair-${dim.key}-${normalizeForMatch(item.a).replace(/\s/g,'_')}`;
        // Get source breakdown for each group
        const membersA=groups.get(item.a)||[];
        const membersB=groups.get(item.b)||[];
        function sourceBreakdown(members,label){
          return members.map(m=>`<div class="matched-row-source"><span class="src-tag ${m.source==='target'?'src-target':'src-current'}">${m.source==='target'?'Target':'Current'}</span><span class="src-val">${esc(m.value)}</span><span class="src-count">${m.count.toLocaleString()}</span></div>`).join('');
        }
        html+=`<div class="review-card">
          <span class="review-card-flag flag-similar">Possible duplicate</span>
          <div class="review-card-value">"${esc(item.a)}" and "${esc(item.b)}"</div>
          <div class="review-card-meta">These look similar. Should they be combined?</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:var(--space-400);margin-bottom:var(--space-400)">
            <div style="background:var(--bg-subtle);border-radius:var(--radius-md);padding:var(--space-300)">
              <div style="font-weight:600;font-size:13px;margin-bottom:var(--space-200)">${esc(item.a)} <span style="color:var(--text-subtle);font-weight:400">(${item.countA.toLocaleString()})</span></div>
              ${sourceBreakdown(membersA)}
            </div>
            <div style="background:var(--bg-subtle);border-radius:var(--radius-md);padding:var(--space-300)">
              <div style="font-weight:600;font-size:13px;margin-bottom:var(--space-200)">${esc(item.b)} <span style="color:var(--text-subtle);font-weight:400">(${item.countB.toLocaleString()})</span></div>
              ${sourceBreakdown(membersB)}
            </div>
          </div>
          <div class="review-card-actions">
            <select id="${mergeId}" data-merge-a="${esc(item.a)}" data-merge-b="${esc(item.b)}">
              <option value="">Keep as separate groups</option>
              <option value="a">Merge into "${esc(item.a)}"</option>
              <option value="b">Merge into "${esc(item.b)}"</option>
            </select>
          </div>
        </div>`;
      }
    }
  }else{
    html+=`<div style="background:var(--bg-success-subtle);border-radius:var(--radius-md);padding:var(--space-400) var(--space-600);margin-bottom:var(--space-400);font-size:14px;color:var(--text-success);font-weight:500">All values matched automatically. No action needed.</div>`;
  }

  // ===== SECTION 2: MATCHED GROUPS (collapsed) =====
  const matchedCount=sortedGroups.length;
  const totalMatched=sortedGroups.reduce((s,[,members])=>s+members.reduce((s2,m)=>s2+m.count,0),0);
  const toggleId=`dc-matched-toggle-${dim.key}`;
  const detailsId=`dc-matched-details-${dim.key}`;

  html+=`<div class="matched-toggle" id="${toggleId}" onclick="dcToggleMatched('${dim.key}')">
    <div class="matched-toggle-left">
      <span class="matched-toggle-label">${matchedCount} matched groups</span>
      <span class="matched-toggle-count">${totalMatched.toLocaleString()} total accounts</span>
    </div>
    <span class="matched-toggle-arrow">&#9660;</span>
  </div>
  <div class="matched-details" id="${detailsId}">`;

  for(let gi=0;gi<sortedGroups.length;gi++){
    const[masterLabel,members]=sortedGroups[gi];
    const totalCount=members.reduce((s,m)=>s+m.count,0);
    const groupId=`dc-grp-${dim.key}-${gi}`;
    const targetVals=members.filter(m=>m.source==='target').map(m=>m.value);
    const currentVals=members.filter(m=>m.source==='current').map(m=>m.value);
    const valSummary=[...targetVals,...currentVals.filter(v=>!targetVals.includes(v))].join(', ');

    const targetMembers=members.filter(m=>m.source==='target');
    const currentMembers=members.filter(m=>m.source==='current');
    const targetTotal=targetMembers.reduce((s,m)=>s+m.count,0);
    const currentTotal2=currentMembers.reduce((s,m)=>s+m.count,0);

    html+=`<div class="matched-row">
      <div class="matched-row-header">
        <input type="checkbox" class="matched-row-check" data-dim="${dim.key}" data-group="${esc(masterLabel)}" onchange="dcUpdateMergeBar('${dim.key}')">
        <input type="text" class="matched-row-name" id="${groupId}" value="${esc(masterLabel)}" data-original="${esc(masterLabel)}">
        <span class="matched-row-total">${totalCount.toLocaleString()} total</span>
      </div>
      <div class="matched-row-sources">
        ${targetMembers.map(m=>`<div class="matched-row-source"><span class="src-tag src-target">Target</span><span class="src-val">${esc(m.value)}</span><span class="src-count">${m.count.toLocaleString()}</span></div>`).join('')}
        ${currentMembers.map(m=>`<div class="matched-row-source"><span class="src-tag src-current">Current</span><span class="src-val">${esc(m.value)}${m.confidence!=='exact'?` <span class="recon-member-badge badge-${m.confidence==='auto'?'auto':'suggest'}">${m.confidence}</span>`:''}</span><span class="src-count">${m.count.toLocaleString()}</span></div>`).join('')}
        ${!currentMembers.length?'<div class="matched-row-source"><span class="src-tag src-current">Current</span><span class="src-val" style="color:var(--text-disabled)">—</span><span class="src-count">0</span></div>':''}
        ${!targetMembers.length?'<div class="matched-row-source"><span class="src-tag src-target">Target</span><span class="src-val" style="color:var(--text-disabled)">—</span><span class="src-count">0</span></div>':''}
      </div>
    </div>`;
  }

  html+=`<div class="merge-bar" id="dc-merge-bar-${dim.key}">
    <span class="merge-bar-text">Merge selected groups:</span>
    <button class="btn btn-primary" onclick="dcMergeSelected('${dim.key}')">Merge</button>
  </div>`;

  html+=`</div>`;
  return html;
}

function dcToggleMatched(dimKey){
  const toggle=document.getElementById('dc-matched-toggle-'+dimKey);
  const details=document.getElementById('dc-matched-details-'+dimKey);
  toggle.classList.toggle('open');
  details.classList.toggle('open');
}

function dcUpdateMergeBar(dimKey){
  const checks=document.querySelectorAll(`.matched-row-check[data-dim="${dimKey}"]:checked`);
  const bar=document.getElementById('dc-merge-bar-'+dimKey);
  bar.classList.toggle('visible',checks.length>=2);
}

function dcMergeSelected(dimKey){
  const checks=[...document.querySelectorAll(`.matched-row-check[data-dim="${dimKey}"]:checked`)];
  if(checks.length<2)return;
  const groupNames=checks.map(c=>c.dataset.group);
  // Keep the first selected group's name, merge others into it
  const keepName=groupNames[0];
  const mergeInto=groupNames.slice(1);

  // Store merge decisions
  if(!DC._manualMerges)DC._manualMerges={};
  if(!DC._manualMerges[dimKey])DC._manualMerges[dimKey]={};
  for(const name of mergeInto){
    DC._manualMerges[dimKey][name]=keepName;
  }

  // Re-render this dimension
  const recon=DC.reconciliations[dimKey];
  // Apply merges to the reconciliation data
  for(const m of recon.mappings){
    if(mergeInto.includes(m.masterLabel)){
      m.masterLabel=keepName;
    }
  }
  // Update master labels
  recon.masterLabels=recon.masterLabels.filter(ml=>!mergeInto.includes(ml));

  // Re-render
  const dim=DC.pendingDims.find(d=>d.key===dimKey);
  const container=document.getElementById('dc-recon-dim-'+dimKey);
  container.innerHTML=dcRenderReconPanel(dim,recon);
}

/* ========== STEP 3→4: CONFIRM AND RUN ========== */
function dcConfirmAndRun(){
  const dims=DC.pendingDims;

  const dimensions=[];
  for(const dim of dims){
    const valueMap={};  // original value → final display label
    const recon=DC.reconciliations[dim.key];

    if(recon){
      // 1. Build rename map: original master label → user-edited name
      const renameMap={};
      const groupInputs=document.querySelectorAll(`[id^="dc-grp-${dim.key}-"]`);
      groupInputs.forEach(input=>{
        const original=input.dataset.original;
        const renamed=input.value.trim();
        if(original&&renamed)renameMap[original]=renamed;
      });

      // 2. Build merge map from similar-pair dropdowns
      const mergeMap={};
      const simPairSelects=document.querySelectorAll(`[id^="dc-simpair-${dim.key}-"]`);
      simPairSelects.forEach(sel=>{
        const choice=sel.value;
        const a=sel.dataset.mergeA;
        const b=sel.dataset.mergeB;
        if(choice==='a')mergeMap[b]=a;
        else if(choice==='b')mergeMap[a]=b;
      });
      // Add manual merges from checkbox+merge button
      if(DC._manualMerges&&DC._manualMerges[dim.key]){
        Object.assign(mergeMap,DC._manualMerges[dim.key]);
      }

      // 3. Resolve merge chains: if A merges into B and B merges into C, A should end up in C
      function resolveTarget(label){
        const seen=new Set();
        let current=label;
        while(mergeMap[current]&&!seen.has(current)){
          seen.add(current);
          current=mergeMap[current];
        }
        return current;
      }

      for(const m of recon.mappings){
        if(m.confidence==='manual'){
          // Unmatched — read from dropdown
          const selId=`dc-recon-${dim.key}-${normalizeForMatch(m.value).replace(/\s/g,'_')}`;
          const sel=document.getElementById(selId);
          if(sel){
            const selected=sel.value;
            const resolved=resolveTarget(selected);
            valueMap[m.value]=renameMap[resolved]||resolved;
          }else{
            valueMap[m.value]=m.value;
          }
        }else{
          // Matched value — apply merge first, then rename
          const resolved=resolveTarget(m.masterLabel);
          const finalLabel=renameMap[resolved]||resolved;
          valueMap[m.value]=finalLabel;
        }
      }
    }

    const targetDist=buildDistributionWithMap(DC.targetRaw,dim.targetCol,dim.label,valueMap,dim.key);
    const currentDist=buildDistributionWithMap(DC.currentRaw,dim.currentCol,dim.label,valueMap,dim.key);

    // Compare with no mapFn — values are already reconciled via valueMap
    const comp=compareSections(targetDist,currentDist,null);
    dimensions.push({key:dim.key,name:dim.label,comparison:comp,target:targetDist,current:currentDist,showMapped:false});
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

/* ========== RENDER RESULTS (Step 4) ========== */

// Store results data for sorting
DC._resultDimensions=[];

function dcRenderResults(dimensions){
  DC._resultDimensions=dimensions;
  const tTotal=DC.targetRaw.length;
  const cTotal=DC.currentRaw.length;

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
    html+='<tr>';
    html+=`<td>${esc(row.targetLabel)}</td>`;
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
