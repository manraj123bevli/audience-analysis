/* ========== COLUMN MAPPING ========== */

const CONTACT_FIELDS=[
  {key:'company',label:'Company / Account Name',required:true,
   hints:['company name','company','account name','account','organization','org name','organisation','employer','company_name','account_name']},
  {key:'title',label:'Job Title',required:true,
   hints:['job title','title','job_title','jobtitle','position','designation','role title']},
  {key:'firstName',label:'First Name',required:false,
   hints:['first name','firstname','first_name','given name','fname']},
  {key:'lastName',label:'Last Name',required:false,
   hints:['last name','lastname','last_name','surname','family name','lname']},
  {key:'fullName',label:'Full Name',required:false,
   hints:['full name','fullname','name','contact name','full_name','contact_name']},
  {key:'email',label:'Email',required:false,
   hints:['email','e-mail','email address','email_address','work email','primary email']},
  {key:'role',label:'Employment Role / Function',required:false,
   hints:['employment role','role','function','department','job function','job_function','dept']},
  {key:'seniority',label:'Employment Seniority / Level',required:false,
   hints:['employment seniority','seniority','level','job level','seniority level','management level','career level']},
  {key:'country',label:'Country / Region',required:false,
   hints:['country','country/region','country_region','region','geography','geo','location country','nation']},
  {key:'city',label:'City',required:false,
   hints:['city','town','metro','location city']},
  {key:'state',label:'State / Region',required:false,
   hints:['state','state/region','state_region','province','territory']},
  {key:'industry',label:'Industry',required:false,
   hints:['industry','sector','vertical','industry group','industry_group']},
  {key:'leadStatus',label:'Lead Status',required:false,
   hints:['lead status','lead_status','status','contact status','prospect status']},
  {key:'lifecycle',label:'Lifecycle Stage',required:false,
   hints:['lifecycle stage','lifecycle_stage','lifecycle','stage','funnel stage']},
  {key:'owner',label:'Contact Owner',required:false,
   hints:['contact owner','owner','assigned to','rep','sales rep','account executive','contact_owner']},
];

const ACCOUNT_FIELDS=[
  {key:'company',label:'Company / Account Name',required:true,
   hints:['company name','company','account name','account','organization','org name','organisation','name','company_name','account_name']},
  {key:'industry',label:'Industry',required:false,
   hints:['industry','industry group','sector','vertical','industry_group']},
  {key:'country',label:'Country / Region',required:false,
   hints:['country','country/region','country_region','region','geography','geo','location country']},
  {key:'employees',label:'Employee Count / Range',required:false,
   hints:['employee','employees','employee range','employee_range','number of employees','headcount','company size','size','revised: employee range','employee count','num employees']},
  {key:'revenue',label:'Revenue',required:false,
   hints:['revenue','annual revenue','annual_revenue','arr','total revenue','revenue range']},
];

function autoDetect(cols,fields,samples){
  const mapping={};
  const used=new Set();
  fields.forEach(field=>{
    let bestCol=null,bestScore=0;
    cols.forEach(col=>{
      if(used.has(col))return;
      const cn=col.toLowerCase().trim();
      for(const hint of field.hints){
        if(cn===hint){if(10>bestScore){bestScore=10;bestCol=col}return}
        if(cn.includes(hint)){if(8>bestScore){bestScore=8;bestCol=col}return}
        if(hint.includes(cn)&&cn.length>2){if(5>bestScore){bestScore=5;bestCol=col}return}
        const cWords=new Set(cn.split(/[\s_\-\/]+/));
        const hWords=hint.split(/[\s_\-\/]+/);
        const overlap=hWords.filter(w=>cWords.has(w)&&w.length>2).length;
        if(overlap>=2&&overlap/hWords.length>0.5){if(4>bestScore){bestScore=4;bestCol=col}}
        else if(overlap>=1&&cn.length<20){if(2>bestScore){bestScore=2;bestCol=col}}
      }
    });
    mapping[field.key]={col:bestCol,score:bestScore,auto:bestScore>=4};
    if(bestCol&&bestScore>=4)used.add(bestCol);
  });
  return mapping;
}

function renderMappingCard(containerId,title,fields,cols,mapping,samples,prefix){
  const el=document.getElementById(containerId);
  let html=`<h3>${title}</h3>`;
  fields.forEach(field=>{
    const m=mapping[field.key];
    const selId=prefix+'_'+field.key;
    html+=`<div class="map-row">
      <span class="map-label">${field.label} ${field.required?'<span class="map-required">*</span>':'<span class="map-optional">(optional)</span>'}</span>
      <select class="map-select" id="${selId}" onchange="updatePreview('${selId}','${prefix}')">
        <option value="">— not mapped —</option>
        ${cols.map(c=>`<option value="${esc(c)}" ${m.col===c?'selected':''}>${esc(c)}</option>`).join('')}
      </select>
      ${m.col&&m.auto?'<span class="map-badge map-auto">auto</span>':m.col?'<span class="map-badge map-manual">suggest</span>':''}
    </div>
    <div class="preview-row" id="prev_${selId}">${m.col&&samples[m.col]?'e.g. '+samples[m.col].join(' | '):''}</div>`;
  });
  el.innerHTML=html;
}

function updatePreview(selId,prefix){
  const sel=document.getElementById(selId);
  const col=sel.value;
  const samples=prefix==='c'?CA.contactSamples:CA.accountSamples;
  const prev=document.getElementById('prev_'+selId);
  prev.textContent=col&&samples[col]?'e.g. '+samples[col].join(' | '):'';
}

function getMap(prefix,key){const el=document.getElementById(prefix+'_'+key);return el?el.value:''}
function getVal(row,col){return col?String(row[col]||'').trim():''}
