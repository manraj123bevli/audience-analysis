/* ========== CONTACT & ACCOUNT ANALYSIS ========== */
window.CA={contactsRaw:[],accountsRaw:[],contactCols:[],accountCols:[],contactSamples:{},accountSamples:{},contacts:[]};

function caRunAnalysis(){
  const CM={};CONTACT_FIELDS.forEach(f=>{CM[f.key]=getMap('c',f.key)});
  const AM={};ACCOUNT_FIELDS.forEach(f=>{AM[f.key]=getMap('a',f.key)});

  if(!CM.company){alert('Please map the Company/Account Name column for contacts.');return}
  if(!AM.company){alert('Please map the Company/Account Name column for accounts.');return}
  if(!CM.title){alert('Please map the Job Title column for contacts.');return}

  const acctMap=new Map();
  CA.accountsRaw.forEach(a=>{
    const name=getVal(a,AM.company).toLowerCase();
    if(name)acctMap.set(name,a);
  });
  const totalAccounts=acctMap.size;

  const contacts=CA.contactsRaw.map(c=>{
    const company=getVal(c,CM.company);
    const normCo=company.toLowerCase();
    const matched=acctMap.has(normCo);
    const title=getVal(c,CM.title);
    const role=getVal(c,CM.role);
    const senField=getVal(c,CM.seniority);
    const fr=classifyFunc(title,role);
    const sr=classifySen(title,senField);
    let name='';
    if(CM.fullName)name=getVal(c,CM.fullName);
    if(!name&&(CM.firstName||CM.lastName))name=(getVal(c,CM.firstName)+' '+getVal(c,CM.lastName)).trim();
    return{
      name,title,email:getVal(c,CM.email),company,normCo,matched,
      country:getVal(c,CM.country),city:getVal(c,CM.city),state:getVal(c,CM.state),
      industry:getVal(c,CM.industry),leadStatus:getVal(c,CM.leadStatus),
      lifecycle:getVal(c,CM.lifecycle),owner:getVal(c,CM.owner),
      func:fr.fn,funcSource:fr.source,funcMethod:fr.method,
      seniority:sr.tier,senSource:sr.source,senMethod:sr.method,
      existingRole:role,existingSeniority:senField,
    };
  });
  CA.contacts=contacts;

  const total=contacts.length;
  const matched=contacts.filter(c=>c.matched);
  const unmatched=contacts.filter(c=>!c.matched);
  const acctsWithContacts=new Set();
  matched.forEach(c=>acctsWithContacts.add(c.normCo));
  const covered=acctsWithContacts.size;
  const uncovered=totalAccounts-covered;
  const cPerAcct=new Map();
  matched.forEach(c=>{cPerAcct.set(c.normCo,(cPerAcct.get(c.normCo)||0)+1)});
  const uniqueCos=new Set(contacts.map(c=>c.normCo).filter(Boolean));

  const funcDist=freq(contacts.map(c=>c.func));
  const senDist=freq(contacts.map(c=>c.seniority));
  const funcNames=funcDist.map(d=>d[0]);
  const senTiers=['Exec Buyers','Decision Influencers','Practitioners','Unknown'];
  const leadDist=freq(contacts.map(c=>c.leadStatus));

  const funcByCRM=contacts.filter(c=>c.funcMethod==='crm').length;
  const funcByTitle=contacts.filter(c=>c.funcMethod==='title').length;
  const senByCRM=contacts.filter(c=>c.senMethod==='crm').length;
  const senByTitle=contacts.filter(c=>c.senMethod==='title').length;

  const cross={};funcNames.forEach(fn=>{cross[fn]={};senTiers.forEach(s=>{cross[fn][s]=0})});
  contacts.forEach(c=>{if(cross[c.func])cross[c.func][c.seniority]=(cross[c.func][c.seniority]||0)+1});
  const fMatch={};funcNames.forEach(fn=>{fMatch[fn]={m:0,u:0}});
  contacts.forEach(c=>{if(fMatch[c.func]){c.matched?fMatch[c.func].m++:fMatch[c.func].u++}});
  const sMatch={};senDist.forEach(([s])=>{sMatch[s]={m:0,u:0}});
  contacts.forEach(c=>{if(sMatch[c.seniority]){c.matched?sMatch[c.seniority].m++:sMatch[c.seniority].u++}});

  const cpaVals=[...cPerAcct.values()];
  const buckets={'1 contact':0,'2-3':0,'4-5':0,'6-10':0,'11-20':0,'20+':0};
  cpaVals.forEach(v=>{if(v===1)buckets['1 contact']++;else if(v<=3)buckets['2-3']++;else if(v<=5)buckets['4-5']++;else if(v<=10)buckets['6-10']++;else if(v<=20)buckets['11-20']++;else buckets['20+']++});

  const singleContact=[...cPerAcct.values()].filter(v=>v===1).length;
  const avgCPA=covered?(matched.length/covered).toFixed(1):'0';
  const execInTarget=matched.filter(c=>c.seniority==='Exec Buyers').length;
  const inflInTarget=matched.filter(c=>c.seniority==='Decision Influencers').length;
  const practInTarget=matched.filter(c=>c.seniority==='Practitioners').length;

  const mappedContact=CONTACT_FIELDS.filter(f=>CM[f.key]).map(f=>`${f.label} → "${CM[f.key]}"`);
  const mappedAccount=ACCOUNT_FIELDS.filter(f=>AM[f.key]).map(f=>`${f.label} → "${AM[f.key]}"`);

  goToStep(3);

  // KPIs
  document.getElementById('ca-kpi-row').innerHTML=`
    <div class="kpi"><div class="label">Total Contacts</div><div class="val">${total.toLocaleString()}</div><div class="sub">${uniqueCos.size.toLocaleString()} unique companies</div></div>
    <div class="kpi"><div class="label">Target Accounts</div><div class="val">${totalAccounts.toLocaleString()}</div></div>
    <div class="kpi"><div class="label">Contacts In Target</div><div class="val" style="color:var(--text-success)">${matched.length.toLocaleString()}</div><div class="sub">${pct(matched.length,total)} of contacts</div></div>
    <div class="kpi"><div class="label">Contacts Outside</div><div class="val" style="color:var(--text-error)">${unmatched.length.toLocaleString()}</div><div class="sub">${pct(unmatched.length,total)} of contacts</div></div>
    <div class="kpi"><div class="label">Accounts Covered</div><div class="val" style="color:var(--text-success)">${covered.toLocaleString()}</div><div class="sub">${pct(covered,totalAccounts)} of target</div></div>
    <div class="kpi"><div class="label">Accounts w/o Contacts</div><div class="val" style="color:var(--text-warning)">${uncovered.toLocaleString()}</div><div class="sub">${pct(uncovered,totalAccounts)} of target</div></div>`;

  // Tabs
  const tabDefs=[
    {key:'ca-summary',label:'Summary'},{key:'ca-overview',label:'Overview'},
    {key:'ca-coverage',label:'Account Coverage'},{key:'ca-function',label:'Function Analysis'},
    {key:'ca-seniority',label:'Seniority Analysis'},{key:'ca-dimensions',label:'Dimensional Slicing'},
  ];
  document.getElementById('ca-tabs').innerHTML=tabDefs.map((t,i)=>`<button class="tab-btn${i===0?' active':''}" onclick="switchTab('${t.key}',this,'ca-results-area')">${t.label}</button>`).join('');

  const P=document.getElementById('ca-panels');
  P.innerHTML='';

  const senBCls={'Exec Buyers':'b-red','Decision Influencers':'b-ylw','Practitioners':'b-blue','Unknown':'b-cyn'};
  const senColors={'Exec Buyers':'var(--bg-error-strong)','Decision Influencers':'var(--text-warning)','Practitioners':'#005bbb','Unknown':'var(--text-disabled)'};

  /* SUMMARY */
  const coveragePct=pctR(covered,totalAccounts);
  const matchPct=pctR(matched.length,total);
  let s=`<div class="tab-panel active" id="panel-ca-summary">`;
  s+=`<div class="card"><h2>Executive Summary</h2>
    <p style="margin-bottom:var(--space-400);color:var(--text-base);font-size:14px">Analysis of <strong>${total.toLocaleString()}</strong> contacts against <strong>${totalAccounts.toLocaleString()}</strong> target accounts.</p>
    <div class="callout ${coveragePct>=.5?'callout-green':coveragePct>=.3?'callout-yellow':'callout-red'}">
      <span class="callout-title">Account Coverage: ${pct(covered,totalAccounts)}</span>
      <strong>${covered}</strong> of ${totalAccounts} target accounts have at least one contact. <strong>${uncovered}</strong> accounts have zero contacts.
      ${singleContact?` Of covered accounts, <strong>${singleContact}</strong> (${pct(singleContact,covered)}) have only 1 contact — single-threaded risk.`:''}
    </div>
    <div class="callout ${matchPct>=.6?'callout-green':matchPct>=.4?'callout-yellow':'callout-red'}">
      <span class="callout-title">Contact Match Rate: ${pct(matched.length,total)}</span>
      <strong>${matched.length.toLocaleString()}</strong> contacts belong to a target account. <strong>${unmatched.length.toLocaleString()}</strong> (${pct(unmatched.length,total)}) are outside the target list.
    </div>
    <div class="callout callout-blue">
      <span class="callout-title">Contact Depth: ${avgCPA} avg contacts per covered account</span>
      Multi-threading (3+ contacts) is ideal for enterprise deals.
    </div></div>`;
  s+=`<div class="three-col">`;
  s+=`<div class="card"><h2>Function Breakdown</h2><div class="table-wrap"><table><thead><tr><th>Function</th><th class="num">#</th><th class="num">%</th></tr></thead><tbody>`;
  funcDist.slice(0,8).forEach(([k,v])=>{s+=`<tr><td>${esc(k)}</td><td class="num">${v.toLocaleString()}</td><td class="num">${pct(v,total)}</td></tr>`});
  s+=`</tbody></table></div><div class="method-box"><h4>How classified</h4><p>${pct(funcByCRM,total)} from CRM field, ${pct(funcByTitle,total)} inferred from job title.</p></div></div>`;
  s+=`<div class="card"><h2>Seniority Breakdown</h2><div class="table-wrap"><table><thead><tr><th>Tier</th><th class="num">#</th><th class="num">%</th></tr></thead><tbody>`;
  senDist.forEach(([k,v])=>{s+=`<tr><td><span class="badge ${senBCls[k]||'b-cyn'}">${esc(k)}</span></td><td class="num">${v.toLocaleString()}</td><td class="num">${pct(v,total)}</td></tr>`});
  s+=`</tbody></table></div><div class="method-box"><h4>How classified</h4><p>${pct(senByCRM,total)} from CRM field, ${pct(senByTitle,total)} from title patterns.</p></div></div>`;
  s+=`<div class="card"><h2>In-Target Persona Mix</h2><p style="margin-bottom:var(--space-300)">${matched.length.toLocaleString()} contacts in target accounts:</p><div class="table-wrap"><table><thead><tr><th>Tier</th><th class="num">#</th><th class="num">%</th></tr></thead><tbody>
    <tr><td><span class="badge b-red">Exec Buyers</span></td><td class="num">${execInTarget}</td><td class="num">${pct(execInTarget,matched.length)}</td></tr>
    <tr><td><span class="badge b-ylw">Decision Influencers</span></td><td class="num">${inflInTarget}</td><td class="num">${pct(inflInTarget,matched.length)}</td></tr>
    <tr><td><span class="badge b-blue">Practitioners</span></td><td class="num">${practInTarget}</td><td class="num">${pct(practInTarget,matched.length)}</td></tr>
  </tbody></table></div><div class="method-box"><h4>Why this matters</h4><p>Exec buyers for budget, influencers for evaluation, practitioners for adoption.</p></div></div>`;
  s+=`</div>`;
  s+=`<div class="card"><h2>Key Observations & Risks</h2>
    <div class="callout callout-red"><span class="callout-title">Coverage Gap</span>${pct(uncovered,totalAccounts)} of target accounts (${uncovered.toLocaleString()}) have <strong>zero contacts</strong>.</div>
    <div class="callout callout-yellow"><span class="callout-title">Single-Threaded Accounts</span>${singleContact.toLocaleString()} covered accounts have only 1 contact. Aim for 3+.</div>
    <div class="callout callout-blue"><span class="callout-title">Outside-Target Contacts</span>${unmatched.length.toLocaleString()} contacts (${pct(unmatched.length,total)}) don't map to any target account.</div></div>`;
  s+=`<div class="card"><h2>Column Mappings Used</h2><div class="two-col">
    <div><h3>Contacts</h3><ul>${mappedContact.map(m=>`<li>${esc(m)}</li>`).join('')}</ul></div>
    <div><h3>Accounts</h3><ul>${mappedAccount.map(m=>`<li>${esc(m)}</li>`).join('')}</ul></div></div></div>`;
  s+=`</div>`;
  P.innerHTML+=s;

  /* OVERVIEW */
  const maxLead=leadDist[0]?leadDist[0][1]:1;
  let o=`<div class="tab-panel" id="panel-ca-overview"><div class="two-col">`;
  o+=`<div class="card"><h2>Contact-Account Match Summary</h2><div class="table-wrap"><table>
    <tr><td>Total unique contacts</td><td class="num">${total.toLocaleString()}</td></tr>
    <tr><td>Matched to target account</td><td class="num"><span class="badge b-grn">${matched.length.toLocaleString()} (${pct(matched.length,total)})</span></td></tr>
    <tr><td>Not in target</td><td class="num"><span class="badge b-red">${unmatched.length.toLocaleString()} (${pct(unmatched.length,total)})</span></td></tr>
    <tr><td>Target accounts with contacts</td><td class="num"><span class="badge b-grn">${covered.toLocaleString()} (${pct(covered,totalAccounts)})</span></td></tr>
    <tr><td>Target accounts with 0 contacts</td><td class="num"><span class="badge b-ylw">${uncovered.toLocaleString()} (${pct(uncovered,totalAccounts)})</span></td></tr>
  </table></div><div class="method-box"><h4>Matching Method</h4><p>Case-insensitive exact match on company name.</p></div></div>`;
  o+=`<div class="card"><h2>Lead Status Distribution</h2><div class="table-wrap scroll-table"><table><thead><tr><th>Status</th><th class="num">#</th><th class="num">%</th><th>Bar</th></tr></thead><tbody>`;
  leadDist.forEach(([k,v])=>{o+=`<tr><td>${esc(k)}</td><td class="num">${v.toLocaleString()}</td><td class="num">${pct(v,total)}</td><td>${makeBar(v,maxLead,'var(--bg-brand)')}</td></tr>`});
  o+=`</tbody></table></div></div></div></div>`;
  P.innerHTML+=o;

  /* ACCOUNT COVERAGE */
  const maxBkt=Math.max(...Object.values(buckets));
  let cv=`<div class="tab-panel" id="panel-ca-coverage"><div class="two-col">`;
  cv+=`<div class="card"><h2>Contacts per Account</h2><div class="table-wrap"><table><thead><tr><th>Bucket</th><th class="num"># Accounts</th><th class="num">%</th><th>Bar</th></tr></thead><tbody>`;
  Object.entries(buckets).forEach(([k,v])=>{cv+=`<tr><td>${k}</td><td class="num">${v}</td><td class="num">${pct(v,covered)}</td><td>${makeBar(v,maxBkt,'#005bbb')}</td></tr>`});
  cv+=`</tbody></table></div></div>`;
  const topA=[...cPerAcct.entries()].sort((a,b)=>b[1]-a[1]).slice(0,25);
  const maxTA=topA[0]?topA[0][1]:1;
  cv+=`<div class="card"><h2>Top 25 Accounts by Contacts</h2><div class="table-wrap scroll-table"><table><thead><tr><th>Account</th><th class="num">#</th><th>Bar</th></tr></thead><tbody>`;
  topA.forEach(([norm,cnt])=>{const a=acctMap.get(norm);cv+=`<tr><td>${esc(a?getVal(a,AM.company):norm)}</td><td class="num">${cnt}</td><td>${makeBar(cnt,maxTA,'var(--bg-success-strong)')}</td></tr>`});
  cv+=`</tbody></table></div></div></div>`;
  const uncovList=[];acctMap.forEach((a,n)=>{if(!acctsWithContacts.has(n))uncovList.push(a)});
  const uncByInd=freq(uncovList.map(a=>AM.industry?getVal(a,AM.industry):''));
  const maxUI=uncByInd[0]?uncByInd[0][1]:1;
  cv+=`<div class="card"><h2>Uncovered Accounts by Industry (${uncovList.length})</h2><div class="table-wrap"><table><thead><tr><th>Industry</th><th class="num">#</th><th class="num">%</th><th>Bar</th></tr></thead><tbody>`;
  uncByInd.forEach(([k,v])=>{cv+=`<tr><td>${esc(k)}</td><td class="num">${v}</td><td class="num">${pct(v,uncovList.length)}</td><td>${makeBar(v,maxUI,'var(--text-warning)')}</td></tr>`});
  cv+=`</tbody></table></div></div></div>`;
  P.innerHTML+=cv;

  /* FUNCTION ANALYSIS */
  const maxF=funcDist[0]?funcDist[0][1]:1;
  let fn=`<div class="tab-panel" id="panel-ca-function"><div class="card"><h2>Function Distribution</h2>
    <div class="table-wrap"><table><thead><tr><th>Function</th><th class="num">#</th><th class="num">%</th><th>Bar</th><th class="num">In Target</th><th class="num">Outside</th>`;
  senTiers.forEach(st=>{fn+=`<th class="num">${st}</th>`});
  fn+=`</tr></thead><tbody>`;
  funcDist.forEach(([f,cnt])=>{
    fn+=`<tr><td><strong>${esc(f)}</strong></td><td class="num">${cnt.toLocaleString()}</td><td class="num">${pct(cnt,total)}</td><td>${makeBar(cnt,maxF,'var(--bg-brand)')}</td><td class="num">${fMatch[f].m}</td><td class="num">${fMatch[f].u}</td>`;
    senTiers.forEach(st=>{fn+=`<td class="num">${cross[f][st]||0}</td>`});fn+=`</tr>`;
  });
  fn+=`<tr class="total-row"><td>Total</td><td class="num">${total.toLocaleString()}</td><td class="num">100%</td><td></td><td class="num">${matched.length}</td><td class="num">${unmatched.length}</td>`;
  senTiers.forEach(st=>{fn+=`<td class="num">${contacts.filter(c=>c.seniority===st).length}</td>`});
  fn+=`</tr></tbody></table></div>`;
  fn+=`<div class="method-box"><h4>Classification Methodology</h4>
    <p><strong>Step 1 — CRM field (${pct(funcByCRM,total)}):</strong> If "${CM.role||'Employment Role'}" is populated, map directly.</p>
    <p style="margin-top:var(--space-100)"><strong>Step 2 — Title keywords (${pct(funcByTitle,total)}):</strong> Scan "${CM.title}" against ordered keyword lists.</p>
    <p style="margin-top:var(--space-100)"><strong>Unclassified:</strong> ${contacts.filter(c=>c.func==='Unknown').length} with no role or title.</p></div></div>`;
  fn+=`<div class="card"><h2>Sample Classifications — Title-Inferred (50)</h2><div class="table-wrap scroll-table"><table><thead><tr><th>Name</th><th>Job Title</th><th>Function</th><th>Reasoning</th></tr></thead><tbody>`;
  contacts.filter(c=>c.funcMethod==='title'&&c.title).slice(0,50).forEach(c=>{
    fn+=`<tr><td>${esc(c.name)}</td><td>${esc(c.title)}</td><td><span class="badge b-pur">${esc(c.func)}</span></td><td style="font-size:12px;color:var(--text-subtle)">${esc(c.funcSource)}</td></tr>`;
  });
  fn+=`</tbody></table></div></div></div>`;
  P.innerHTML+=fn;

  /* SENIORITY ANALYSIS */
  const maxSn=senDist[0]?senDist[0][1]:1;
  let sn=`<div class="tab-panel" id="panel-ca-seniority"><div class="card"><h2>Seniority Distribution</h2>
    <div class="table-wrap"><table><thead><tr><th>Tier</th><th class="num">#</th><th class="num">%</th><th>Bar</th><th class="num">In Target</th><th class="num">Outside</th></tr></thead><tbody>`;
  senDist.forEach(([st,cnt])=>{
    sn+=`<tr><td><span class="badge ${senBCls[st]||'b-cyn'}">${esc(st)}</span></td><td class="num">${cnt.toLocaleString()}</td><td class="num">${pct(cnt,total)}</td><td>${makeBar(cnt,maxSn,senColors[st]||'var(--bg-brand)')}</td><td class="num">${sMatch[st].m}</td><td class="num">${sMatch[st].u}</td></tr>`;
  });
  sn+=`</tbody></table></div>`;
  sn+=`<div class="method-box"><h4>Classification Methodology</h4>
    <p><strong>Step 1 — CRM (${pct(senByCRM,total)}):</strong> Executive/VP/Owner/Partner → <span class="badge b-red">Exec Buyers</span>, Director/Senior/Manager → <span class="badge b-ylw">Decision Influencers</span>, Employee/Entry → <span class="badge b-blue">Practitioners</span>.</p>
    <p style="margin-top:var(--space-100)"><strong>Step 2 — Title patterns (${pct(senByTitle,total)}):</strong> Regex: exec → influencer → practitioner.</p></div></div>`;
  sn+=`<div class="card"><h2>Seniority x Function Matrix</h2><div class="table-wrap"><table><thead><tr><th>Function</th>`;
  senTiers.forEach(st=>{sn+=`<th class="num">${st}</th>`});
  sn+=`<th class="num">Total</th></tr></thead><tbody>`;
  funcDist.forEach(([f,cnt])=>{sn+=`<tr><td><strong>${esc(f)}</strong></td>`;senTiers.forEach(st=>{sn+=`<td class="num">${cross[f][st]||0}</td>`});sn+=`<td class="num">${cnt}</td></tr>`});
  sn+=`</tbody></table></div></div>`;
  sn+=`<div class="card"><h2>Sample Seniority Classifications (50)</h2><div class="table-wrap scroll-table"><table><thead><tr><th>Name</th><th>Job Title</th><th>Tier</th><th>Reasoning</th></tr></thead><tbody>`;
  contacts.filter(c=>c.senMethod==='title'&&c.title).slice(0,50).forEach(c=>{
    sn+=`<tr><td>${esc(c.name)}</td><td>${esc(c.title)}</td><td><span class="badge ${senBCls[c.seniority]||'b-cyn'}">${esc(c.seniority)}</span></td><td style="font-size:12px;color:var(--text-subtle)">${esc(c.senSource)}</td></tr>`;
  });
  sn+=`</tbody></table></div></div></div>`;
  P.innerHTML+=sn;

  /* DIMENSIONAL SLICING */
  let dm=`<div class="tab-panel" id="panel-ca-dimensions">
    <div class="card"><h2>Slice & Dice</h2><div class="filter-row">
      <label>Match:</label><select id="f-match"><option value="all">All</option><option value="matched">In Target</option><option value="unmatched">Outside</option></select>
      <label>Function:</label><select id="f-func"><option value="all">All</option>${funcNames.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join('')}</select>
      <label>Seniority:</label><select id="f-sen"><option value="all">All</option>${senTiers.map(st=>`<option value="${esc(st)}">${esc(st)}</option>`).join('')}</select>
      <label>Lead Status:</label><select id="f-lead"><option value="all">All</option>${leadDist.map(([k])=>`<option value="${esc(k)}">${esc(k)}</option>`).join('')}</select>
      <button class="filter-btn" onclick="caApplyFilters()">Apply</button>
    </div></div><div id="ca-dim-results"></div></div>`;
  P.innerHTML+=dm;
  caApplyFilters();
}

function caApplyFilters(){
  let d=[...CA.contacts];
  const g=id=>document.getElementById(id).value;
  if(g('f-match')==='matched')d=d.filter(c=>c.matched);
  if(g('f-match')==='unmatched')d=d.filter(c=>!c.matched);
  if(g('f-func')!=='all')d=d.filter(c=>c.func===g('f-func'));
  if(g('f-sen')!=='all')d=d.filter(c=>c.seniority===g('f-sen'));
  if(g('f-lead')!=='all')d=d.filter(c=>c.leadStatus===g('f-lead'));
  caRenderDim(d);
}

function caRenderDim(data){
  const el=document.getElementById('ca-dim-results');
  const n=data.length;
  function dt(title,dist,color){
    const mx=dist[0]?dist[0][1]:1;
    let h=`<div class="card"><h3>${title} <span style="font-weight:400;color:var(--text-subtle);font-size:12px">(${n.toLocaleString()})</span></h3><div class="table-wrap scroll-table"><table><thead><tr><th>Value</th><th class="num">#</th><th class="num">%</th><th>Bar</th></tr></thead><tbody>`;
    dist.forEach(([k,v])=>{h+=`<tr><td>${esc(k)}</td><td class="num">${v.toLocaleString()}</td><td class="num">${pct(v,n)}</td><td>${makeBar(v,mx,color)}</td></tr>`});
    return h+`</tbody></table></div></div>`;
  }
  let h=`<div style="background:var(--bg-container);border:1px solid var(--border-base);border-radius:var(--radius-md);padding:var(--space-300) var(--space-400);margin-bottom:var(--space-400);font-size:14px"><strong>${n.toLocaleString()}</strong> contacts match filters</div>`;
  h+=`<div class="two-col">`+dt('By Industry',freq(data.map(c=>c.industry)),'var(--bg-brand)')+dt('By Country',freq(data.map(c=>c.country)),'#005bbb')+`</div>`;
  h+=`<div class="two-col">`+dt('By Function',freq(data.map(c=>c.func)),'#5b3da0')+dt('By Seniority',freq(data.map(c=>c.seniority)),'var(--bg-error-strong)')+`</div>`;
  h+=`<div class="two-col">`+dt('By Lead Status',freq(data.map(c=>c.leadStatus)),'var(--bg-success-strong)')+dt('By City (Top 30)',freq(data.map(c=>c.city)).slice(0,30),'#006b7a')+`</div>`;
  h+=dt('By Company (Top 30)',freq(data.map(c=>c.company)).slice(0,30),'var(--text-warning)');
  el.innerHTML=h;
}
