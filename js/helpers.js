/* ========== SHARED HELPERS ========== */
function esc(s){const d=document.createElement('div');d.textContent=s||'';return d.innerHTML}
function pct(n,t){return t?((n/t)*100).toFixed(1)+'%':'0%'}
function pctR(n,t){return t?n/t:0}
function makeBar(v,mx,color){const w=mx?Math.min(v/mx*100,100):0;return`<div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${color}"></div></div></div>`}
function makeCumBar(cumPct,color){const w=Math.min(cumPct*100,100);return`<div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${w}%;background:${color}"></div></div><span style="font-size:11px;color:var(--text-subtle);min-width:42px;text-align:right">${(cumPct*100).toFixed(1)}%</span></div>`}
function freq(arr){const m=new Map();arr.forEach(v=>{const k=v||'(blank)';m.set(k,(m.get(k)||0)+1)});return[...m.entries()].sort((a,b)=>b[1]-a[1])}
function normCompany(s){if(!s)return'';return s.toLowerCase().trim()}

function switchTab(key,btn,containerId){
  const container=document.getElementById(containerId);
  container.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  container.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  container.querySelector('#panel-'+key).classList.add('active');
}
