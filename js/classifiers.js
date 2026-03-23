/* ========== FUNCTION CLASSIFICATION ========== */
const FUNC_RULES=[
  {fn:'Security',kw:['security','infosec','ciso','soc ','penetration','threat','vulnerability','appsec','devsecops','identity and access','iam ','zero trust']},
  {fn:'Engineering',kw:['engineer','developer','dev ','devops','sre ','site reliability','software','architect','frontend','backend','full stack','fullstack','full-stack','platform','infrastructure','programmer','tech lead','technical lead','cto','vp of engineering','vp engineering','head of engineering','director of engineering','r&d','research and development']},
  {fn:'Product',kw:['product','ux ','user experience','ui ','design','product manager','product owner','product lead','head of product','vp product','chief product']},
  {fn:'IT',kw:[' it ','information technology','sysadmin','system admin','systems admin','helpdesk','help desk','it manager','it director','cio','network admin','it operations','it infrastructure']},
  {fn:'Operations',kw:['operations','ops ','coo','chief operating','business operations','project manag','program manag']},
  {fn:'Data/Analytics',kw:['data scientist','data engineer','data analyst','analytics','machine learning','ml engineer','ai ','artificial intelligence','deep learning']},
  {fn:'Sales',kw:['sales','account executive','business development','bdr','sdr','revenue','cro','chief revenue']},
  {fn:'Marketing',kw:['marketing','cmo','growth','content','brand','communications','demand gen','digital marketing']},
  {fn:'Finance',kw:['finance','cfo','accounting','controller','treasurer','financial']},
  {fn:'Legal',kw:['legal','counsel','attorney','lawyer','compliance','privacy officer','data protection officer','dpo']},
  {fn:'HR',kw:['human resources','hr ','people','talent','recruiting','recruiter','chief people']},
];

const ROLE_MAP={
  'Engineering':'Engineering','Product':'Product','Operations':'Operations',
  'Information Technology':'IT','Sales':'Sales','Marketing':'Marketing',
  'Finance':'Finance','Legal':'Legal','Human Resources':'HR',
  'Research':'Data/Analytics','Design':'Product','Business Development':'Sales',
  'Consulting':'Operations','Support':'Operations','Customer Service':'Operations',
  'Education':'Other','Quality Assurance':'Engineering','Administrative':'Operations',
  'Entrepreneurship':'Executive/Founder','Project Management':'Operations'
};

function classifyFunc(title,role){
  if(role&&role.trim()){return{fn:ROLE_MAP[role.trim()]||'Other',source:'CRM field: "'+role.trim()+'"',method:'crm'}}
  if(!title)return{fn:'Unknown',source:'No title or role',method:'none'};
  const t=' '+title.toLowerCase()+' ';
  for(const rule of FUNC_RULES){for(const kw of rule.kw){if(t.includes(kw))return{fn:rule.fn,source:'Keyword "'+kw.trim()+'" in "'+title+'"',method:'title'}}}
  if(/\b(ceo|founder|co-founder|cofounder|owner|president|chief executive)\b/i.test(title))return{fn:'Executive/Founder',source:'Exec/founder title',method:'title'};
  return{fn:'Other',source:'No match in "'+title+'"',method:'title'};
}

/* ========== SENIORITY CLASSIFICATION ========== */
const SEN_EXEC=['\\bceo\\b','\\bcto\\b','\\bcio\\b','\\bciso\\b','\\bcfo\\b','\\bcoo\\b','\\bcmo\\b','\\bcpo\\b','\\bcro\\b','\\bchief\\b','\\bpresident\\b','\\bfounder\\b','\\bco-founder\\b','\\bcofounder\\b','\\bowner\\b','\\bpartner\\b','\\bvp\\b','vice president','\\bsvp\\b','\\bevp\\b','general manager','head of'];
const SEN_INFL=['\\bdirector\\b','\\bsr\\.?\\s','\\bsenior\\b','\\bprincipal\\b','\\blead\\b','\\bstaff\\b','\\bmanager\\b','team lead','group lead'];
const SEN_PRAC=['engineer','developer','analyst','specialist','coordinator','associate','administrator','consultant','designer','architect','scientist','intern','\\bjunior\\b','\\bjr\\.?\\b'];
const SEN_CRM={'Executive':'Exec Buyers','VP':'Exec Buyers','Owner':'Exec Buyers','Partner':'Exec Buyers','Director':'Decision Influencers','Senior':'Decision Influencers','Manager':'Decision Influencers','Employee':'Practitioners','Entry':'Practitioners'};

function classifySen(title,sen){
  if(sen&&sen.trim()){return{tier:SEN_CRM[sen.trim()]||'Unknown',source:'CRM: "'+sen.trim()+'"',method:'crm'}}
  if(!title)return{tier:'Unknown',source:'No title/seniority',method:'none'};
  const t=title.toLowerCase();
  for(const p of SEN_EXEC){if(new RegExp(p,'i').test(t))return{tier:'Exec Buyers',source:'/'+p+'/ in "'+title+'"',method:'title'}}
  for(const p of SEN_INFL){if(new RegExp(p,'i').test(t))return{tier:'Decision Influencers',source:'/'+p+'/ in "'+title+'"',method:'title'}}
  for(const p of SEN_PRAC){if(new RegExp(p,'i').test(t))return{tier:'Practitioners',source:'/'+p+'/ in "'+title+'"',method:'title'}}
  return{tier:'Unknown',source:'No pattern in "'+title+'"',method:'title'};
}
