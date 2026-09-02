(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.LeadScoring=api})(typeof self!=='undefined'?self:this,function(){
  const FIELD_DEFS=[
    {key:'fit',label:'客户匹配度',hint:'行业、规模、地区',weight:30,aliases:['匹配度','客户匹配度','fit','company fit']},
    {key:'intent',label:'购买意向',hint:'互动、询价、访问',weight:25,aliases:['购买意向','意向','intent','buying intent']},
    {key:'budget',label:'预算',hint:'预算充足程度',weight:20,aliases:['预算','budget']},
    {key:'authority',label:'决策权',hint:'联系人影响力',weight:15,aliases:['决策权','决策能力','authority']},
    {key:'urgency',label:'紧迫度',hint:'预计成交时间',weight:10,aliases:['紧迫度','紧急程度','urgency']}
  ];
  const COMPANY_ALIASES=['公司名称','客户名称','公司','企业名称','company','account','lead'];
  function normalize(v){return String(v??'').trim().toLowerCase().replace(/[\s_-]/g,'')}
  function findColumn(headers,aliases){const normalized=headers.map(normalize);return aliases.map(normalize).reduce((found,a)=>found>=0?found:normalized.indexOf(a),-1)}
  function clampScore(v){if(v===null||v===undefined||v==='')return 0;const raw=typeof v==='string'?v.replace('%','').trim():v;const n=Number(raw);if(!Number.isFinite(n))return 0;return Math.max(0,Math.min(100,n<=5?n*20:n<=10?n*10:n))}
  function priority(score){return score>=80?'A - 立即跟进':score>=60?'B - 重点培育':score>=40?'C - 持续观察':'D - 暂缓投入'}
  function reason(parts){return parts.sort((a,b)=>b.contribution-a.contribution).slice(0,2).filter(x=>x.value>0).map(x=>x.label+' '+Math.round(x.value)+'分').join('；')||'评分信息不足'}
  function scoreRow(row,mapping,weights){const totalWeight=FIELD_DEFS.reduce((s,f)=>s+(Number(weights[f.key])||0),0)||1;const parts=FIELD_DEFS.map(f=>{const value=mapping[f.key]>=0?clampScore(row[mapping[f.key]]):0;return{label:f.label,value,contribution:value*(Number(weights[f.key])||0)/totalWeight}});const score=Math.round(parts.reduce((s,p)=>s+p.contribution,0));return{score,priority:priority(score),reason:reason(parts)}}
  function detect(headers){const mapping={company:findColumn(headers,COMPANY_ALIASES)};FIELD_DEFS.forEach(f=>mapping[f.key]=findColumn(headers,f.aliases));return mapping}
  return{FIELD_DEFS,COMPANY_ALIASES,findColumn,clampScore,priority,scoreRow,detect};
});
