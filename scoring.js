(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.LeadScoring=api})(typeof self!=='undefined'?self:this,function(){
  const FIELD_DEFS=[
    {key:'fit',label:'Customer fit',hint:'Industry, size, and region',weight:30,aliases:['\u5339\u914d\u5ea6','\u5ba2\u6237\u5339\u914d\u5ea6','fit','company fit','customer fit']},
    {key:'intent',label:'Buying intent',hint:'Engagement, inquiries, and visits',weight:25,aliases:['\u8d2d\u4e70\u610f\u5411','\u610f\u5411','intent','buying intent']},
    {key:'budget',label:'Budget',hint:'Available budget level',weight:20,aliases:['\u9884\u7b97','budget']},
    {key:'authority',label:'Decision authority',hint:'Contact influence',weight:15,aliases:['\u51b3\u7b56\u6743','\u51b3\u7b56\u80fd\u529b','authority','decision authority']},
    {key:'urgency',label:'Urgency',hint:'Expected purchase timeline',weight:10,aliases:['\u7d27\u8feb\u5ea6','\u7d27\u6025\u7a0b\u5ea6','urgency']}
  ];
  const COMPANY_ALIASES=['\u516c\u53f8\u540d\u79f0','\u5ba2\u6237\u540d\u79f0','\u516c\u53f8','\u4f01\u4e1a\u540d\u79f0','company','account','lead'];
  function normalize(v){return String(v??'').trim().toLowerCase().replace(/[\s_-]/g,'')}
  function findColumn(headers,aliases){const normalized=headers.map(normalize);return aliases.map(normalize).reduce((found,a)=>found>=0?found:normalized.indexOf(a),-1)}
  function clampScore(v){if(v===null||v===undefined||v==='')return 0;const raw=typeof v==='string'?v.replace('%','').trim():v;const n=Number(raw);if(!Number.isFinite(n))return 0;return Math.max(0,Math.min(100,n<=5?n*20:n<=10?n*10:n))}
  function priority(score){return score>=80?'A - Follow up now':score>=60?'B - Nurture actively':score>=40?'C - Monitor':'D - Deprioritize'}
  function reason(parts){return parts.sort((a,b)=>b.contribution-a.contribution).slice(0,2).filter(x=>x.value>0).map(x=>x.label+' '+Math.round(x.value)).join('; ')||'Insufficient scoring data'}
  function scoreRow(row,mapping,weights){const totalWeight=FIELD_DEFS.reduce((s,f)=>s+(Number(weights[f.key])||0),0)||1;const parts=FIELD_DEFS.map(f=>{const value=mapping[f.key]>=0?clampScore(row[mapping[f.key]]):0;return{label:f.label,value,contribution:value*(Number(weights[f.key])||0)/totalWeight}});const score=Math.round(parts.reduce((s,p)=>s+p.contribution,0));return{score,priority:priority(score),reason:reason(parts)}}
  function detect(headers){const mapping={company:findColumn(headers,COMPANY_ALIASES)};FIELD_DEFS.forEach(f=>mapping[f.key]=findColumn(headers,f.aliases));return mapping}
  return{FIELD_DEFS,COMPANY_ALIASES,findColumn,clampScore,priority,scoreRow,detect};
});
