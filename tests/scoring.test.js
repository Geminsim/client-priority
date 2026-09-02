const assert=require('assert');const s=require('../scoring');
assert.equal(s.clampScore('80%'),80);assert.equal(s.clampScore(4),80);assert.equal(s.clampScore(12),12);assert.equal(s.priority(80),'A - 立即跟进');assert.equal(s.priority(59),'C - 持续观察');
const headers=['公司名称','匹配度','购买意向','预算','决策权','紧迫度'];const map=s.detect(headers);const weights=Object.fromEntries(s.FIELD_DEFS.map(f=>[f.key,f.weight]));const result=s.scoreRow(['星海科技',90,80,70,60,50],map,weights);assert.equal(result.score,75);assert.equal(result.priority,'B - 重点培育');assert.ok(result.reason.includes('客户匹配度'));
assert.equal(s.detect(['Company','Buying Intent','Budget']).company,0);console.log('scoring tests passed');
