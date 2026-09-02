const assert=require('assert');const s=require('../scoring');
assert.equal(s.clampScore('80%'),80);assert.equal(s.clampScore(4),80);assert.equal(s.clampScore(12),12);assert.equal(s.priority(80),'A - Follow up now');assert.equal(s.priority(59),'C - Monitor');
const headers=['Company','Customer Fit','Buying Intent','Budget','Authority','Urgency'];const map=s.detect(headers);const weights=Object.fromEntries(s.FIELD_DEFS.map(f=>[f.key,f.weight]));const result=s.scoreRow(['Northstar Technology',90,80,70,60,50],map,weights);assert.equal(result.score,75);assert.equal(result.priority,'B - Nurture actively');assert.ok(result.reason.includes('Customer fit'));
assert.equal(s.detect(['Company','Buying Intent','Budget']).company,0);console.log('scoring tests passed');
