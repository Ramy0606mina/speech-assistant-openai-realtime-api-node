import test from 'node:test';
import assert from 'node:assert/strict';
import {MorningBrief,briefSlot,easternMidnight,gatherBrief,parseBriefReport,renderBriefHtml} from '../src/morning-brief.js';
const reportJson=JSON.stringify({title:'Morning Executive Report',date:'2026-09-07',priorities:[{priority:'1',item:'Review quote',hint:'Needs approval',status:'Pending',due:'Today'}],calendar:[],emails:[],tasks:[],risks:[],completed:[]});
test('morning schedule observes Eastern daylight saving and the catch-up window',()=>{
 assert.equal(briefSlot(new Date('2026-09-07T11:29:00Z')).due,false);
 assert.equal(briefSlot(new Date('2026-09-07T11:30:00Z')).due,true);
 assert.equal(briefSlot(new Date('2026-01-07T12:30:00Z')).due,true);
 assert.equal(briefSlot(new Date('2026-09-07T16:00:00Z')).due,false);
 assert.equal(easternMidnight('2026-03-08'),'2026-03-08T05:00:00.000Z');
 assert.equal(easternMidnight('2026-03-09'),'2026-03-09T04:00:00.000Z');
});
function harness(){
 const claims=new Set();const sent=[];
 const graph={principalMailbox:'owner@example.com',listPrincipalInbox:async()=>[],listPrincipalCalendar:async()=>[],listFollowUps:async()=>[],sendMail:async m=>sent.push(m)};
 const guard={check:async k=>claims.has(k),claim:async k=>{if(claims.has(k))return false;claims.add(k);return true;},complete:async()=>{}};
 const deps={graph,guard,openai:{respond:async()=>({text:reportJson})},dropbox:{saveReport:async()=>({path:'report.pdf'})}};
 return {deps,sent};
}
test('daily report sends once across concurrent workers and restart, then once next day',async()=>{
 const {deps,sent}=harness();const now=new Date('2026-09-07T11:30:00Z');
 await Promise.all([new MorningBrief(deps).tick(now),new MorningBrief(deps).tick(now)]);
 await new MorningBrief(deps).tick(now);assert.equal(sent.length,1);
 await new MorningBrief(deps).tick(new Date('2026-09-08T11:30:00Z'));assert.equal(sent.length,2);
});
test('ambiguous morning send is never repeated after restart',async()=>{
 const {deps}=harness();let sends=0;deps.graph.sendMail=async()=>{sends++;throw new Error('timeout');};const now=new Date('2026-09-07T11:30:00Z');
 await assert.rejects(new MorningBrief(deps).tick(now),/timeout/);
 await new MorningBrief(deps).tick(now);assert.equal(sends,1);
});
test('brief preserves missing-source status rather than treating failure as empty',async()=>{
 const {deps}=harness();deps.graph.listFollowUps=async()=>{throw new Error('denied');};
 const data=await gatherBrief(deps.graph,new Date('2026-09-07T11:30:00Z'));assert.equal(data.sources.followUps.available,false);assert.equal(data.sources.calendar.available,true);
});
test('brief JSON renders a compact colored HTML table',()=>{
 const report=parseBriefReport(reportJson);const html=renderBriefHtml(report);
 assert.match(html,/Top priorities/);assert.match(html,/Review quote/);assert.match(html,/background:#fff2cc/);assert.doesNotMatch(html,/undefined/);
});
