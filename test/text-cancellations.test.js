import test from 'node:test';
import assert from 'node:assert/strict';
import {TextCancellations} from '../src/text-cancellations.js';
function fixture(){
 const records=new Map();let writes=0;
 const event={id:'event-1',subject:'Budget review',start:{dateTime:'2030-09-11T14:00:00',timeZone:'Eastern Standard Time'},end:{dateTime:'2030-09-11T15:00:00',timeZone:'Eastern Standard Time'},organizer:{emailAddress:{address:'owner@example.com'}},attendees:[{emailAddress:{address:'alex@example.com'}}],isOrganizer:true,isCancelled:false,type:'singleInstance'};
 const graph={principalMailbox:'owner@example.com',listPrincipalCalendar:async()=>[event],getPrincipalCalendarEvent:async()=>event,cancelVoiceMeeting:async()=>{writes++;return {cancelled:true,cancellationSent:true};}};
 const dropbox={readDeliveryRecord:async key=>records.get(key),createDeliveryRecord:async(key,value)=>{if(records.has(key))return false;records.set(key,value);return true;}};
 const openai={respond:async()=>({text:JSON.stringify({startIso:'2030-09-11T00:00:00-04:00',endIso:'2030-09-12T00:00:00-04:00',selector:'Budget review'})})};
 const now=()=>new Date('2030-09-10T12:00:00Z'),make=()=>new TextCancellations({graph,dropbox,openai,now});
 const request={owner:graph.principalMailbox,requestKey:'prepare',text:'Cancel the Budget review meeting tomorrow',receivedAt:now().toISOString(),maxReplyLength:480};
 return {records,event,graph,openai,make,request,writes:()=>writes};
}
const confirm=(f,reply,extra={})=>({...f.request,text:reply.match(/CONFIRM CANCEL [a-f0-9]{12}/)[0],requestKey:'confirm',...extra});
test('SMS/email cancellation previews exact event and requires a separate owner confirmation',async()=>{
 const f=fixture(),reply=await f.make().handle(f.request);assert.match(reply,/Budget review/);assert.match(reply,/alex@example.com/);assert.equal(f.writes(),0);
 assert.match(await f.make().handle(confirm(f,reply)),/^Cancelled:/);assert.equal(f.writes(),1);
 assert.match(await f.make().handle(confirm(f,reply,{requestKey:'again'})),/already attempted/);assert.equal(f.writes(),1);
});
test('wrong owner and same source cannot cancel',async()=>{
 const f=fixture(),reply=await f.make().handle(f.request);
 await assert.rejects(f.make().handle(confirm(f,reply,{owner:'other@example.com'})),/Authenticated/);
 assert.match(await f.make().handle(confirm(f,reply,{requestKey:'prepare'})),/No matching/);assert.equal(f.writes(),0);
});
test('changed event, non-organizer and multiple matches never cancel',async()=>{
 const f=fixture(),reply=await f.make().handle(f.request);f.event.subject='Changed';
 assert.match(await f.make().handle(confirm(f,reply)),/changed/);
 f.event.subject='Budget review';f.event.isOrganizer=false;assert.match(await f.make().handle({...f.request,requestKey:'other'}),/not the organizer/);
 f.event.isOrganizer=true;f.graph.listPrincipalCalendar=async()=>[f.event,{...f.event,id:'another'}];
 assert.match(await f.make().handle({...f.request,requestKey:'ambiguous'}),/2 matching/);assert.equal(f.writes(),0);
});
test('expired and uncertain cancellation cannot execute or retry',async()=>{
 const f=fixture(),reply=await f.make().handle(f.request);
 const proposal=[...f.records.values()].find(v=>v.event);proposal.expiresAt='2020-01-01';
 assert.match(await f.make().handle(confirm(f,reply)),/expired/);
 proposal.expiresAt='2035-01-01';let attempts=0;f.graph.cancelVoiceMeeting=async()=>{attempts++;throw Error('timeout');};
 assert.match(await f.make().handle(confirm(f,reply)),/not verified/);
 assert.match(await f.make().handle(confirm(f,reply,{requestKey:'again'})),/already attempted/);assert.equal(attempts,1);
});
test('invented selector and broad date windows do not access the calendar',async()=>{
 const f=fixture();f.graph.listPrincipalCalendar=async()=>assert.fail('invalid lookup');
 f.openai.respond=async()=>({text:JSON.stringify({selector:'Invented',startIso:'2030-01-01',endIso:'2031-01-01'})});
 assert.match(await f.make().handle(f.request),/Which meeting/);assert.equal(f.writes(),0);
});
