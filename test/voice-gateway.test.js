import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizePhone, isAuthorizedCaller, buildIncomingCallTwiML } from '../src/voice-gateway.js';

test('normalizes North American phone formatting', () => {
  assert.equal(normalizePhone('(514) 814-3333'), '+15148143333');
  assert.equal(normalizePhone('+1 514 814 3333'), '+15148143333');
});

test('authorizes only the configured principal caller', () => {
  assert.equal(isAuthorizedCaller('+1 514 814 3333', '+15148143333'), true);
  assert.equal(isAuthorizedCaller('+1 514 555 0000', '+15148143333'), false);
});

test('incoming call TwiML connects only to the generated tokenized media route', () => {
  const xml = buildIncomingCallTwiML({ host: 'london-ai-pr-1.onrender.com', streamToken: 'abc-123' });
  assert.match(xml, /wss:\/\/london-ai-pr-1\.onrender\.com\/media-stream\/abc-123/);
  assert.match(xml, /<Connect>/);
});

import {EventEmitter} from 'node:events';
import {registerVoiceRoutes} from '../src/voice-gateway.js';
test('voice uses current protocol and waits for both session and phone stream before greeting',async()=>{
 let incoming,media,upstream;
 class Socket extends EventEmitter {constructor(url,options){super();this.readyState=1;this.sent=[];this.options=options;upstream=this;}send(x){this.sent.push(JSON.parse(x));}close(){this.readyState=3;}}
 const app={all:(path,handler)=>incoming=handler,register:fn=>fn({get:(path,options,handler)=>media=handler})};
 registerVoiceRoutes(app,{openAiApiKey:'test',principalPhone:'+15145550000',WebSocketImpl:Socket});
 let xml;const reply={type:()=>reply,send:x=>xml=x};
 await incoming({body:{From:'+15145550000'},headers:{host:'example.com'}},reply);
 const token=xml.match(/media-stream\/([^" ]+)/)[1];
 const phone=new EventEmitter();phone.readyState=1;phone.send=()=>{};phone.close=()=>{};
 media(phone,{params:{streamToken:token}});
 assert.equal(upstream.options.headers['OpenAI-Beta'],undefined);
 upstream.emit('open');
 phone.emit('message',JSON.stringify({event:'media',media:{payload:'audio'}}));
 assert.equal(upstream.sent.filter(e=>e.type==='input_audio_buffer.append').length,0);
 upstream.emit('message',JSON.stringify({type:'session.updated'}));
 assert.equal(upstream.sent.filter(e=>e.type==='response.create').length,0);
 phone.emit('message',JSON.stringify({event:'start',start:{streamSid:'stream'}}));
 assert.equal(upstream.sent.filter(e=>e.type==='response.create').length,1);
 assert.equal(upstream.sent.filter(e=>e.type==='input_audio_buffer.append').length,1);
 upstream.emit('message',JSON.stringify({type:'session.updated'}));
 assert.equal(upstream.sent.filter(e=>e.type==='response.create').length,1);
});
