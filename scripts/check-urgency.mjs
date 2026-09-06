// Run explicitly with the deployed OpenAI configuration. No mailbox, SMS, or ledger writes.
import fs from 'node:fs';
import {loadConfig} from '../src/config.js';
import {OpenAIClient} from '../src/openai-client.js';
import {urgencyInstructions} from '../src/urgent-alerts.js';
const config=loadConfig();
const openai=new OpenAIClient({apiKey:config.openai.apiKey,model:config.openai.taskModel});
const cases=JSON.parse(fs.readFileSync(new URL('../test/fixtures/urgency-cases.json',import.meta.url),'utf8'));
let failures=0;
for(const item of cases){
 const result=await openai.respond({instructions:urgencyInstructions,input:JSON.stringify({subject:item.subject,preview:item.preview,received:new Date().toISOString()})});
 const actual=result.text.trim().toUpperCase();
 const passed=actual===item.expected;
 if(!passed)failures++;
 console.log(JSON.stringify({case:item.name,expected:item.expected,actual,passed}));
}
console.log(JSON.stringify({cases:cases.length,failures,realMessagesSent:0}));
process.exitCode=failures?1:0;
