const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const script = fs.readFileSync(path.join(__dirname,'../calendar.js'),'utf8');
const html = fs.readFileSync(path.join(__dirname,'../index.html'),'utf8');
for(const match of html.matchAll(/<script>([\s\S]*?)<\/script>/g))new vm.Script(match[1]);
function harness(events=[]) {
  const storage = new Map([['sm.calendar.v2',JSON.stringify({events,from:'2026-11-16',to:'2026-12-01',syncedAt:new Date().toISOString()})]]);
  const elements = new Map();
  const element=()=>({value:'',innerHTML:'',textContent:'',open:false,querySelector(s){if(!elements.has(s))elements.set(s,element());return elements.get(s);},querySelectorAll(){return [];},showModal(){this.open=true;}});
  const host=element(),app=element();let mounted=false;const board={append(){mounted=true;}};
  const context={console,Intl,Date,URL,URLSearchParams,localStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},
    DOMParser:class{parseFromString(s){return {body:{textContent:s.replace(/<[^>]*>/g,'')}};}},
    MutationObserver:class{observe(){}},setInterval(){},location:{href:'https://example.test/v2/'},
    document:{body:{append(){}},createElement:tag=>tag==='section'?host:element(),getElementById:id=>id==='app'?app:(mounted?host:null),querySelector(){return board;}},
    window:{addEventListener(){}},cfg:{names:['Maia','Thaís']},now:()=>new Date('2026-11-20T07:30:00+10:00')};
  vm.createContext(context);vm.runInContext(script,context);return {api:context.window.schoolCalendar,context,host,storage};
}
function event(subject,start,end,extra={}) {return {id:subject,subject,start,end,body:'',allDay:false,private:false,cancelled:false,...extra};}
const tests=[];const test=(name,run)=>tests.push({name,run});
test('Brisbane date crosses UTC midnight correctly',()=>{
  const {api}=harness();assert.equal(api.dateKey(new Date('2026-11-19T15:00:00Z')),'2026-11-20');assert.equal(api.dateShift('2026-12-31',1),'2027-01-01');
});
test('Only overlapping, non-private, matching activities appear',()=>{
  const list=[event('Swimming Maia','2026-11-19T23:00:00Z','2026-11-20T00:00:00Z'),event('Swimming cancelled','2026-11-19T23:00:00Z','2026-11-20T00:00:00Z',{cancelled:true}),event('Swimming private','2026-11-19T23:00:00Z','2026-11-20T00:00:00Z',{private:true}),event('Adult appointment','2026-11-19T23:00:00Z','2026-11-20T00:00:00Z'),event('Swimming yesterday','2026-11-19T12:00:00Z','2026-11-19T14:00:00Z')];
  const {api}=harness(list);assert.equal(api.eventsOn('2026-11-20').length,1);
});
test('All-day end date is exclusive',()=>{
  const {api}=harness([event('School event','','',{allDay:true,startDate:'2026-11-20',endDate:'2026-11-21'})]);assert.equal(api.eventsOn('2026-11-20').length,1);assert.equal(api.eventsOn('2026-11-21').length,0);
});
test('UTC all-day boundaries convert to Brisbane dates',()=>{
  const {api}=harness();const result=api.normalise({id:'all-day',subject:'School',isAllDay:true,start:{dateTime:'2026-11-19T14:00:00',timeZone:'UTC'},end:{dateTime:'2026-11-20T14:00:00',timeZone:'UTC'}});assert.equal(result.startDate,'2026-11-20');assert.equal(result.endDate,'2026-11-21');
});
test('Children are not guessed for unnamed events',()=>{
  const {api}=harness();assert.equal(api.childFor({subject:'Swimming Thais'}),'Thaís');assert.equal(api.childFor({subject:'Learn to swim'}),'Check who is attending');
});
test('Surf instructions differ from generic swim suggestions',()=>{
  const {api}=harness();assert.equal(api.packing({subject:'Surf Skills',body:'Students to wear their swimming gear to school'})[0],'Wear swimmers and rash shirt to school');assert.match(api.packing({subject:'Swimming',body:''})[0],/suggested/);
});
test('Event titles are escaped in generated markup',()=>{
  const {host}=harness([event('Swimming <img src=x onerror=alert(1)>','2026-11-19T23:00:00Z','2026-11-20T00:00:00Z')]);
  assert.ok(host.innerHTML.includes('&lt;img'));assert.ok(!host.innerHTML.includes('<img'));
});
test('Graph UTC date fields and unsupported zones are handled explicitly',()=>{
  const {api}=harness();const base={id:'x',subject:'Swimming',start:{dateTime:'2026-11-20T00:00:00',timeZone:'UTC'},end:{dateTime:'2026-11-20T01:00:00',timeZone:'UTC'}};
  assert.equal(api.normalise(base).start,'2026-11-20T00:00:00Z');assert.throws(()=>api.normalise({...base,start:{dateTime:'2026-11-20T00:00:00',timeZone:'Unknown'}}));
});
// Expose the async integration only inside this test VM, leaving production globals minimal.
test('Graph pagination collects all pages without sending tokens elsewhere',async()=>{
  const {context}=harness();vm.runInContext(script.replace('window.schoolCalendar={','window.schoolCalendar={graphAll,sync,'),context);
  let count=0;context.fetch=async url=>({ok:true,json:async()=>++count===1?{value:[1],'@odata.nextLink':'https://graph.microsoft.com/v1.0/next'}:{value:[2]}});
  assert.equal((await context.window.schoolCalendar.graphAll('/me/calendarView','test-token')).join(','),'1,2');
  context.fetch=async()=>({ok:true,json:async()=>({value:[],'@odata.nextLink':'https://other.test/steal'})});
  await assert.rejects(context.window.schoolCalendar.graphAll('/me/calendarView','test-token'),/Unexpected/);
});
test('Failed refresh preserves the last saved calendar',async()=>{
  const {context,storage}=harness();vm.runInContext(script.replace('window.schoolCalendar={','window.schoolCalendar={sync,'),context);
  const before=storage.get('sm.calendar.v2');await context.window.schoolCalendar.sync(false);assert.equal(storage.get('sm.calendar.v2'),before);
});
(async()=>{for(const t of tests){await t.run();console.log('PASS '+t.name);}console.log(`${tests.length} tests passed; inline dashboard JavaScript parsed.`);})().catch(error=>{console.error(error);process.exit(1);});
