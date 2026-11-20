/* Calendar data stays in this browser. No events or tokens are sent to GitHub. */
(function () {
  'use strict';
  const TZ = 'Australia/Brisbane';
  const KEY = 'sm.calendar.v2';
  const SETTINGS = 'sm.calendar.settings.v2';
  const TICKS = 'sm.calendar.ticks.v2';
  const MSAL_URL = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.30.0/+esm';
  const SCOPES = ['Calendars.Read'];
  const escape = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const read = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } };
  const save = (key, value) => localStorage.setItem(key, JSON.stringify(value));
  let settings = read(SETTINGS, {clientId:'', calendarId:'', keywords:'swim, surf, gymnastics, year 6, experience day, school, excursion'});
  if (!settings.clientId) settings.clientId = 'a3115d87-966f-40f2-89a3-751905e09378';
  let cache = read(KEY, {events:[], syncedAt:null, from:'', to:''});
  let ticks = read(TICKS, {});
  let weekOffset = 0, auth = null, busy = false, status = '', calendars = [];
  const dateKey = date => new Intl.DateTimeFormat('en-CA', {timeZone:TZ,year:'numeric',month:'2-digit',day:'2-digit'}).format(date);
  const shift = (date, days) => new Date(new Date(date + 'T12:00:00+10:00').getTime() + days * 86400000);
  const dateShift = (date, days) => dateKey(shift(date, days));
  const today = () => dateKey(typeof now === 'function' ? now() : new Date());
  const formatDate = (date, options) => new Intl.DateTimeFormat('en-AU', {timeZone:TZ,...options}).format(new Date(date));
  function timestamp(value) {
    if (!value) return '';
    if (typeof value === 'string') return value;
    const text = value.dateTime;
    if (!text) return '';
    if (/Z$|[+-]\d\d:\d\d$/.test(text)) return text;
    if (value.timeZone !== 'UTC') throw new Error('Calendar returned an unsupported timezone. Refresh using UTC.');
    return text + 'Z';
  }
  function visible(event) {
    const words = settings.keywords.split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
    return !event.cancelled && !event.private && words.some(word => event.subject.toLowerCase().includes(word));
  }
  function normalise(event) {
    const body = event.body?.content || event.bodyPreview || '';
    const parser = new DOMParser();
    return {id:String(event.id), subject:String(event.subject || 'Activity').slice(0,300), start:timestamp(event.start), end:timestamp(event.end),
      allDay:!!(event.isAllDay ?? event.is_all_day),
      startDate:(event.isAllDay ?? event.is_all_day) ? dateKey(new Date(timestamp(event.start))) : null,
      endDate:(event.isAllDay ?? event.is_all_day) ? dateKey(new Date(timestamp(event.end))) : null,
      body:parser.parseFromString(body,'text/html').body.textContent,
      cancelled:!!(event.isCancelled ?? event.is_cancelled), private:['private','confidential'].includes(event.sensitivity)};
  }
  function eventsOn(date) {
    return cache.events.filter(event => visible(event) && (event.allDay
      ? event.startDate <= date && event.endDate > date
      : new Date(event.start) < new Date(dateShift(date,1)+'T00:00:00+10:00') && new Date(event.end) > new Date(date+'T00:00:00+10:00')))
      .sort((a,b)=>a.start.localeCompare(b.start));
  }
  function childFor(event) {
    const subject = event.subject.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
    const names = typeof cfg !== 'undefined' ? cfg.names : ['Maia','Thaís'];
    const matches = names.filter(name => subject.includes(name.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()));
    return matches.length ? matches.join(' & ') : 'Check who is attending';
  }
  function packing(event) {
    const title = event.subject.toLowerCase(), body = event.body.toLowerCase();
    if (/surf/.test(title)) return [
      /wear their swimming gear|wear swimming gear/.test(body) ? 'Wear swimmers and rash shirt to school' : 'Pack swimmers and rash shirt',
      'Pack school uniform for afterwards', 'Pack hat and suitable footwear', 'Apply sunscreen', 'Pack towel (suggested)'];
    if (/swim/.test(title)) return ['Pack swimmers, towel and goggles (suggested)', 'Check the school or lesson packing list'];
    if (/gymnast/.test(title)) return ['Pack gymnastics clothes and water bottle (suggested)'];
    if (/experience day|excursion/.test(title)) return ['Check arrival time and school instructions'];
    return [];
  }
  function eventCard(event, date, checklist) {
    const title = event.subject.replace(/^\d+[A-Z]?\s*-\s*/,'').split('\n')[0];
    const time = event.allDay ? 'All day' : formatDate(event.start,{hour:'numeric',minute:'2-digit'});
    return '<article class="cal-event"><div class="cal-event-top"><b>'+escape(title)+'</b><span>'+escape(time)+'</span></div>'+
      '<small>'+escape(childFor(event))+'</small>'+(checklist?packing(event).map((label,index)=>{
        const id = date+'|'+event.id+'|'+index;
        return '<label class="cal-check"><input type="checkbox" data-cal-tick="'+escape(id)+'" '+(ticks[id]?'checked':'')+'><span>'+escape(label)+'</span></label>';
      }).join(''):'')+'</article>';
  }
  const transportPrefix = 'sm.transport.v2.';
  const children = () => typeof cfg === 'undefined' ? ['Maia','Thaís'] : cfg.names;
  function tripsFor(date) {
    const trips = [{id:'school', title:'School', legs:['Drop-off','Pick-up / after-school care']}];
    for (const event of eventsOn(date)) {
      trips.push({id:'event:'+event.id, title:event.subject, legs:['Taking','Collecting']});
    }
    return trips;
  }
  function tripFields(date, trip, editable) {
    const plan = read(transportPrefix+date, {});
    return children().map((name,child) => {
      const fields = trip.legs.map((leg,index) => {
        const key = JSON.stringify([trip.id,child,index]);
        const value = plan[key] || '';
        return editable ? '<label>'+escape(leg)+'<input data-trip="'+escape(key)+'" maxlength="80" placeholder="Needs confirming" value="'+escape(value)+'"></label>'
          : '<div>'+escape(leg)+' · '+escape(value || 'Needs confirming')+'</div>';
      });
      return '<article class="cal-event"><b>'+escape(name)+'</b>'+fields.join('')+'</article>';
    }).join('');
  }
  function transportPanels(date) {
    const [school,...activities] = tripsFor(date);
    return '<div class="cal-today"><section><h3>School drop-off &amp; pick-up</h3>'+tripFields(date,school,false)+'</section>'+
      '<section><h3>Activity transport</h3>'+activities.map(trip=>'<h4>'+escape(trip.title)+'</h4>'+tripFields(date,trip,false)).join('')+
      (activities.length?'':'<p>'+(cache.syncedAt && date>=cache.from && date<cache.to?'No matching activities for today.':'Load Outlook to see today’s activities.')+'</p>')+'</section></div>';
  }
  const tripDialog = document.createElement('dialog');
  tripDialog.className = 'cal-settings';
  document.body.append(tripDialog);
  function editTrips(date) {
    tripDialog.innerHTML = '<form method="dialog"><h2>Family transport</h2><label>Plan for<input type="date" value="'+escape(date)+'"></label>'+
      tripsFor(date).map(trip=>'<h3>'+escape(trip.title)+'</h3>'+tripFields(date,trip,true)).join('')+
      '<p>Enter a parent, carer, care arrangement or “Not attending”. Changes save as you type.</p><button>Done</button></form>';
    tripDialog.querySelector('[type=date]').onchange = event => {if(event.target.value) editTrips(event.target.value);};
    tripDialog.querySelectorAll('[data-trip]').forEach(input => input.oninput = () => {
      const key = transportPrefix+date;
      const plan = read(key,{});
      plan[input.dataset.trip] = input.value.trim();
      save(key,plan);
      draw();
    });
    if (!tripDialog.open) tripDialog.showModal();
  }
  function draw() {
    let host = document.getElementById('calendar-section');
    if (!host) return;
    const date = today(), tomorrow = dateShift(date,1), weekday = shift(date,0).getUTCDay();
    const monday = dateShift(date, -(weekday === 0 ? 6 : weekday-1) + weekOffset*7);
    const current = eventsOn(date), next = eventsOn(tomorrow);
    const covered = cache.syncedAt && date >= cache.from && date < cache.to;
    const stale = cache.syncedAt && Date.now()-Date.parse(cache.syncedAt)>3600000;
    host.innerHTML = '<div class="cal-heading"><div><small>'+escape(formatDate(date+'T12:00:00+10:00',{weekday:'long',day:'numeric',month:'long'}))+'</small><h2>Today & this week</h2></div><div><a class="cal-control-link" href="?view=control">Tick routine tasks</a> <button type="button" data-cal-open>Calendar settings</button></div></div>'+
      '<p class="cal-status" role="status">'+escape(status || (!cache.syncedAt ? 'Connect Outlook to see school activities here.' : !covered ? 'Calendar needs refreshing for today.' : 'Updated '+formatDate(cache.syncedAt,{weekday:'short',hour:'numeric',minute:'2-digit'})+(stale?' · Saved copy — refresh before relying on it.':'')))+'</p>'+
      transportPanels(date)+
      '<div class="cal-today" style="margin-top:18px"><section><h3>Today · remember these</h3>'+(current.length?current.map(e=>eventCard(e,date,true)).join(''):'<p>'+(covered?'No matching school activities in the saved calendar.':'Calendar not loaded for today.')+'</p>')+'</section><section><h3>Prepare for tomorrow</h3>'+(next.length?next.map(e=>eventCard(e,tomorrow,true)).join(''):'<p>'+(cache.syncedAt && tomorrow<cache.to?'No matching activities in the saved calendar.':'Refresh to check tomorrow.')+'</p>')+'</section></div>'+
      '<div class="cal-nav"><button data-cal-week="-1" aria-label="Previous week">←</button><b>Week of '+escape(formatDate(monday+'T12:00:00+10:00',{day:'numeric',month:'short'}))+'</b><button data-cal-week="0">This week</button><button data-cal-week="1" aria-label="Next week">→</button></div>'+
      '<div class="cal-week">'+Array.from({length:7},(_,i)=>{const d=dateShift(monday,i);return '<section class="cal-day '+(d===date?'is-today':'')+'"><h3>'+escape(formatDate(d+'T12:00:00+10:00',{weekday:'short',day:'numeric'}))+'</h3>'+(eventsOn(d).map(e=>eventCard(e,d,false)).join('') || '<p>'+(cache.syncedAt&&d>=cache.from&&d<cache.to?'No matching activities':'Not loaded')+'</p>')+'</section>';}).join('')+'</div>';
    host.querySelector('[data-cal-open]').onclick = openSettings;
    host.querySelectorAll('[data-cal-week]').forEach(button=>button.onclick=()=>{const n=Number(button.dataset.calWeek);weekOffset=n===0?0:weekOffset+n;draw();});
    host.querySelectorAll('[data-cal-tick]').forEach(input=>input.onchange=()=>{ticks[input.dataset.calTick]=input.checked;save(TICKS,ticks);});
  }
  function mount() {
    const bar = document.querySelector('.control .ctl-bar');
    if (bar && !document.getElementById('edit-transport')) {
      const button = document.createElement('button');
      button.id = 'edit-transport';
      button.textContent = 'Family transport';
      button.className = 'skipbtn';
      button.style = 'width:auto;padding:8px 12px';
      button.onclick = () => editTrips(today());
      bar.append(button);
    }
    if(document.getElementById('calendar-section')) return;
    const board = document.querySelector('.board');
    if(!board) return;
    const section = document.createElement('section'); section.id='calendar-section'; board.append(section); draw();
  }
  const modal = document.createElement('dialog'); modal.className='cal-settings';
  modal.innerHTML='<form method="dialog"><h2>School calendar</h2><p>Only activities matching your chosen words appear on the board. Private events are hidden. Sign in separately on each device.</p><label>Microsoft application ID<input name="clientId" placeholder="Application (client) ID" autocomplete="off"></label><p class="cal-help">One-time Microsoft app registration is needed before the dashboard can sign in. Never enter a password or client secret here.</p><label>Activity words (separate with commas)<textarea name="keywords"></textarea></label><label>Calendar<select name="calendarId"><option value="">Main calendar</option></select></label><p id="cal-dialog-status" role="status"></p><div class="cal-actions"><button type="button" data-cal-save>Save settings</button><button type="button" data-cal-connect>Sign in to Outlook</button><button type="button" data-cal-refresh>Refresh</button><button type="button" data-cal-forget>Disconnect & clear calendar</button><button value="close">Close</button></div></form>';
  document.body.append(modal);
  const field = name => modal.querySelector('[name="'+name+'"]');
  function say(message) {status=message;modal.querySelector('#cal-dialog-status').textContent=message;draw();}
  function openSettings() {
    field('clientId').value=settings.clientId;field('keywords').value=settings.keywords;
    field('calendarId').innerHTML='<option value="">Main calendar</option>'+calendars.map(c=>'<option value="'+escape(c.id)+'">'+escape(c.name)+'</option>').join('');
    if(settings.calendarId && !calendars.some(c=>c.id===settings.calendarId))field('calendarId').insertAdjacentHTML('beforeend','<option value="'+escape(settings.calendarId)+'">Selected calendar</option>');
    field('calendarId').value=settings.calendarId;modal.showModal();
  }
  async function saveSettings() {
    const clientId=field('clientId').value.trim();
    if(clientId && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId))throw new Error('Enter a valid Microsoft application ID.');
    const changed = clientId!==settings.clientId || field('calendarId').value!==settings.calendarId;
    if(clientId!==settings.clientId)auth=null;
    settings={clientId,calendarId:field('calendarId').value,keywords:field('keywords').value};save(SETTINGS,settings);
    if(changed){cache={events:[],syncedAt:null,from:'',to:''};localStorage.removeItem(KEY);}
    say('Settings saved on this device.');
  }
  async function getAuth() {
    if(!settings.clientId)throw new Error('Microsoft application ID is needed before signing in.');
    if(!auth){
      const {PublicClientApplication}=await import(MSAL_URL);
      const instance=new PublicClientApplication({auth:{clientId:settings.clientId,authority:'https://login.microsoftonline.com/consumers',redirectUri:new URL('auth.html',location.href).href},cache:{cacheLocation:'sessionStorage'}});
      await instance.initialize();auth=instance;
    }
    return auth;
  }
  async function getToken(interactive) {
    const instance=await getAuth();
    let account=instance.getActiveAccount() || instance.getAllAccounts()[0];
    if(!account){
      if(!interactive)throw new Error('Open Calendar settings and sign in to Outlook.');
      const result=await instance.loginPopup({scopes:SCOPES,prompt:'select_account'});account=result.account;instance.setActiveAccount(account);return result.accessToken;
    }
    try{return (await instance.acquireTokenSilent({scopes:SCOPES,account})).accessToken;}
    catch(error){if(!interactive)throw new Error('Outlook needs you to sign in again. Open Calendar settings.');return (await instance.acquireTokenPopup({scopes:SCOPES,account})).accessToken;}
  }
  async function graphAll(path,token) {
    let url='https://graph.microsoft.com/v1.0'+path,items=[];
    while(url){
      if(!url.startsWith('https://graph.microsoft.com/v1.0/'))throw new Error('Unexpected calendar response.');
      const response=await fetch(url,{headers:{Authorization:'Bearer '+token,Prefer:'outlook.timezone="UTC"'}});
      if(!response.ok)throw new Error(response.status===429?'Outlook is busy. Try again later.':'Outlook could not refresh ('+response.status+'). Your saved calendar has been kept.');
      const data=await response.json();items.push(...(data.value||[]));url=data['@odata.nextLink']||null;
    }
    return items;
  }
  async function sync(interactive=false) {
    if(busy)return;busy=true;
    try{
      say('Refreshing Outlook…');const token=await getToken(interactive);
      const from=dateShift(today(),-7),to=dateShift(today(),22);
      const query=new URLSearchParams({startDateTime:from+'T00:00:00+10:00',endDateTime:to+'T00:00:00+10:00','$top':'100','$select':'id,subject,start,end,isAllDay,isCancelled,sensitivity,body'});
      const path=settings.calendarId?'/me/calendars/'+encodeURIComponent(settings.calendarId)+'/calendarView':'/me/calendarView';
      const raw=await graphAll(path+'?'+query,token);
      const events=raw.map(normalise).filter(visible);
      cache={events,syncedAt:new Date().toISOString(),from,to};save(KEY,cache);
      calendars=await graphAll('/me/calendars?$select=id,name',token);
      if(modal.open){const selected=settings.calendarId;field('calendarId').innerHTML='<option value="">Main calendar</option>'+calendars.map(c=>'<option value="'+escape(c.id)+'">'+escape(c.name)+'</option>').join('');field('calendarId').value=selected;}
      status='';modal.querySelector('#cal-dialog-status').textContent='Calendar refreshed.';draw();
    }catch(error){say(error.message || 'Calendar could not refresh.');}finally{busy=false;}
  }
  const action=(selector,callback)=>modal.querySelector(selector).onclick=()=>Promise.resolve().then(callback).catch(error=>say(error.message));
  action('[data-cal-save]',saveSettings);
  action('[data-cal-connect]',async()=>{await saveSettings();await sync(true);});
  action('[data-cal-refresh]',async()=>{await saveSettings();await sync(false);});
  action('[data-cal-forget]',async()=>{
    if(busy)throw new Error('Wait for the current refresh to finish.');
    if(auth){await auth.clearCache();auth=null;}
    cache={events:[],syncedAt:null,from:'',to:''};ticks={};localStorage.removeItem(KEY);localStorage.removeItem(TICKS);say('Calendar disconnected and saved events cleared from this device.');
  });
  new MutationObserver(mount).observe(document.getElementById('app'),{childList:true,subtree:true});mount();
  window.addEventListener('storage',event=>{if(event.key && event.key.startsWith(transportPrefix)) draw();if(event.key===KEY){cache=read(KEY,{events:[],syncedAt:null,from:'',to:''});draw();}});
  setInterval(()=>{if(settings.clientId)sync(false);},15*60*1000);
  if(settings.clientId)sync(false);
  // Pure helpers are exposed for deterministic timezone and rule tests.
  window.schoolCalendar={dateKey,dateShift,normalise,packing,childFor,eventsOn};
})();

