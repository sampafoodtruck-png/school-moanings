(() => {
  'use strict';
  const prefix = 'sm.school-updates.';
  const dayKey = date => prefix + date;
  const weekday = date => new Date(date + 'T12:00:00Z').getUTCDay();
  const weeklyKey = date => prefix + 'weekday.' + weekday(date);
  const today = () => window.schoolCalendar.dateKey(now());
  const read = key => {
    try { return JSON.parse(localStorage.getItem(key)) || {}; }
    catch { return {}; }
  };
  function write(key, field, value) {
    const record = read(key);
    if (value === undefined) delete record[field];
    else record[field] = value;
    localStorage.setItem(key, JSON.stringify(record));
  }

  const dialog = document.createElement('dialog');
  dialog.className = 'cal-settings';
  dialog.setAttribute('aria-label','Parent updates');
  dialog.innerHTML = '<form method="dialog"><h2>Parent updates</h2><label>School date<input type="date" required></label>' +
    '<label>Classes for<select><option value="date">This date only</option><option value="weekly">Every week on this day</option></select></label>' +
    '<div data-classes></div><label>Family message for this date<input data-note maxlength="140" placeholder="e.g. Bring your library bag"></label>' +
    '<p>Updates save as you type on this laptop and appear on the TV on the selected date.</p><button>Done</button></form>';
  document.body.append(dialog);
  const dateField = dialog.querySelector('[type=date]');
  const scope = dialog.querySelector('select');
  const note = dialog.querySelector('[data-note]');
  const rows = dialog.querySelector('[data-classes]');
  function loadForm() {
    const date = dateField.value;
    rows.replaceChildren();
    note.disabled = !date;
    note.value = date ? read(dayKey(date)).note || '' : '';
    if (!date) return;
    const weekly = read(weeklyKey(date));
    const recurring = scope.value === 'weekly';
    const key = recurring ? weeklyKey(date) : dayKey(date);
    const record = read(key);
    cfg.names.forEach((name,i) => {
      const label = document.createElement('label');
      label.textContent = name + ' · specialist classes';
      const input = document.createElement('input');
      input.maxLength = 80;
      input.value = record[i] || '';
      input.placeholder = (!recurring && weekly[i]) || 'e.g. Library, PE';
      label.append(input);
      rows.append(label);
      input.oninput = () => {
        write(key,i,input.value.trim() || undefined);
        refresh();
      };
      if (!recurring) {
        const skipLabel = document.createElement('label');
        const skip = document.createElement('input');
        skip.type = 'checkbox';
        skip.style.cssText = 'display:inline;width:auto;margin:0 8px 0 0';
        skip.checked = record[i] === null;
        input.disabled = skip.checked;
        skipLabel.append(skip,'No specialist class on this date');
        rows.append(skipLabel);
        skip.onchange = () => {
          write(key,i,skip.checked ? null : input.value.trim() || undefined);
          input.disabled = skip.checked;
          refresh();
        };
      }
    });
  }
  dateField.oninput = scope.onchange = loadForm;
  note.oninput = () => {
    if (dateField.value) write(dayKey(dateField.value),'note',note.value.trim() || undefined);
    refresh();
  };
  document.addEventListener('click', event => {
    if (!event.target.closest('[data-parent-updates]')) return;
    dateField.value = today();
    scope.value = 'date';
    loadForm();
    dialog.showModal();
  });

  let lastMessage = '';
  function refresh() {
    const banner = document.querySelector('[data-school-banner]');
    if (!banner) return;
    const date = today();
    const day = read(dayKey(date)), weekly = read(weeklyKey(date));
    const items = cfg.names.map((name,i) => {
      const classes = Object.hasOwn(day,i) ? day[i] : weekly[i];
      return classes ? name + ' · ' + classes : '';
    }).filter(Boolean);
    if (day.note) items.push(day.note);
    const message = items.join('   •   ');
    if (banner.textContent !== message) banner.textContent = message;
    banner.hidden = !message;
    if (message && lastMessage !== message && !matchMedia('(prefers-reduced-motion: reduce)').matches)
      banner.animate([{opacity:0.35},{opacity:1}],{duration:700});
    lastMessage = message;
  }
  new MutationObserver(refresh).observe(document.getElementById('app'),{childList:true});
  window.addEventListener('storage',event => {if (!event.key || event.key.startsWith(prefix)) refresh();});
  refresh();
})();
