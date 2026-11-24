(() => {
  const $ = (s, root = document) => root.querySelector(s);
  const form = $('.facets');
  const stateKey = 'fedoras-reading-browse';
  if (!form) {
    const back = $('[data-back]');
    try {
      const last = sessionStorage.getItem(stateKey);
      if (back && last && last.startsWith('?') && last.endsWith('#activities')) back.href = '../' + last;
    } catch (_) { /* a direct visit still has a browse link */ }
    return;
  }
  const input = $('#query'), status = $('#count');
  const buttons = [...form.querySelectorAll('[data-facet]')];
  const keys = [...new Set(buttons.map(b => b.dataset.facet))];
  const cards = [...document.querySelectorAll('[data-activity]')].map(el => ({
    el, label: el.textContent.toLowerCase(), id: el.dataset.activity,
    facets: Object.fromEntries(keys.map(k => [k, JSON.parse(el.getAttribute('data-' + k))]))
  }));
  let searchIndex, searchPromise;
  const index = () => searchPromise ||= fetch('./search-index.json').then(res => {
    if (!res.ok) throw new Error('Search unavailable');
    return res.json();
  }).then(data => { searchIndex = data; render(); }).catch(() => {});
  const render = () => {
    const q = input.value.trim().toLowerCase();
    const filters = Object.fromEntries(keys.map(k => [k, buttons.filter(b => b.dataset.facet === k && b.getAttribute('aria-pressed') === 'true').map(b => b.value)]));
    let count = 0;
    for (const c of cards) {
      const match = (!q || (searchIndex?.[c.id]?.toLowerCase() || c.label).includes(q)) && keys.every(k => !filters[k].length || filters[k].some(v => c.facets[k].includes(v)));
      c.el.hidden = !match;
      if (match) count++;
    }
    status.textContent = `${count} ${count === 1 ? 'activity' : 'activities'}`;
  };
  const restore = () => {
    const params = new URLSearchParams(location.search);
    input.value = params.get('q') || '';
    for (const b of buttons) b.setAttribute('aria-pressed', params.getAll(b.dataset.facet).includes(b.value));
    render(); if (input.value.trim()) index();
  };
  const update = () => {
    const params = new URLSearchParams();
    if (input.value.trim()) params.set('q', input.value.trim());
    for (const b of buttons) if (b.getAttribute('aria-pressed') === 'true') params.append(b.dataset.facet, b.value);
    const qs = params.size ? '?' + params.toString() : '';
    history.replaceState(null, '', location.pathname + qs + '#activities');
    render(); if (input.value.trim()) index();
  };
  form.hidden = false;
  form.addEventListener('submit', e => e.preventDefault());
  form.addEventListener('reset', e => { e.preventDefault(); input.value = ''; for (const b of buttons) b.setAttribute('aria-pressed', 'false'); update(); });
  for (const b of buttons) b.addEventListener('click', () => { b.setAttribute('aria-pressed', b.getAttribute('aria-pressed') !== 'true'); update(); });
  input.addEventListener('input', update);
  $('#cards').addEventListener('click', e => {
    if (!e.target.closest('a')) return;
    try { if (location.search) sessionStorage.setItem(stateKey, location.search + '#activities'); else sessionStorage.removeItem(stateKey); } catch (_) {}
  });
  window.addEventListener('popstate', restore);
  restore();
})();
// Allow keyboard scrolling when a source equation exceeds the reading column.
for (const el of document.querySelectorAll('.formula')) {
  const sync = () => { if (el.scrollWidth > el.clientWidth) el.tabIndex = 0; else el.removeAttribute('tabindex'); };
  sync(); if ('ResizeObserver' in window) new ResizeObserver(sync).observe(el);
}

(() => {
  const start = document.getElementById('start');
  const activities = document.getElementById('activities');
  if (!start || !activities) return;
  const links = [...document.querySelectorAll('[data-rail]')];
  let pending = false;
  const mark = (area) => {
    for (const link of links) {
      if (link.dataset.rail === area) link.setAttribute('aria-current', 'location');
      else link.removeAttribute('aria-current');
    }
  };
  const sync = () => {
    pending = false;
    const threshold = Math.min(100, window.innerHeight / 5);
    mark(activities.getBoundingClientRect().top <= threshold ||
      (location.hash === '#activities' && document.documentElement.scrollHeight <= window.innerHeight + window.scrollY + 2)
      ? 'activities' : 'start');
  };
  const queue = () => { if (!pending) { pending = true; requestAnimationFrame(sync); } };
  for (const link of links.filter(item => item.hash)) link.addEventListener('click', () => mark(link.dataset.rail));
  window.addEventListener('scroll', queue, {passive:true});
  window.addEventListener('hashchange', queue);
  window.addEventListener('resize', queue);
  window.addEventListener('pageshow', queue);
  queue();
})();
