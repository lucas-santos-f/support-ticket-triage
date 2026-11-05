(function () {
  'use strict';

  const T = window.Triage;
  const SVG = 'http://www.w3.org/2000/svg';
  const LEVEL = {
    critical: { label: 'Critical', color: 'var(--critical)', due: 'answer today' },
    high: { label: 'High', color: 'var(--serious)', due: 'within 24 hours' },
    medium: { label: 'Medium', color: 'var(--warning)', due: 'within 2 business days' },
    low: { label: 'Low', color: 'var(--low)', due: 'when there is time' },
  };
  const ORDER = ['critical', 'high', 'medium', 'low'];

  let source = null;      // { name, rows }
  let result = null;
  let selected = null;    // item id
  const edits = new Map(); // id -> reply text changed by hand
  const filter = { level: 'all', topic: 'all', text: '' };

  const $ = (id) => document.getElementById(id);

  function el(tag, attrs, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === false || v == null) continue;
      node.setAttribute(k, v === true ? '' : v);
    }
    for (const child of children.flat()) {
      if (child == null) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  }

  function svg(tag, attrs, text) {
    const node = document.createElementNS(SVG, tag);
    for (const [k, v] of Object.entries(attrs || {})) node.setAttribute(k, v);
    if (text != null) node.textContent = text;
    return node;
  }

  function icon(level) {
    const color = LEVEL[level].color;
    const s = svg('svg', { width: 14, height: 14, viewBox: '0 0 14 14', 'aria-hidden': 'true' });
    if (level === 'critical') s.append(svg('path', { d: 'M7 1 13 12H1Z', fill: color }));
    else if (level === 'high') s.append(svg('path', { d: 'M7 1 13 7 7 13 1 7Z', fill: color }));
    else if (level === 'medium') s.append(svg('circle', { cx: 7, cy: 7, r: 5, fill: color }));
    else s.append(svg('circle', { cx: 7, cy: 7, r: 4.25, fill: 'none', stroke: 'var(--muted)', 'stroke-width': 1.5 }));
    return s;
  }

  function status(level, extra) {
    return el('span', { class: 'status' }, icon(level), LEVEL[level].label + (extra || ''));
  }

  function download(name, data, type) {
    const url = URL.createObjectURL(new Blob([data], { type }));
    const a = el('a', { href: url, download: name });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ------------------------------------------------------------ loading

  async function inflate(bytes) {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function readFile(file) {
    if (/\.xlsx$/i.test(file.name)) {
      if (typeof DecompressionStream === 'undefined') throw new Error('This browser cannot open .xlsx files. Save the sheet as CSV and try again.');
      return T.readXLSX(await file.arrayBuffer(), inflate);
    }
    return T.parseCSV(await file.text());
  }

  async function load(file) {
    if (!file) return;
    try {
      const rows = await readFile(file);
      source = { name: file.name, rows };
      edits.clear();
      selected = null;
      $('drop').querySelector('[data-name]').textContent = file.name;
      $('as-of').value = '';
      run();
    } catch (err) {
      $('errors').textContent = err.message;
    }
  }

  function useSample() {
    source = { name: 'messages.csv (sample)', rows: T.parseCSV(window.SAMPLE.csv) };
    edits.clear();
    selected = null;
    $('drop').querySelector('[data-name]').textContent = source.name;
    $('as-of').value = window.SAMPLE.asOf;
    run();
  }

  function run() {
    if (!source) return;
    const read = T.readItems(source.rows);
    const errors = read.errors;
    $('errors').textContent = errors.length ? errors.slice(0, 5).join(' · ') + (errors.length > 5 ? ` · and ${errors.length - 5} more` : '') : '';
    result = T.triage(read.items, { asOf: $('as-of').value || null, team: $('team').value.trim() || 'Customer Care' });
    if (!$('as-of').value) $('as-of').value = result.asOf;
    $('loaded').textContent = `${read.items.length} messages`;
    if (!selected || !result.items.some((i) => i.id === selected)) selected = result.items.length ? result.items[0].id : null;
    render();
  }

  // ------------------------------------------------------------ rendering

  function render() {
    renderTiles();
    renderLegend();
    renderChart();
    renderFilters();
    renderQueue();
    renderDetail();
  }

  function renderTiles() {
    const s = result.summary;
    const reviews = result.items.filter((i) => i.public).length;
    const oldest = result.items.reduce((m, i) => (i.waiting > m.waiting ? i : m), result.items[0] || { waiting: 0 });
    const tiles = [
      { label: 'Messages', value: String(s.total), sub: `${s.total - reviews} tickets, ${reviews} public reviews` },
      { label: 'Critical', level: 'critical', value: String(s.levels.critical), sub: LEVEL.critical.due },
      { label: 'High', level: 'high', value: String(s.levels.high), sub: LEVEL.high.due },
      { label: 'Waiting longest', value: `${s.oldest} days`, sub: oldest && oldest.customer ? oldest.customer : '' },
      { label: 'Average rating', value: s.avgRating == null ? '–' : `${s.avgRating.toFixed(1)} / 5`, sub: `from ${s.reviews} rated messages` },
    ];
    $('kpis').replaceChildren(...tiles.map((t) => el('div', { class: 'tile' },
      el('div', { class: 'label' }, t.level ? icon(t.level) : null, t.label),
      el('div', { class: 'value' }, t.value),
      el('div', { class: 'sub' }, t.sub))));
  }

  function renderLegend() {
    $('legend').replaceChildren(...ORDER.map((k) => el('span', { class: 'legend-item' }, icon(k), LEVEL[k].label)));
  }

  function renderChart() {
    const host = $('topics-chart');
    const topics = result.topics;
    const draw = () => {
      host.textContent = '';
      const W = Math.max(280, host.clientWidth);
      const narrow = W < 520;
      const labelW = narrow ? 142 : 168;
      const totalW = 40;
      const row = 30;
      const bh = 14;
      const H = topics.length * row + 4;
      const max = Math.max(...topics.map((t) => t.total), 1);
      const span = W - labelW - totalW;
      const root = svg('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, class: 'bars', role: 'img',
        'aria-label': 'Messages per topic by urgency: ' + topics.map((t) => `${t.label} ${t.total}`).join(', ') });
      root.append(svg('line', { x1: labelW, x2: labelW, y1: 0, y2: H, stroke: 'var(--axis)' }));
      const tip = el('div', { class: 'tooltip', hidden: true });
      topics.forEach((t, i) => {
        const y = i * row + (row - bh) / 2 + 2;
        const g = svg('g', { tabindex: 0, role: 'button', class: filter.topic === t.key ? 'on' : '',
          'aria-label': `${t.label}: ${t.total} messages, ` + ORDER.map((k) => `${t.levels[k]} ${LEVEL[k].label.toLowerCase()}`).join(', ') + '. Filter the queue.' });
        g.append(svg('text', { x: labelW - 8, y: y + bh / 2 + 4, 'text-anchor': 'end', class: 'cat', style: `font-size:${narrow ? 12 : 13}px` }, t.label));
        const parts = ORDER.filter((k) => t.levels[k] > 0);
        let x = labelW + 1;
        const end = labelW + 1 + (span - 2) * (t.total / max);
        parts.forEach((k, j) => {
          const w = (end - labelW - 1) * (t.levels[k] / t.total);
          const last = j === parts.length - 1;
          const ww = Math.max(1, last ? w : w - 2);
          if (last) {
            const r = Math.min(4, ww / 2, bh / 2);
            g.append(svg('path', { class: 'bar', fill: LEVEL[k].color,
              d: `M${x},${y}H${x + ww - r}Q${x + ww},${y} ${x + ww},${y + r}V${y + bh - r}Q${x + ww},${y + bh} ${x + ww - r},${y + bh}H${x}Z` }));
          } else {
            g.append(svg('rect', { class: 'bar', x, y, width: ww, height: bh, fill: LEVEL[k].color }));
          }
          x += w;
        });
        g.append(svg('text', { x: end + 8, y: y + bh / 2 + 4, class: 'value' }, t.total));
        g.append(svg('rect', { class: 'hit', x: 0, y: i * row, width: W, height: row }));
        const show = () => {
          tip.replaceChildren(el('div', { class: 'when' }, `${t.label} · ${t.total}`),
            ...ORDER.map((k) => el('div', { class: 'row' }, icon(k), el('strong', null, t.levels[k]), el('span', { class: 'name' }, LEVEL[k].label))));
          tip.hidden = false;
          const left = Math.min(end + 40, W - tip.offsetWidth);
          tip.style.left = Math.max(0, left) + 'px';
          tip.style.top = Math.max(0, Math.min(i * row - 10, H - tip.offsetHeight)) + 'px';
        };
        const pick = () => {
          filter.topic = filter.topic === t.key ? 'all' : t.key;
          renderChart();
          renderFilters();
          renderQueue();
        };
        g.addEventListener('pointerenter', show);
        g.addEventListener('focus', show);
        g.addEventListener('pointerleave', () => (tip.hidden = true));
        g.addEventListener('blur', () => (tip.hidden = true));
        g.addEventListener('click', pick);
        g.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });
        root.append(g);
      });
      host.append(root, tip);
    };
    draw();
    if (!host._observer) {
      let lastW = host.clientWidth;
      host._observer = new ResizeObserver(() => {
        if (host.clientWidth === lastW) return;
        lastW = host.clientWidth;
        requestAnimationFrame(() => result && renderChart());
      });
      host._observer.observe(host);
    }
  }

  function renderFilters() {
    const counts = result.summary.levels;
    const buttons = [['all', `All ${result.summary.total}`]].concat(ORDER.map((k) => [k, `${LEVEL[k].label} ${counts[k]}`]));
    $('level-filter').replaceChildren(...buttons.map(([k, text]) => {
      const b = el('button', { type: 'button', 'aria-pressed': String(filter.level === k) }, k === 'all' ? null : icon(k), text);
      b.addEventListener('click', () => { filter.level = k; renderFilters(); renderQueue(); });
      return b;
    }));
    const sel = $('topic-filter');
    sel.replaceChildren(el('option', { value: 'all' }, 'All topics'),
      ...result.topics.map((t) => el('option', { value: t.key }, `${t.label} (${t.total})`)));
    sel.value = result.topics.some((t) => t.key === filter.topic) ? filter.topic : 'all';
  }

  function visible() {
    const q = filter.text.toLowerCase();
    return result.items.filter((i) => (filter.level === 'all' || i.level === filter.level) &&
      (filter.topic === 'all' || i.topic === filter.topic) &&
      (!q || `${i.customer} ${i.subject} ${i.message} ${i.id}`.toLowerCase().includes(q)));
  }

  function preview(i) {
    const s = i.subject || i.message;
    return s.length > 54 ? s.slice(0, 52) + '…' : s;
  }

  function renderQueue() {
    const rows = visible();
    $('queue-count').textContent = rows.length === result.items.length ? 'Most urgent first.' : `${rows.length} of ${result.items.length} shown. Most urgent first.`;
    const columns = [
      { label: 'Urgency', value: (i) => status(i.level) },
      { label: 'Waiting', num: true, value: (i) => (i.waiting ? i.waiting + ' d' : 'today') },
      { label: 'Topic', value: (i) => [i.topicLabel, i.followUp ? el('span', { class: 'badge' }, 'follow-up') : null] },
      { label: 'Customer', cls: 'customer', value: (i) => [i.customer, i.public ? el('span', { class: 'badge' }, 'public') : null] },
      { label: 'Subject', cls: 'subject', value: (i) => preview(i) },
    ];
    const head = el('tr', null, columns.map((c) => el('th', { class: c.num ? 'num' : null }, c.label)));
    const body = rows.length ? rows.map((i) => {
      const tr = el('tr', { class: 'pick' + (i.id === selected ? ' selected' : ''), tabindex: 0, 'data-id': i.id },
        columns.map((c) => el('td', { class: [c.num ? 'num' : '', c.cls || ''].join(' ').trim() || null }, c.value(i))));
      const pick = () => {
        selected = i.id;
        $('queue').querySelectorAll('tr.selected').forEach((r) => r.classList.remove('selected'));
        tr.classList.add('selected');
        renderDetail();
        if (window.matchMedia('(max-width: 899px)').matches) $('detail').scrollIntoView({ behavior: 'smooth', block: 'start' });
      };
      tr.addEventListener('click', pick);
      tr.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });
      return tr;
    }) : [el('tr', null, el('td', { colspan: columns.length, class: 'note' }, 'Nothing matches these filters.'))];
    $('queue').replaceChildren(el('thead', null, head), el('tbody', null, body));
  }

  function replyOf(i) {
    return edits.has(i.id) ? edits.get(i.id) : i.reply;
  }

  function renderDetail() {
    const box = $('detail');
    const i = result.items.find((x) => x.id === selected);
    if (!i) {
      box.replaceChildren(el('h2', { id: 'detail-title' }, 'Message'), el('p', { class: 'note' }, 'Pick a message in the queue.'));
      return;
    }
    const order = T.orderNumber(`${i.subject} ${i.message}`);
    const meta = [i.customer, i.channel, i.date || 'no date', i.rating != null ? `${i.rating}/5` : null, order ? `order ${order}` : null, i.id]
      .filter(Boolean).join(' · ');
    const area = el('textarea', { id: 'reply', rows: 10, spellcheck: 'true', 'aria-label': 'Suggested reply' });
    area.value = replyOf(i);
    area.addEventListener('input', () => {
      if (area.value === i.reply) edits.delete(i.id);
      else edits.set(i.id, area.value);
      reset.hidden = !edits.has(i.id);
    });
    const copy = el('button', { type: 'button', class: 'btn' }, 'Copy reply');
    copy.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(area.value);
      } catch (e) {
        area.select();
        document.execCommand('copy');
      }
      copy.textContent = 'Copied';
      setTimeout(() => (copy.textContent = 'Copy reply'), 1400);
    });
    const reset = el('button', { type: 'button', class: 'btn', hidden: !edits.has(i.id) }, 'Undo my changes');
    reset.addEventListener('click', () => { edits.delete(i.id); area.value = i.reply; reset.hidden = true; });
    const next = el('button', { type: 'button', class: 'btn' }, 'Next message');
    next.addEventListener('click', () => {
      const rows = visible();
      const k = rows.findIndex((x) => x.id === i.id);
      const n = rows[k + 1] || rows[0];
      if (!n) return;
      selected = n.id;
      renderQueue();
      renderDetail();
      const tr = $('queue').querySelector(`tr[data-id="${CSS.escape(n.id)}"]`);
      if (tr) tr.scrollIntoView({ block: 'nearest' });
    });
    box.replaceChildren(
      el('h2', { id: 'detail-title' }, i.subject || `Message from ${i.customer}`),
      el('p', { class: 'meta' }, meta),
      el('p', { class: 'verdict' }, status(i.level, ` · ${i.score} pts`), el('span', { class: 'topic' }, i.topicLabel + (i.followUp ? ' (follow-up)' : ''))),
      i.reasons.length ? el('ul', { class: 'chips', 'aria-label': 'Why' }, i.reasons.map((r) => el('li', null, r))) : null,
      el('blockquote', null, i.message || i.subject),
      el('div', { class: 'reply-head' }, el('label', { for: 'reply' }, el('strong', null, 'Suggested reply')),
        i.public ? el('span', { class: 'note' }, 'Public review: short answer, details in private.') : null),
      area,
      el('div', { class: 'actions' }, copy, next, reset));
  }

  // ------------------------------------------------------------ wiring

  const drop = $('drop');
  $('file').addEventListener('change', (e) => load(e.target.files[0]));
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('over');
    load(e.dataTransfer.files[0]);
  });
  $('use-sample').addEventListener('click', useSample);
  $('get-sample').addEventListener('click', () => {
    const bin = atob(window.SAMPLE.xlsx);
    const bytes = Uint8Array.from(bin, (c) => c.charCodeAt(0));
    download('messages.xlsx', bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  });
  $('as-of').addEventListener('change', run);
  $('team').addEventListener('change', run);
  $('topic-filter').addEventListener('change', (e) => { filter.topic = e.target.value; renderChart(); renderQueue(); });
  $('search').addEventListener('input', (e) => { filter.text = e.target.value.trim(); renderQueue(); });
  $('export').addEventListener('click', () => {
    if (!result) return;
    download(`triage-${result.asOf}.csv`, T.toCSV(visible(), [
      { label: 'id', value: (i) => i.id },
      { label: 'date', value: (i) => i.date },
      { label: 'customer', value: (i) => i.customer },
      { label: 'channel', value: (i) => i.channel },
      { label: 'rating', value: (i) => i.rating },
      { label: 'topic', value: (i) => i.topicLabel },
      { label: 'urgency', value: (i) => i.levelLabel },
      { label: 'score', value: (i) => i.score },
      { label: 'waiting_days', value: (i) => i.waiting },
      { label: 'why', value: (i) => i.reasons.join('; ') },
      { label: 'subject', value: (i) => i.subject },
      { label: 'message', value: (i) => i.message },
      { label: 'reply', value: (i) => replyOf(i) },
    ]), 'text/csv;charset=utf-8');
  });
  useSample();
})();
