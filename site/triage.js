// Sorts customer messages (tickets and reviews) by topic and urgency and drafts a reply for each.
// Runs in the browser (window.Triage) and under Node (require).
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Triage = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- reading

  function detectDelimiter(line) {
    const count = (ch) => line.split(ch).length - 1;
    return count(';') > count(',') ? ';' : count('\t') > count(',') ? '\t' : ',';
  }

  function parseCSV(text) {
    text = String(text).replace(/^﻿/, '');
    const delim = detectDelimiter(text.slice(0, text.search(/\r?\n|$/)));
    const rows = [];
    let row = [];
    let field = '';
    let quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (ch === '"') quoted = false;
        else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === delim) { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return tableToObjects(rows);
  }

  function tableToObjects(rows) {
    const nonEmpty = rows.filter((r) => r.some((v) => String(v).trim() !== ''));
    if (!nonEmpty.length) return [];
    const header = nonEmpty[0].map((h) => String(h).trim());
    return nonEmpty.slice(1).map((r) => Object.fromEntries(header.map((h, k) => [h, String(r[k] == null ? '' : r[k]).trim()])));
  }

  // Minimal .xlsx reader: first worksheet, shared and inline strings, numbers and dates.
  // `inflate(bytes) -> Promise<Uint8Array>` decompresses raw deflate data.
  async function readXLSX(buffer, inflate) {
    const bytes = new Uint8Array(buffer);
    const files = unzipIndex(bytes);
    const text = async (name) => {
      const f = files[name];
      if (!f) return null;
      const data = bytes.subarray(f.start, f.start + f.size);
      const out = f.method === 0 ? data : f.method === 8 ? await inflate(data) : null;
      if (!out) throw new Error('unsupported compression in ' + name);
      return new TextDecoder('utf-8').decode(out);
    };
    const workbook = await text('xl/workbook.xml');
    if (!workbook) throw new Error('not an .xlsx file');
    const rels = (await text('xl/_rels/workbook.xml.rels')) || '';
    const firstSheet = /<sheet\b[^>]*\br:id="([^"]+)"/.exec(workbook);
    let sheetPath = 'xl/worksheets/sheet1.xml';
    if (firstSheet) {
      const rel = new RegExp('<Relationship\\b[^>]*Id="' + firstSheet[1] + '"[^>]*Target="([^"]+)"').exec(rels) ||
        new RegExp('<Relationship\\b[^>]*Target="([^"]+)"[^>]*Id="' + firstSheet[1] + '"').exec(rels);
      if (rel) sheetPath = rel[1].startsWith('/') ? rel[1].slice(1) : 'xl/' + rel[1].replace(/^\.\//, '');
    }
    const shared = [];
    const sst = await text('xl/sharedStrings.xml');
    if (sst) {
      for (const si of sst.match(/<si\b[\s\S]*?<\/si>/g) || []) {
        shared.push((si.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map((t) => xmlText(t.replace(/<[^>]+>/g, ''))).join(''));
      }
    }
    const sheet = await text(sheetPath);
    if (!sheet) throw new Error('worksheet not found');
    const rows = [];
    for (const rowXml of sheet.match(/<row\b[\s\S]*?<\/row>/g) || []) {
      const row = [];
      for (const c of rowXml.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || []) {
        const ref = /\br="([A-Z]+)\d+"/.exec(c);
        const col = ref ? colIndex(ref[1]) : row.length;
        const type = (/\bt="([^"]+)"/.exec(c) || [])[1];
        const v = /<v>([\s\S]*?)<\/v>/.exec(c);
        let value = '';
        if (type === 's' && v) value = shared[Number(v[1])] || '';
        else if (type === 'inlineStr') value = ((c.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map((t) => xmlText(t.replace(/<[^>]+>/g, '')))).join('');
        else if (v) value = xmlText(v[1]);
        row[col] = value;
      }
      rows.push(Array.from(row, (v) => (v == null ? '' : v)));
    }
    return tableToObjects(rows);
  }

  function colIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function xmlText(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d))).replace(/&amp;/g, '&');
  }

  function unzipIndex(bytes) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('not a zip file');
    const count = view.getUint16(eocd + 10, true);
    let p = view.getUint32(eocd + 16, true);
    const files = {};
    for (let n = 0; n < count; n++) {
      if (view.getUint32(p, true) !== 0x02014b50) break;
      const method = view.getUint16(p + 10, true);
      const size = view.getUint32(p + 20, true);
      const nameLen = view.getUint16(p + 28, true);
      const extraLen = view.getUint16(p + 30, true);
      const commentLen = view.getUint16(p + 32, true);
      const local = view.getUint32(p + 42, true);
      const name = new TextDecoder('utf-8').decode(bytes.subarray(p + 46, p + 46 + nameLen));
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      files[name] = { method, size, start };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }

  // Excel stores dates as days since 1899-12-30.
  function parseDate(value) {
    const s = String(value || '').trim();
    let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(s)) * 86400000);
      return d.toISOString().slice(0, 10);
    }
    return null;
  }

  function daysBetween(a, b) {
    return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);
  }

  const ALIASES = {
    id: ['id', 'ticket', 'ticket_id', 'number', 'ref'],
    date: ['date', 'created', 'created_at', 'received', 'opened'],
    channel: ['channel', 'source', 'via'],
    customer: ['customer', 'name', 'customer_name', 'from', 'requester'],
    rating: ['rating', 'stars', 'score'],
    subject: ['subject', 'title', 'summary'],
    message: ['message', 'body', 'text', 'review', 'description', 'comment'],
  };

  function pick(row, key) {
    const keys = Object.keys(row);
    for (const alias of ALIASES[key]) {
      const k = keys.find((x) => x.toLowerCase().replace(/[\s-]+/g, '_') === alias);
      if (k != null && row[k] !== '') return row[k];
    }
    return '';
  }

  function readItems(rows) {
    const items = [];
    const errors = [];
    rows.forEach((row, i) => {
      const line = i + 2;
      const message = pick(row, 'message');
      const subject = pick(row, 'subject');
      if (!message && !subject) { errors.push(`row ${line}: no message`); return; }
      const date = parseDate(pick(row, 'date'));
      if (!date) errors.push(`row ${line}: date not recognised, treated as today`);
      const ratingRaw = pick(row, 'rating');
      const rating = ratingRaw === '' ? null : Number(String(ratingRaw).replace(',', '.'));
      items.push({
        id: pick(row, 'id') || String(line - 1),
        date,
        channel: pick(row, 'channel') || 'Email',
        customer: pick(row, 'customer') || 'Customer',
        rating: Number.isFinite(rating) && rating >= 1 && rating <= 5 ? Math.round(rating) : null,
        subject,
        message,
      });
    });
    return { items, errors };
  }

  // ---------------------------------------------------------------- topics

  const TOPICS = [
    { key: 'billing', label: 'Billing and refunds', words: ['charged', 'charge', 'refund', 'refunded', 'invoice', 'billing', 'billed', 'payment', 'card', 'overcharged', 'twice', 'double', 'receipt', 'coupon', 'discount', 'price', 'paid'] },
    { key: 'delivery', label: 'Delivery', words: ['delivery', 'deliver', 'delivered', 'shipping', 'shipped', 'tracking', 'package', 'parcel', 'courier', 'arrived', 'arrive', 'late', 'delayed', 'lost', 'address', 'dispatch', 'dispatched'] },
    { key: 'defect', label: 'Product problem', words: ['broken', 'damaged', 'defective', 'faulty', 'cracked', 'stopped working', 'not working', "doesn't work", "won't turn on", "doesn't turn on", 'missing part', 'leaking', 'leaks', 'noise', 'noisy', 'scratched', 'dead', 'overheating', 'smell', 'broke', 'breaks', 'smoke', 'sparks', 'burnt', 'burned', 'melted'] },
    { key: 'returns', label: 'Returns and exchanges', words: ['return', 'returns', 'exchange', 'send it back', 'send back', 'return label', 'warranty', 'replacement', 'wrong size', 'wrong item', 'wrong color'] },
    { key: 'account', label: 'Account and login', words: ['login', 'log in', 'password', 'account', 'locked', 'sign in', 'verification code', 'two-factor', '2fa', 'reset link', 'email address', 'username'] },
    { key: 'website', label: 'Website and app', words: ['app', 'website', 'site', 'error', 'crash', 'crashes', 'bug', 'checkout', 'page', 'loading', 'button', 'update'] },
    { key: 'cancel', label: 'Cancellation', words: ['cancel', 'cancellation', 'unsubscribe', 'close my account', 'subscription', 'stop the order'] },
    { key: 'praise', label: 'Praise', words: ['thank', 'thanks', 'great', 'love', 'excellent', 'amazing', 'perfect', 'recommend', 'happy', 'awesome', 'fantastic', 'quick delivery', 'fast delivery'] },
    { key: 'request', label: 'Suggestion', words: ['would be nice', 'please add', 'suggestion', 'suggest', 'feature', 'wish', 'could you add', 'it would help'] },
  ];
  const OTHER = { key: 'other', label: 'Other' };
  const NEGATIVE = /\b(not|never|no|broken|damaged|refund|late|worst|bad|terrible|disappointed|angry|still|again|wrong|missing|charged)\b/i;

  function countHits(text, word) {
    const re = new RegExp('\\b' + word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+') + '\\b', 'gi');
    return (text.match(re) || []).length;
  }

  function classify(item) {
    const subject = (item.subject || '').toLowerCase();
    const body = (item.message || '').toLowerCase();
    let best = null;
    for (const t of TOPICS) {
      let score = 0;
      const found = [];
      for (const w of t.words) {
        const n = countHits(body, w) + 2 * countHits(subject, w);
        if (n) { score += n * (w.includes(' ') ? 2 : 1); found.push(w); }
      }
      if (t.key === 'praise') {
        // praise only counts when nothing negative is said
        if (NEGATIVE.test(subject + ' ' + body)) score = 0;
        if (item.rating >= 4) score += 2;
      }
      if (t.key === 'request' && NEGATIVE.test(subject + ' ' + body)) score = Math.floor(score / 2);
      if (score > 0 && (!best || score > best.score)) best = { topic: t, score, found };
    }
    if (!best) return { topic: OTHER.key, label: OTHER.label, words: [] };
    return { topic: best.topic.key, label: best.topic.label, words: best.found };
  }

  // ---------------------------------------------------------------- urgency

  const LEVELS = [
    { key: 'critical', label: 'Critical', min: 70 },
    { key: 'high', label: 'High', min: 40 },
    { key: 'medium', label: 'Medium', min: 15 },
    { key: 'low', label: 'Low', min: -Infinity },
  ];
  const SIGNALS = [
    { re: /\b(smoke|smoking|fire|caught fire|burn|burned|burnt|sparks?|electric shock|shocked|injur(y|ed)|hurt|melted|explod(ed|e))\b/i, points: 100, reason: 'safety risk' },
    { re: /\b(lawyer|legal action|sue|chargeback|dispute (the|this) charge|consumer protection|small claims|report you)\b/i, points: 40, reason: 'threat to escalate' },
    { re: /\b(charged twice|double charged|charged two times|overcharged|unauthori[sz]ed|did not authori[sz]e)\b/i, points: 30, reason: 'money taken in error' },
    { re: /\b(again|still|third time|second time|no reply|no response|nobody (answered|replied)|already (contacted|wrote|called|emailed)|weeks? ago)\b/i, points: 15, reason: 'repeat contact' },
    { re: /\b(urgent|urgently|asap|immediately|right now|today)\b/i, points: 10, reason: 'asks for urgency' },
    { re: /\b(before it ships|before shipping|before it is shipped|before dispatch)\b/i, points: 10, reason: 'time-sensitive' },
  ];
  // some topics need an answer sooner than others even when the customer is calm
  const TOPIC_BASE = { cancel: 15, defect: 10, delivery: 5, billing: 5, account: 5 };

  function urgency(item, ctx) {
    const text = `${item.subject || ''} ${item.message || ''}`;
    const topic = ctx && ctx.topic;
    let score = TOPIC_BASE[topic] || 0;
    const reasons = [];
    for (const s of SIGNALS) {
      if (s.re.test(text)) { score += s.points; reasons.push(s.reason); }
    }
    if (item.rating === 1) { score += 20; reasons.push('1-star rating'); }
    else if (item.rating === 2) { score += 10; reasons.push('2-star rating'); }
    if (isPublic(item) && item.rating != null && item.rating <= 2) { score += 10; reasons.push('public review'); }
    const waiting = ctx && ctx.asOf && item.date ? Math.max(0, daysBetween(item.date, ctx.asOf)) : 0;
    if (waiting > 5) { score += 20; reasons.push(`waiting ${waiting} days`); }
    else if (waiting > 2) { score += 10; reasons.push(`waiting ${waiting} days`); }
    const caps = (text.match(/\b[A-Z]{3,}\b/g) || []).filter((w) => !/^(USB|LED|SKU|PDF|USA|ASAP|FAQ|TV|HDMI|OK)$/.test(w)).length;
    if (caps >= 3 || /!{3,}/.test(text)) { score += 10; reasons.push('angry tone'); }
    const prior = ctx && ctx.contacts ? (ctx.contacts[contactKey(item)] || 1) : 1;
    if (prior >= 2) { score += 15; reasons.push(`${ordinal(prior)} message from this customer`); }
    if ((topic === 'praise' || topic === 'request') && score < 40) score = Math.min(score, 10);
    const level = reasons.includes('safety risk') ? LEVELS[0] : LEVELS.find((l) => score >= l.min);
    return { level: level.key, levelLabel: level.label, score, reasons, waiting };
  }

  function ordinal(n) {
    return n + (n % 10 === 1 && n % 100 !== 11 ? 'st' : n % 10 === 2 && n % 100 !== 12 ? 'nd' : n % 10 === 3 && n % 100 !== 13 ? 'rd' : 'th');
  }

  function contactKey(item) {
    return String(item.customer || '').trim().toLowerCase();
  }

  function isPublic(item) {
    return /review|store|app store|play|google|trustpilot|marketplace|social/i.test(item.channel || '');
  }

  // ---------------------------------------------------------------- replies

  function firstName(name) {
    const first = String(name || '').trim().split(/\s+/)[0] || '';
    return /^[A-Za-zÀ-ÿ'-]+$/.test(first) && first.length > 1 ? first : '';
  }

  function orderNumber(text) {
    const m = /(?:order|ord|#)\s*(?:no\.?|number|nr\.?)?\s*#?\s*([A-Z]{0,3}-?\d{4,})/i.exec(text || '');
    return m ? m[1] : null;
  }

  // Replies are plain templates picked by topic and urgency; there is no AI involved.
  // {name} first name, {customer} full name, {order} order number, {team} who signs.
  // [text] is left out when a placeholder inside it is empty; [text|other] uses other instead.
  const DEFAULT_TEMPLATES = {
    greeting: 'Hi[ {name}],',
    signoff: 'Best regards,\n{team}',
    safety: 'Please stop using the product and unplug it now. Your safety comes first, and we take reports like this very seriously.',
    repeat: 'I am sorry you had to contact us more than once about this. I am taking it over personally until it is solved.',
    sorry: 'I am sorry for the trouble this has caused.',
    topics: {
      billing: 'I checked the charge[ on order {order}]. If you were billed more than once or for the wrong amount, the difference goes back to the original card within 3 to 5 business days, and I will send you the refund confirmation as soon as it is issued.',
      delivery: 'I looked up the shipment[ for order {order}] and opened a trace with the carrier. If the package has not moved by tomorrow, we will ship a replacement at no cost, without waiting for the trace to close.',
      defect: 'That is not how it should work, and I am sorry it happened. I can send a replacement right away or refund you in full[ for order {order}], whichever you prefer. A prepaid return label is on the way, and you do not need to send the item back before the replacement ships.',
      returns: 'I have started the return[ for order {order}] and a prepaid label will reach your inbox shortly. Once the carrier scans the parcel, the refund or exchange is processed within 3 business days.',
      account: 'I have sent a new sign-in link to the email address on your account. It is valid for 30 minutes. If it does not arrive, check the spam folder or reply here and I will update the address for you.',
      website: 'Thank you for the details. I have passed them to our web team with your steps to reproduce it. In the meantime, the order can be placed by phone or from a different browser, and I am happy to place it for you.',
      cancel: 'I have cancelled [order {order}|it] as you asked. If anything was already charged, the refund goes back to the original payment method within 3 to 5 business days.',
      praise: 'Thank you for taking the time to write this. I have shared it with the team, and it made their day.',
      request: 'Thanks for the suggestion. I have logged it with the product team, who review requests like this every month.',
      other: 'Thank you for getting in touch. I have read your message and I am looking into it now.',
    },
    promise: {
      critical: 'A team lead will contact you today.',
      high: 'You will hear back from us within 24 hours.',
      medium: 'We will get back to you within two business days.',
      low: '',
    },
    publicOpening: 'Hi[ {name}], we are sorry to read this.',
    publicSafety: 'Please stop using the product and unplug it.',
    publicAsk: 'We want to put it right: please reply to our order confirmation email with your order number so we can follow up directly.',
    publicPraise: 'Thank you[, {name}]! We are glad it worked out, and we have passed your words on to the team.',
    publicSignoff: '— {team}',
  };

  function fillTemplate(text, vars) {
    const filled = (s) => !(s.match(/\{(\w+)\}/g) || []).some((p) => !vars[p.slice(1, -1)]);
    const out = String(text == null ? '' : text).replace(/\[([^[\]]*)\]/g, (_, inner) => {
      const bar = inner.indexOf('|');
      const first = bar < 0 ? inner : inner.slice(0, bar);
      const other = bar < 0 ? '' : inner.slice(bar + 1);
      return filled(first) ? first : other;
    });
    return out.replace(/\{(\w+)\}/g, (m, key) => (key in vars ? String(vars[key] || '') : m));
  }

  // user templates may be partial or come from an older file: anything missing falls back to the default
  function mergeTemplates(custom) {
    const c = custom && typeof custom === 'object' ? custom : {};
    const out = {};
    for (const [k, v] of Object.entries(DEFAULT_TEMPLATES)) {
      if (v && typeof v === 'object') {
        out[k] = {};
        for (const kk of Object.keys(v)) out[k][kk] = c[k] && typeof c[k][kk] === 'string' ? c[k][kk] : v[kk];
      } else {
        out[k] = typeof c[k] === 'string' ? c[k] : v;
      }
    }
    return out;
  }

  function suggestReply(item, analysis, opts) {
    const t = mergeTemplates(opts && opts.templates);
    const team = (opts && opts.team) || 'Customer Care';
    const vars = { name: firstName(item.customer), customer: String(item.customer || '').trim(), order: orderNumber(`${item.subject} ${item.message}`) || '', team };
    const f = (text) => fillTemplate(text, vars).trim();
    const level = analysis.level;
    const topic = analysis.topic;
    const promise = f(t.promise[level] || '');
    const safety = analysis.reasons.includes('safety risk');
    if (isPublic(item)) {
      if (topic === 'praise') return `${f(t.publicPraise)}\n${f(t.publicSignoff)}`;
      const parts = [f(t.publicOpening), safety ? f(t.publicSafety) : '', f(t.publicAsk), promise].filter(Boolean);
      return parts.join(' ') + `\n${f(t.publicSignoff)}`;
    }
    const lines = [];
    if (safety) lines.push(f(t.safety));
    else if (level === 'critical' || level === 'high') {
      lines.push(analysis.reasons.includes('repeat contact') || analysis.reasons.some((r) => /message from this customer/.test(r)) ? f(t.repeat) : f(t.sorry));
    }
    lines.push(f(t.topics[topic] != null ? t.topics[topic] : t.topics.other));
    lines.push(promise);
    return `${f(t.greeting)}\n\n${lines.filter(Boolean).join(' ')}\n\n${f(t.signoff)}`;
  }

  // ---------------------------------------------------------------- triage

  function triage(items, options) {
    const opts = options || {};
    const asOf = opts.asOf || items.reduce((m, i) => (i.date && i.date > m ? i.date : m), '') || new Date().toISOString().slice(0, 10);
    // count messages per customer in date order, so the 2nd message is flagged and the 1st is not
    const seen = {};
    const ordered = items.slice().sort((a, b) => ((a.date || asOf) < (b.date || asOf) ? -1 : (a.date || asOf) > (b.date || asOf) ? 1 : 0));
    const nth = new Map();
    for (const it of ordered) {
      const k = contactKey(it);
      seen[k] = (seen[k] || 0) + 1;
      nth.set(it, seen[k]);
    }
    // a message that says nothing about its subject ("following up...") takes the topic of the customer's earlier one
    const topicOf = new Map(items.map((it) => [it, classify(it)]));
    const lastTopic = {};
    for (const it of ordered) {
      const k = contactKey(it);
      const c = topicOf.get(it);
      if (c.topic === OTHER.key && lastTopic[k]) topicOf.set(it, Object.assign({}, lastTopic[k], { followUp: true }));
      else if (c.topic !== OTHER.key) lastTopic[k] = c;
    }
    const out = items.map((item) => {
      const c = topicOf.get(item);
      const u = urgency(item, { asOf, topic: c.topic, contacts: { [contactKey(item)]: nth.get(item) } });
      const analysis = { topic: c.topic, topicLabel: c.label, words: c.words, followUp: !!c.followUp, level: u.level, levelLabel: u.levelLabel, score: u.score, reasons: u.reasons, waiting: u.waiting };
      return Object.assign({}, item, analysis, { reply: suggestReply(item, analysis, opts), public: isPublic(item) });
    });
    const rank = Object.fromEntries(LEVELS.map((l, i) => [l.key, i]));
    out.sort((a, b) => rank[a.level] - rank[b.level] || b.score - a.score || b.waiting - a.waiting);
    const topics = TOPICS.concat(OTHER).map((t) => {
      const these = out.filter((i) => i.topic === t.key);
      return { key: t.key, label: t.label, total: these.length, levels: Object.fromEntries(LEVELS.map((l) => [l.key, these.filter((i) => i.level === l.key).length])) };
    }).filter((t) => t.total).sort((a, b) => b.total - a.total);
    const rated = out.filter((i) => i.rating != null);
    const summary = {
      total: out.length,
      levels: Object.fromEntries(LEVELS.map((l) => [l.key, out.filter((i) => i.level === l.key).length])),
      avgRating: rated.length ? rated.reduce((s, i) => s + i.rating, 0) / rated.length : null,
      reviews: rated.length,
      oldest: out.reduce((m, i) => Math.max(m, i.waiting), 0),
    };
    return { asOf, items: out, topics, summary };
  }

  function toCSV(rows, columns) {
    const esc = (v) => {
      const s = v == null ? '' : String(v);
      return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return [columns.map((c) => esc(c.label)).join(',')].concat(rows.map((r) => columns.map((c) => esc(c.value(r))).join(','))).join('\n') + '\n';
  }

  return { parseCSV, readXLSX, readItems, parseDate, daysBetween, classify, urgency, suggestReply, triage, toCSV, orderNumber, TOPICS, LEVELS, OTHER, DEFAULT_TEMPLATES, fillTemplate, mergeTemplates };
});
