const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const T = require('../site/triage.js');

const inflate = async (bytes) => new Uint8Array(zlib.inflateRawSync(bytes));
const item = (o) => Object.assign({ id: '1', date: '2026-11-05', channel: 'Email', customer: 'Ana Silva', rating: null, subject: '', message: '' }, o);

// Builds a small .xlsx in memory: deflated entries, central directory, end record.
function makeZip(entries) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of Object.entries(entries)) {
    const raw = Buffer.from(text, 'utf8');
    const data = zlib.deflateRawSync(raw);
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = zlib.crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    const dir = Buffer.alloc(46);
    dir.writeUInt32LE(0x02014b50, 0);
    dir.writeUInt16LE(20, 4);
    dir.writeUInt16LE(20, 6);
    dir.writeUInt16LE(8, 10);
    dir.writeUInt32LE(crc, 16);
    dir.writeUInt32LE(data.length, 20);
    dir.writeUInt32LE(raw.length, 24);
    dir.writeUInt16LE(nameBuf.length, 28);
    dir.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, data);
    central.push(dir, nameBuf);
    offset += local.length + nameBuf.length + data.length;
  }
  const dirBuf = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(central.length / 2, 8);
  end.writeUInt16LE(central.length / 2, 10);
  end.writeUInt32LE(dirBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, dirBuf, end]);
}

test('parseCSV reads quoted fields and semicolon files', () => {
  const rows = T.parseCSV('id;message\n1;"Hello; ""there""\nsecond line"\n');
  assert.deepEqual(rows, [{ id: '1', message: 'Hello; "there"\nsecond line' }]);
});

test('readXLSX reads shared strings, inline strings and Excel dates', async () => {
  const zip = makeZip({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Tickets" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>date</t></si><si><t>message</t></si><si><r><t>Box </t></r><r><t>arrived &amp; broken</t></r></si></sst>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>rating</t></is></c></row>' +
      '<row r="2"><c r="A2"><v>46331</v></c><c r="B2" t="s"><v>2</v></c><c r="C2"><v>2</v></c></row>' +
      '</sheetData></worksheet>',
  });
  const rows = await T.readXLSX(zip, inflate);
  assert.deepEqual(rows, [{ date: '46331', message: 'Box arrived & broken', rating: '2' }]);
  const { items } = T.readItems(rows);
  assert.equal(items[0].date, '2026-11-05');
  assert.equal(items[0].rating, 2);
});

test('the sample .xlsx and .csv give the same rows', async () => {
  const dir = path.join(__dirname, '..', 'data');
  const fromCsv = T.parseCSV(fs.readFileSync(path.join(dir, 'messages.csv'), 'utf8'));
  const fromXlsx = await T.readXLSX(fs.readFileSync(path.join(dir, 'messages.xlsx')), inflate);
  assert.equal(fromXlsx.length, fromCsv.length);
  assert.deepEqual(T.readItems(fromXlsx).items, T.readItems(fromCsv).items);
});

test('readItems maps other column names and reports bad rows', () => {
  const { items, errors } = T.readItems([
    { Ticket: 'A-1', Created: '03/11/2026', From: 'Ben Ito', Stars: '4,0', Body: 'Thanks!' },
    { Ticket: 'A-2', Created: '', From: '', Stars: '', Body: '' },
  ]);
  assert.equal(items.length, 1);
  assert.deepEqual([items[0].id, items[0].date, items[0].customer, items[0].rating], ['A-1', '2026-11-03', 'Ben Ito', 4]);
  assert.match(errors[0], /row 3: no message/);
});

test('classify picks the topic with the strongest words', () => {
  assert.equal(T.classify(item({ subject: 'Charged twice', message: 'I was billed twice for one order.' })).topic, 'billing');
  assert.equal(T.classify(item({ message: 'The kettle started to smoke.' })).topic, 'defect');
  assert.equal(T.classify(item({ message: 'Tracking has not updated and the parcel is late.' })).topic, 'delivery');
  assert.equal(T.classify(item({ message: 'Store opening hours on Sunday?' })).topic, 'other');
});

test('praise does not count when the message also complains', () => {
  assert.equal(T.classify(item({ rating: 5, message: 'Great service, thank you!' })).topic, 'praise');
  assert.notEqual(T.classify(item({ message: 'Great, the box arrived damaged. Not happy.' })).topic, 'praise');
});

test('a safety word makes a message critical whatever the score', () => {
  const u = T.urgency(item({ message: 'It sparks when I plug it in.' }), { asOf: '2026-11-05', topic: 'defect' });
  assert.equal(u.level, 'critical');
  assert.ok(u.reasons.includes('safety risk'));
});

test('urgency adds up escalation, money, rating, public review and waiting time', () => {
  const u = T.urgency(item({ date: '2026-10-29', channel: 'Store review', rating: 1, message: 'Charged twice. I will file a chargeback.' }),
    { asOf: '2026-11-05', topic: 'billing' });
  assert.deepEqual(u.reasons, ['threat to escalate', 'money taken in error', '1-star rating', 'public review', 'waiting 7 days']);
  assert.equal(u.score, 5 + 40 + 30 + 20 + 10 + 20);
  assert.equal(u.level, 'critical');
});

test('praise stays low', () => {
  const u = T.urgency(item({ rating: 5, message: 'Thanks, arrived today!' }), { asOf: '2026-11-05', topic: 'praise' });
  assert.equal(u.level, 'low');
});

test('triage sorts the most urgent first and flags repeat customers in date order', () => {
  const r = T.triage([
    item({ id: 'a', date: '2026-11-01', customer: 'Joe Park', message: 'Where is my parcel? Tracking is stuck.' }),
    item({ id: 'b', date: '2026-11-04', customer: 'Joe Park', subject: 'Following up', message: 'Any news?' }),
    item({ id: 'c', date: '2026-11-05', customer: 'Kim Lee', rating: 5, message: 'Love it, thank you.' }),
    item({ id: 'd', date: '2026-11-05', customer: 'Max Roy', message: 'The charger melted.' }),
  ], { asOf: '2026-11-05' });
  assert.equal(r.items[0].id, 'd');
  assert.equal(r.items[r.items.length - 1].id, 'c');
  const b = r.items.find((i) => i.id === 'b');
  assert.equal(b.topic, 'delivery');
  assert.equal(b.followUp, true);
  assert.ok(b.reasons.includes('2nd message from this customer'));
  assert.ok(!r.items.find((i) => i.id === 'a').reasons.some((x) => /message from this customer/.test(x)));
  assert.equal(r.summary.total, 4);
  assert.equal(r.topics.reduce((s, t) => s + t.total, 0), 4);
});

test('as-of date defaults to the latest message date', () => {
  const r = T.triage([item({ date: '2026-10-02', message: 'Hello' }), item({ id: '2', date: '2026-10-09', message: 'Hi' })]);
  assert.equal(r.asOf, '2026-10-09');
});

test('replies use the first name, the order number and the team name', () => {
  const [a] = T.triage([item({ customer: 'Grace Okafor', message: 'Order #55120 has not arrived.' })], { asOf: '2026-11-05', team: 'Help Desk' }).items;
  assert.match(a.reply, /^Hi Grace,/);
  assert.match(a.reply, /order 55120/);
  assert.match(a.reply, /Help Desk$/);
});

test('safety replies tell the customer to stop using the product', () => {
  const [a] = T.triage([item({ message: 'The heater started to smoke.' })], { asOf: '2026-11-05' }).items;
  assert.match(a.reply, /stop using the product/);
  assert.match(a.reply, /team lead will contact you today/);
});

test('public reviews get a short public answer', () => {
  const [a] = T.triage([item({ channel: 'Store review', rating: 1, customer: 'Mia Ward', message: 'Arrived broken.' })], { asOf: '2026-11-05' }).items;
  assert.ok(a.public);
  assert.match(a.reply, /^Hi Mia, we are sorry to read this\./);
  assert.ok(!a.reply.includes('Best regards'));
});

test('fillTemplate drops a bracketed part when its placeholder is empty, or uses the fallback', () => {
  assert.equal(T.fillTemplate('Hi[ {name}],', { name: 'Ana' }), 'Hi Ana,');
  assert.equal(T.fillTemplate('Hi[ {name}],', { name: '' }), 'Hi,');
  assert.equal(T.fillTemplate('cancelled [order {order}|it].', { order: '' }), 'cancelled it.');
  assert.equal(T.fillTemplate('cancelled [order {order}|it].', { order: 'A-1234' }), 'cancelled order A-1234.');
  assert.equal(T.fillTemplate('Hello {nmae}', { name: 'Ana' }), 'Hello {nmae}');
});

test('custom templates change the reply and missing ones fall back to the default', () => {
  const templates = { topics: { delivery: 'Your parcel[ ({order})] is on its way.' }, promise: { high: '' }, signoff: 'Thanks,\n{team}' };
  const merged = T.mergeTemplates(templates);
  assert.equal(merged.topics.billing, T.DEFAULT_TEMPLATES.topics.billing);
  assert.equal(merged.greeting, T.DEFAULT_TEMPLATES.greeting);
  const [a] = T.triage([item({ customer: 'Grace Okafor', message: 'Order #55120 has not arrived.' })], { asOf: '2026-11-05', team: 'Help Desk', templates }).items;
  assert.equal(a.reply, 'Hi Grace,\n\nYour parcel (55120) is on its way.\n\nThanks,\nHelp Desk');
  assert.deepEqual(T.mergeTemplates({ topics: 'bad', greeting: 42 }), T.mergeTemplates(null));
});

test('toCSV quotes fields with commas, quotes and line breaks', () => {
  const csv = T.toCSV([{ a: 'x, "y"', b: 'one\ntwo' }], [{ label: 'a', value: (r) => r.a }, { label: 'b', value: (r) => r.b }]);
  assert.equal(csv, 'a,b\n"x, ""y""","one\ntwo"\n');
});
