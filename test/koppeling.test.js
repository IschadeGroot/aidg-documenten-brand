// node test/koppeling.test.js — rekent de sha256 en de sleutelvorm van het koppelscript na.
const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
let haak = null;
global.window = {
  location: { pathname: '/x/index.html' },
  localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  crypto: { getRandomValues: (a) => crypto.randomFillSync(a) },
  __aidgKoppelingTest: (h) => { haak = h; },
};
global.btoa = (s) => Buffer.from(s, 'binary').toString('base64');
global.document = { readyState: 'complete' };
new Function(fs.readFileSync(path.join(__dirname, '..', 'koppeling', 'aidg-koppeling.js'), 'utf8'))();
assert.ok(haak, 'test-haakje');
for (const t of ['', 'abc', 'aidgo_' + 'x'.repeat(43), 'é€😀', 'a'.repeat(55), 'a'.repeat(56), 'a'.repeat(64), 'a'.repeat(1000)]) {
  assert.strictEqual(haak.sha256hex(t), crypto.createHash('sha256').update(t, 'utf8').digest('hex'), JSON.stringify(t).slice(0, 20));
}
for (let i = 0; i < 50; i++) {
  const s = haak.nieuweSleutel();
  assert.match(s, /^aidgo_[A-Za-z0-9_-]{43}$/);
  assert.strictEqual(haak.sha256hex(s), crypto.createHash('sha256').update(s).digest('hex'));
}
console.log('koppeling ok');
