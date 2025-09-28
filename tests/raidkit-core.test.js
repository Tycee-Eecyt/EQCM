const assert = require('assert');
const rk = require('../src/raidkit-core');

function test(name, fn){
  try { fn(); console.log('ok - ' + name); }
  catch(e){ console.error('not ok - ' + name + ' -> ' + (e && e.message || e)); process.exitCode = 1; }
}

test('merge includes defaults by default', () => {
  const merged = rk.getMergedRaidKit({ raidKitItems: [], raidKitHidden: [] });
  assert(merged.find(x => x.name === 'Vial of Velium Vapors'));
});

test('hide default removes from merged', () => {
  const merged = rk.getMergedRaidKit({ raidKitItems: [], raidKitHidden: ['Vial of Velium Vapors'] });
  assert(!merged.find(x => x.name === 'Vial of Velium Vapors'));
});

test('add custom present item shows up in merged', () => {
  const merged = rk.getMergedRaidKit({ raidKitItems: [{ name:'Custom Item', mode:'present' }], raidKitHidden: [] });
  assert(merged.find(x => x.name === 'Custom Item'));
});

test('add custom count item has regex pattern and mode count', () => {
  const merged = rk.getMergedRaidKit({ raidKitItems: [{ name:'Foo Bar', mode:'count' }], raidKitHidden: [] });
  const it = merged.find(x => x.name === 'Foo Bar');
  assert(it && it.mode === 'count');
  assert(new RegExp(it.pattern, 'i').test('Foo Bar'));
});

test('summary maps new fixed columns and removed Velium Vial', () => {
  const items = [
    { Name:'10 Dose Potion of Stinging Wort', Count: 3 },
    { Name:'Pegasus Feather Cloak', Count: 1 }
  ];
  const s = rk.getRaidKitSummary(items);
  assert.strictEqual(s.tenDosePotionOfStingingWortCount, 3);
  assert.strictEqual(s.pegasusFeatherCloak, 'Y');
  assert.strictEqual(s.veliumVialCount, undefined);
});

test('inventory count respects regex and sums Count', () => {
  const settings = { raidKitItems: [{ name:'My Gem', mode:'count' }], raidKitHidden: [] };
  const out = rk.countRaidKitForInventory(settings, [
    { Name:'My Gem', Count: 2 },
    { Name:'My Gem', Count: 3 }
  ]);
  const row = out.find(r => r.name === 'My Gem');
  assert(row && row.count === 5 && row.present === true);
});

