// Core raid kit logic extracted for unit testing and reuse

function escapeRe(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

const DEFAULT_RAID_KIT = [
  { name: 'Vial of Velium Vapors', mode: 'present', pattern: '^Vial of Velium Vapors$' },
  { name: 'Leatherfoot Raider Skullcap', mode: 'present', pattern: '^Leatherfoot Raider Skullcap$' },
  { name: 'Shiny Brass Idol',      mode: 'present', pattern: '^Shiny Brass Idol$' },
  { name: 'Ring of Shadows',       mode: 'count',   pattern: '^Ring of Shadows$' },
  { name: 'Reaper of the Dead',    mode: 'present', pattern: '^Reaper of the Dead$' },
  { name: 'Pearl',                 mode: 'count',   pattern: '^Pearl$' },
  { name: 'Peridot',               mode: 'count',   pattern: '^Peridot$' },
  { name: 'Mana Battery - Class Five', mode: 'count', pattern: '^Mana Battery - Class Five$' },
  { name: 'Mana Battery - Class Four', mode: 'count', pattern: '^Mana Battery - Class Four$' },
  { name: 'Mana Battery - Class Three', mode: 'count', pattern: '^Mana Battery - Class Three$' },
  { name: 'Mana Battery - Class Two', mode: 'count', pattern: '^Mana Battery - Class Two$' },
  { name: 'Mana Battery - Class One', mode: 'count', pattern: '^Mana Battery - Class One$' },
  { name: '10 Dose Potion of Stinging Wort', mode: 'count', pattern: '^10 Dose Potion of Stinging Wort$' },
  { name: 'Pegasus Feather Cloak', mode: 'present', pattern: '^Pegasus Feather Cloak$' },
  { name: "Larrikan's Mask",      mode: 'present', pattern: "^Larrikan'?s Mask$" }
];

const FIXED_RK_NAMES = new Set([
  'Vial of Velium Vapors','Leatherfoot Raider Skullcap','Shiny Brass Idol',
  'Ring of Shadows','Reaper of the Dead','Pearl','Peridot',
  'Mana Battery - Class Five','Mana Battery - Class Four','Mana Battery - Class Three','Mana Battery - Class Two','Mana Battery - Class One',
  '10 Dose Potion of Stinging Wort','Pegasus Feather Cloak',
  "Larrikan's Mask"
]);

function normalizeUserItems(userItems){
  const list = Array.isArray(userItems) ? userItems : [];
  // Dedup by name; last write wins
  const map = new Map();
  for (const it of list){
    if (!it || !it.name) continue;
    const name = String(it.name);
    const mode = (it.mode === 'count') ? 'count' : 'present';
    const pattern = it.pattern || ('^' + escapeRe(name) + '$');
    map.set(name, { name, mode, pattern });
  }
  return Array.from(map.values());
}

function getMergedRaidKit(settings){
  const hiddenSet = new Set(((settings && settings.raidKitHidden) || []).map(String));
  const merged = [];
  for (const d of DEFAULT_RAID_KIT){ if (!hiddenSet.has(d.name)) merged.push({ ...d }); }
  for (const u of normalizeUserItems((settings && settings.raidKitItems) || [])){
    merged.push(u);
  }
  return merged;
}

function getRaidKitSummary(items){
  const list = Array.isArray(items) ? items : [];
  const count = (namePattern) => {
    const re = new RegExp(namePattern, 'i');
    let n = 0; for (const it of list){ if (re.test(it.Name || '')) n += Number(it.Count || 0) || 0; }
    return n;
  };
  const has = (namePattern) => { const re = new RegExp(namePattern, 'i'); return list.some(it => re.test(it.Name || '')); };
  return {
    vialVeliumVapors: has('^Vial of Velium Vapors$') ? 'Y' : 'N',
    leatherfootSkullcap: has('^Leatherfoot Raider Skullcap$') ? 'Y' : 'N',
    shinyBrassIdol: has('^Shiny Brass Idol$') ? 'Y' : 'N',
    ringOfShadowsCount: count('^Ring of Shadows$'),
    reaperOfTheDead: has('^Reaper of the Dead$') ? 'Y' : 'N',
    pearlCount: count('^Pearl$'),
    peridotCount: count('^Peridot$'),
    mbClassFive: count('^Mana Battery - Class Five$'),
    mbClassFour: count('^Mana Battery - Class Four$'),
    mbClassThree: count('^Mana Battery - Class Three$'),
    mbClassTwo: count('^Mana Battery - Class Two$'),
    mbClassOne: count('^Mana Battery - Class One$'),
    tenDosePotionOfStingingWortCount: count('^10 Dose Potion of Stinging Wort$'),
    pegasusFeatherCloak: has('^Pegasus Feather Cloak$') ? 'Y' : 'N',
    larrikansMask: has("^Larrikan'?s Mask$") ? 'Y' : 'N'
  };
}

function countRaidKitForInventory(settings, inventory){
  const items = Array.isArray(inventory) ? inventory : [];
  const kit = getMergedRaidKit(settings);
  const out = [];
  for (const k of kit){
    const re = new RegExp(k.pattern || ('^' + escapeRe(k.name) + '$'), 'i');
    let present = false, count = 0;
    for (const it of items){ if (re.test(String(it.Name||''))){ present = true; count += Number(it.Count||0); } }
    out.push({ name: k.name, mode: k.mode, present, count });
  }
  return out;
}

module.exports = {
  DEFAULT_RAID_KIT,
  FIXED_RK_NAMES,
  getMergedRaidKit,
  getRaidKitSummary,
  countRaidKitForInventory,
  normalizeUserItems
};
