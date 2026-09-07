const { sheetRows, check } = require('./harness');

function rowsToObjects(name) {
  const rows = sheetRows(name);
  const headers = rows[0];
  return rows.slice(1).filter(r => r.some(v => v !== null)).map(r => {
    const o = {};
    headers.forEach((h, i) => { o[h] = r[i] === undefined ? null : r[i]; });
    return o;
  });
}

// TAGLIST_DATABASE
{
  const objs = rowsToObjects('taglist');
  const gen = {};
  objs.forEach(o => { gen[o.id] = o.info; });
  check('TAGLIST_DATABASE', gen);
}

// TALENTLIST_DATABASE
{
  const objs = rowsToObjects('talentlist');
  console.log('talentlist rows sample:', JSON.stringify(objs.slice(0,3)));
}

// TALENTLIST_DATABASE
{
  const objs = rowsToObjects('talentlist');
  const gen = {};
  objs.forEach(o => { gen[o.group] = { info: o.info, unlockJob: o.unlock_job }; });
  check('TALENTLIST_DATABASE', gen);
}

// CHARACTER_STAT_DATABASE
{
  const objs = rowsToObjects('character');
  const fieldMap = { HP:'hp', ATK:'atk', MATK:'matk', DEF:'def', MDEF:'mdef', CRI:'cri', MP:'mp',
    cri_resist:'criResist', AGI:'agi', damage_reflect:'damageReflect', cri_damage:'criDamage',
    cri_damage_resist:'criDamageResist', skill_rate:'skillRate', skill_resist:'skillResist',
    damage_increase:'damageIncrease', damage_reduce:'damageReduce', mp_recove:'mpRecover',
    shield_rate:'shieldRate', permeate:'permeate', atypical_rate:'atypicalRate', heal_rate:'healRate' };
  const gen = {};
  objs.forEach(o => {
    const entry = { limit: o.limite };
    Object.keys(fieldMap).forEach(col => { if (o[col] !== null && o[col] !== undefined) entry[fieldMap[col]] = o[col]; });
    gen[o.tpye] = entry;
  });
  check('CHARACTER_STAT_DATABASE', gen);
}

// NATURE_DATABASE
{
  const objs = rowsToObjects('nature');
  const statCols = ['int','vit','agi','str','dex','spr'];
  const gen = objs.map(o => {
    const requires = {};
    statCols.forEach(c => { if (o[c] !== null && o[c] !== undefined) requires[c] = o[c]; });
    return { name: o.nature_info, requires, buffId: o.buff_id };
  });
  check('NATURE_DATABASE', gen);
}

// helper: excel serial date -> "YYYY-MM-DD HH:mm"
function excelDateToStr(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60; total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60*60));
  const minutes = Math.floor(total_seconds / 60) % 60;
  const pad = n => String(n).padStart(2,'0');
  return `${date_info.getUTCFullYear()}-${pad(date_info.getUTCMonth()+1)}-${pad(date_info.getUTCDate())} ${pad(hours)}:${pad(minutes)}`;
}

// MAPLIST_DATABASE
{
  const objs = rowsToObjects('maplist');
  const keyByInfo = { '神格試煉':'god_trial', '試煉之塔':'trial_tower', '無限地城':'infinite_dungeon', '殲滅任務':'annihilation_quest' };
  const gen = objs.map(o => ({
    key: keyByInfo[o.info],
    type: o.type === null || o.type === undefined ? null : o.type,
    info: o.info,
    beginTime: excelDateToStr(o.begin_time),
    endTime: excelDateToStr(o.end_time),
  }));
  check('MAPLIST_DATABASE', gen);
}

// START_GIFT_DATABASE
{
  const objs = rowsToObjects('start_gift');
  const gen = objs.map(o => ({ id: o.id, amount: o.id_amount }));
  check('START_GIFT_DATABASE', gen);
}

// generic: find header row (first row where more than half cells non-null and looks like header)
function rowsToObjectsSkipBlank(name) {
  const rows = sheetRows(name);
  let headerIdx = rows.findIndex(r => r.some(v => v !== null));
  const headers = rows[headerIdx];
  return rows.slice(headerIdx+1).filter(r => r.some(v => v !== null)).map(r => {
    const o = {};
    headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
    return o;
  });
}

// COST_DATABASE
{
  const objs = rowsToObjectsSkipBlank('cost').filter(o => typeof o.id === 'number');
  const gen = {};
  objs.forEach(o => {
    if (!gen[o.id]) gen[o.id] = {};
    gen[o.id][o.level] = {
      item1Id: o.item1_id,
      item1Amount: o.item1_amount,
      item2Id: o.item2_id === null || o.item2_id === undefined ? null : o.item2_id,
      item2Amount: o.item2_amount === null || o.item2_amount === undefined ? 0 : o.item2_amount,
      rate: o.rate,
    };
  });
  check('COST_DATABASE', gen);
}


// REFINE_DATABASE
{
  const statFieldMap = { HP:'hp', ATK:'atk', MATK:'matk', DEF:'def', MDEF:'mdef', MP:'mp', AGI:'agi', CRI:'cri',
    cri_resist:'criResist', cri_damage:'criDamage', cri_damage_resist:'criDamageResist',
    skill_rate:'skillRate', skill_resist:'skillResist', damage_increase:'damageIncrease',
    damage_reduce:'damageReduce', damage_reflect:'damageReflect' };
  const objs = rowsToObjectsSkipBlank('refine').filter(o => typeof o.id === 'number');
  const gen = {};
  objs.forEach(o => {
    if (!gen[o.id]) gen[o.id] = {};
    const entry = {};
    Object.keys(statFieldMap).forEach(col => { if (o[col] !== null && o[col] !== undefined) entry[statFieldMap[col]] = o[col]; });
    gen[o.id][o.level] = entry;
  });
  check('REFINE_DATABASE', gen);
}

// ITEM_DATABASE (icon not in xlsx -> keep as hardcoded override map, sourced from current data)
{
  const { currentValue } = require('./harness');
  const cur = currentValue('ITEM_DATABASE');
  const ICON_MAP = {};
  Object.keys(cur).forEach(id => { ICON_MAP[id] = cur[id].icon; });
  // print for embedding into loader later
  require('fs').writeFileSync('/tmp/item_icon_map.json', JSON.stringify(ICON_MAP, null, 2));

  const objs = rowsToObjects('item'); // header row 0: id, 名稱, usage
  const gen = {};
  objs.forEach(o => {
    gen[o.id] = { id: o.id, name: o['名稱'], usage: o.usage, icon: ICON_MAP[o.id] || '📦' };
  });
  check('ITEM_DATABASE', gen);
}

// LEVELCURVE_DATABASE
{
  const objs = rowsToObjects('levelcurve');
  const gen = objs.map(o => ({ level: o.level, next: o.next === null ? 0 : o.next, cumExp: o['累计获得经验'] }));
  check('LEVELCURVE_DATABASE', gen);
}

// TALENT_CURVE
{
  const objs = rowsToObjects('curve').filter(o => o.type === 1);
  const gen = objs.map(o => ({ level: o.level, point: o.point, total: o['#總點數'] }));
  check('TALENT_CURVE', gen);
}

// IDLEZONE_DATABASE
{
  const objs = rowsToObjects('idlezone');
  const splitIds = s => String(s).split(',').map(x => parseInt(x.trim(), 10));
  const gen = {};
  objs.forEach(o => {
    const rewards = [];
    for (let n = 1; n <= 4; n++) {
      const idKey = `idlereward${n}_id`, amtKey = `idlereward${n}_amount`, rateKey = `idlereward${n}_rate`;
      if (o[idKey] !== null && o[idKey] !== undefined) {
        rewards.push({ ids: splitIds(o[idKey]), amount: o[amtKey], rate: o[rateKey] });
      }
    }
    gen[o.mapid] = { mapid: o.mapid, randomPlay: splitIds(o.random_play), level: o.level, rewards };
  });
  check('IDLEZONE_DATABASE', gen);
}

// GACHALIST_DATABASE
{
  function excelDateToStrSec(serial) {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60; total_seconds -= seconds;
    const hours = Math.floor(total_seconds / (60*60));
    const minutes = Math.floor(total_seconds / 60) % 60;
    const pad = n => String(n).padStart(2,'0');
    return `${date_info.getUTCFullYear()}-${pad(date_info.getUTCMonth()+1)}-${pad(date_info.getUTCDate())} ${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
  }
  const objs = rowsToObjects('gachalist');
  const gen = objs.map(o => ({
    id: o.id, info: o.info,
    cost1Id: o.cost1_id, cost1Amount: o.cost1_amount,
    cost10Id: o.cost10_id, cost10Amount: o.cost10_amount,
    beginTime: excelDateToStrSec(o.begin_time), endTime: excelDateToStrSec(o.end_time),
  }));
  check('GACHALIST_DATABASE', gen);
}

// DUNGEON_LEVEL_DATABASE
{
  const objs = rowsToObjects('dungeon');
  const gen = {};
  objs.forEach(o => {
    gen[o.level] = { level: o.level, rateAdd: o.rate_add, guarantMin: o.guarant_min, guarantMax: o.guarant_max, buffMax: o.buff_max };
  });
  check('DUNGEON_LEVEL_DATABASE', gen);
}

// DUNGEONPREFEB_DATABASE
{
  const objs = rowsToObjects('dungeonprefeb');
  const gen = objs.map(o => ({
    prefebMap: o.prefeb_map,
    slots: String(o.list).split(',').map(s => s.trim().replace(/^\{/, '').replace(/\}$/, '')),
    power: (o.power === null || o.power === undefined) ? null : o.power,
    isEnd: !!o.is_end,
  }));
  check('DUNGEONPREFEB_DATABASE', gen);
}

// TALENT_DATABASE + TALENT_GROUP_GRID
{
  const objs = rowsToObjects('talent');
  const parseFront = v => {
    if (v === null || v === undefined) return [];
    return String(v).split(',').map(s => parseInt(s.trim(), 10));
  };
  const genTalent = {};
  const genGrid = {};
  objs.forEach(o => {
    const entry = { id: o.id, group: o.group, name: o.group_info, front: parseFront(o.front), cost: o.cost, buffPasiveId: o.buff_pasive };
    if (o.role_card_unlock !== null && o.role_card_unlock !== undefined) entry.roleCardUnlock = o.role_card_unlock;
    genTalent[o.id] = entry;
    if (o.display_position !== null && o.display_position !== undefined) {
      const grid = String(o.display_position).split('\n').map(line =>
        line.split(',').map(cell => parseInt(cell.trim().replace(/^\{/, '').replace(/\}$/, ''), 10))
      );
      genGrid[o.group] = grid;
    }
  });
  check('TALENT_DATABASE', genTalent);
  check('TALENT_GROUP_GRID', genGrid);
}

// BUFF_DATABASE + BUFF_SET_TIERS
{
  const { currentValue } = require('./harness');
  const curBuff = currentValue('BUFF_DATABASE');
  const ICON_MAP = {};
  Object.keys(curBuff).forEach(id => { ICON_MAP[id] = { icon: curBuff[id].icon, isUp: curBuff[id].isUp }; });
  require('fs').writeFileSync('/tmp/buff_icon_map.json', JSON.stringify(ICON_MAP));

  const BUFF_STAT_FIELD_MAP = {
    ATK:'atk', MATK:'matk', DEF:'def', MDEF:'mdef', CRI:'cri', MP:'mp',
    cri_resist:'criResist', AGI:'agi', damage_reflect:'damageReflect', cri_damage:'criDamage',
    cri_damage_resist:'criDamageResist', skill_rate:'skillRate', skill_resist:'skillResist',
    damage_increase:'damageIncrease', damage_reduce:'damageReduce',
  };
  function statFieldsFrom(o) {
    const entry = {};
    Object.keys(BUFF_STAT_FIELD_MAP).forEach(col => { if (o[col] !== null && o[col] !== undefined) entry[BUFF_STAT_FIELD_MAP[col]] = o[col]; });
    return entry;
  }

  const rows = rowsToObjects('buff').filter(o => typeof o.buff_id === 'number');
  const byId = {};
  rows.forEach(o => { (byId[o.buff_id] = byId[o.buff_id] || []).push(o); });

  const genBuff = {};
  const genSetTiers = {};
  Object.entries(byId).forEach(([idStr, rs]) => {
    const id = Number(idStr);
    const iconInfo = ICON_MAP[id] || { icon: '✨', isUp: true };
    const plainRows = rs.filter(r => r.set_effective_count === null || r.set_effective_count === undefined);
    if (plainRows.length > 0) {
      const o = plainRows[0];
      const rawName = o['buff名稱'] === null || o['buff名稱'] === undefined ? '' : o['buff名稱'];
      const name = rawName.startsWith('#') ? rawName.slice(1) : rawName;
      const entry = { id, name, icon: iconInfo.icon, isUp: iconInfo.isUp,
        reflashAble: !!o.reflashAble, maxStack: o.max_stack === null || o.max_stack === undefined ? 1 : o.max_stack,
        duration: o['持續回合數'] === null || o['持續回合數'] === undefined ? 1 : o['持續回合數'], ...statFieldsFrom(o) };
      if (o.invisible) entry.invisible = true;
      genBuff[id] = entry;
    } else {
      // set-tier-only id: base entry = lowest effectiveCount row
      const tierRows = [...rs].sort((a, b) => a.set_effective_count - b.set_effective_count);
      const base = tierRows[0];
      genBuff[id] = { id, name: base['buff名稱'], icon: iconInfo.icon, isSetBase: true, setInfo: base.set_info,
        ...statFieldsFrom(base), invisible: true, reflashAble: !!base.reflashAble, maxStack: base.max_stack, duration: base['持續回合數'] };
      genSetTiers[id] = {
        setInfo: base.set_info,
        tiers: tierRows.map(o => ({ effectiveCount: o.set_effective_count, name: o['buff名稱'], ...statFieldsFrom(o) })),
      };
    }
  });
  check('BUFF_DATABASE', genBuff);
  check('BUFF_SET_TIERS', genSetTiers);
}

// SKILL_DATABASE
{
  function rowsToObjectsAuto(name) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    if (headerIdx === -1) return [];
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    });
  }

  const SKILL_TYPE_MAP = { phycial: 'AD', magic: 'AP', heal: 'HEAL', shield: 'SHIELD' };
  const SKILL_ADD_BUFF_ID_OVERRIDE = { 400018: [4001] }; // xlsx data typo "4001.4002" (decimal, should be comma list); original game data only used 4001

  function parseIdList(v) {
    if (v === null || v === undefined) return undefined;
    if (typeof v === 'number') return [v];
    return String(v).split(',').map(s => Number(s.trim()));
  }

  const objs = rowsToObjectsAuto('skill').filter(o => typeof o.skill_id === 'number');
  const gen = {};
  objs.forEach(o => {
    const id = o.skill_id;
    const entry = {
      id,
      name: o['技能名稱'],
      ad: o['AD(技能係數)'] === null || o['AD(技能係數)'] === undefined ? 0 : o['AD(技能係數)'],
      ap: o['AP(技能係數)'] === null || o['AP(技能係數)'] === undefined ? 0 : o['AP(技能係數)'],
      cost: o.MP_cost === null || o.MP_cost === undefined ? 0 : Math.round(o.MP_cost * 1e6) / 1e6,
      type: SKILL_TYPE_MAP[o.type] || 'BUFF',
    };
    const addBuffId = id === 400018 ? SKILL_ADD_BUFF_ID_OVERRIDE[id] : parseIdList(o.add_buff_id);
    if (addBuffId !== undefined) entry.add_buff_id = addBuffId;
    if (o.add_buff_rate_overwrite !== null && o.add_buff_rate_overwrite !== undefined) entry.add_buff_rate_overwrite = o.add_buff_rate_overwrite;
    const selfAddBuffId = parseIdList(o.self_add_buff_id);
    if (selfAddBuffId !== undefined) entry.self_add_buff_id = selfAddBuffId;
    entry.target = `0,${o['攻擊對象數量']},${o['是否為隨機'] ? 1 : 0},${o['是否自身施放'] ? 1 : 0}`;
    if (o.prefeb_atk !== null && o.prefeb_atk !== undefined) entry.prefebAtk = o.prefeb_atk;
    const prefebHit1 = parseIdList; // not used
    if (o.prefeb_hit1 !== null && o.prefeb_hit1 !== undefined) entry.prefebHit1 = [o.prefeb_hit1];
    if (o.prefeb_hit2 !== null && o.prefeb_hit2 !== undefined) entry.prefebHit2 = [o.prefeb_hit2];
    const removeBuffId = parseIdList(o.remove_buff_id);
    if (removeBuffId !== undefined) entry.remove_buff_id = removeBuffId;
    if (o.remove_buff_rate_overwrite !== null && o.remove_buff_rate_overwrite !== undefined) entry.remove_buff_rate_overwrite = o.remove_buff_rate_overwrite;
    if (o.self_buff_rate_overwrite !== null && o.self_buff_rate_overwrite !== undefined) entry.self_buff_rate_overwrite = o.self_buff_rate_overwrite;
    entry.desc = o['說明'];
    if (o.MP_revocer !== null && o.MP_revocer !== undefined) entry.mp_recover = o.MP_revocer;
    if (o.cooldown !== null && o.cooldown !== undefined) entry.cooldown = o.cooldown;
    if (o['攻擊次數'] !== null && o['攻擊次數'] !== undefined && o['攻擊次數'] !== 1) entry.hits = o['攻擊次數'];
    if (o.current_hp !== null && o.current_hp !== undefined) entry.current_hp = o.current_hp;
    gen[id] = entry;
  });
  check('SKILL_DATABASE', gen);
}

// CARD_DATABASE
{
  function rowsToObjectsAutoFiltered(name, idKey) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    }).filter(o => typeof o[idKey] === 'number');
  }

  const CARD_ICON_MAP = require('./transforms/card_icon_map.json');
  const CARD_ICONBG_MAP = require('./transforms/card_iconbg_map.json');
  const CARD_POSITION_MAP = require('./transforms/card_position_map.json');
  const CARD_SYSTEM_OVERRIDE = { 60001: '20.0' }; // original literal quirk (all other state cards render as plain "N", this one has "N.0")

  function parseNumList(v) {
    if (v === null || v === undefined) return undefined;
    return String(v).split(',').map(s => Number(s.trim()));
  }
  function parseBraceList(v) {
    if (v === null || v === undefined) return undefined;
    const matches = String(v).match(/\{([^}]*)\}/g) || [];
    return matches.map(m => Number(m.slice(1, -1)));
  }

  const objs = rowsToObjectsAutoFiltered('card', '角色id');
  const gen = objs.map(s => {
    const id = s['角色id'];
    const o = {
      id,
      rare: s.rare === null ? '' : s.rare,
      name: (s['角色名稱'] === null ? '' : s['角色名稱']).trim(),
      icon: CARD_ICON_MAP[id] !== undefined ? CARD_ICON_MAP[id] : '',
      iconBg: CARD_ICONBG_MAP[id] !== undefined ? CARD_ICONBG_MAP[id] : '#000000',
    };
    if (s.tag !== null) o.tag = parseNumList(s.tag);
    if (s.element !== null) o.element = s.element;
    o.hp = s.hp === null ? 0 : s.hp;
    o.atk = s.atk === null ? 0 : s.atk;
    o.matk = s.matk === null ? 0 : s.matk;
    o.def = s.def === null ? 0 : s.def;
    o.mdef = s.mdef === null ? 0 : s.mdef;
    o.mp = s.mp === null ? 0 : s.mp;
    o.agi = s.agi === null ? 0 : s.agi;
    o.cd = s['重生回合數'] === null || s['重生回合數'] === undefined ? null : s['重生回合數'];
    if (s.skill0_id !== null) o.s0 = s.skill0_id;
    if (s.skill1_id !== null) o.s1 = s.skill1_id;
    if (s.skill2_id !== null) o.s2 = s.skill2_id;
    o.isRole = s.type === 'role';
    if (s.exchange_id !== null) o.exchangeId = s.exchange_id;
    if (s.exchange_amount !== null) o.exchangeAmount = s.exchange_amount;
    if (s.star_refine_id !== null) o.starRefineId = Number(s.star_refine_id);
    if (s.star_cost_id !== null) o.starCostId = s.star_cost_id;
    if (s.equipment_type !== null) o.equipmentType = parseNumList(s.equipment_type);
    const ss1 = {};
    [1, 2, 3, 4, 5].forEach(n => { const v = s['star' + n + '_skill1']; if (v !== null) ss1[n] = v; });
    if (Object.keys(ss1).length) o.starSkill1 = ss1;
    const ss2 = {};
    [1, 2, 3, 4, 5].forEach(n => { const v = s['star' + n + '_skill2']; if (v !== null) ss2[n] = v; });
    if (Object.keys(ss2).length) o.starSkill2 = ss2;
    if (s['體系'] !== null) o.system = CARD_SYSTEM_OVERRIDE[id] !== undefined ? CARD_SYSTEM_OVERRIDE[id] : (typeof s['體系'] === 'number' ? String(s['體系']) : s['體系']);
    if (CARD_POSITION_MAP[id] !== undefined) o.position = CARD_POSITION_MAP[id];
    if (s.resist_buff !== null) o.resistBuff = parseNumList(s.resist_buff);
    if (s['role_ addition'] !== null) o.roleAddition = parseBraceList(s['role_ addition']);
    return o;
  });
  check('CARD_DATABASE', gen);
}

// EQUIPMENT_DATABASE
{
  function rowsToObjectsAutoFilteredE(name, idKey) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    }).filter(o => typeof o[idKey] === 'number');
  }

  const EQUIP_ICON_MAP = require('./transforms/equip_icon_map.json');
  const EQUIP_ICONBG_MAP = require('./transforms/equip_iconbg_map.json');
  const EQUIP_PASSIVE_MAP = require('./transforms/equip_passive_map.json');

  function parseNumListE(v) {
    if (v === null || v === undefined || v === '') return undefined;
    return String(v).split(',').map(s => Number(s.trim()));
  }
  function numOrZero(v) { return (v === null || v === undefined || v === '') ? 0 : v; }

  const objs = rowsToObjectsAutoFilteredE('equipment', 'id');
  const gen = {};
  objs.forEach(s => {
    const id = s.id;
    const recipe = parseNumListE(s.recipe);
    const build = parseNumListE(s.build);
    const passive = EQUIP_PASSIVE_MAP[id] || { key: null, value: 0, desc: '' };
    gen[id] = {
      id,
      type: s.type,
      name: s.info,
      rare: s.rare,
      tag: (s.tag === null || s.tag === undefined) ? '' : s.tag,
      hp: numOrZero(s.HP),
      atk: numOrZero(s.ATK),
      matk: numOrZero(s.MATK),
      def: numOrZero(s.DEF),
      mdef: numOrZero(s.MDEF),
      mp: numOrZero(s.MP),
      agi: numOrZero(s.AGI),
      recipe: recipe === undefined ? null : recipe,
      recipeRate: (s.recipe_rate === null || s.recipe_rate === undefined) ? 1 : s.recipe_rate,
      buffPasiveIds: parseNumListE(s.buff_pasive),
      refineId: s.refine_id,
      costId: s.cost_id,
      build: build === undefined ? null : build,
      icon: EQUIP_ICON_MAP[id] !== undefined ? EQUIP_ICON_MAP[id] : '❔',
      iconBg: EQUIP_ICONBG_MAP[id] !== undefined ? EQUIP_ICONBG_MAP[id] : '#000000',
      passiveKey: passive.key,
      passiveValue: passive.value,
      passiveDesc: passive.desc,
    };
  });
  check('EQUIPMENT_DATABASE', gen);
}

// GACHA_DATABASE
{
  const rows = sheetRows('gacha');
  const headerIdx = rows.findIndex(r => r.some(v => v !== null));
  const dataRows = rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null) && typeof r[0] === 'number');
  const gen = dataRows.map(r => ({ id: r[0], power: r[1], cardId: r[2] }));
  check('GACHA_DATABASE', gen);
}

// BUILD_DATABASE
{
  function rowsToObjectsAutoFilteredB(name, idKey) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    }).filter(o => typeof o[idKey] === 'number');
  }

  const BUILD_STAT_FIELD_MAP = {
    HP: 'hp', ATK: 'atk', MATK: 'matk', DEF: 'def', MDEF: 'mdef', MP: 'mp', AGI: 'agi', CRI: 'cri',
    cri_resist: 'criResist', cri_damage: 'criDamage', cri_damage_resist: 'criDamageResist',
    skill_rate: 'skillRate', skill_resist: 'skillResist', damage_increase: 'damageIncrease',
    damage_reduce: 'damageReduce', damage_reflect: 'damageReflect', mp_recove: 'mpRecover',
    shield_rate: 'shieldRate', permeate: 'permeate', atypical_rate: 'atypicalRate', heal_rate: 'healRate',
  };
  function buildStatsFrom(o) {
    const s = {};
    Object.keys(BUILD_STAT_FIELD_MAP).forEach(col => { if (o[col] !== null && o[col] !== undefined) s[BUILD_STAT_FIELD_MAP[col]] = o[col]; });
    return s;
  }

  const objs = rowsToObjectsAutoFilteredB('build', 'group_id');
  const gen = {};
  objs.forEach(o => {
    const gid = o.group_id;
    (gen[gid] = gen[gid] || []).push({
      power: o.power,
      stats: buildStatsFrom(o),
      rare: o.rare === null ? 1 : o.rare,
      valueOffset: o.value_offset === null ? 0 : o.value_offset,
    });
  });
  check('BUILD_DATABASE', gen);
}

// QUEST_DATABASE + GOD_TRIAL_VARIANTS
{
  function rowsToObjectsAutoQ(name) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    });
  }

  function parseQuestTokens(questStr) {
    if (typeof questStr !== 'string') return null;
    const tokens = questStr.match(/\{[^}]*\}/g) || [];
    return tokens.map(t => {
      const inner = t.slice(1, -1);
      const parts = inner.split(',').map(s => s.trim()).filter(Boolean);
      const obj = {};
      parts.forEach(p => {
        const sepIdx = p.search(/[:=]/);
        if (sepIdx === -1) return;
        const k = p.slice(0, sepIdx).trim();
        let v = p.slice(sepIdx + 1).trim();
        v = v.replace(/^"(.*)"$/, '$1');
        obj[k] = v;
      });
      return obj;
    });
  }
  function qNum(v) { return v === undefined ? undefined : Number(v); }
  function qParseNumList(v) {
    if (v === null || v === undefined || v === '') return undefined;
    return String(v).split(',').map(x => Number(x.trim()));
  }
  function qBuildEvent(name, type, buff, time, target, rate) {
    if (name === null || name === undefined) return null;
    const buffIds = typeof buff === 'number' ? [buff] : qParseNumList(buff);
    return {
      name, type,
      buffIds: buffIds === undefined ? undefined : buffIds,
      times: qParseNumList(time),
      rate: (rate === null || rate === undefined) ? 1 : rate,
      target: (target === null || target === undefined || target === '') ? null : target,
    };
  }

  function buildQuestEntry(s) {
    const o = { mapId: s.mapid, capability: null, type: s.type, chapter: s.chapter, hint: s.hint, level: s.level };
    o.capability = (s.Capability === null || s.Capability === undefined || s.Capability === '') ? null : s.Capability;
    o.requiresArrangement = !!s.arrangement;

    const tokens = parseQuestTokens(s.quest);
    let requiresPlayerRole = false, forcedPlayerRoleId = null;
    let playerWeaponSlots = 0, playerPresetUnits = [];
    let enemyIds = [], enemyUnitOverrides = [], waves;

    if (tokens) {
      const roleTok = tokens.find(t => 'role' in t);
      if (roleTok) {
        const rv = qNum(roleTok.role);
        if (rv === -1) { requiresPlayerRole = true; forcedPlayerRoleId = null; }
        else { requiresPlayerRole = false; forcedPlayerRoleId = -rv; }
      }
      const idTokens = tokens.filter(t => 'id' in t);
      const playerTokens = idTokens.filter(t => qNum(t.id) < 0);
      const enemyTokens = idTokens.filter(t => qNum(t.id) >= 0);

      playerWeaponSlots = playerTokens.length;
      playerTokens.forEach((t, idx) => {
        const idVal = qNum(t.id);
        if (idVal === -1) return;
        const overrides = {};
        Object.keys(t).forEach(k => { if (k !== 'id') overrides[k] = qNum(t[k]); });
        playerPresetUnits.push({ slotIndex: idx, cardId: -idVal, overrides });
      });

      const hasWave = enemyTokens.some(t => 'wave' in t);
      if (hasWave) {
        const waveMap = new Map();
        enemyTokens.forEach(t => {
          const w = t.wave;
          if (!waveMap.has(w)) waveMap.set(w, []);
          waveMap.get(w).push(qNum(t.id));
        });
        const waveKeys = [...waveMap.keys()];
        waves = waveKeys.map(k => waveMap.get(k));
        enemyIds = waves[0];
        const firstWaveTokens = enemyTokens.filter(t => t.wave === waveKeys[0]);
        enemyUnitOverrides = firstWaveTokens.map(t => {
          const ov = {};
          Object.keys(t).forEach(k => { if (k !== 'id' && k !== 'wave') ov[k] = qNum(t[k]); });
          return ov;
        });
      } else {
        enemyIds = enemyTokens.map(t => qNum(t.id));
        enemyUnitOverrides = enemyTokens.map(t => {
          const ov = {};
          Object.keys(t).forEach(k => { if (k !== 'id') ov[k] = qNum(t[k]); });
          return ov;
        });
      }
    }

    o.requiresPlayerRole = requiresPlayerRole;
    o.forcedPlayerRoleId = forcedPlayerRoleId;
    o.playerWeaponSlots = playerWeaponSlots;
    o.enemyIds = enemyIds;
    o.enemyTeamSize = enemyIds.length;
    o.enemyRoleId = null;
    o.loot1Id = s.loot1_id; o.loot1Amount = s.loot1_amount; o.loot1Rate = s.loot1_rate;
    o.loot2Id = s.loot2_id; o.loot2Amount = s.loot2_amount; o.loot2Rate = s.loot2_rate;
    o.playerPresetUnits = playerPresetUnits;
    o.enemyUnitOverrides = enemyUnitOverrides;
    if (waves !== undefined) o.waves = waves;
    o.event1 = qBuildEvent(s.event1_name, s.event1_type, s.event1_buff, s.event1_time, s.event1_target, s.event1_rate);
    o.event2 = qBuildEvent(s.event2_name, s.event2_type, s.event2_buff, s.event2_time, s.event2_target, s.event2_rate);
    return o;
  }

  const allRows = rowsToObjectsAutoQ('mapdata').filter(o => typeof o.mapid === 'number');
  const nonGod = allRows.filter(o => o.type !== 'godcharacter');
  const genQuest = nonGod.map(buildQuestEntry);
  check('QUEST_DATABASE', genQuest);

  const god = allRows.filter(o => o.type === 'godcharacter');
  const byGroup = {};
  god.forEach(s => { (byGroup[s.mapid] = byGroup[s.mapid] || []).push(s); });
  const genGodTrial = {};
  Object.keys(byGroup).forEach(gid => {
    genGodTrial[gid] = byGroup[gid].map(s => {
      const tokens = parseQuestTokens(s.quest) || [];
      const idTokens = tokens.filter(t => 'id' in t);
      const enemyTokens = idTokens.filter(t => qNum(t.id) >= 0);
      const enemyIds = enemyTokens.map(t => qNum(t.id));
      return {
        level: s.level,
        tagLimit: s.tag_limit,
        enemyIds,
        loot1Id: s.loot1_id,
        loot1Amount: s.loot1_amount,
        loot1Rate: (s.loot1_rate === null || s.loot1_rate === undefined) ? 1 : s.loot1_rate,
      };
    });
  });
  check('GOD_TRIAL_VARIANTS', genGodTrial);
}

// STORY_LINES (only meaningful once tools/transforms text sheet fix is applied; run against edited xlsx copy)
{
  function rowsToObjectsAutoS(name) {
    const rows = sheetRows(name);
    const headerIdx = rows.findIndex(r => r.some(v => v !== null));
    const headers = rows[headerIdx];
    return rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null)).map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    });
  }
  const rows = rowsToObjectsAutoS('text');
  const map100After = rows.filter(o => o.mapid === 100 && o.type === 'story' && o.timing === 'after');
  const map100Before = rows.filter(o => o.mapid === 100 && o.type === 'story' && o.timing === 'before');
  const guideDialogue = rows.filter(o => o.type === 'guide' && o.mapid === null && !/^\{[a-z_]+\}$/.test(String(o.info || '').trim()));
  const gen = {
    before100: map100Before[0].info.split('\n'),
    after100a: map100After[0].info.split('\n'),
    after100b: map100After[1].info.split('\n'),
    classIntro: map100After[2].info.split('\n'),
    gachaIntro: guideDialogue[0].info.split('\n'),
  };
  check('STORY_LINES', gen);
}
