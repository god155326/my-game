// ============================================================================
// xlsx-loader.js
//
// 在遊戲主邏輯開始前，從 data/不朽之旅.xlsx 讀取資料，
// 覆寫下面這些原本 hardcode 在 index.html 裡的資料表常數。
//
// 用法：以「本機伺服器」模式開啟本遊戲(雙擊 start.bat)，之後只要編輯/替換
// data/不朽之旅.xlsx，重新整理瀏覽器頁面就會套用最新內容 —— 不需要改 index.html。
//
// 目前已支援(其餘資料表仍暫時沿用 index.html 內建的預設值，尚未串接 xlsx)：
//   - TAGLIST_DATABASE        (xlsx: taglist)
//   - TALENTLIST_DATABASE     (xlsx: talentlist)
//   - CHARACTER_STAT_DATABASE (xlsx: character)
//   - NATURE_DATABASE         (xlsx: nature)
//   - MAPLIST_DATABASE        (xlsx: maplist)
//   - START_GIFT_DATABASE     (xlsx: start_gift)
//   - COST_DATABASE           (xlsx: cost)
//   - REFINE_DATABASE         (xlsx: refine)
//   - ITEM_DATABASE           (xlsx: item；icon 欄位不在表格內，沿用 index.html 內建的圖示對照表)
//   - LEVELCURVE_DATABASE     (xlsx: levelcurve)
//   - TALENT_CURVE            (xlsx: curve，type=1)
//   - IDLEZONE_DATABASE       (xlsx: idlezone)
//   - GACHALIST_DATABASE      (xlsx: gachalist)
//   - DUNGEON_LEVEL_DATABASE  (xlsx: dungeon)
//   - DUNGEONPREFEB_DATABASE  (xlsx: dungeonprefeb)
//   - TALENT_DATABASE / TALENT_GROUP_GRID (xlsx: talent)
//   - BUFF_DATABASE / BUFF_SET_TIERS      (xlsx: buff；icon/isUp 欄位不在表格內，沿用內建圖示對照表)
//
// 尚未串接(仍是 index.html 內建 hardcode，之後可再繼續做)：
//   CARD_DATABASE、SKILL_DATABASE、EQUIPMENT_DATABASE、GACHA_DATABASE、
//   BUILD_DATABASE、QUEST_DATABASE、GOD_TRIAL_VARIANTS，以及劇情文字
//   (STORY_LINES / MAINLINE_STORY / STORY_PORTRAITS / BG_ASSETS / CLASS_FLAVOR_TEXT)
// ============================================================================

const GAME_XLSX_PATH = 'data/不朽之旅.xlsx';

// 快取破壞用版本號：每次「重新整理頁面」都會拿到新的值(頁面載入當下的時間戳記)，
// 附加在 xlsx 與所有卡片/技能/裝備/道具/劇情立繪圖片的網址後面(?v=...)，確保瀏覽器
// (以及 GitHub Pages 的 CDN)一定會抓最新版本的檔案，不會因為檔名沒變就一直沿用舊的快取
// —— Wei 編輯 xlsx 或替換圖片後，只要重新整理頁面就一定看得到最新內容，不需要清瀏覽器快取。
const ASSET_VERSION = Date.now();

/** 把 xlsx 分頁轉成物件陣列，用第一個「非全空」列當標題列 */
function xlsxSheetToObjects(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) {
    console.warn('[xlsx-loader] 找不到分頁:', sheetName);
    return [];
  }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const headerIdx = rows.findIndex(r => r.some(v => v !== null));
  if (headerIdx === -1) return [];
  const headers = rows[headerIdx];
  return rows.slice(headerIdx + 1)
    .filter(r => r.some(v => v !== null))
    .map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] === undefined ? null : r[i]; });
      return o;
    });
}

/** Excel 序列日期 -> "YYYY-MM-DD HH:mm" */
function excelDateToStr(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60; total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor(total_seconds / 60) % 60;
  const pad = n => String(n).padStart(2, '0');
  return `${date_info.getUTCFullYear()}-${pad(date_info.getUTCMonth() + 1)}-${pad(date_info.getUTCDate())} ${pad(hours)}:${pad(minutes)}`;
}

const CHARACTER_STAT_FIELD_MAP = {
  HP: 'hp', ATK: 'atk', MATK: 'matk', DEF: 'def', MDEF: 'mdef', CRI: 'cri', MP: 'mp',
  cri_resist: 'criResist', AGI: 'agi', damage_reflect: 'damageReflect', cri_damage: 'criDamage',
  cri_damage_resist: 'criDamageResist', skill_rate: 'skillRate', skill_resist: 'skillResist',
  damage_increase: 'damageIncrease', damage_reduce: 'damageReduce', mp_recove: 'mpRecover',
  shield_rate: 'shieldRate', permeate: 'permeate', atypical_rate: 'atypicalRate', heal_rate: 'healRate',
};

const REFINE_STAT_FIELD_MAP = {
  HP: 'hp', ATK: 'atk', MATK: 'matk', DEF: 'def', MDEF: 'mdef', MP: 'mp', AGI: 'agi', CRI: 'cri',
  cri_resist: 'criResist', cri_damage: 'criDamage', cri_damage_resist: 'criDamageResist',
  skill_rate: 'skillRate', skill_resist: 'skillResist', damage_increase: 'damageIncrease',
  damage_reduce: 'damageReduce', damage_reflect: 'damageReflect',
};

const MAPLIST_KEY_BY_INFO = {
  '神格試煉': 'god_trial', '試煉之塔': 'trial_tower',
  '無限地城': 'infinite_dungeon', '殲滅任務': 'annihilation_quest',
};

function buildTaglistDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'taglist').forEach(o => { gen[o.id] = o.info; });
  return gen;
}

function buildTalentlistDatabase(wb) {
  const gen = {};
  // 2026-09-10新增：talentlist分頁新增了tier欄位(1=初階6職業、2=二轉6職業)，天賦畫面目前只顯示tier=1的分類，
  // tier=2暫不顯示於分類列表(見index.html的getTalentGroups())，這裡把tier也一併讀進來供其判斷用。
  xlsxSheetToObjects(wb, 'talentlist').forEach(o => { gen[o.group] = { info: o.info, unlockJob: o.unlock_job, tier: o.tier }; });
  // 2026-09-09新增：talentlist分頁目前只有group1~6(6個初始職業)6列資料，這次新增的「二轉職業」天賦樹
  // (group7~12，對應90008~90013)還沒有對應的列，若不補上，這6組天賦頁在畫面上會顯示成「分類7」這種
  // 沒有意義的預設標籤、且isTalentGroupUnlocked()查不到unlockJob會被當成「不鎖定」永遠可開啟(不符預期，
  // 二轉職業的天賦頁應該要等玩家真的擁有該二轉職業才能開啟)。這裡只在xlsx缺少該group時才補上預設值
  // (TALENTLIST_DEFAULT_EXTRA，定義在下方TALENT_DEFAULT_DESIGN區塊)，一旦Wei之後在talentlist分頁
  // 補上group7~12這6列，會自動改用xlsx的真實資料。
  Object.entries(TALENTLIST_DEFAULT_EXTRA).forEach(([g, info]) => { if (!gen[g]) gen[g] = info; });
  return gen;
}

function buildCharacterStatDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'character').forEach(o => {
    const entry = { limit: o.limite };
    Object.keys(CHARACTER_STAT_FIELD_MAP).forEach(col => {
      if (o[col] !== null && o[col] !== undefined) entry[CHARACTER_STAT_FIELD_MAP[col]] = o[col];
    });
    gen[o.tpye] = entry;
  });
  return gen;
}

function buildNatureDatabase(wb) {
  const statCols = ['int', 'vit', 'agi', 'str', 'dex', 'spr'];
  return xlsxSheetToObjects(wb, 'nature').map(o => {
    const requires = {};
    statCols.forEach(c => { if (o[c] !== null && o[c] !== undefined) requires[c] = o[c]; });
    return { name: o.nature_info, requires, buffId: o.buff_id };
  });
}

function buildMaplistDatabase(wb) {
  return xlsxSheetToObjects(wb, 'maplist').map(o => ({
    key: MAPLIST_KEY_BY_INFO[o.info],
    type: (o.type === null || o.type === undefined) ? null : o.type,
    info: o.info,
    beginTime: excelDateToStr(o.begin_time),
    endTime: excelDateToStr(o.end_time),
  }));
}

function buildStartGiftDatabase(wb) {
  return xlsxSheetToObjects(wb, 'start_gift').map(o => ({ id: o.id, amount: o.id_amount }));
}

function buildCostDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'cost').filter(o => typeof o.id === 'number').forEach(o => {
    if (!gen[o.id]) gen[o.id] = {};
    gen[o.id][o.level] = {
      item1Id: o.item1_id,
      item1Amount: o.item1_amount,
      item2Id: (o.item2_id === null || o.item2_id === undefined) ? null : o.item2_id,
      item2Amount: (o.item2_amount === null || o.item2_amount === undefined) ? 0 : o.item2_amount,
      rate: o.rate,
    };
  });
  return gen;
}

function buildRefineDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'refine').filter(o => typeof o.id === 'number').forEach(o => {
    if (!gen[o.id]) gen[o.id] = {};
    const entry = {};
    Object.keys(REFINE_STAT_FIELD_MAP).forEach(col => {
      if (o[col] !== null && o[col] !== undefined) entry[REFINE_STAT_FIELD_MAP[col]] = o[col];
    });
    gen[o.id][o.level] = entry;
  });
  return gen;
}

// ---- 圖示對照表(xlsx 沒有圖示欄位，沿用原本 index.html 內建的對照，之後新增的 id 會用預設圖示) ----
const ITEM_ICON_MAP = {"1000": "🪙", "1001": "💎", "2000": "🔹", "2001": "🔸", "2002": "🔶", "2003": "📦", "2004": "📦", "2005": "🗝️", "2006": "📿", "8000": "✨", "8001": "📗", "8002": "🔷", "8003": "🔷", "9101": "📜", "9102": "📜", "9103": "📜", "9104": "📜", "9105": "📜", "9106": "📜", "9107": "📜", "9108": "📜", "9109": "📜", "9110": "📜", "9111": "📜", "9112": "📜", "9113": "📜", "9114": "📜", "9115": "📜", "9116": "📜", "9117": "📜", "9118": "📜", "9119": "📜", "9120": "📜", "9121": "📜", "9122": "📜", "9123": "📜", "9124": "📜", "9125": "📜", "9126": "📜", "9127": "📜", "9128": "📜", "9129": "📜", "9130": "📜", "9131": "📦"};

// ---- 統一的「圖示自動載入 + 找不到就用預設圖示」規則 ----
// 只要把對應的圖片放到 assets/images/<category>/<category>_<id>.png，就會自動顯示；
// 還沒放圖片之前，維持顯示 fallbackHtml(通常是原本的 emoji)，找到圖片後會自動換成圖片，不用改程式碼。
function buildIconWithFallback(category, id, fallbackHtml) {
  return `<span style="position:relative;display:inline-flex;width:100%;height:100%;align-items:center;justify-content:center;">` +
    `<img src="assets/images/${category}/${category}_${id}.png?v=${ASSET_VERSION}" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;pointer-events:none;" ` +
    `onerror="this.style.display='none';this.nextElementSibling.style.display='inline-flex';">` +
    `<span style="display:none;width:100%;height:100%;align-items:center;justify-content:center;">${fallbackHtml}</span>` +
    `</span>`;
}

// ---- 給「穿插在一般文字/句子裡」的小圖示用(例如「消耗 [圖示] 金幣 x100」這種提示文字) ----
// 跟 buildIconWithFallback 的差別：後者是給本身就有固定寬高的方框容器用(圖片鋪滿100%寬高)，
// 這個版本的寬高用 em(相對於當下文字的 font-size)自己定義，不依賴外層容器有沒有設定固定尺寸，
// 所以可以安全地嵌在一句話中間，不會因為外層是不定寬高的 <span> 而讓圖片消失或跑版。
// sizeEm 可微調顯示大小(預設 1.5 倍字高，比純文字 emoji 明顯一點，圖片也才看得清楚細節)。
function buildInlineIconWithFallback(category, id, fallbackHtml, sizeEm) {
  const size = sizeEm || 1.5;
  const dip = (size * 0.22).toFixed(2);
  return `<span style="display:inline-flex;position:relative;width:${size}em;height:${size}em;vertical-align:-${dip}em;align-items:center;justify-content:center;">` +
    `<img src="assets/images/${category}/${category}_${id}.png?v=${ASSET_VERSION}" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;pointer-events:none;" ` +
    `onerror="this.style.display='none';this.nextElementSibling.style.display='inline-flex';">` +
    `<span style="display:none;width:100%;height:100%;align-items:center;justify-content:center;">${fallbackHtml}</span>` +
    `</span>`;
}

function buildItemDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'item').forEach(o => {
    gen[o.id] = { id: o.id, name: o['名稱'], usage: o.usage, icon: buildIconWithFallback('item', o.id, ITEM_ICON_MAP[o.id] || '📦') };
  });
  return gen;
}

function buildLevelcurveDatabase(wb) {
  // 2026-09-09修正：levelcurve分頁的「累计获得经验」欄位過去只有level=1那一列填了0，其餘level 2~100全部是空白，
  // 直接讀取會是null。index.html的getClassLevel()/getWeaponLevel相關查表邏輯是「exp >= row.cumExp」，
  // JS裡「任意數字 >= null」一律視為「>= 0」而恆成立，導致這個迴圈永遠不會在中途break，新玩家0經驗值
  // 就會被迴圈一路跑到最後一列、被誤判成滿級(100級)——這正是「開場就100等(應該從1等開始)」的根本原因，
  // 連帶影響所有依等級縮放的數值(含AGI)看起來都不正常。
  // 修正：不直接信任xlsx這個欄位，一律用「next」欄位(升到下一級所需經驗)由level=1開始往後累加自行算出
  // cumExp，這樣即使xlsx這欄位空白/沒有維護，也不會再發生同樣的問題(自動忽略xlsx裡的舊值，不依賴人工填寫)。
  const rows = xlsxSheetToObjects(wb, 'levelcurve').map(o => ({
    level: o.level, next: o.next === null ? 0 : o.next,
  }));
  rows.sort((a, b) => a.level - b.level);
  let cum = 0;
  rows.forEach(r => { r.cumExp = cum; cum += r.next; });
  return rows;
}

// 2026-09-10 Wei新增per_level欄位規則後更新：curve分頁的type=1(天賦點數,item8002)/type=2(屬性點數,item8003)
// 兩組資料現在都是「級距(level)+每幾等觸發一次(per_level)+每次觸發獲得點數(point)」的通用格式，
// 例如type=1第一列 per_level=1、level=5、point=2，代表「5等以內，每1等產生2點」；type=2 per_level=5、
// level=100、point=1，代表「100等以內，每5等產生1點」。實際計算邏輯見index.html的computeCurvePoints()。
// 這裡的參考預設值僅在xlsx資料不完整時才會用到(見下方buildTalentCurve/buildGodPointCurve的complete判斷)。
const TALENT_CURVE_REFERENCE = [
  { level: 5, point: 2, perLevel: 1 }, { level: 10, point: 3, perLevel: 1 }, { level: 15, point: 5, perLevel: 1 }, { level: 20, point: 7, perLevel: 1 },
  { level: 25, point: 9, perLevel: 1 }, { level: 30, point: 12, perLevel: 1 }, { level: 35, point: 15, perLevel: 1 }, { level: 40, point: 18, perLevel: 1 },
  { level: 45, point: 21, perLevel: 1 }, { level: 50, point: 25, perLevel: 1 }, { level: 55, point: 29, perLevel: 1 }, { level: 60, point: 33, perLevel: 1 },
  { level: 65, point: 37, perLevel: 1 }, { level: 70, point: 41, perLevel: 1 }, { level: 75, point: 46, perLevel: 1 }, { level: 80, point: 51, perLevel: 1 },
  { level: 85, point: 56, perLevel: 1 }, { level: 90, point: 61, perLevel: 1 }, { level: 95, point: 66, perLevel: 1 }, { level: 100, point: 71, perLevel: 1 },
];
function buildTalentCurve(wb) {
  const rows = xlsxSheetToObjects(wb, 'curve').filter(o => o.type === 1)
    .map(o => ({ level: o.level, point: o.point, perLevel: o.per_level, total: o['#總點數'] }));
  const complete = TALENT_CURVE_REFERENCE.every(ref => rows.some(r => r.level === ref.level && r.point !== null && r.point !== undefined));
  if (!complete) {
    console.warn('[xlsx-loader] curve分頁type=1的天賦點數里程碑資料不完整，改用內建參考曲線TALENT_CURVE_REFERENCE，詳見程式碼註解。');
    return TALENT_CURVE_REFERENCE.map(r => ({ level: r.level, point: r.point, perLevel: r.perLevel, total: null }));
  }
  return rows;
}

// type=2：屬性點數(item 8003，用於神格系統)，目前xlsx只有一列(per_level=5、level=100、point=1，
// 即100等以內每5等產生1點)，跟天賦點數共用同一份curve分頁、同一套通用計算邏輯。
const GOD_POINT_CURVE_REFERENCE = [
  { level: 100, point: 1, perLevel: 5 },
];
function buildGodPointCurve(wb) {
  const rows = xlsxSheetToObjects(wb, 'curve').filter(o => o.type === 2)
    .map(o => ({ level: o.level, point: o.point, perLevel: o.per_level, total: o['#總點數'] }));
  const complete = GOD_POINT_CURVE_REFERENCE.every(ref => rows.some(r => r.level === ref.level && r.point !== null && r.point !== undefined && r.perLevel != null));
  if (!complete) {
    console.warn('[xlsx-loader] curve分頁type=2的屬性點數里程碑資料不完整，改用內建參考曲線GOD_POINT_CURVE_REFERENCE，詳見程式碼註解。');
    return GOD_POINT_CURVE_REFERENCE.map(r => ({ level: r.level, point: r.point, perLevel: r.perLevel, total: null }));
  }
  return rows;
}

function buildIdlezoneDatabase(wb) {
  const splitIds = s => String(s).split(',').map(x => parseInt(x.trim(), 10));
  const gen = {};
  xlsxSheetToObjects(wb, 'idlezone').forEach(o => {
    const rewards = [];
    // 掛機獎勵欄位數量：2026-09新手引導整合計畫批次1擴充為8欄(idlereward1~8)，
    // 對應R~SR裝備/武器經驗/職業經驗/黑暗契約卷/魔石碎片/裝備強化石/金幣/神格試煉卷共8個產出類別
    for (let n = 1; n <= 8; n++) {
      const idKey = `idlereward${n}_id`, amtKey = `idlereward${n}_amount`, rateKey = `idlereward${n}_rate`;
      if (o[idKey] !== null && o[idKey] !== undefined) {
        rewards.push({ ids: splitIds(o[idKey]), amount: o[amtKey], rate: o[rateKey] });
      }
    }
    gen[o.mapid] = { mapid: o.mapid, randomPlay: splitIds(o.random_play), level: o.level, rewards };
  });
  return gen;
}

function excelDateToStrSec(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60; total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor(total_seconds / 60) % 60;
  const pad = n => String(n).padStart(2, '0');
  return `${date_info.getUTCFullYear()}-${pad(date_info.getUTCMonth() + 1)}-${pad(date_info.getUTCDate())} ${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
}

function buildGachalistDatabase(wb) {
  return xlsxSheetToObjects(wb, 'gachalist').map(o => ({
    id: o.id, info: o.info,
    cost1Id: o.cost1_id, cost1Amount: o.cost1_amount,
    cost10Id: o.cost10_id, cost10Amount: o.cost10_amount,
    beginTime: excelDateToStrSec(o.begin_time), endTime: excelDateToStrSec(o.end_time),
  }));
}

function buildDungeonLevelDatabase(wb) {
  const gen = {};
  xlsxSheetToObjects(wb, 'dungeon').forEach(o => {
    gen[o.level] = { level: o.level, rateAdd: o.rate_add, guarantMin: o.guarant_min, guarantMax: o.guarant_max, buffMax: o.buff_max };
  });
  return gen;
}

// 2026-09-09(Wei要求「圖片顯示規格改成填表可新增」)：dungeonprefeb分頁新增了`img`欄位(僅填檔名，
// 例如"img_5d6c6ccfd6d1.jpg")，這裡統一補上"assets/images/gate/"資料夾路徑組成完整圖片路徑。
// 之後 Wei 不管是新增gate5/gate6這種同命名慣例的樣板，或是像map_redboss這種完全不同命名的新樣板，
// 只要在xlsx這張表新增一列、img欄填上放進assets/images/gate/資料夾的檔名，畫面就會自動抓到新圖，
// 不需要再改任何程式碼(舊的DUNGEON_GATE_IMG寫死物件已移除，全部改讀這裡)。
const DUNGEON_GATE_IMG_FOLDER = 'assets/images/gate/';
function buildDungeonprefebDatabase(wb) {
  return xlsxSheetToObjects(wb, 'dungeonprefeb').map(o => {
    const imgFile = o.img ? String(o.img).trim() : '';
    if (!imgFile) console.warn('[xlsx-loader] dungeonprefeb分頁提醒：', o.prefeb_map, '沒有填img欄位，這個樣板在無限地城選門畫面將不會顯示門的圖片');
    return {
      prefebMap: o.prefeb_map,
      slots: String(o.list).split(',').map(s => s.trim().replace(/^\{/, '').replace(/\}$/, '')),
      power: (o.power === null || o.power === undefined) ? null : o.power,
      isEnd: !!o.is_end,
      img: imgFile ? (DUNGEON_GATE_IMG_FOLDER + imgFile) : '',
    };
  });
}

// ##TALENT_DEFAULT_DESIGN:START##
// ##TALENT_DEFAULT_DESIGN v2 (2026-09-09 第五次對話)##
// Wei 這次要求：①天賦樹結構要「完整化」，初始職業(6組)用簡單的圖形設計(單一菱形，解鎖二階職業的節點
// 放在圖形正中間，不要放在最邊緣)；②順便設計「二轉職業」自己的天賦樹(6組，共12組)；③cost全部先降到
// 1~15區間方便現在測試，實際數值由Wei之後自行調整，不需要再打「40等剛好解鎖、100等剛好點滿」的算式。
// 每組固定15個節點、9列x5欄的菱形排列：根節點(頂點) -> 兩條平行支線各3節點 -> 中央節點(第5列/正中間，
// 6組初始職業的中央節點=解鎖對應二轉職業；6組二轉職業的中央節點只是普通強化節點，沒有再往下解鎖任何東西)
// -> 兩條平行支線各3節點 -> 底部收尾節點。initial職業組沿用Wei自己已經在xlsx talent分頁group1(id100~109)
// 填好的名稱與cost(根源之力/攻擊精進/守禦意志/解鎖劍豪...)，只是把「劍豪」代換成各組對應的二轉職業名稱，
// 並延續同一種風格設計了110~114收尾的5個節點；二轉職業組(group7~12，對應90008~90013)是全新設計。
// id配置：初始職業6組沿用id 100~189(每組15個，group1=100~114...group6=175~189，與先前批次相同)；
// 二轉職業6組改用id 300~389(每組15個，group7=300~314...group12=375~389，全新範圍不與既有id衝突)。
// buff_pasive配置：初始職業沿用8000~8089(不變)；二轉職業改用8090~8179(全新範圍，180筆合計不衝突)。
const TALENT_DEFAULT_NODES = {
  100: { id: 100, group: 1, name: "根源之力", front: [], cost: 5, buffPasiveId: 8000 },
  101: { id: 101, group: 1, name: "攻擊精進", front: [100], cost: 7, buffPasiveId: 8001 },
  102: { id: 102, group: 1, name: "守禦意志", front: [100], cost: 7, buffPasiveId: 8002 },
  103: { id: 103, group: 1, name: "攻擊精進1", front: [101], cost: 9, buffPasiveId: 8003 },
  104: { id: 104, group: 1, name: "守禦意志1", front: [102], cost: 11, buffPasiveId: 8004 },
  105: { id: 105, group: 1, name: "攻擊精進2", front: [103], cost: 13, buffPasiveId: 8005 },
  106: { id: 106, group: 1, name: "守禦意志2", front: [104], cost: 12, buffPasiveId: 8006 },
  107: { id: 107, group: 1, name: "解鎖劍豪", front: [105, 106], cost: 8, buffPasiveId: 8007, roleCardUnlock: 90008 },
  108: { id: 108, group: 1, name: "持續修練1", front: [107], cost: 10, buffPasiveId: 8008 },
  109: { id: 109, group: 1, name: "持續修練2", front: [107], cost: 15, buffPasiveId: 8009 },
  110: { id: 110, group: 1, name: "持續修練3", front: [108], cost: 9, buffPasiveId: 8010 },
  111: { id: 111, group: 1, name: "持續修練4", front: [109], cost: 11, buffPasiveId: 8011 },
  112: { id: 112, group: 1, name: "持續修練5", front: [110], cost: 12, buffPasiveId: 8012 },
  113: { id: 113, group: 1, name: "持續修練6", front: [111], cost: 14, buffPasiveId: 8013 },
  114: { id: 114, group: 1, name: "究極奧義", front: [112, 113], cost: 15, buffPasiveId: 8014 },
  115: { id: 115, group: 2, name: "根源之力", front: [], cost: 5, buffPasiveId: 8015 },
  116: { id: 116, group: 2, name: "攻擊精進", front: [115], cost: 7, buffPasiveId: 8016 },
  117: { id: 117, group: 2, name: "守禦意志", front: [115], cost: 7, buffPasiveId: 8017 },
  118: { id: 118, group: 2, name: "攻擊精進1", front: [116], cost: 9, buffPasiveId: 8018 },
  119: { id: 119, group: 2, name: "守禦意志1", front: [117], cost: 11, buffPasiveId: 8019 },
  120: { id: 120, group: 2, name: "攻擊精進2", front: [118], cost: 13, buffPasiveId: 8020 },
  121: { id: 121, group: 2, name: "守禦意志2", front: [119], cost: 12, buffPasiveId: 8021 },
  122: { id: 122, group: 2, name: "解鎖狂戰士", front: [120, 121], cost: 8, buffPasiveId: 8022, roleCardUnlock: 90009 },
  123: { id: 123, group: 2, name: "持續修練1", front: [122], cost: 10, buffPasiveId: 8023 },
  124: { id: 124, group: 2, name: "持續修練2", front: [122], cost: 15, buffPasiveId: 8024 },
  125: { id: 125, group: 2, name: "持續修練3", front: [123], cost: 9, buffPasiveId: 8025 },
  126: { id: 126, group: 2, name: "持續修練4", front: [124], cost: 11, buffPasiveId: 8026 },
  127: { id: 127, group: 2, name: "持續修練5", front: [125], cost: 12, buffPasiveId: 8027 },
  128: { id: 128, group: 2, name: "持續修練6", front: [126], cost: 14, buffPasiveId: 8028 },
  129: { id: 129, group: 2, name: "究極奧義", front: [127, 128], cost: 15, buffPasiveId: 8029 },
  130: { id: 130, group: 3, name: "根源之力", front: [], cost: 5, buffPasiveId: 8030 },
  131: { id: 131, group: 3, name: "攻擊精進", front: [130], cost: 7, buffPasiveId: 8031 },
  132: { id: 132, group: 3, name: "守禦意志", front: [130], cost: 7, buffPasiveId: 8032 },
  133: { id: 133, group: 3, name: "攻擊精進1", front: [131], cost: 9, buffPasiveId: 8033 },
  134: { id: 134, group: 3, name: "守禦意志1", front: [132], cost: 11, buffPasiveId: 8034 },
  135: { id: 135, group: 3, name: "攻擊精進2", front: [133], cost: 13, buffPasiveId: 8035 },
  136: { id: 136, group: 3, name: "守禦意志2", front: [134], cost: 12, buffPasiveId: 8036 },
  137: { id: 137, group: 3, name: "解鎖神射手", front: [135, 136], cost: 8, buffPasiveId: 8037, roleCardUnlock: 90010 },
  138: { id: 138, group: 3, name: "持續修練1", front: [137], cost: 10, buffPasiveId: 8038 },
  139: { id: 139, group: 3, name: "持續修練2", front: [137], cost: 15, buffPasiveId: 8039 },
  140: { id: 140, group: 3, name: "持續修練3", front: [138], cost: 9, buffPasiveId: 8040 },
  141: { id: 141, group: 3, name: "持續修練4", front: [139], cost: 11, buffPasiveId: 8041 },
  142: { id: 142, group: 3, name: "持續修練5", front: [140], cost: 12, buffPasiveId: 8042 },
  143: { id: 143, group: 3, name: "持續修練6", front: [141], cost: 14, buffPasiveId: 8043 },
  144: { id: 144, group: 3, name: "究極奧義", front: [142, 143], cost: 15, buffPasiveId: 8044 },
  145: { id: 145, group: 4, name: "根源之力", front: [], cost: 5, buffPasiveId: 8045 },
  146: { id: 146, group: 4, name: "攻擊精進", front: [145], cost: 7, buffPasiveId: 8046 },
  147: { id: 147, group: 4, name: "守禦意志", front: [145], cost: 7, buffPasiveId: 8047 },
  148: { id: 148, group: 4, name: "攻擊精進1", front: [146], cost: 9, buffPasiveId: 8048 },
  149: { id: 149, group: 4, name: "守禦意志1", front: [147], cost: 11, buffPasiveId: 8049 },
  150: { id: 150, group: 4, name: "攻擊精進2", front: [148], cost: 13, buffPasiveId: 8050 },
  151: { id: 151, group: 4, name: "守禦意志2", front: [149], cost: 12, buffPasiveId: 8051 },
  152: { id: 152, group: 4, name: "解鎖刺客", front: [150, 151], cost: 8, buffPasiveId: 8052, roleCardUnlock: 90011 },
  153: { id: 153, group: 4, name: "持續修練1", front: [152], cost: 10, buffPasiveId: 8053 },
  154: { id: 154, group: 4, name: "持續修練2", front: [152], cost: 15, buffPasiveId: 8054 },
  155: { id: 155, group: 4, name: "持續修練3", front: [153], cost: 9, buffPasiveId: 8055 },
  156: { id: 156, group: 4, name: "持續修練4", front: [154], cost: 11, buffPasiveId: 8056 },
  157: { id: 157, group: 4, name: "持續修練5", front: [155], cost: 12, buffPasiveId: 8057 },
  158: { id: 158, group: 4, name: "持續修練6", front: [156], cost: 14, buffPasiveId: 8058 },
  159: { id: 159, group: 4, name: "究極奧義", front: [157, 158], cost: 15, buffPasiveId: 8059 },
  160: { id: 160, group: 5, name: "根源之力", front: [], cost: 5, buffPasiveId: 8060 },
  161: { id: 161, group: 5, name: "魔力精進", front: [160], cost: 7, buffPasiveId: 8061 },
  162: { id: 162, group: 5, name: "守禦意志", front: [160], cost: 7, buffPasiveId: 8062 },
  163: { id: 163, group: 5, name: "魔力精進1", front: [161], cost: 9, buffPasiveId: 8063 },
  164: { id: 164, group: 5, name: "守禦意志1", front: [162], cost: 11, buffPasiveId: 8064 },
  165: { id: 165, group: 5, name: "魔力精進2", front: [163], cost: 13, buffPasiveId: 8065 },
  166: { id: 166, group: 5, name: "守禦意志2", front: [164], cost: 12, buffPasiveId: 8066 },
  167: { id: 167, group: 5, name: "解鎖魔導士", front: [165, 166], cost: 8, buffPasiveId: 8067, roleCardUnlock: 90012 },
  168: { id: 168, group: 5, name: "持續修練1", front: [167], cost: 10, buffPasiveId: 8068 },
  169: { id: 169, group: 5, name: "持續修練2", front: [167], cost: 15, buffPasiveId: 8069 },
  170: { id: 170, group: 5, name: "持續修練3", front: [168], cost: 9, buffPasiveId: 8070 },
  171: { id: 171, group: 5, name: "持續修練4", front: [169], cost: 11, buffPasiveId: 8071 },
  172: { id: 172, group: 5, name: "持續修練5", front: [170], cost: 12, buffPasiveId: 8072 },
  173: { id: 173, group: 5, name: "持續修練6", front: [171], cost: 14, buffPasiveId: 8073 },
  174: { id: 174, group: 5, name: "究極奧義", front: [172, 173], cost: 15, buffPasiveId: 8074 },
  175: { id: 175, group: 6, name: "根源之力", front: [], cost: 5, buffPasiveId: 8075 },
  176: { id: 176, group: 6, name: "魔力精進", front: [175], cost: 7, buffPasiveId: 8076 },
  177: { id: 177, group: 6, name: "守禦意志", front: [175], cost: 7, buffPasiveId: 8077 },
  178: { id: 178, group: 6, name: "魔力精進1", front: [176], cost: 9, buffPasiveId: 8078 },
  179: { id: 179, group: 6, name: "守禦意志1", front: [177], cost: 11, buffPasiveId: 8079 },
  180: { id: 180, group: 6, name: "魔力精進2", front: [178], cost: 13, buffPasiveId: 8080 },
  181: { id: 181, group: 6, name: "守禦意志2", front: [179], cost: 12, buffPasiveId: 8081 },
  182: { id: 182, group: 6, name: "解鎖神官", front: [180, 181], cost: 8, buffPasiveId: 8082, roleCardUnlock: 90013 },
  183: { id: 183, group: 6, name: "持續修練1", front: [182], cost: 10, buffPasiveId: 8083 },
  184: { id: 184, group: 6, name: "持續修練2", front: [182], cost: 15, buffPasiveId: 8084 },
  185: { id: 185, group: 6, name: "持續修練3", front: [183], cost: 9, buffPasiveId: 8085 },
  186: { id: 186, group: 6, name: "持續修練4", front: [184], cost: 11, buffPasiveId: 8086 },
  187: { id: 187, group: 6, name: "持續修練5", front: [185], cost: 12, buffPasiveId: 8087 },
  188: { id: 188, group: 6, name: "持續修練6", front: [186], cost: 14, buffPasiveId: 8088 },
  189: { id: 189, group: 6, name: "究極奧義", front: [187, 188], cost: 15, buffPasiveId: 8089 },
};
const TALENT_DEFAULT_GRID = {
  1: [[0,0,100,0,0], [0,101,0,102,0], [0,103,0,104,0], [0,105,0,106,0], [0,0,107,0,0], [0,108,0,109,0], [0,110,0,111,0], [0,112,0,113,0], [0,0,114,0,0]],
  2: [[0,0,115,0,0], [0,116,0,117,0], [0,118,0,119,0], [0,120,0,121,0], [0,0,122,0,0], [0,123,0,124,0], [0,125,0,126,0], [0,127,0,128,0], [0,0,129,0,0]],
  3: [[0,0,130,0,0], [0,131,0,132,0], [0,133,0,134,0], [0,135,0,136,0], [0,0,137,0,0], [0,138,0,139,0], [0,140,0,141,0], [0,142,0,143,0], [0,0,144,0,0]],
  4: [[0,0,145,0,0], [0,146,0,147,0], [0,148,0,149,0], [0,150,0,151,0], [0,0,152,0,0], [0,153,0,154,0], [0,155,0,156,0], [0,157,0,158,0], [0,0,159,0,0]],
  5: [[0,0,160,0,0], [0,161,0,162,0], [0,163,0,164,0], [0,165,0,166,0], [0,0,167,0,0], [0,168,0,169,0], [0,170,0,171,0], [0,172,0,173,0], [0,0,174,0,0]],
  6: [[0,0,175,0,0], [0,176,0,177,0], [0,178,0,179,0], [0,180,0,181,0], [0,0,182,0,0], [0,183,0,184,0], [0,185,0,186,0], [0,187,0,188,0], [0,0,189,0,0]],
};
// 2026-09-09新增：talentlist分頁目前只有group1~6(對應初始職業)6列資料，二轉職業6組(group7~12)
// 是這次新增的系統，xlsx還沒有對應列(Wei之後可以自行在talentlist分頁補上這6列覆蓋掉這份預設值，
// 欄位是 group/info/unlock_job，例如 7 / 劍豪天賦 / 90008)，這裡先在程式端補上預設值確保新分頁
// 能正常顯示標籤與解鎖判斷(isTalentGroupUnlocked仍是查gameState.ownedClasses是否擁有unlock_job這個
// 職業id，二轉職業的擁有狀態則是透過學習對應初始職業天賦樹中央的解鎖節點後自動加入，邏輯完全共用不需要改動)。
const TALENTLIST_DEFAULT_EXTRA = {
  7: { info: "劍豪天賦", unlockJob: 90007, tier: 2 },
  8: { info: "狂戰士天賦", unlockJob: 90008, tier: 2 },
  9: { info: "神射手天賦", unlockJob: 90009, tier: 2 },
  10: { info: "刺客天賦", unlockJob: 90010, tier: 2 },
  11: { info: "魔導士天賦", unlockJob: 90011, tier: 2 },
  12: { info: "神官天賦", unlockJob: 90012, tier: 2 },
};
// ##TALENT_DEFAULT_DESIGN:END##

function buildTalentDatabaseAndGrid(wb) {
  const parseFront = v => {
    if (v === null || v === undefined) return [];
    return String(v).split(',').map(s => parseInt(s.trim(), 10));
  };
  const genTalent = {};
  const genGrid = {};
  // 2026-09-09防呆修正：talent分頁目前每個group只有第一列(該group的根節點)填了id，同一group底下其餘
  // 每個節點各自的列全部id留空——過去這裡不論o.id是什麼都直接genTalent[o.id]=entry，id留空的列全部會
  // 撞在同一個「null」/「undefined」key上互相覆蓋，最後只會剩下最後一筆蓋過去的殘影，導致天賦畫面
  // 幾乎完全沒有節點可顯示(Wei回報「天賦點的天賦丟失，完全沒顯示內容物」)。這裡先加上防呆：id缺漏的列
  // 直接跳過並印出警告，不再讓它們互相覆蓋、汙染到其他正常資料；但這只是避免資料衝突的防呆，並不能讓
  // 天賦樹恢復正常顯示——每個節點本來就需要自己獨立的id(display_position格線裡已經列出每個group該有
  // 哪些id，例如group1需要100~114共15個)，這部分的資料本身還沒有逐列補上，需要請Wei參照格線把每一列
  // 對應的id、buff_pasive(全部節點目前也都是空白，代表就算id補齊了、點下去也不會有任何實際加成效果)
  // 補齊到xlsx裡，程式端目前沒有足夠資訊能自動還原這份對照關係。
  xlsxSheetToObjects(wb, 'talent').forEach(o => {
    if (o.id === null || o.id === undefined) {
      console.warn('[xlsx-loader] talent分頁有節點缺少id，已略過(不會顯示在天賦樹)：', o);
      return;
    }
    const entry = { id: o.id, group: o.group, name: o.group_info, front: parseFront(o.front), cost: o.cost, buffPasiveId: o.buff_pasive };
    if (o.role_card_unlock !== null && o.role_card_unlock !== undefined) entry.roleCardUnlock = o.role_card_unlock;
    genTalent[o.id] = entry;
    if (o.display_position !== null && o.display_position !== undefined) {
      genGrid[o.group] = String(o.display_position).split('\n').map(line =>
        line.split(',').map(cell => parseInt(cell.trim().replace(/^\{/, '').replace(/\}$/, ''), 10))
      );
    }
  });
  // 2026-09-09新增：每個group應該要有15個節點，上面這段防呆頂多只能「跳過缺id的列」，如果xlsx某個
  // group讀到的節點數還是不足15個，代表這個group的資料本來就還沒補齊，這裡改成整組直接改用
  // TALENT_DEFAULT_NODES/TALENT_DEFAULT_GRID頂替，確保天賦樹在Wei填齊xlsx之前也能正常顯示與運作。
  // 一旦Wei之後把xlsx talent分頁裡某個group的15個節點id都補齊，這裡會自動偵測到並改用xlsx的真實資料。
  // 2026-09-09再次更新：g從原本1~6擴充為1~12——group1~6是6個初始職業，group7~12是這次新增的「二轉職業」
  // 專屬天賦樹(對應90008~90013)，xlsx talent分頁目前完全沒有group7~12的任何資料，所以這6組現在一定會
  // 套用TALENT_DEFAULT_NODES/GRID裡的預設設計；等Wei之後在xlsx補上這6組各自15個節點，會自動改用xlsx資料。
  // 2026-09-10再加強：光看節點數量是否滿15個還不足以保證資料真的可用——實測發現Wei在talent分頁補資料
  // 時，group2~6的id欄位填的是「每組各自重新從頭編號」(例如group2=1~15、group3=16~30...一路往下累加)，
  // 但同一批資料的front/display_position欄位卻是原封不動沿用本來設計稿的絕對編號(group2=115~129、
  // group3=130~144...)，兩邊對不起來，導致front永遠查不到同組內任何一個真實存在的id、grid裡的節點座標
  // 也對應不到任何節點，天賦樹畫面因此完全空白(Wei回報「弓劍手天賦沒有顯示對應天賦表」，弓劍手正好是
  // 受影響的group3)。這裡在原本的節點數量檢查之外，多加一道「每個節點的front是否都能在同一組內找到
  // 對應id」的檢查，只要有任何一個front指向的id不屬於本組，就代表這組資料前後不一致(不論是id編號方式
  // 對不上、還是純粹填錯字打錯數字)，整組直接改用內建預設值頂替，不會讓玩家看到一片空白或殘缺不全的
  // 天賦樹；一旦Wei之後把某組15個節點的id改成跟front/display_position互相對得起來，這裡會自動偵測到
  // 並改回讀取xlsx的真實資料。
  for (let g = 1; g <= 12; g++) {
    const nodesInGroup = Object.values(genTalent).filter(n => n.group === g);
    const idSet = new Set(nodesInGroup.map(n => n.id));
    const hasDanglingFront = nodesInGroup.some(n => (n.front || []).some(f => !idSet.has(f)));
    if (nodesInGroup.length < 15 || hasDanglingFront) {
      if (hasDanglingFront && nodesInGroup.length >= 15) {
        console.warn(`[xlsx-loader] talent分頁group${g}的節點數量足夠，但front欄位引用的id跟本組實際id對不上(可能id編號方式跟front/display_position原本設計的絕對編號不一致)，已改用內建預設值，不使用xlsx這組資料：`, nodesInGroup);
      }
      Object.keys(genTalent).forEach(k => { if (genTalent[k].group === g) delete genTalent[k]; });
      Object.entries(TALENT_DEFAULT_NODES).forEach(([id, node]) => { if (node.group === g) genTalent[id] = node; });
      genGrid[g] = TALENT_DEFAULT_GRID[g];
    }
  }
  return { talent: genTalent, grid: genGrid };
}

const BUFF_ICON_MAP = {"1001": {"icon": "⚔️▲", "isUp": true}, "1002": {"icon": "🛡️▲", "isUp": true}, "1003": {"icon": "🔮▲", "isUp": true}, "1004": {"icon": "🥷▲", "isUp": true}, "1005": {"icon": "💥▲", "isUp": true}, "1006": {"icon": "💥⚔️▲", "isUp": true}, "1007": {"icon": "📜▲", "isUp": true}, "1008": {"icon": "🛡️💥▲", "isUp": true}, "1009": {"icon": "💥🛡️▲", "isUp": true}, "1010": {"icon": "📜🛡️▲", "isUp": true}, "1011": {"icon": "🗡️▲", "isUp": true}, "1012": {"icon": "🧊▲", "isUp": true}, "1996": {"icon": "✨▲", "isUp": true}, "1997": {"icon": "🔁", "isUp": true}, "1998": {"icon": "✨▲", "isUp": true}, "1999": {"icon": "🛡️", "isUp": true}, "2001": {"icon": "⚔️▼", "isUp": false}, "2002": {"icon": "🛡️▼", "isUp": false}, "2003": {"icon": "🔮▼", "isUp": false}, "2004": {"icon": "🥷▼", "isUp": false}, "2005": {"icon": "💥▼", "isUp": false}, "2006": {"icon": "💥⚔️▼", "isUp": false}, "2007": {"icon": "📜▼", "isUp": false}, "2008": {"icon": "🛡️💥▼", "isUp": false}, "2009": {"icon": "💥🛡️▼", "isUp": false}, "2010": {"icon": "📜🛡️▼", "isUp": false}, "2011": {"icon": "🗡️▼", "isUp": false}, "2012": {"icon": "🧊▼", "isUp": false}, "3001": {"icon": "💚", "isUp": true}, "3002": {"icon": "💙", "isUp": true}, "3003": {"icon": "💨", "isUp": true}, "3004": {"icon": "🕊️", "isUp": true}, "3005": {"icon": "💢🛡️", "isUp": true}, "3006": {"icon": "💢🔮", "isUp": true}, "3007": {"icon": "👹", "isUp": true}, "3008": {"icon": "✨▲", "isUp": true}, "3009": {"icon": "✨▲", "isUp": true}, "3010": {"icon": "✨▲", "isUp": true}, "3011": {"icon": "✨▲", "isUp": true}, "3012": {"icon": "✨▲", "isUp": true}, "3013": {"icon": "✨▲", "isUp": true}, "3014": {"icon": "✨▲", "isUp": true}, "3015": {"icon": "✨▼", "isUp": true}, "3016": {"icon": "✨▲", "isUp": true}, "3017": {"icon": "✨▲", "isUp": true}, "3018": {"icon": "✨▲", "isUp": true}, "3019": {"icon": "✨", "isUp": true}, "3020": {"icon": "⚔️👑▲", "isUp": true}, "4001": {"icon": "🔥", "isUp": false}, "4002": {"icon": "❄️", "isUp": false}, "4003": {"icon": "🩸", "isUp": false}, "4004": {"icon": "☠️", "isUp": false}, "4005": {"icon": "🚫💚", "isUp": false}, "5001": {"icon": "⛓️", "isUp": false}, "5002": {"icon": "💫", "isUp": false}, "5003": {"icon": "🧊", "isUp": false}, "5004": {"icon": "⏹️", "isUp": false}, "5005": {"icon": "🌀", "isUp": true}, "5006": {"icon": "💥", "isUp": true}, "6001": {"icon": "🧩"}, "6002": {"icon": "🧩"}, "6003": {"icon": "🧩"}, "6004": {"icon": "🧩"}, "6005": {"icon": "🧩"}, "6006": {"icon": "🧩"}, "6007": {"icon": "✨▲"}, "6008": {"icon": "✨▲"}, "6009": {"icon": "🧩"}, "6010": {"icon": "🧩"}, "6011": {"icon": "🧩"}, "7001": {"icon": "✨▲", "isUp": true}, "7002": {"icon": "✨▲", "isUp": true}, "7003": {"icon": "✨▲", "isUp": true}, "7004": {"icon": "✨▲", "isUp": true}, "7005": {"icon": "✨▲", "isUp": true}, "7006": {"icon": "✨▲", "isUp": true}, "7007": {"icon": "✨▲", "isUp": true}, "7008": {"icon": "✨▲", "isUp": true}, "7009": {"icon": "✨▲", "isUp": true}, "7010": {"icon": "✨▲", "isUp": true}, "7011": {"icon": "✨▲", "isUp": true}, "7012": {"icon": "✨▲", "isUp": true}, "7013": {"icon": "✨▲", "isUp": true}, "7014": {"icon": "✨▲", "isUp": true}, "7015": {"icon": "✨▲", "isUp": true}, "7016": {"icon": "✨▲", "isUp": true}, "7017": {"icon": "✨▲", "isUp": true}, "7018": {"icon": "✨▲", "isUp": true}, "7019": {"icon": "✨▲", "isUp": true}, "7020": {"icon": "✨▲", "isUp": true}, "7021": {"icon": "✨▲", "isUp": true}, "7022": {"icon": "✨▲", "isUp": true}, "7023": {"icon": "✨▲", "isUp": true}, "7024": {"icon": "✨▲", "isUp": true}, "7025": {"icon": "✨▲", "isUp": true}, "7026": {"icon": "✨▲", "isUp": true}, "7027": {"icon": "✨▲", "isUp": true}, "7028": {"icon": "✨▲", "isUp": true}, "7029": {"icon": "✨▲", "isUp": true}, "7030": {"icon": "✨▲", "isUp": true}, "7031": {"icon": "✨▲", "isUp": true}, "7032": {"icon": "✨▲", "isUp": true}, "7033": {"icon": "✨▲", "isUp": true}, "7034": {"icon": "✨▲", "isUp": true}, "7035": {"icon": "✨▲", "isUp": true}, "7036": {"icon": "✨▲", "isUp": true}, "7037": {"icon": "✨▲", "isUp": true}, "7038": {"icon": "✨▲", "isUp": true}, "7039": {"icon": "✨▲", "isUp": true}, "7040": {"icon": "✨▲", "isUp": true}, "7041": {"icon": "✨▲", "isUp": true}, "7042": {"icon": "✨▲", "isUp": true}, "7043": {"icon": "✨▲", "isUp": true}, "7044": {"icon": "✨▲", "isUp": true}, "7045": {"icon": "✨▲", "isUp": true}, "7046": {"icon": "✨▲", "isUp": true}, "7047": {"icon": "✨▲", "isUp": true}, "7048": {"icon": "✨▲", "isUp": true}, "7049": {"icon": "✨▲", "isUp": true}, "7050": {"icon": "✨▲", "isUp": true}, "7051": {"icon": "✨▲", "isUp": true}, "7052": {"icon": "✨▲", "isUp": true}, "7053": {"icon": "✨▲", "isUp": true}, "7054": {"icon": "✨▲", "isUp": true}, "7055": {"icon": "✨▲", "isUp": true}, "7056": {"icon": "✨▲", "isUp": true}, "7057": {"icon": "✨▲", "isUp": true}, "7058": {"icon": "✨▲", "isUp": true}, "7059": {"icon": "✨▲", "isUp": true}, "7060": {"icon": "✨▲", "isUp": true}, "7061": {"icon": "✨▲", "isUp": true}, "7062": {"icon": "✨▲", "isUp": true}, "7063": {"icon": "✨▲", "isUp": true}, "7064": {"icon": "✨▲", "isUp": true}, "7065": {"icon": "✨▲", "isUp": true}, "7066": {"icon": "✨▲", "isUp": true}, "7067": {"icon": "✨▲", "isUp": true}, "7068": {"icon": "✨▲", "isUp": true}, "7069": {"icon": "✨▲", "isUp": true}, "7070": {"icon": "✨▲", "isUp": true}, "7071": {"icon": "✨▲", "isUp": true}, "7072": {"icon": "✨▲", "isUp": true}, "7073": {"icon": "✨▲", "isUp": true}, "7074": {"icon": "✨▲", "isUp": true}, "7075": {"icon": "✨▲", "isUp": true}, "7076": {"icon": "✨▲", "isUp": true}, "7077": {"icon": "✨▲", "isUp": true}, "7078": {"icon": "✨▲", "isUp": true}, "7079": {"icon": "✨▲", "isUp": true}, "7080": {"icon": "✨▲", "isUp": true}, "8000": {"icon": "✨▲", "isUp": true}, "8001": {"icon": "✨▲", "isUp": true}, "8002": {"icon": "✨▲", "isUp": true}, "8003": {"icon": "✨▲", "isUp": true}, "8004": {"icon": "✨▲", "isUp": true}, "8005": {"icon": "✨▲", "isUp": true}, "8006": {"icon": "✨▲", "isUp": true}, "8007": {"icon": "✨▲", "isUp": true}, "8008": {"icon": "✨▲", "isUp": true}, "8009": {"icon": "✨▲", "isUp": true}, "8010": {"icon": "✨▲", "isUp": true}, "8011": {"icon": "✨▲", "isUp": true}, "8012": {"icon": "✨▲", "isUp": true}, "8013": {"icon": "✨▲", "isUp": true}, "8014": {"icon": "✨▲", "isUp": true}, "8015": {"icon": "✨▲", "isUp": true}, "8016": {"icon": "✨▲", "isUp": true}, "8017": {"icon": "✨▲", "isUp": true}, "8018": {"icon": "✨▲", "isUp": true}, "8019": {"icon": "✨▲", "isUp": true}, "8020": {"icon": "✨▲", "isUp": true}, "8021": {"icon": "✨▲", "isUp": true}, "8022": {"icon": "✨▲", "isUp": true}, "8023": {"icon": "✨▲", "isUp": true}, "8024": {"icon": "✨▲", "isUp": true}, "8025": {"icon": "✨▲", "isUp": true}, "8026": {"icon": "✨▲", "isUp": true}, "8027": {"icon": "✨▲", "isUp": true}, "8028": {"icon": "✨▲", "isUp": true}, "8029": {"icon": "✨▲", "isUp": true}, "8030": {"icon": "✨▲", "isUp": true}, "8031": {"icon": "✨▲", "isUp": true}, "8032": {"icon": "✨▲", "isUp": true}, "8033": {"icon": "✨▲", "isUp": true}, "8034": {"icon": "✨▲", "isUp": true}, "8035": {"icon": "✨▲", "isUp": true}, "8036": {"icon": "✨▲", "isUp": true}, "8037": {"icon": "✨▲", "isUp": true}, "8038": {"icon": "✨▲", "isUp": true}, "8039": {"icon": "✨▲", "isUp": true}, "8040": {"icon": "✨▲", "isUp": true}, "8041": {"icon": "✨▲", "isUp": true}, "8042": {"icon": "✨▲", "isUp": true}, "8043": {"icon": "✨▲", "isUp": true}, "8044": {"icon": "✨▲", "isUp": true}, "8045": {"icon": "✨▲", "isUp": true}, "8046": {"icon": "✨▲", "isUp": true}, "8047": {"icon": "✨▲", "isUp": true}, "8048": {"icon": "✨▲", "isUp": true}, "8049": {"icon": "✨▲", "isUp": true}, "8050": {"icon": "✨▲", "isUp": true}, "8051": {"icon": "✨▲", "isUp": true}, "8052": {"icon": "✨▲", "isUp": true}, "8053": {"icon": "✨▲", "isUp": true}, "8054": {"icon": "✨▲", "isUp": true}, "8055": {"icon": "✨▲", "isUp": true}, "8056": {"icon": "✨▲", "isUp": true}, "8057": {"icon": "✨▲", "isUp": true}, "8058": {"icon": "✨▲", "isUp": true}, "8059": {"icon": "✨▲", "isUp": true}, "8060": {"icon": "✨▲", "isUp": true}, "8061": {"icon": "✨▲", "isUp": true}, "8062": {"icon": "✨▲", "isUp": true}, "8063": {"icon": "✨▲", "isUp": true}, "8064": {"icon": "✨▲", "isUp": true}, "8065": {"icon": "✨▲", "isUp": true}, "8066": {"icon": "✨▲", "isUp": true}, "8067": {"icon": "✨▲", "isUp": true}, "8068": {"icon": "✨▲", "isUp": true}, "8069": {"icon": "✨▲", "isUp": true}, "8070": {"icon": "✨▲", "isUp": true}, "8071": {"icon": "✨▲", "isUp": true}, "8072": {"icon": "✨▲", "isUp": true}, "8073": {"icon": "✨▲", "isUp": true}, "8074": {"icon": "✨▲", "isUp": true}, "8075": {"icon": "✨▲", "isUp": true}, "8076": {"icon": "✨▲", "isUp": true}, "8077": {"icon": "✨▲", "isUp": true}, "8078": {"icon": "✨▲", "isUp": true}, "8079": {"icon": "✨▲", "isUp": true}, "8080": {"icon": "✨▲", "isUp": true}, "8081": {"icon": "✨▲", "isUp": true}, "8082": {"icon": "✨▲", "isUp": true}, "8083": {"icon": "✨▲", "isUp": true}, "8084": {"icon": "✨▲", "isUp": true}, "8085": {"icon": "✨▲", "isUp": true}, "8086": {"icon": "✨▲", "isUp": true}, "8087": {"icon": "✨▲", "isUp": true}, "8088": {"icon": "✨▲", "isUp": true}, "8089": {"icon": "✨▲", "isUp": true}, "9001": {"icon": "✨▲", "isUp": true}, "9002": {"icon": "✨▲", "isUp": true}, "9003": {"icon": "✨▼", "isUp": false}, "9004": {"icon": "✨▼", "isUp": false}, "9005": {"icon": "✨▲", "isUp": true}, "9006": {"icon": "✨▼", "isUp": true}, "9007": {"icon": "✨▼", "isUp": false}, "9008": {"icon": "✨▼", "isUp": false}, "9009": {"icon": "✨▲", "isUp": true}, "9010": {"icon": "✨▼", "isUp": false}, "9011": {"icon": "✨▼", "isUp": false}, "9012": {"icon": "✨▼", "isUp": false}, "9900": {"icon": "🌟", "isUp": true}, "9901": {"icon": "🌟", "isUp": true}, "9902": {"icon": "🌟", "isUp": true}, "9903": {"icon": "🌟", "isUp": true}, "9904": {"icon": "🌟", "isUp": true}, "9905": {"icon": "🌟", "isUp": true}, "9906": {"icon": "🌟", "isUp": true}, "9907": {"icon": "🌟", "isUp": true}, "9908": {"icon": "🌟", "isUp": true}, "9909": {"icon": "🌟", "isUp": true}, "9910": {"icon": "🌟", "isUp": true}, "9911": {"icon": "🌟", "isUp": true}, "9912": {"icon": "🌟", "isUp": true}, "9913": {"icon": "🌟", "isUp": true}, "9914": {"icon": "🌟", "isUp": true}, "9915": {"icon": "🌟", "isUp": true}};

const BUFF_STAT_FIELD_MAP = {
  ATK: 'atk', MATK: 'matk', DEF: 'def', MDEF: 'mdef', CRI: 'cri', MP: 'mp',
  cri_resist: 'criResist', AGI: 'agi', damage_reflect: 'damageReflect', cri_damage: 'criDamage',
  cri_damage_resist: 'criDamageResist', skill_rate: 'skillRate', skill_resist: 'skillResist',
  damage_increase: 'damageIncrease', damage_reduce: 'damageReduce',
};
function buffStatFields(o) {
  const entry = {};
  Object.keys(BUFF_STAT_FIELD_MAP).forEach(col => { if (o[col] !== null && o[col] !== undefined) entry[BUFF_STAT_FIELD_MAP[col]] = o[col]; });
  return entry;
}

// 2026-09-10修正(Wei回報暈眩/控制效果整批壞掉)：麻痺(5001)/暈眩(5002)/冰凍(5003)/停止(5004)/凝聚(5005)
// 這5個buff在xlsx `buff`分頁裡的buff_info欄位其實都寫得很清楚是「1回合無法行動」，但整套BUFF_DATABASE
// 產生流程(不論是這個xlsx讀取路徑、還是index.html裡xlsx讀取失敗時的內建備援陣列)從頭到尾都沒有任何地方
// 真正把這個「會讓單位這回合跳過行動」的旗標(index.html的getControlBuff()實際檢查的info.isControl)寫進
// BUFF_DATABASE——這不是這次不小心被改掉，而是這個旗標本身從未真正被賦值過，跳過行動的判斷式(getControlBuff)
// 永遠找不到任何buff的isControl為true，等於這5個控制效果的「附加成功」跟「回合結算扣血/持續時間」都正常，
// 唯獨「讓對象真的跳過這回合行動」這個核心效果完全沒有生效。這裡比照BUFF_ICON_MAP(icon/isUp)同樣的做法，
// 額外新增一份「控制類buff」的中繼資料清單，這份清單不是xlsx `buff`分頁的任何一個既有欄位、後續也不會被
// 分頁資料覆蓋掉(如同icon一樣，是本檔案內固定的分類設定，不需要也不會出現在xlsx裡)。
const CONTROL_BUFF_IDS = new Set([5001, 5002, 5003, 5004, 5005]); // 麻痺/暈眩/冰凍/停止/凝聚
// 2026-09-10修正(Wei回報持續傷害類效果又壞了)：跟上面控制類buff的isControl一模一樣的問題——燒傷(4001)/
// 凍傷(4002)/裂傷(4003)/中毒(4004)這4個持續傷害buff，index.html的processBuffTriggerEffects()是靠
// info.dotStat(要用ATK還是MATK當基準)+info.rate(每層每回合造成的傷害比例)這兩個欄位來計算每回合傷害，
// 這裡新增分類名單(同樣不是xlsx欄位，是固定的分類設定)，比照CONTROL_BUFF_IDS的做法，把dotStat/rate正確
// 賦值上去：rate取ATK/MATK欄位的絕對值(xlsx裡填的是負值，真正計算真實傷害時要用正值相乘)。
const DOT_BUFF_STAT = { 4001: 'matk', 4002: 'matk', 4003: 'atk', 4004: 'atk' }; // 灼傷/凍傷/裂傷/中毒
// 持續恢復(3001,HP)/持續回魔(3002,MP)：跟DOT一樣，這兩個buff的MATK欄位其實已經填好比例(0.4/0.02)，
// 只是同樣沒有被標記成「這是持續恢復類buff」，導致「每回合真的回血/回魔」也沒有生效，一併修正。
const REGEN_BUFF_TYPE = { 3001: 'hp', 3002: 'mp' }; // 持續恢復/持續回魔，皆以MATK欄位為基準比例
const REGEN_BUFF_STAT = { 3001: 'matk', 3002: 'matk' };

function buildBuffDatabaseAndSetTiers(wb) {
  const rows = xlsxSheetToObjects(wb, 'buff').filter(o => typeof o.buff_id === 'number');
  const byId = {};
  rows.forEach(o => { (byId[o.buff_id] = byId[o.buff_id] || []).push(o); });

  const genBuff = {};
  const genSetTiers = {};
  Object.entries(byId).forEach(([idStr, rs]) => {
    const id = Number(idStr);
    const iconInfo = BUFF_ICON_MAP[id] || { icon: '✨', isUp: true };
    const plainRows = rs.filter(r => r.set_effective_count === null || r.set_effective_count === undefined);
    if (plainRows.length > 0) {
      const o = plainRows[0];
      const rawName = o['buff名稱'] === null || o['buff名稱'] === undefined ? '' : o['buff名稱'];
      const name = rawName.startsWith('#') ? rawName.slice(1) : rawName;
      const entry = { id, name, icon: iconInfo.icon, isUp: iconInfo.isUp, info: o.buff_info,
        reflashAble: !!o.reflashAble, maxStack: o.max_stack === null || o.max_stack === undefined ? 1 : o.max_stack,
        duration: o['持續回合數'] === null || o['持續回合數'] === undefined ? 1 : o['持續回合數'], ...buffStatFields(o) };
      if (o.invisible) entry.invisible = true;
      if (CONTROL_BUFF_IDS.has(id)) entry.isControl = true;
      if (DOT_BUFF_STAT[id]) {
        entry.dotStat = DOT_BUFF_STAT[id];
        entry.rate = Math.abs(entry[entry.dotStat] || 0);
      }
      if (REGEN_BUFF_TYPE[id]) {
        entry.regenType = REGEN_BUFF_TYPE[id];
        entry.rate = Math.abs(entry[REGEN_BUFF_STAT[id]] || 0);
      }
      genBuff[id] = entry;
    } else {
      const tierRows = [...rs].sort((a, b) => a.set_effective_count - b.set_effective_count);
      const base = tierRows[0];
      genBuff[id] = { id, name: base['buff名稱'], icon: iconInfo.icon, isSetBase: true, setInfo: base.set_info, info: base.buff_info,
        ...buffStatFields(base), invisible: true, reflashAble: !!base.reflashAble, maxStack: base.max_stack, duration: base['持續回合數'] };
      genSetTiers[id] = {
        setInfo: base.set_info,
        tiers: tierRows.map(o => ({ effectiveCount: o.set_effective_count, name: o['buff名稱'], ...buffStatFields(o) })),
      };
    }
  });
  // 2026-09-10 Wei指示：天賦被動效果(buff_id 8000~8179)全部改成直接讀取xlsx buff分頁的真實內容，
  // 不再用程式碼裡虛構的百分比數值/名稱覆蓋——即使目前這批buff在xlsx裡大多還是「名稱空白、統一ATK+1%」
  // 的預留佔位內容，也應該照實呈現，等Wei之後自己在buff分頁填上每個節點真正想要的效果/數值即可生效，
  // 不需要程式碼這邊另外維護一份對照表。「天賦被動解鎖後永久生效」這件事本來就是由 applyTalentBuffs()
  // 在套用時直接固定給 duration:9999(不讀取buff表本身的持續回合數欄位)來保證，跟這裡是否覆寫無關，
  // 拿掉這段覆寫不會讓天賦buff變回「打一回合就消失」。
  return { buff: genBuff, setTiers: genSetTiers };
}


/** CARD_DATABASE 相關：icon/iconBg/position 為角色美術資源，xlsx 未包含，沿用內建對照表 */
// CARD_ICON_STYLE_MAP：每張卡片圖示縮放比例(依原始美術素材裁切比例微調，非公式可推導)。
// 找不到對應 id 時預設用 80%。圖片路徑一律由 id 直接組成，不需要另外維護路徑對照表：
// 只要把圖放到 assets/images/card/card_<角色id>.png，遊戲就會自動讀到。
const CARD_ICON_STYLE_MAP = {"10001":80.5,"10002":80.5,"10003":80.5,"10004":80.5,"10005":80.5,"10006":80.5,"10007":80.5,"10008":80.5,"10009":80.5,"10010":80.5,"10011":80.5,"10012":80.5,"10013":63,"10014":63,"10015":63,"10016":63,"10017":63,"10018":63,"10019":63,"10020":63,"10021":63,"10022":70,"10023":70,"10024":70,"10025":70,"10026":70,"10027":70,"10028":70,"10029":70,"10030":70,"10031":70,"40001":80,"40002":80,"40003":80,"40004":80,"40005":80,"40006":80,"40007":80,"40008":80,"40009":80,"40010":80,"40011":80,"40012":80,"40013":80,"40014":80,"40015":80,"50000":80,"50001":72,"50002":72,"50003":72,"90001":80,"90002":80,"90003":80,"90004":80,"90005":80,"90006":80,"90007":80,"90008":80,"90009":80,"90010":80,"90011":80,"90012":80,"90013":80};
// 目前支援的三種常駐卡面特效(card分頁card_prefeb欄位可填的值)：漂浮/金屬流光/變色霓虹，樣式定義見index.html
const CARD_PREFEB_CLASSES = new Set(['card_floating', 'card_metal', 'card_magic']);
function buildCardIconHtml(id, prefeb) {
  const pct = CARD_ICON_STYLE_MAP[id] !== undefined ? CARD_ICON_STYLE_MAP[id] : 80;
  const src = `assets/images/card/card_${id}.png?v=${ASSET_VERSION}`;
  const imgTag = `<img src='${src}' style='max-width:${pct}%;max-height:${pct}%;width:auto;height:auto;object-fit:contain;pointer-events:none;' onerror="this.style.display='none'">`;
  // 2026-09-10新增：card分頁的card_prefeb欄位(card_floating/card_metal/card_magic)，讓卡面本身疊加一層
  // 常駐(不需要觸發，永遠在播放)的視覺特效(流光劃過/呼吸輝光/變色)。沒有填/填的值不是這三種之一，
  // 就完全比照原本純<img>的輸出，不受影響。這裡統一忽略前後空白，避免xlsx儲存格不小心多打空格導致比對失敗。
  const prefebClass = CARD_PREFEB_CLASSES.has(String(prefeb || '').trim()) ? String(prefeb).trim().replace(/^card_/, '') : null;
  if (!prefebClass) return imgTag;
  // 2026-09-11修正：曾經用<span style="display:contents">包裝，讓瀏覽器「當作這層包裝不存在」、直接沿用
  // 父層既有的position:relative當定位基準——但手機實機測試發現display:contents在部分瀏覽器上不可靠，
  // 流光疊層偶爾會抓不到正確的父層基準，一路往上抓到很遠的祖先元素，變成一個巨大、位置錯誤、忽隱忽現的
  // 半透明武器圖案飄在戰鬥畫面背景。改成不依賴display:contents，讓<span>本身直接position:absolute+inset:0
  // (不管父層是「一般高度」還是「用padding撐出來的特殊正方形」，這個寫法都能正確貼齊父層的實際顯示範圍，
  // 這正是原本.weapon-slot-badge img那條規則的做法)，並自己用display:flex置中裡面的武器圖，讓這個<span>
  // 本身就是穩固、不需要看瀏覽器臉色的定位基準。
  return `<span class="card-icon-wrap card-prefeb-${prefebClass}" style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;">`
    + imgTag
    + `<div class="card-prefeb-shine" style="--card-mask-url:url('${src}')"></div>`
    + `</span>`;
}
const CARD_ICONBG_MAP = {"10001":"#000000","10002":"#000000","10003":"#000000","10004":"#000000","10005":"#000000","10006":"#000000","10007":"#000000","10008":"#000000","10009":"#000000","10010":"#000000","10011":"#000000","10012":"#000000","10013":"#000000","10014":"#000000","10015":"#000000","10016":"#000000","10017":"#000000","10018":"#000000","10019":"#000000","10020":"#000000","10021":"#000000","10022":"#000000","10023":"#000000","10024":"#000000","10025":"#000000","10026":"#000000","10027":"#000000","10028":"#000000","10029":"#000000","10030":"#000000","10031":"#000000","40001":"#3a2a1a","40002":"#3a3a3a","40003":"#4a4436","40004":"#4a3a3a","40005":"#1a3a1a","40006":"#3a3a2a","40007":"#1a3a3a","40008":"#3a2a3a","40009":"#4a2010","40010":"#1a3a1a","40011":"#1a4a2a","40012":"#4a1a1a","40013":"#1a3a4a","40014":"#3a3020","40015":"#3a3a3a","40016":"#000000","40017":"#000000","40018":"#000000","40019":"#000000","40020":"#000000","40021":"#000000","40022":"#000000","40023":"#000000","40024":"#000000","40025":"#000000","40026":"#000000","40027":"#000000","40028":"#000000","40029":"#000000","40030":"#000000","40031":"#000000","40032":"#000000","40033":"#000000","40034":"#000000","40035":"#000000","40036":"#000000","40037":"#000000","40038":"#000000","40039":"#000000","40040":"#000000","40041":"#000000","40042":"#000000","40043":"#000000","40044":"#000000","40045":"#000000","40046":"#000000","40047":"#000000","40048":"#000000","40049":"#000000","40050":"#000000","40051":"#000000","40052":"#000000","40053":"#000000","40054":"#000000","40055":"#000000","40056":"#000000","40057":"#000000","40058":"#000000","40059":"#000000","40060":"#000000","40061":"#000000","40062":"#000000","40063":"#000000","40064":"#000000","40065":"#000000","40066":"#000000","40067":"#000000","40068":"#000000","40069":"#000000","40070":"#000000","40071":"#000000","40072":"#000000","40073":"#000000","40074":"#000000","40075":"#000000","40076":"#000000","40077":"#000000","40078":"#000000","40079":"#000000","40080":"#000000","50000":"#2a1020","50001":"#2a2a2a","50002":"#4a3a10","50003":"#4a1a10","50004":"#000000","50005":"#000000","50006":"#000000","50007":"#000000","50008":"#000000","50009":"#000000","50010":"#000000","50011":"#000000","50012":"#000000","50013":"#000000","50014":"#000000","50015":"#000000","60001":"#000000","60002":"#000000","60003":"#000000","60004":"#000000","60005":"#000000","60006":"#000000","60007":"#000000","60008":"#000000","60009":"#000000","60010":"#000000","60011":"#000000","60012":"#000000","90001":"#000000","90002":"#000000","90003":"#000000","90004":"#000000","90005":"#000000","90006":"#000000","90007":"#1a1a2a","90008":"#000000","90009":"#000000","90010":"#000000","90011":"#000000","90012":"#000000","90013":"#000000"};
const CARD_POSITION_MAP = {"90001":"打","90002":"坦","90003":"打","90004":"坦","90005":"打","90006":"補","90007":"打","90008":"打","90009":"坦","90010":"打","90011":"坦","90012":"打","90013":"補"};
const CARD_SYSTEM_OVERRIDE = {60001:"20.0"}; // 原始資料的既有格式差異（其餘 state 卡為純數字字串）

function parseNumList(v) {
  if (v === null || v === undefined || v === "") return undefined;
  return String(v).split(",").map(s => Number(s.trim()));
}
function parseBraceList(v) {
  if (v === null || v === undefined) return undefined;
  const matches = String(v).match(/\{([^}]*)\}/g) || [];
  return matches.map(m => Number(m.slice(1, -1)));
}

// 2026-09-09新增：二階職業(90008~90013)在card分頁裡其實已經有列(名稱/rare/type/體系/equipment_type/
// skill0~2都填好了，且對應的skill_id在skill分頁裡也都確實存在)，但這6列的「角色id」欄位本身是空白，
// 導致原本的篩選條件(typeof o["角色id"]==="number")直接把這6列整個濾掉，遊戲裡完全找不到這6張卡
// (Wei回報「職業和職業的技能全部...有問題」時看到的，其實就是只剩下6個一階職業、二階職業完全不存在)。
// 這裡先用「名稱」把這6列比對回Wei訊息裡給的對照表(90008劍豪/90009狂戰士/90010神射手/90011刺客/
// 90012魔導士/90013神官)，直接補上id——這個對照關係是Wei親自在訊息裡列出來的，不是我自行判斷。
const ADVANCED_CLASS_NAME_TO_ID = { "劍豪": 90008, "狂戰士": 90009, "神射手": 90010, "刺客": 90011, "魔導士": 90012, "神官": 90013 };
// 二階對應的一階職業id(用來在數值完全空白時借用一階基礎值*1.3當作暫定數值，避免「有id/有技能但hp=atk=0」
// 這種比原本完全不存在還糟的殘廢卡片狀態——玩家透過天賦樹解鎖後至少能正常上場，不會出現0血0攻擊的怪異卡)
const ADVANCED_CLASS_BASE_ID = { 90008: 90001, 90009: 90002, 90010: 90003, 90011: 90004, 90012: 90005, 90013: 90006 };

function buildCardDatabase(wb) {
  const rawObjs = xlsxSheetToObjects(wb, "card");
  rawObjs.forEach(o => {
    if ((o["角色id"] === null || o["角色id"] === undefined) && o.type === "role") {
      const name = (o["角色名稱"] || "").trim();
      if (ADVANCED_CLASS_NAME_TO_ID[name] !== undefined) o["角色id"] = ADVANCED_CLASS_NAME_TO_ID[name];
    }
  });
  const objs = rawObjs.filter(o => typeof o["角色id"] === "number");
  const list = objs.map(s => {
    const id = s["角色id"];
    const o = {
      id,
      rare: s.rare === null ? "" : s.rare,
      name: (s["角色名稱"] === null ? "" : s["角色名稱"]).trim(),
      icon: buildCardIconHtml(id, s.card_prefeb),
      iconBg: CARD_ICONBG_MAP[id] !== undefined ? CARD_ICONBG_MAP[id] : "#000000",
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
    o.cd = (s["重生回合數"] === null || s["重生回合數"] === undefined) ? null : s["重生回合數"];
    if (s.skill0_id !== null) o.s0 = s.skill0_id;
    if (s.skill1_id !== null) o.s1 = s.skill1_id;
    if (s.skill2_id !== null) o.s2 = s.skill2_id;
    o.isRole = s.type === "role";
    if (s.exchange_id !== null) o.exchangeId = s.exchange_id;
    if (s.exchange_amount !== null) o.exchangeAmount = s.exchange_amount;
    if (s.star_refine_id !== null) o.starRefineId = Number(s.star_refine_id);
    if (s.star_cost_id !== null) o.starCostId = s.star_cost_id;
    if (s.equipment_type !== null) o.equipmentType = parseNumList(s.equipment_type);
    const ss1 = {};
    [1, 2, 3, 4, 5].forEach(n => { const v = s["star" + n + "_skill1"]; if (v !== null) ss1[n] = v; });
    if (Object.keys(ss1).length) o.starSkill1 = ss1;
    const ss2 = {};
    [1, 2, 3, 4, 5].forEach(n => { const v = s["star" + n + "_skill2"]; if (v !== null) ss2[n] = v; });
    if (Object.keys(ss2).length) o.starSkill2 = ss2;
    if (s["體系"] !== null) o.system = CARD_SYSTEM_OVERRIDE[id] !== undefined ? CARD_SYSTEM_OVERRIDE[id] : (typeof s["體系"] === "number" ? String(s["體系"]) : s["體系"]);
    if (CARD_POSITION_MAP[id] !== undefined) o.position = CARD_POSITION_MAP[id];
    if (s.resist_buff !== null) o.resistBuff = parseNumList(s.resist_buff);
    if (s["role_ addition"] !== null) o.roleAddition = parseBraceList(s["role_ addition"]);
    // 捕捉機率(card表 capture_rate 欄位)：只有魔物(monster)才會填此欄位，有填才能被捕捉，沒填(null)則不可捕捉
    if (s.capture_rate !== null && s.capture_rate !== undefined) o.captureRate = s.capture_rate;
    // 2026-09-10新增：atk_motion(近戰動作類型，跟著武器/卡片走，不是跟著技能走)。採用overwrite規則：
    // 沒填就是原本的動作(刺/atk_stab)，有填才套用對應的新動作，見index.html的ATK_MOTION_CLASS對照表。
    if (s.atk_motion !== null && s.atk_motion !== undefined && String(s.atk_motion).trim() !== '') {
      o.atkMotion = String(s.atk_motion).trim();
    }
    return o;
  });
  // 二階職業(90008~90013)這6張卡雖然透過上面的名稱比對補回了id，但card分頁本身完全沒有填hp/atk/matk/
  // def/mdef/agi(全部空白)，會被上面統一套用「null就當0」，若不處理，玩家透過天賦樹解鎖後上場會是一張
  // 0血0攻擊、比原本完全拿不到還要糟的殘廢卡。這裡只在「數值確實全部是0(=xlsx全空白，不是Wei刻意設計
  // 成0)」時才借用對應一階職業(見ADVANCED_CLASS_BASE_ID)的基礎數值*1.3當暫定值，讓卡片至少能正常上場，
  // 一旦Wei之後在card分頁把這6張卡自己的hp/atk/matk/def/mdef/agi填上實際數字，這裡會偵測到不再是
  // 全0而不覆蓋，自動改用xlsx的真實數值。
  const byId = {};
  list.forEach(c => { byId[c.id] = c; });
  Object.entries(ADVANCED_CLASS_BASE_ID).forEach(([advId, baseId]) => {
    const adv = byId[advId];
    const base = byId[baseId];
    if (!adv || !base) return;
    const allZero = !adv.hp && !adv.atk && !adv.matk && !adv.def && !adv.mdef;
    if (allZero) {
      adv.hp = Math.round(base.hp * 1.3);
      adv.atk = Math.round(base.atk * 1.3);
      adv.matk = Math.round(base.matk * 1.3);
      adv.def = Math.round(base.def * 1.3);
      adv.mdef = Math.round(base.mdef * 1.3);
      adv.agi = base.agi + 2;
      adv._statsArePlaceholder = true; // 供UI或未來檢查用：標記這是暫定數值、不是Wei自己填的正式數字
    }
  });
  return list;
}

/** EQUIPMENT_DATABASE 相關：icon/iconBg/passive* 為美術與被動效果描述，xlsx 未包含，沿用內建對照表 */
const EQUIP_ICON_MAP = {"70001":"🔥","70002":"🔥","70003":"🔥","70004":"🔥","70005":"🔥","70006":"🔥","70007":"🔥","70008":"🔥","70009":"🔥","70010":"🔥","70011":"✨","70012":"✨","70013":"✨","70014":"✨","70015":"✨","70016":"✨","70017":"✨","70018":"✨","70019":"✨","70020":"✨","70021":"💠","70022":"💠","70023":"💠","70024":"💠","70025":"💠","70026":"💠","70027":"💠","70028":"💠","70029":"💠","70030":"💠","70031":"💎","70032":"💎","70033":"💎","70034":"💎","70035":"💎","70036":"💎","70037":"💎","70038":"💎","70039":"💎","70040":"💎","70041":"💎","70042":"💎","70043":"💎","70044":"💎","70045":"💎","70046":"💎","70047":"💎","70048":"💎","70049":"💎","70050":"💎","70051":"💎","70052":"💎","70053":"💎","70054":"💎","70055":"💎","70056":"💎","70057":"💎","70058":"💎","70059":"💎","70060":"💎","70061":"💎","70062":"💎","70063":"💎","70064":"💎","70065":"💎","70066":"💎","70067":"💎","70068":"💎","70069":"💎","70070":"💎","70071":"💎","70072":"💎","70073":"💎","70074":"💎","70075":"💎","70076":"💎","70077":"💎","70078":"💎","70079":"💎","70080":"💎","80001":"🗡️","80002":"🗡️","80003":"🗡️","80004":"🗡️","80005":"🗡️","80006":"🗡️","80007":"🗡️","80008":"🗡️","80009":"🗡️","80010":"🗡️","80011":"🗡️","80012":"🗡️","81001":"👒","81002":"🎩","81003":"👒","81004":"👒","81005":"👒","81006":"👒","81007":"👒","81008":"👒","81009":"👒","81010":"👒","81011":"👒","81012":"👒","82001":"👗","82002":"👕","82003":"👗","82004":"👗","82005":"👗","82006":"👗","82007":"👗","82008":"👗","82009":"👗","82010":"👗","82011":"👗","82012":"👗","83001":"👞","83002":"👞","83003":"👞","83004":"👞","83005":"👞","83006":"👞","83007":"👞","83008":"👞","83009":"👞","83010":"👞","83011":"👞","83012":"👞","84001":"🧣","84002":"🧣","84003":"🧣","84004":"🧣","84005":"🧣","84006":"🧣","84007":"🧣","84008":"🧣","84009":"🧣","84010":"🧣","84011":"🧣","84012":"🧣","85001":"👖","85002":"👖","85003":"👖","85004":"👖","85005":"👖","85006":"👖","85007":"👖","85008":"👖","85009":"👖","85010":"👖","85011":"👖","85012":"👖","88000":"🗡️","88001":"🗡️","88002":"🗡️","88003":"👒","88004":"👗","88005":"👞","88006":"🧣","88007":"👖","88008":"🗡️","88009":"👒","88010":"👗","88011":"👞","88012":"🧣","88013":"👖","88014":"🗡️","88015":"👒","88016":"👗","88017":"👞","88018":"🧣","88019":"👖","88020":"🗡️","88021":"👒","88022":"👗","88023":"👞","88024":"🧣","88025":"👖","88026":"🗡️","88027":"👒","88028":"👗","88029":"👞","88030":"🧣","88031":"👖","88032":"🗡️","88033":"🗡️","88034":"🗡️","88035":"👒","88036":"👗","88037":"👞","88038":"🧣","88039":"👖","88040":"🗡️","88041":"🗡️"};
const EQUIP_ICONBG_MAP = {"70001":"#4a2410","70002":"#4a2410","70003":"#4a2410","70004":"#4a2410","70005":"#4a2410","70006":"#4a2410","70007":"#4a2410","70008":"#4a2410","70009":"#4a2410","70010":"#4a2410","70011":"#164f2a","70012":"#164f2a","70013":"#164f2a","70014":"#164f2a","70015":"#164f2a","70016":"#164f2a","70017":"#164f2a","70018":"#164f2a","70019":"#164f2a","70020":"#164f2a","70021":"#10304a","70022":"#10304a","70023":"#10304a","70024":"#10304a","70025":"#10304a","70026":"#10304a","70027":"#10304a","70028":"#10304a","70029":"#10304a","70030":"#10304a","70031":"#2a2a4a","70032":"#2a2a4a","70033":"#2a2a4a","70034":"#2a2a4a","70035":"#2a2a4a","70036":"#2a2a4a","70037":"#2a2a4a","70038":"#2a2a4a","70039":"#2a2a4a","70040":"#2a2a4a","70041":"#2a2a4a","70042":"#2a2a4a","70043":"#2a2a4a","70044":"#2a2a4a","70045":"#2a2a4a","70046":"#2a2a4a","70047":"#2a2a4a","70048":"#2a2a4a","70049":"#2a2a4a","70050":"#2a2a4a","70051":"#2a2a4a","70052":"#2a2a4a","70053":"#2a2a4a","70054":"#2a2a4a","70055":"#2a2a4a","70056":"#2a2a4a","70057":"#2a2a4a","70058":"#2a2a4a","70059":"#2a2a4a","70060":"#2a2a4a","70061":"#2a2a4a","70062":"#2a2a4a","70063":"#2a2a4a","70064":"#2a2a4a","70065":"#2a2a4a","70066":"#2a2a4a","70067":"#2a2a4a","70068":"#2a2a4a","70069":"#2a2a4a","70070":"#2a2a4a","70071":"#2a2a4a","70072":"#2a2a4a","70073":"#2a2a4a","70074":"#2a2a4a","70075":"#2a2a4a","70076":"#2a2a4a","70077":"#2a2a4a","70078":"#2a2a4a","70079":"#2a2a4a","70080":"#2a2a4a","80001":"#4a2410","80002":"#4a3520","80003":"#4a2410","80004":"#4a2410","80005":"#4a2410","80006":"#4a2410","80007":"#4a2410","80008":"#4a2410","80009":"#4a2410","80010":"#4a2410","80011":"#4a2410","80012":"#4a2410","81001":"#4a3a12","81002":"#4a3520","81003":"#4a3a12","81004":"#4a3a12","81005":"#4a3a12","81006":"#4a3a12","81007":"#4a3a12","81008":"#4a3a12","81009":"#4a3a12","81010":"#4a3a12","81011":"#4a3a12","81012":"#4a3a12","82001":"#164f2a","82002":"#4a3520","82003":"#164f2a","82004":"#164f2a","82005":"#164f2a","82006":"#164f2a","82007":"#164f2a","82008":"#164f2a","82009":"#164f2a","82010":"#164f2a","82011":"#164f2a","82012":"#164f2a","83001":"#3a2f24","83002":"#4a3520","83003":"#3a2f24","83004":"#3a2f24","83005":"#3a2f24","83006":"#3a2f24","83007":"#3a2f24","83008":"#3a2f24","83009":"#3a2f24","83010":"#3a2f24","83011":"#3a2f24","83012":"#3a2f24","84001":"#3a1f4a","84002":"#4a3520","84003":"#3a1f4a","84004":"#3a1f4a","84005":"#3a1f4a","84006":"#3a1f4a","84007":"#3a1f4a","84008":"#3a1f4a","84009":"#3a1f4a","84010":"#3a1f4a","84011":"#3a1f4a","84012":"#3a1f4a","85001":"#1f2f4a","85002":"#4a3520","85003":"#1f2f4a","85004":"#1f2f4a","85005":"#1f2f4a","85006":"#1f2f4a","85007":"#1f2f4a","85008":"#1f2f4a","85009":"#1f2f4a","85010":"#1f2f4a","85011":"#1f2f4a","85012":"#1f2f4a","88000":"#4a2410","88001":"#4a2410","88002":"#4a2410","88003":"#4a3a12","88004":"#164f2a","88005":"#3a2f24","88006":"#3a1f4a","88007":"#1f2f4a","88008":"#4a2410","88009":"#4a3a12","88010":"#164f2a","88011":"#3a2f24","88012":"#3a1f4a","88013":"#1f2f4a","88014":"#4a2410","88015":"#4a3a12","88016":"#164f2a","88017":"#3a2f24","88018":"#3a1f4a","88019":"#1f2f4a","88020":"#4a2410","88021":"#4a3a12","88022":"#164f2a","88023":"#3a2f24","88024":"#3a1f4a","88025":"#1f2f4a","88026":"#4a2410","88027":"#4a3a12","88028":"#164f2a","88029":"#3a2f24","88030":"#3a1f4a","88031":"#1f2f4a","88032":"#4a2410","88033":"#4a2410","88034":"#4a2410","88035":"#4a3a12","88036":"#164f2a","88037":"#3a2f24","88038":"#3a1f4a","88039":"#1f2f4a","88040":"#4a2410","88041":"#4a2410"};
const EQUIP_PASSIVE_MAP = {"70001":{"key":"atkPct","value":0.01,"desc":"戰鬥開始時裝備武器的攻擊力+1%"},"70002":{"key":"atkPct","value":0.02,"desc":"戰鬥開始時裝備武器的攻擊力+2%"},"70003":{"key":"atkPct","value":0.03,"desc":"戰鬥開始時裝備武器的攻擊力+3%"},"70004":{"key":"atkPct","value":0.04,"desc":"戰鬥開始時裝備武器的攻擊力+4%"},"70005":{"key":"atkPct","value":0.05,"desc":"戰鬥開始時裝備武器的攻擊力+5%"},"70006":{"key":"atkPct","value":0.06,"desc":"戰鬥開始時裝備武器的攻擊力+6%"},"70007":{"key":"atkPct","value":0.07,"desc":"戰鬥開始時裝備武器的攻擊力+7%"},"70008":{"key":"atkPct","value":0.08,"desc":"戰鬥開始時裝備武器的攻擊力+8%"},"70009":{"key":"atkPct","value":0.09,"desc":"戰鬥開始時裝備武器的攻擊力+9%"},"70010":{"key":"atkPct","value":0.1,"desc":"戰鬥開始時裝備武器的攻擊力+10%"},"70011":{"key":"matkPct","value":0.01,"desc":"戰鬥開始時裝備武器的魔法攻擊力+1%"},"70012":{"key":"matkPct","value":0.02,"desc":"戰鬥開始時裝備武器的魔法攻擊力+2%"},"70013":{"key":"matkPct","value":0.03,"desc":"戰鬥開始時裝備武器的魔法攻擊力+3%"},"70014":{"key":"matkPct","value":0.04,"desc":"戰鬥開始時裝備武器的魔法攻擊力+4%"},"70015":{"key":"matkPct","value":0.05,"desc":"戰鬥開始時裝備武器的魔法攻擊力+5%"},"70016":{"key":"matkPct","value":0.06,"desc":"戰鬥開始時裝備武器的魔法攻擊力+6%"},"70017":{"key":"matkPct","value":0.07,"desc":"戰鬥開始時裝備武器的魔法攻擊力+7%"},"70018":{"key":"matkPct","value":0.08,"desc":"戰鬥開始時裝備武器的魔法攻擊力+8%"},"70019":{"key":"matkPct","value":0.09,"desc":"戰鬥開始時裝備武器的魔法攻擊力+9%"},"70020":{"key":"matkPct","value":0.1,"desc":"戰鬥開始時裝備武器的魔法攻擊力+10%"},"70021":{"key":"defMdefPct","value":0.01,"desc":"戰鬥開始時裝備武器的防禦力與魔防+1%"},"70022":{"key":"defMdefPct","value":0.02,"desc":"戰鬥開始時裝備武器的防禦力與魔防+2%"},"70023":{"key":"defMdefPct","value":0.03,"desc":"戰鬥開始時裝備武器的防禦力與魔防+3%"},"70024":{"key":"defMdefPct","value":0.04,"desc":"戰鬥開始時裝備武器的防禦力與魔防+4%"},"70025":{"key":"defMdefPct","value":0.05,"desc":"戰鬥開始時裝備武器的防禦力與魔防+5%"},"70026":{"key":"defMdefPct","value":0.06,"desc":"戰鬥開始時裝備武器的防禦力與魔防+6%"},"70027":{"key":"defMdefPct","value":0.07,"desc":"戰鬥開始時裝備武器的防禦力與魔防+7%"},"70028":{"key":"defMdefPct","value":0.08,"desc":"戰鬥開始時裝備武器的防禦力與魔防+8%"},"70029":{"key":"defMdefPct","value":0.09,"desc":"戰鬥開始時裝備武器的防禦力與魔防+9%"},"70030":{"key":"defMdefPct","value":0.1,"desc":"戰鬥開始時裝備武器的防禦力與魔防+10%"},"70031":{"key":"defPct","value":0.02,"desc":"戰鬥開始時裝備者的物防增加2%"},"70032":{"key":"defPct","value":0.03,"desc":"戰鬥開始時裝備者的物防增加3%"},"70033":{"key":"defPct","value":0.04,"desc":"戰鬥開始時裝備者的物防增加4%"},"70034":{"key":"defPct","value":0.05,"desc":"戰鬥開始時裝備者的物防增加5%"},"70035":{"key":"defPct","value":0.06,"desc":"戰鬥開始時裝備者的物防增加6%"},"70036":{"key":"defPct","value":0.07,"desc":"戰鬥開始時裝備者的物防增加7%"},"70037":{"key":"defPct","value":0.09,"desc":"戰鬥開始時裝備者的物防增加9.%"},"70038":{"key":"defPct","value":0.11,"desc":"戰鬥開始時裝備者的物防增加11.%"},"70039":{"key":"defPct","value":0.13,"desc":"戰鬥開始時裝備者的物防增加13%"},"70040":{"key":"defPct","value":0.15,"desc":"戰鬥開始時裝備者的物防增加15%"},"70041":{"key":"skillRatePct","value":0.05,"desc":"戰鬥開始時裝備者的技能傷害增加5%"},"70042":{"key":"skillRatePct","value":0.075,"desc":"戰鬥開始時裝備者的技能傷害增加7.5%"},"70043":{"key":"skillRatePct","value":0.1,"desc":"戰鬥開始時裝備者的技能傷害增加10%"},"70044":{"key":"skillRatePct","value":0.125,"desc":"戰鬥開始時裝備者的技能傷害增加12.5%"},"70045":{"key":"skillRatePct","value":0.175,"desc":"戰鬥開始時裝備者的技能傷害增加17.5%"},"70046":{"key":"skillRatePct","value":0.225,"desc":"戰鬥開始時裝備者的技能傷害增加22.5%"},"70047":{"key":"skillRatePct","value":0.275,"desc":"戰鬥開始時裝備者的技能傷害增加27.5%"},"70048":{"key":"skillRatePct","value":0.345,"desc":"戰鬥開始時裝備者的技能傷害增加34.5%"},"70049":{"key":"skillRatePct","value":0.415,"desc":"戰鬥開始時裝備者的技能傷害增加41.5%"},"70050":{"key":"skillRatePct","value":0.485,"desc":"戰鬥開始時裝備者的技能傷害增加48.5%"},"70051":{"key":"skillResistPct","value":0.08,"desc":"戰鬥開始時裝備者的技能傷害抗性增加8%"},"70052":{"key":"skillResistPct","value":0.11,"desc":"戰鬥開始時裝備者的技能傷害抗性增加11%"},"70053":{"key":"skillResistPct","value":0.14,"desc":"戰鬥開始時裝備者的技能傷害抗性增加14.%"},"70054":{"key":"skillResistPct","value":0.17,"desc":"戰鬥開始時裝備者的技能傷害抗性增加17%"},"70055":{"key":"skillResistPct","value":0.2,"desc":"戰鬥開始時裝備者的技能傷害抗性增加20%"},"70056":{"key":"skillResistPct","value":0.25,"desc":"戰鬥開始時裝備者的技能傷害抗性增加25%"},"70057":{"key":"skillResistPct","value":0.3,"desc":"戰鬥開始時裝備者的技能傷害抗性增加30%"},"70058":{"key":"skillResistPct","value":0.37,"desc":"戰鬥開始時裝備者的技能傷害抗性增加37%"},"70059":{"key":"skillResistPct","value":0.44,"desc":"戰鬥開始時裝備者的技能傷害抗性增加44%"},"70060":{"key":"skillResistPct","value":0.51,"desc":"戰鬥開始時裝備者的技能傷害抗性增加51%"},"70061":{"key":"damageIncreasePct","value":0.05,"desc":"戰鬥開始時裝備者的爆擊傷害增加5%"},"70062":{"key":"damageIncreasePct","value":0.075,"desc":"戰鬥開始時裝備者的爆擊傷害增加7.5%"},"70063":{"key":"damageIncreasePct","value":0.1,"desc":"戰鬥開始時裝備者的爆擊傷害增加10%"},"70064":{"key":"damageIncreasePct","value":0.125,"desc":"戰鬥開始時裝備者的爆擊傷害增加12.5%"},"70065":{"key":"damageIncreasePct","value":0.175,"desc":"戰鬥開始時裝備者的爆擊傷害增加17.5%"},"70066":{"key":"damageIncreasePct","value":0.225,"desc":"戰鬥開始時裝備者的爆擊傷害增加22.5%"},"70067":{"key":"damageIncreasePct","value":0.275,"desc":"戰鬥開始時裝備者的爆擊傷害增加27.5%"},"70068":{"key":"damageIncreasePct","value":0.345,"desc":"戰鬥開始時裝備者的爆擊傷害增加34.5%"},"70069":{"key":"damageIncreasePct","value":0.415,"desc":"戰鬥開始時裝備者的爆擊傷害增加41.5%"},"70070":{"key":"damageIncreasePct","value":0.485,"desc":"戰鬥開始時裝備者的爆擊傷害增加48.5%"},"70071":{"key":"damageReducePct","value":0.08,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加8%"},"70072":{"key":"damageReducePct","value":0.11,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加11%"},"70073":{"key":"damageReducePct","value":0.14,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加14.%"},"70074":{"key":"damageReducePct","value":0.17,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加17%"},"70075":{"key":"damageReducePct","value":0.2,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加20%"},"70076":{"key":"damageReducePct","value":0.25,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加25%"},"70077":{"key":"damageReducePct","value":0.3,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加30%"},"70078":{"key":"damageReducePct","value":0.37,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加37%"},"70079":{"key":"damageReducePct","value":0.44,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加44%"},"70080":{"key":"damageReducePct","value":0.51,"desc":"戰鬥開始時裝備者的爆擊傷害抗性增加51%"},"80001":{"key":null,"value":0,"desc":""},"80002":{"key":null,"value":0,"desc":""},"80003":{"key":null,"value":0,"desc":""},"80004":{"key":null,"value":0,"desc":""},"80005":{"key":null,"value":0,"desc":""},"80006":{"key":null,"value":0,"desc":""},"80007":{"key":null,"value":0,"desc":""},"80008":{"key":null,"value":0,"desc":""},"80009":{"key":null,"value":0,"desc":""},"80010":{"key":null,"value":0,"desc":""},"80011":{"key":null,"value":0,"desc":""},"80012":{"key":null,"value":0,"desc":""},"81001":{"key":null,"value":0,"desc":""},"81002":{"key":null,"value":0,"desc":""},"81003":{"key":null,"value":0,"desc":""},"81004":{"key":null,"value":0,"desc":""},"81005":{"key":null,"value":0,"desc":""},"81006":{"key":null,"value":0,"desc":""},"81007":{"key":null,"value":0,"desc":""},"81008":{"key":null,"value":0,"desc":""},"81009":{"key":null,"value":0,"desc":""},"81010":{"key":null,"value":0,"desc":""},"81011":{"key":null,"value":0,"desc":""},"81012":{"key":null,"value":0,"desc":""},"82001":{"key":null,"value":0,"desc":""},"82002":{"key":null,"value":0,"desc":""},"82003":{"key":null,"value":0,"desc":""},"82004":{"key":null,"value":0,"desc":""},"82005":{"key":null,"value":0,"desc":""},"82006":{"key":null,"value":0,"desc":""},"82007":{"key":null,"value":0,"desc":""},"82008":{"key":null,"value":0,"desc":""},"82009":{"key":null,"value":0,"desc":""},"82010":{"key":null,"value":0,"desc":""},"82011":{"key":null,"value":0,"desc":""},"82012":{"key":null,"value":0,"desc":""},"83001":{"key":null,"value":0,"desc":""},"83002":{"key":null,"value":0,"desc":""},"83003":{"key":null,"value":0,"desc":""},"83004":{"key":null,"value":0,"desc":""},"83005":{"key":null,"value":0,"desc":""},"83006":{"key":null,"value":0,"desc":""},"83007":{"key":null,"value":0,"desc":""},"83008":{"key":null,"value":0,"desc":""},"83009":{"key":null,"value":0,"desc":""},"83010":{"key":null,"value":0,"desc":""},"83011":{"key":null,"value":0,"desc":""},"83012":{"key":null,"value":0,"desc":""},"84001":{"key":null,"value":0,"desc":""},"84002":{"key":null,"value":0,"desc":""},"84003":{"key":null,"value":0,"desc":""},"84004":{"key":null,"value":0,"desc":""},"84005":{"key":null,"value":0,"desc":""},"84006":{"key":null,"value":0,"desc":""},"84007":{"key":null,"value":0,"desc":""},"84008":{"key":null,"value":0,"desc":""},"84009":{"key":null,"value":0,"desc":""},"84010":{"key":null,"value":0,"desc":""},"84011":{"key":null,"value":0,"desc":""},"84012":{"key":null,"value":0,"desc":""},"85001":{"key":null,"value":0,"desc":""},"85002":{"key":null,"value":0,"desc":""},"85003":{"key":null,"value":0,"desc":""},"85004":{"key":null,"value":0,"desc":""},"85005":{"key":null,"value":0,"desc":""},"85006":{"key":null,"value":0,"desc":""},"85007":{"key":null,"value":0,"desc":""},"85008":{"key":null,"value":0,"desc":""},"85009":{"key":null,"value":0,"desc":""},"85010":{"key":null,"value":0,"desc":""},"85011":{"key":null,"value":0,"desc":""},"85012":{"key":null,"value":0,"desc":""},"88000":{"key":null,"value":0,"desc":""},"88001":{"key":null,"value":0,"desc":""},"88002":{"key":null,"value":0,"desc":""},"88003":{"key":null,"value":0,"desc":""},"88004":{"key":null,"value":0,"desc":""},"88005":{"key":null,"value":0,"desc":""},"88006":{"key":null,"value":0,"desc":""},"88007":{"key":null,"value":0,"desc":""},"88008":{"key":null,"value":0,"desc":""},"88009":{"key":null,"value":0,"desc":""},"88010":{"key":null,"value":0,"desc":""},"88011":{"key":null,"value":0,"desc":""},"88012":{"key":null,"value":0,"desc":""},"88013":{"key":null,"value":0,"desc":""},"88014":{"key":null,"value":0,"desc":""},"88015":{"key":null,"value":0,"desc":""},"88016":{"key":null,"value":0,"desc":""},"88017":{"key":null,"value":0,"desc":""},"88018":{"key":null,"value":0,"desc":""},"88019":{"key":null,"value":0,"desc":""},"88020":{"key":null,"value":0,"desc":""},"88021":{"key":null,"value":0,"desc":""},"88022":{"key":null,"value":0,"desc":""},"88023":{"key":null,"value":0,"desc":""},"88024":{"key":null,"value":0,"desc":""},"88025":{"key":null,"value":0,"desc":""},"88026":{"key":null,"value":0,"desc":""},"88027":{"key":null,"value":0,"desc":""},"88028":{"key":null,"value":0,"desc":""},"88029":{"key":null,"value":0,"desc":""},"88030":{"key":null,"value":0,"desc":""},"88031":{"key":null,"value":0,"desc":""},"88032":{"key":null,"value":0,"desc":""},"88033":{"key":null,"value":0,"desc":""},"88034":{"key":null,"value":0,"desc":""},"88035":{"key":null,"value":0,"desc":""},"88036":{"key":null,"value":0,"desc":""},"88037":{"key":null,"value":0,"desc":""},"88038":{"key":null,"value":0,"desc":""},"88039":{"key":null,"value":0,"desc":""},"88040":{"key":null,"value":0,"desc":""},"88041":{"key":null,"value":0,"desc":""}};

function buildEquipmentDatabase(wb) {
  const objs = xlsxSheetToObjects(wb, "equipment").filter(o => typeof o.id === "number");
  const gen = {};
  objs.forEach(s => {
    const id = s.id;
    const recipe = parseNumList(s.recipe);
    const build = parseNumList(s.build);
    const passive = EQUIP_PASSIVE_MAP[id] || { key: null, value: 0, desc: "" };
    gen[id] = {
      id, type: s.type, name: s.info, rare: s.rare,
      tag: (s.tag === null || s.tag === undefined) ? "" : s.tag,
      hp: (s.HP === null || s.HP === "") ? 0 : s.HP,
      atk: (s.ATK === null || s.ATK === "") ? 0 : s.ATK,
      matk: (s.MATK === null || s.MATK === "") ? 0 : s.MATK,
      def: (s.DEF === null || s.DEF === "") ? 0 : s.DEF,
      mdef: (s.MDEF === null || s.MDEF === "") ? 0 : s.MDEF,
      mp: (s.MP === null || s.MP === "") ? 0 : s.MP,
      agi: (s.AGI === null || s.AGI === "") ? 0 : s.AGI,
      recipe: recipe === undefined ? null : recipe,
      recipeRate: (s.recipe_rate === null || s.recipe_rate === undefined) ? 1 : s.recipe_rate,
      buffPasiveIds: parseNumList(s.buff_pasive),
      refineId: s.refine_id, costId: s.cost_id,
      build: build === undefined ? null : build,
      icon: buildIconWithFallback('equipment', id, EQUIP_ICON_MAP[id] !== undefined ? EQUIP_ICON_MAP[id] : "❔"),
      iconBg: EQUIP_ICONBG_MAP[id] !== undefined ? EQUIP_ICONBG_MAP[id] : "#000000",
      passiveKey: passive.key, passiveValue: passive.value, passiveDesc: passive.desc,
    };
  });
  return gen;
}

function buildGachaDatabase(wb) {
  const ws = wb.Sheets["gacha"];
  if (!ws) { console.warn("[xlsx-loader] 找不到分頁: gacha"); return []; }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const headerIdx = rows.findIndex(r => r.some(v => v !== null));
  const dataRows = rows.slice(headerIdx + 1).filter(r => r.some(v => v !== null) && typeof r[0] === "number");
  return dataRows.map(r => ({ id: r[0], power: r[1], cardId: r[2] }));
}

const BUILD_STAT_FIELD_MAP = {
  HP: "hp", ATK: "atk", MATK: "matk", DEF: "def", MDEF: "mdef", MP: "mp", AGI: "agi", CRI: "cri",
  cri_resist: "criResist", cri_damage: "criDamage", cri_damage_resist: "criDamageResist",
  skill_rate: "skillRate", skill_resist: "skillResist", damage_increase: "damageIncrease",
  damage_reduce: "damageReduce", damage_reflect: "damageReflect", mp_recove: "mpRecover",
  shield_rate: "shieldRate", permeate: "permeate", atypical_rate: "atypicalRate", heal_rate: "healRate",
};
function buildStatsFromRow(o) {
  const s = {};
  Object.keys(BUILD_STAT_FIELD_MAP).forEach(col => { if (o[col] !== null && o[col] !== undefined) s[BUILD_STAT_FIELD_MAP[col]] = o[col]; });
  return s;
}
function buildBuildDatabase(wb) {
  const objs = xlsxSheetToObjects(wb, "build").filter(o => typeof o.group_id === "number");
  const gen = {};
  objs.forEach(o => {
    const gid = o.group_id;
    (gen[gid] = gen[gid] || []).push({
      power: o.power,
      stats: buildStatsFromRow(o),
      rare: o.rare === null ? 1 : o.rare,
      valueOffset: o.value_offset === null ? 0 : o.value_offset,
    });
  });
  return gen;
}

/** SKILL_DATABASE */
const SKILL_TYPE_MAP = { phycial: "AD", magic: "AP", heal: "HEAL", shield: "SHIELD" };
const SKILL_ADD_BUFF_ID_OVERRIDE = { 400018: [4001] }; // xlsx 資料筆誤（"4001.4002" 應為逗號分隔，原始遊戲資料僅使用 4001）
function buildSkillDatabase(wb) {
  const objs = xlsxSheetToObjects(wb, "skill").filter(o => typeof o.skill_id === "number");
  const gen = {};
  objs.forEach(o => {
    const id = o.skill_id;
    const entry = {
      id, name: o["技能名稱"],
      ad: (o["AD(技能係數)"] === null || o["AD(技能係數)"] === undefined) ? 0 : o["AD(技能係數)"],
      ap: (o["AP(技能係數)"] === null || o["AP(技能係數)"] === undefined) ? 0 : o["AP(技能係數)"],
      cost: (o.MP_cost === null || o.MP_cost === undefined) ? 0 : Math.round(o.MP_cost * 1e6) / 1e6,
      type: SKILL_TYPE_MAP[o.type] || "BUFF",
    };
    const addBuffId = id === 400018 ? SKILL_ADD_BUFF_ID_OVERRIDE[id] : parseNumList(o.add_buff_id);
    if (addBuffId !== undefined) entry.add_buff_id = addBuffId;
    if (o.add_buff_rate_overwrite !== null && o.add_buff_rate_overwrite !== undefined) entry.add_buff_rate_overwrite = o.add_buff_rate_overwrite;
    const selfAddBuffId = parseNumList(o.self_add_buff_id);
    if (selfAddBuffId !== undefined) entry.self_add_buff_id = selfAddBuffId;
    entry.target = `0,${o["攻擊對象數量"]},${o["是否為隨機"] ? 1 : 0},${o["是否自身施放"] ? 1 : 0}`;
    if (o.prefeb_atk !== null && o.prefeb_atk !== undefined) entry.prefebAtk = o.prefeb_atk;
    if (o.prefeb_hit1 !== null && o.prefeb_hit1 !== undefined) entry.prefebHit1 = [o.prefeb_hit1];
    if (o.prefeb_hit2 !== null && o.prefeb_hit2 !== undefined) entry.prefebHit2 = [o.prefeb_hit2];
    const removeBuffId = parseNumList(o.remove_buff_id);
    if (removeBuffId !== undefined) entry.remove_buff_id = removeBuffId;
    if (o.remove_buff_rate_overwrite !== null && o.remove_buff_rate_overwrite !== undefined) entry.remove_buff_rate_overwrite = o.remove_buff_rate_overwrite;
    if (o.self_buff_rate_overwrite !== null && o.self_buff_rate_overwrite !== undefined) entry.self_buff_rate_overwrite = o.self_buff_rate_overwrite;
    entry.desc = o["說明"];
    if (o.MP_revocer !== null && o.MP_revocer !== undefined) entry.mp_recover = o.MP_revocer;
    if (o.cooldown !== null && o.cooldown !== undefined) entry.cooldown = o.cooldown;
    if (o["攻擊次數"] !== null && o["攻擊次數"] !== undefined && o["攻擊次數"] !== 1) entry.hits = o["攻擊次數"];
    if (o.current_hp !== null && o.current_hp !== undefined) entry.current_hp = o.current_hp;
    gen[id] = entry;
  });
  return gen;
}


/** QUEST_DATABASE + GOD_TRIAL_VARIANTS：mapdata 分頁的 quest 欄位使用內嵌小型語法
 * (例："{role:-1},{id:-1},{id:-1},...,{id:10002},{id:10004},...")
 * {role:N}：N===-1 表示玩家可自選職業；N<-1 表示強制指定職業 -N
 * {id:N}：N<0（且非 -1）表示玩家出戰卡欄位預先指定卡片 -N；N===-1 表示空欄位（僅計數）；N>=0 表示敵方單位 id
 * token 內其餘 key（level/atk/def/wave 等）為該單位的數值覆寫；wave="N" 用於連續多波次副本(annihilation)分組
 * {state:N} token 目前未對應任何輸出欄位，解析後捨棄（與原始資料行為一致）*/
function parseQuestTokens(questStr) {
  if (typeof questStr !== "string") return null;
  const tokens = questStr.match(/\{[^}]*\}/g) || [];
  return tokens.map(t => {
    const inner = t.slice(1, -1);
    const parts = inner.split(",").map(s => s.trim()).filter(Boolean);
    const obj = {};
    parts.forEach(p => {
      const sepIdx = p.search(/[:=]/);
      if (sepIdx === -1) return;
      const k = p.slice(0, sepIdx).trim();
      let v = p.slice(sepIdx + 1).trim();
      v = v.replace(/^"(.*)"$/, "$1");
      obj[k] = v;
    });
    return obj;
  });
}
function qNum(v) { return v === undefined ? undefined : Number(v); }
function qParseNumList(v) {
  if (v === null || v === undefined || v === "") return undefined;
  return String(v).split(",").map(x => Number(x.trim()));
}
function qBuildEvent(name, type, buff, time, target, rate) {
  if (name === null || name === undefined) return null;
  const buffIds = typeof buff === "number" ? [buff] : qParseNumList(buff);
  return {
    name, type,
    buffIds: buffIds === undefined ? undefined : buffIds,
    times: qParseNumList(time),
    rate: (rate === null || rate === undefined) ? 1 : rate,
    target: (target === null || target === undefined || target === "") ? null : target,
  };
}
function buildQuestEntry(s) {
  const o = { mapId: s.mapid, capability: null, type: s.type, chapter: s.chapter, hint: s.hint, level: s.level };
  o.capability = (s.Capability === null || s.Capability === undefined || s.Capability === "") ? null : s.Capability;
  o.requiresArrangement = !!s.arrangement;

  const tokens = parseQuestTokens(s.quest);
  let requiresPlayerRole = false, forcedPlayerRoleId = null;
  let playerWeaponSlots = 0, playerPresetUnits = [];
  let enemyIds = [], enemyUnitOverrides = [], waves;

  if (tokens) {
    const roleTok = tokens.find(t => "role" in t);
    if (roleTok) {
      const rv = qNum(roleTok.role);
      if (rv === -1) { requiresPlayerRole = true; forcedPlayerRoleId = null; }
      else { requiresPlayerRole = false; forcedPlayerRoleId = -rv; }
    }
    const idTokens = tokens.filter(t => "id" in t);
    const playerTokens = idTokens.filter(t => qNum(t.id) < 0);
    const enemyTokens = idTokens.filter(t => qNum(t.id) >= 0);

    playerWeaponSlots = playerTokens.length;
    playerTokens.forEach((t, idx) => {
      const idVal = qNum(t.id);
      if (idVal === -1) return;
      const overrides = {};
      Object.keys(t).forEach(k => { if (k !== "id") overrides[k] = qNum(t[k]); });
      playerPresetUnits.push({ slotIndex: idx, cardId: -idVal, overrides });
    });

    const hasWave = enemyTokens.some(t => "wave" in t);
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
        Object.keys(t).forEach(k => { if (k !== "id" && k !== "wave") ov[k] = qNum(t[k]); });
        return ov;
      });
    } else {
      enemyIds = enemyTokens.map(t => qNum(t.id));
      enemyUnitOverrides = enemyTokens.map(t => {
        const ov = {};
        Object.keys(t).forEach(k => { if (k !== "id") ov[k] = qNum(t[k]); });
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
  // loot3/loot4：2026-09新手引導整合計畫批次1b新增，主線首通獎勵(loot3=鑽石, loot4=職業經驗)用；
  // 其餘關卡類型該分頁原本沒有這兩欄資料(undefined)，quest.loot3Id等自然是undefined，showResultModal判斷時視同沒有掉落，不影響既有行為
  o.loot3Id = s.loot3_id; o.loot3Amount = s.loot3_amount; o.loot3Rate = s.loot3_rate;
  o.loot4Id = s.loot4_id; o.loot4Amount = s.loot4_amount; o.loot4Rate = s.loot4_rate;
  o.playerPresetUnits = playerPresetUnits;
  o.enemyUnitOverrides = enemyUnitOverrides;
  if (waves !== undefined) o.waves = waves;
  o.event1 = qBuildEvent(s.event1_name, s.event1_type, s.event1_buff, s.event1_time, s.event1_target, s.event1_rate);
  o.event2 = qBuildEvent(s.event2_name, s.event2_type, s.event2_buff, s.event2_time, s.event2_target, s.event2_rate);
  return o;
}
function buildQuestDatabaseAndGodTrial(wb) {
  const allRows = xlsxSheetToObjects(wb, "mapdata").filter(o => typeof o.mapid === "number");
  const nonGod = allRows.filter(o => o.type !== "godcharacter");
  const quest = nonGod.map(buildQuestEntry);

  const god = allRows.filter(o => o.type === "godcharacter");
  const byGroup = {};
  god.forEach(s => { (byGroup[s.mapid] = byGroup[s.mapid] || []).push(s); });
  const godTrial = {};
  Object.keys(byGroup).forEach(gid => {
    godTrial[gid] = byGroup[gid].map(s => {
      const tokens = parseQuestTokens(s.quest) || [];
      const idTokens = tokens.filter(t => "id" in t);
      const enemyTokens = idTokens.filter(t => qNum(t.id) >= 0);
      const enemyIds = enemyTokens.map(t => qNum(t.id));
      return {
        level: s.level,
        tagLimit: s.tag_limit,
        enemyIds,
        loot1Id: s.loot1_id,
        loot1Amount: s.loot1_amount,
        loot1Rate: (s.loot1_rate === null || s.loot1_rate === undefined) ? 1 : s.loot1_rate,
        // loot2/loot3：批次5新增，神格試煉補上item2008(破碎的轉職推薦書)與item8003(屬性點數)掉落
        loot2Id: s.loot2_id,
        loot2Amount: s.loot2_amount,
        loot2Rate: (s.loot2_rate === null || s.loot2_rate === undefined) ? 1 : s.loot2_rate,
        loot3Id: s.loot3_id,
        loot3Amount: s.loot3_amount,
        loot3Rate: (s.loot3_rate === null || s.loot3_rate === undefined) ? 1 : s.loot3_rate,
      };
    });
  });
  return { quest, godTrial };
}


/** STORY_LINES：text 分頁驅動的開場劇情文字（mapId=100 教學劇情 + 職業/召喚介紹） */
function buildStoryLines(wb) {
  const rows = xlsxSheetToObjects(wb, "text");
  // 修正：text分頁裡mapid=100/after的第3列(row7)其實是純備註列(info本身是null，只有操作備註有文字)，
  // 不是真正的劇情文本，過去這裡沒有排除null info，一旦排到它就會整個buildStoryLines()丟例外、
  // 連帶讓「它後面」的資料表(包含批次3新增的GUIDE_DATABASE)全部讀取失敗、退回內建預設值——這裡改成只取有實際文字的列
  const map100After = rows.filter(o => o.mapid === 100 && o.type === "story" && o.timing === "after" && o.info);
  const map100Before = rows.filter(o => o.mapid === 100 && o.type === "story" && o.timing === "before" && o.info);
  // guideDialogue：mapid=null的type=guide列裡，真正的「對話文本」(教官職業介紹/艾琳召喚武器前置對話)，
  // 排除{battle}/{start_gift}/{name}/{arrangement}這類純標記列(info只是一個{xxx}標籤，不是要顯示給玩家的文字，
  // 2026-09-10已把這幾個標記從text分頁的info欄位清空，這條regex保留下來當作未來萬一又混入標記時的安全網)。
  // 依text分頁裡的實際排列順序，第1筆是教官/職業介紹對話(對應classIntro)，第2筆才是艾琳/召喚武器前置對話(對應gachaIntro)。
  const guideDialogue = rows.filter(o => o.type === "guide" && o.mapid === null && o.info && !/^\{[a-z_]+\}$/.test(String(o.info || "").trim()));
  // 只回傳xlsx裡實際有文字內容的欄位；缺少的欄位則不覆寫，由呼叫端(loadGameDataFromXlsx)以Object.assign合併，
  // 繼續沿用index.html內建的預設文本，避免覆寫成空字串
  const out = {};
  if (map100Before[0]) out.before100 = map100Before[0].info.split("\n");
  if (map100After[0]) out.after100a = map100After[0].info.split("\n");
  if (map100After[1]) out.after100b = map100After[1].info.split("\n");
  // 2026-09-10新增：mapid=100這幾列的background欄位過去完全沒有被讀取，教學劇情的背景圖一直是寫死的
  // bg:'castle'/'central'(對照BG_ASSETS)，即使Wei在text分頁這幾列填了background欄位也不會生效。
  // 這裡比照MAINLINE_STORY_XLSX的bgFile做法一併讀出，index.html會優先採用這裡的值，找不到才退回寫死的bg key。
  if (map100Before[0] && map100Before[0].background) out.before100Bg = String(map100Before[0].background).trim();
  if (map100After[0] && map100After[0].background) out.after100aBg = String(map100After[0].background).trim();
  if (map100After[1] && map100After[1].background) out.after100bBg = String(map100After[1].background).trim();
  // 2026-09-10修正(Wei要求把引導文字統一接回text表)：這裡過去把map100After[2]（永遠不存在，
  // mapid=100/after的story列實際上只有2筆有文字）指派給classIntro，導致classIntro從未真正讀到xlsx內容、
  // 永遠退回index.html內建預設值；同時guideDialogue[0]（教官/職業介紹對話）被誤接到gachaIntro，
  // 顯示成錯誤的對話內容。改成依text分頁實際排列順序正確對應：guideDialogue[0]=classIntro、
  // guideDialogue[1]=gachaIntro，兩者現在都會真正讀取xlsx的info內容，Wei之後直接改這兩列文字即可生效。
  if (guideDialogue[0]) out.classIntro = guideDialogue[0].info.split("\n");
  if (guideDialogue[1]) out.gachaIntro = guideDialogue[1].info.split("\n");
  return out;
}

/**
 * 2026-09-10新增：MAINLINE_STORY_XLSX —— text分頁裡「mapid不是100」的type=story列，
 * 對應各主線關卡自己的劇情(Wei陸續在xlsx新增，例如mapid=1009/1029/1030...)，過去這些列完全沒有
 * 程式碼讀取(buildStoryLines()只處理mapid=100)，導致Wei填了新關卡的劇情卻永遠不會被觸發。
 * Wei目前指示「暫時均改使用after功能」：不論該列原本的timing欄位填的是before還是after，
 * 這裡統一歸類到「戰鬥勝利、結算彈窗關閉後才播放」的after清單裡(同一個mapId若有多列，依xlsx原本
 * 排列順序依序串接播放)；之後如果要恢復區分before/after，只需要把下面這行判斷式拆開即可。
 * 2026-09-10再新增：text分頁新增了background欄位(劇情用背景圖檔名，例如"forest.jpg")，讀取
 * assets/images/background/<檔名> 當作該劇情播放時的背景圖；只是新增讀取功能，其餘播放行為不變。
 * 回傳 { mapId: { after: [文字陣列], bgFile: "檔名或undefined" } }，index.html的getMainlineStoryFor()
 * 會優先採用這裡的資料，找不到時才退回內建的MAINLINE_STORY(僅涵蓋mapId=200/209等舊版關卡編號，供舊資料相容)。
 */
function buildMainlineStoryDatabase(wb) {
  const rows = xlsxSheetToObjects(wb, "text");
  const out = {};
  rows.forEach(o => {
    if (o.type !== "story") return;
    if (o.mapid == null || o.mapid === 100) return; // mapid=100(教學劇情)由buildStoryLines()專門處理，這裡不重複收錄
    if (!o.info) return; // 純備註列(info為null)略過，同buildStoryLines()的既有防呆
    if (!out[o.mapid]) out[o.mapid] = { after: [] };
    out[o.mapid].after = out[o.mapid].after.concat(String(o.info).split("\n"));
    if (o.background != null && String(o.background).trim() !== '') {
      out[o.mapid].bgFile = String(o.background).trim();
    }
  });
  return out;
}

/**
 * 批次3：GUIDE_DATABASE —— 16個「系統開放檢查點」的引導說明文字，由 text 分頁驅動。
 * 沿用既有 type=guide 的欄位慣例，用 timing 欄位存放 checkpoint key(既有的教學步驟(戰鬥/命名/職業選擇等)
 * 的guide列timing欄位一律是null，藉此天然區隔出「新版checkpoint引導」而不影響舊有教學系統)。
 * 只要 timing 有值就視為一個checkpoint key，回傳 { key: { mapId, info } }；
 * Wei若要暫時關閉某個checkpoint的引導文字，直接把該列從xlsx刪掉即可，不需要改任何程式碼。
 */
function buildFeatureGuideDatabase(wb) {
  const rows = xlsxSheetToObjects(wb, "text");
  const out = {};
  rows.forEach(o => {
    if (o.type !== "guide") return;
    if (o.timing === null || o.timing === undefined || String(o.timing).trim() === "") return;
    out[String(o.timing).trim()] = { mapId: o.mapid, info: o.info };
  });
  // 2026-09-10新增：text分頁這裡的mapid欄位只是給Wei對照「這則引導對應哪一關」用的文件資訊，
  // 引導實際何時觸發完全由index.html的FEATURE_UNLOCK_META決定，兩者並不連動讀取彼此。
  // 過去發生過FEATURE_UNLOCK_META的檢查點時程調整後，忘記同步更新text分頁的mapid欄位，
  // 導致Wei在xlsx裡看到的「這關對應的引導文字」跟遊戲實際觸發的關卡對不上(2026-09-10已修正這批舊資料)。
  // 這裡在載入時做一次一致性檢查，兩者不一致就印出警告，方便日後檢查點時程再調整時能及早發現、
  // 純粹是診斷用的console.warn，不影響遊戲實際運作(即使FEATURE_UNLOCK_META_MAP尚未定義也不會報錯)。
  try {
    if (typeof FEATURE_UNLOCK_META_MAP !== 'undefined' && FEATURE_UNLOCK_META_MAP) {
      Object.keys(out).forEach(key => {
        const meta = FEATURE_UNLOCK_META_MAP[key];
        if (meta && out[key].mapId != null && meta.mapId !== out[key].mapId) {
          console.warn(`[text分頁提醒] 引導「${key}」在text分頁標記的mapid(${out[key].mapId})跟遊戲實際觸發的mapid(${meta.mapId})不一致，建議把text分頁該列的mapid欄位改成${meta.mapId}`);
        }
      });
    }
  } catch (e) {}
  return out;
}

/**
 * 讀取 xlsx 並覆寫已支援的資料表常數。
 * 若讀取/解析失敗，會印出警告並保留 index.html 內建的預設值，遊戲仍可正常開啟。
 */
async function loadGameDataFromXlsx() {
  try {
    const resp = await fetch(`${GAME_XLSX_PATH}?v=${ASSET_VERSION}`, { cache: 'no-store' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });

    TAGLIST_DATABASE = buildTaglistDatabase(wb);
    TALENTLIST_DATABASE = buildTalentlistDatabase(wb);
    CHARACTER_STAT_DATABASE = buildCharacterStatDatabase(wb);
    NATURE_DATABASE = buildNatureDatabase(wb);
    MAPLIST_DATABASE = buildMaplistDatabase(wb);
    START_GIFT_DATABASE = buildStartGiftDatabase(wb);
    COST_DATABASE = buildCostDatabase(wb);
    REFINE_DATABASE = buildRefineDatabase(wb);
    ITEM_DATABASE = buildItemDatabase(wb);
    LEVELCURVE_DATABASE = buildLevelcurveDatabase(wb);
    TALENT_CURVE = buildTalentCurve(wb);
    GOD_POINT_CURVE = buildGodPointCurve(wb);
    IDLEZONE_DATABASE = buildIdlezoneDatabase(wb);
    GACHALIST_DATABASE = buildGachalistDatabase(wb);
    DUNGEON_LEVEL_DATABASE = buildDungeonLevelDatabase(wb);
    DUNGEONPREFEB_DATABASE = buildDungeonprefebDatabase(wb);
    const talentResult = buildTalentDatabaseAndGrid(wb);
    TALENT_DATABASE = talentResult.talent;
    TALENT_GROUP_GRID = talentResult.grid;
    const buffResult = buildBuffDatabaseAndSetTiers(wb);
    BUFF_DATABASE = buffResult.buff;
    BUFF_SET_TIERS = buffResult.setTiers;
    SKILL_DATABASE = buildSkillDatabase(wb);
    CARD_DATABASE = buildCardDatabase(wb);
    EQUIPMENT_DATABASE = buildEquipmentDatabase(wb);
    GACHA_DATABASE = buildGachaDatabase(wb);
    BUILD_DATABASE = buildBuildDatabase(wb);
    const questResult = buildQuestDatabaseAndGodTrial(wb);
    QUEST_DATABASE = questResult.quest;
    GOD_TRIAL_VARIANTS = questResult.godTrial;
    STORY_LINES = Object.assign({}, STORY_LINES, buildStoryLines(wb));
    MAINLINE_STORY_XLSX = buildMainlineStoryDatabase(wb);
    GUIDE_DATABASE = buildFeatureGuideDatabase(wb);

    console.log('[xlsx-loader] 已從', GAME_XLSX_PATH, '載入資料表');
  } catch (err) {
    console.warn('[xlsx-loader] 讀取 xlsx 失敗，改用 index.html 內建的預設資料。' +
      '（若是雙擊 index.html 開啟，請改用 start.bat 啟動本機伺服器）', err);
  }
}
