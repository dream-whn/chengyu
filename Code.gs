/**
 * 成語練習室：Google 試算表後端
 * 1. 改下面的老師密碼
 * 2. 選擇函式「setup」按「執行」一次（第一次會要求授權）
 * 3. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人）
 * 已經用緊舊版？選擇函式「assignLevels」執行一次，就會加上「級別」欄並把 100 個成語分成 5 關。
 */
const TEACHER_PASSWORD = 'cisabc123';

const SHEET_IDIOMS = '成語';
const SHEET_LOG = '作答紀錄';
const IDIOM_HEAD = ['id', '成語', '解釋', '例句', '分類', '讀音', '建立時間', '級別'];
const LOG_HEAD = ['時間', '學生編號', '班別學號', '姓名', '成語ID', '成語', '題型', '結果', '所選答案', '關卡', '場次'];

/* ---------- 初次設定 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET_IDIOMS) || ss.insertSheet(SHEET_IDIOMS);
  if (s.getLastRow() === 0) {
    s.appendRow(IDIOM_HEAD);
    const now = Date.now();
    const rows = SEED.map((r, i) => ['c' + String(i + 1).padStart(3, '0'), r[0], r[1], '', r[2], '', now + i, LEVELS[r[0]] || '']);
    s.getRange(2, 1, rows.length, IDIOM_HEAD.length).setValues(rows);
  }
  ensureHeaders_();
  s.setFrozenRows(1);
  let l = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (l.getLastRow() === 0) l.appendRow(LOG_HEAD);
  l.setFrozenRows(1);
  const first = ss.getSheets()[0];
  if (first.getName() !== SHEET_IDIOMS && first.getName() !== SHEET_LOG && first.getLastRow() === 0) ss.deleteSheet(first);
}

/* 舊版試算表補上新欄位標題 */
function ensureHeaders_() {
  const ss = SpreadsheetApp.getActive();
  [[SHEET_IDIOMS, IDIOM_HEAD], [SHEET_LOG, LOG_HEAD]].forEach(([name, head]) => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 0) sh.getRange(1, 1, 1, head.length).setValues([head]);
  });
}

/* 把初始 100 個成語分成 5 關（只填空白的級別，老師改過的不會被覆蓋） */
function assignLevels() {
  setup();
  const s = sheet_(SHEET_IDIOMS);
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  const words = s.getRange(2, 2, n, 1).getValues();
  const lv = s.getRange(2, 8, n, 1).getValues();
  let changed = 0;
  words.forEach((w, i) => {
    const k = LEVELS[String(w[0]).trim()];
    if (k && !lv[i][0]) { lv[i][0] = k; changed++; }
  });
  s.getRange(2, 8, n, 1).setValues(lv);
  Logger.log('已設定 ' + changed + ' 個成語的級別');
}

/* ---------- 讀取 ---------- */
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // 防止被當成公式
  return s;
}

function readIdioms_() {
  const s = sheet_(SHEET_IDIOMS);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const vals = s.getRange(2, 1, n, IDIOM_HEAD.length).getValues();
  const out = [];
  let fixed = false;
  vals.forEach((r, i) => {
    if (!String(r[1]).trim()) return;
    if (!String(r[0]).trim()) { // 老師直接在試算表加的成語：補上 id
      r[0] = 'r' + Date.now().toString(36) + i;
      s.getRange(i + 2, 1).setValue(r[0]);
      if (!r[6]) { r[6] = Date.now() + i; s.getRange(i + 2, 7).setValue(r[6]); }
      fixed = true;
    }
    out.push({ id: String(r[0]), word: String(r[1]).trim(), meaning: String(r[2]), example: String(r[3]),
      category: String(r[4]), reading: String(r[5]), createdAt: Number(r[6]) || (i + 1), level: Number(r[7]) || 0 });
  });
  return out;
}

function readLog_() {
  const s = sheet_(SHEET_LOG);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, LOG_HEAD.length).getValues().map(r => ({
    t: r[0] instanceof Date ? r[0].getTime() : Number(r[0]) || 0,
    sid: String(r[1]), cls: String(r[2]), name: String(r[3]), i: String(r[4]), w: String(r[5]),
    m: String(r[6]), ok: r[7] === '對' ? 1 : 0, a: String(r[8]), lv: Number(r[9]) || 0, ss: String(r[10] || '')
  }));
}

/* ---------- 首三名：已掌握成語數 → 正確率 → 作答題數 ---------- */
function top_(n) {
  const ids = new Set(readIdioms_().map(x => x.id));
  const m = {};
  readLog_().sort((a, b) => a.t - b.t).forEach(r => {
    const s = m[r.sid] = m[r.sid] || { name: r.name, cls: r.cls, total: 0, correct: 0, by: {} };
    s.name = r.name || s.name; s.cls = r.cls || s.cls;
    s.total++; if (r.ok) s.correct++;
    const b = s.by[r.i] = s.by[r.i] || { c: 0, w: 0, last: 0 };
    if (r.ok) b.c++; else b.w++;
    b.last = r.ok;
  });
  return Object.values(m).map(s => ({
    name: s.name, cls: s.cls, total: s.total,
    acc: s.total ? Math.round(s.correct / s.total * 100) : 0,
    mastered: Object.keys(s.by).filter(id => ids.has(id) && s.by[id].c >= s.by[id].w && s.by[id].last).length
  })).filter(s => s.mastered > 0)
    .sort((a, b) => b.mastered - a.mastered || b.acc - a.acc || b.total - a.total)
    .slice(0, n);
}

/* ---------- 網頁入口 ---------- */
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'idioms') return out_({ ok: true, idioms: readIdioms_() });
    if (p.action === 'top') return out_({ ok: true, top: top_(3) });
    if (p.action === 'myRecord') return out_({ ok: true, rows: readLog_().filter(r => r.sid === p.sid) });
    return out_({ ok: true, hello: '成語練習室後端運作中' });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '格式錯誤' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'answers') return out_(saveAnswers_(b.rows || []));
    if (b.pw !== TEACHER_PASSWORD) return out_({ ok: false, error: '密碼不正確' });
    switch (b.action) {
      case 'login': return out_({ ok: true });
      case 'teacherData': return out_({ ok: true, idioms: readIdioms_(), log: readLog_() });
      case 'saveIdiom': return out_(saveIdiom_(b.idiom));
      case 'deleteIdiom': return out_(deleteIdiom_(b.id));
      case 'bulkIdioms': return out_(bulkIdioms_(b.items || []));
      case 'clearStudent': return out_(clearStudent_(b.sid));
    }
    return out_({ ok: false, error: '未知指令' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

/* ---------- 寫入 ---------- */
function saveAnswers_(rows) {
  rows = rows.slice(0, 200);
  if (!rows.length) return { ok: true, saved: 0 };
  const vals = rows.map(r => [new Date(Number(r.t) || Date.now()), clean_(r.sid, 80), clean_(r.cls, 20), clean_(r.name, 30),
    clean_(r.i, 40), clean_(r.w, 20), clean_(r.m, 20), r.ok ? '對' : '錯', clean_(r.a, 20), Number(r.lv) || '', clean_(r.ss, 40)]);
  const s = sheet_(SHEET_LOG);
  s.getRange(s.getLastRow() + 1, 1, vals.length, LOG_HEAD.length).setValues(vals);
  return { ok: true, saved: vals.length };
}

function findRow_(s, id) {
  const n = s.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = s.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function saveIdiom_(d) {
  if (!d || !d.word || !d.meaning) return { ok: false, error: '要有成語和解釋' };
  const s = sheet_(SHEET_IDIOMS);
  const id = d.id || ('i' + Date.now().toString(36) + Math.floor(Math.random() * 1e4));
  const row = [id, clean_(d.word, 20), clean_(d.meaning), clean_(d.example), clean_(d.category, 40), clean_(d.reading, 80), Number(d.createdAt) || Date.now(), Number(d.level) || ''];
  const at = d.id ? findRow_(s, d.id) : -1;
  if (at > 0) s.getRange(at, 1, 1, row.length).setValues([row]);
  else {
    if (readIdioms_().some(x => x.word === String(d.word).trim())) return { ok: false, error: '成語庫已有「' + d.word + '」' };
    s.appendRow(row);
  }
  return { ok: true, id: id };
}

function deleteIdiom_(id) {
  const s = sheet_(SHEET_IDIOMS);
  const at = findRow_(s, id);
  if (at > 0) s.deleteRow(at);
  return { ok: true };
}

function bulkIdioms_(items) {
  const have = new Set(readIdioms_().map(x => x.word));
  const now = Date.now();
  const rows = [];
  items.forEach((p, k) => {
    const w = String(p.word || '').trim();
    if (!w || !p.meaning || have.has(w)) return;
    have.add(w);
    rows.push(['i' + now.toString(36) + k, clean_(w, 20), clean_(p.meaning), clean_(p.example), clean_(p.category, 40), clean_(p.reading, 80), now + k, Number(p.level) || '']);
  });
  if (rows.length) {
    const s = sheet_(SHEET_IDIOMS);
    s.getRange(s.getLastRow() + 1, 1, rows.length, IDIOM_HEAD.length).setValues(rows);
  }
  return { ok: true, added: rows.length, skipped: items.length - rows.length };
}

function clearStudent_(sid) {
  const s = sheet_(SHEET_LOG);
  const n = s.getLastRow() - 1;
  if (n < 1) return { ok: true };
  const vals = s.getRange(2, 1, n, LOG_HEAD.length).getValues();
  const keep = vals.filter(r => String(r[1]) !== String(sid));
  s.getRange(2, 1, n, LOG_HEAD.length).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, LOG_HEAD.length).setValues(keep);
  return { ok: true, removed: n - keep.length };
}

/* ---------- 初始 100 個成語 ---------- */
const SEED = [
  ["一心一意", "心思專一。", "數字與幾何"],
  ["三心二意", "意志不堅定，猶豫不決。", "數字與幾何"],
  ["五花八門", "比喻花樣繁多或變幻多端。", "數字與幾何"],
  ["七上八下", "形容心裡慌亂不安。", "數字與幾何"],
  ["九牛一毛", "比喻極大數量中的極小部分。", "數字與幾何"],
  ["十全十美", "比喻圓滿美好，毫無缺陷。", "數字與幾何"],
  ["千方百計", "想盡或用盡一切辦法。", "數字與幾何"],
  ["萬無一失", "形容絕對不會出差錯。", "數字與幾何"],
  ["四面八方", "周圍各個方面或各個地方。", "數字與幾何"],
  ["獨一無二", "唯一的，沒有第二個。", "數字與幾何"],
  ["守株待兔", "比喻拘泥守成，不知變通；或妄想不勞而獲。", "動物寓言"],
  ["畫蛇添足", "比喻多此一舉，反將事情弄糟。", "動物寓言"],
  ["亡羊補牢", "比喻犯錯後及時更正，尚能補救。", "動物寓言"],
  ["井底之蛙", "比喻見識淺薄的人。", "動物寓言"],
  ["狐假虎威", "比喻憑恃他人的勢力來欺壓人。", "動物寓言"],
  ["鶴立雞群", "比喻人的儀表或才能超群出眾。", "動物寓言"],
  ["對牛彈琴", "比喻對不懂道理的人講道理。", "動物寓言"],
  ["盲人摸象", "比喻以偏概全，不能了解真相。", "動物寓言"],
  ["如魚得水", "比喻遇到意氣相投的人或環境。", "動物寓言"],
  ["膽小如鼠", "形容非常膽小。", "動物寓言"],
  ["春暖花開", "春氣和煦，花兒盛開。", "自然與景色"],
  ["風和日麗", "微風和煦，陽光映照。", "自然與景色"],
  ["萬里無雲", "天空晴朗，視野遼闊。", "自然與景色"],
  ["傾盆大雨", "形容雨勢極大。", "自然與景色"],
  ["山清水秀", "形容風景優美。", "自然與景色"],
  ["鳥語花香", "形容景色的美好。", "自然與景色"],
  ["秋高氣爽", "秋空高朗，氣候涼爽。", "自然與景色"],
  ["冰天雪地", "形容氣候嚴寒。", "自然與景色"],
  ["依山傍水", "形容地理位置極佳（多指建築）。", "自然與景色"],
  ["繁星點點", "形容夜晚天空星星很多。", "自然與景色"],
  ["廢寢忘食", "形容專心努力工作或學習。", "學習與態度"],
  ["全神貫注", "將心思精神完全集中在某事物上。", "學習與態度"],
  ["持之以恆", "有恆心地堅持下去。", "學習與態度"],
  ["循序漸進", "按部就班，由淺入深。", "學習與態度"],
  ["溫故知新", "複習舊學，能有所新領悟。", "學習與態度"],
  ["勤能補拙", "勤奮努力可以彌補天資的不足。", "學習與態度"],
  ["不恥下問", "不以向身分較低或學問較差的人求教為恥。", "學習與態度"],
  ["學富五車", "形容人書讀很多，學識淵博。", "學習與態度"],
  ["名列前茅", "形容成績極好。", "學習與態度"],
  ["半途而廢", "做事有始無終。", "學習與態度"],
  ["興高采烈", "形容興致勃勃，情緒熱烈的樣子。", "情緒與表情"],
  ["垂頭喪氣", "形容失意沮喪的樣子。", "情緒與表情"],
  ["怒髮衝冠", "形容盛怒的樣子。", "情緒與表情"],
  ["眉開眼笑", "形容非常高興。", "情緒與表情"],
  ["心花怒放", "形容心情極其快活。", "情緒與表情"],
  ["提心吊膽", "形容心理上驚恐擔憂。", "情緒與表情"],
  ["悶悶不樂", "心情憂鬱不快樂。", "情緒與表情"],
  ["喜出望外", "因意想不到的喜事而特別高興。", "情緒與表情"],
  ["大吃一驚", "形容非常驚訝。", "情緒與表情"],
  ["手舞足蹈", "形容高興到了極點。", "情緒與表情"],
  ["和藹可親", "態度溫和，容易親近。", "外貌與神態"],
  ["笑容滿面", "臉上充滿笑容。", "外貌與神態"],
  ["骨瘦如柴", "形容非常消瘦。", "外貌與神態"],
  ["濃眉大眼", "形容五官明顯，長相俊美。", "外貌與神態"],
  ["精神奕奕", "形容精神飽滿，容光煥發。", "外貌與神態"],
  ["彬彬有禮", "形容文雅有禮貌的樣子。", "外貌與神態"],
  ["目不轉睛", "凝神注視，眼珠一動也不動。", "外貌與神態"],
  ["呆若木雞", "形容愚笨或受驚嚇而發愣的樣子。", "外貌與神態"],
  ["恍然大悟", "對某事突然明白。", "外貌與神態"],
  ["自言自語", "自己對自己說話。", "外貌與神態"],
  ["人山人海", "形容人聚集得非常多。", "場景與狀況"],
  ["井井有條", "形容整齊有條理。", "場景與狀況"],
  ["亂七八糟", "形容毫無條理，亂得一塌糊塗。", "場景與狀況"],
  ["迫不及待", "急迫得不能等待。", "場景與狀況"],
  ["車水馬龍", "形容繁華熱鬧的景象。", "場景與狀況"],
  ["鴉雀無聲", "形容非常寂靜。", "場景與狀況"],
  ["震耳欲聾", "形容聲音很大。", "場景與狀況"],
  ["川流不息", "形容像水流一樣連續不斷（多指人車）。", "場景與狀況"],
  ["煥然一新", "形容呈現出嶄新的面貌。", "場景與狀況"],
  ["美不勝收", "美好的東西太多，無法盡收眼底。", "場景與狀況"],
  ["助人為樂", "以幫助別人為快樂。", "品德與修養"],
  ["大公無私", "秉公處理，不偏私。", "品德與修養"],
  ["拾金不昧", "撿到財物不據為己有。", "品德與修養"],
  ["見義勇為", "看到正義的事，勇敢地去做。", "品德與修養"],
  ["禮尚往來", "指用對方對待自己的態度和方式去對待對方。", "品德與修養"],
  ["誠實守信", "真誠老實，遵守承諾。", "品德與修養"],
  ["知錯能改", "知道犯錯後能夠改正。", "品德與修養"],
  ["飲水思源", "比喻不忘本。", "品德與修養"],
  ["尊師重道", "尊敬師長，重視道統。", "品德與修養"],
  ["相親相愛", "形容關係密切，感情深厚。", "品德與修養"],
  ["自相矛盾", "比喻行事或言語先後不相應、互相牴觸。", "邏輯與處事"],
  ["掩耳盜鈴", "比喻自欺欺人。", "邏輯與處事"],
  ["刻舟求劍", "比喻拘泥固執，不知變通。", "邏輯與處事"],
  ["雪中送炭", "比喻在人艱困時給予援助。", "邏輯與處事"],
  ["錦上添花", "比喻美上加美，好上加好。", "邏輯與處事"],
  ["火上加油", "比喻使情況更惡化。", "邏輯與處事"],
  ["異想天開", "比喻想法荒唐怪誕，不切實際。", "邏輯與處事"],
  ["得意忘形", "形容高興得失去常態。", "邏輯與處事"],
  ["班門弄斧", "比喻在行家面前賣弄本事。", "邏輯與處事"],
  ["杞人憂天", "比喻無謂的憂慮。", "邏輯與處事"],
  ["光陰似箭", "形容時間過得很快。", "時間與行動"],
  ["日月如梭", "形容時光流逝迅速。", "時間與行動"],
  ["曇花一現", "比喻美好事物不久留。", "時間與行動"],
  ["爭分奪秒", "形容充分利用時間，不浪費一分一秒。", "時間與行動"],
  ["不約而同", "彼此未經商量而行為、看法一致。", "時間與行動"],
  ["自告奮勇", "主動要求擔任某項工作。", "時間與行動"],
  ["隨手可得", "容易取得。", "時間與行動"],
  ["滿載而歸", "形容收穫豐富。", "時間與行動"],
  ["無可奈何", "毫無辦法。", "時間與行動"],
  ["莫名其妙", "形容事情奇怪，不合常理。", "時間與行動"]
];

/* ---------- 5 關分級（1 最易，5 最難） ---------- */
const LEVELS = {"一心一意": 1, "三心二意": 1, "四面八方": 1, "十全十美": 1, "七上八下": 1, "人山人海": 1, "春暖花開": 1, "鳥語花香": 1, "笑容滿面": 1, "大吃一驚": 1, "自言自語": 1, "亂七八糟": 1, "傾盆大雨": 1, "萬里無雲": 1, "冰天雪地": 1, "繁星點點": 1, "手舞足蹈": 1, "眉開眼笑": 1, "相親相愛": 1, "助人為樂": 1, "五花八門": 2, "獨一無二": 2, "千方百計": 2, "山清水秀": 2, "風和日麗": 2, "秋高氣爽": 2, "興高采烈": 2, "垂頭喪氣": 2, "悶悶不樂": 2, "膽小如鼠": 2, "骨瘦如柴": 2, "濃眉大眼": 2, "和藹可親": 2, "彬彬有禮": 2, "拾金不昧": 2, "知錯能改": 2, "誠實守信": 2, "光陰似箭": 2, "滿載而歸": 2, "迫不及待": 2, "萬無一失": 3, "九牛一毛": 3, "依山傍水": 3, "心花怒放": 3, "提心吊膽": 3, "喜出望外": 3, "目不轉睛": 3, "精神奕奕": 3, "井井有條": 3, "車水馬龍": 3, "鴉雀無聲": 3, "震耳欲聾": 3, "川流不息": 3, "煥然一新": 3, "全神貫注": 3, "半途而廢": 3, "名列前茅": 3, "不約而同": 3, "自告奮勇": 3, "隨手可得": 3, "守株待兔": 4, "畫蛇添足": 4, "亡羊補牢": 4, "井底之蛙": 4, "狐假虎威": 4, "對牛彈琴": 4, "盲人摸象": 4, "如魚得水": 4, "鶴立雞群": 4, "自相矛盾": 4, "掩耳盜鈴": 4, "刻舟求劍": 4, "杞人憂天": 4, "雪中送炭": 4, "錦上添花": 4, "火上加油": 4, "恍然大悟": 4, "呆若木雞": 4, "日月如梭": 4, "爭分奪秒": 4, "怒髮衝冠": 5, "美不勝收": 5, "廢寢忘食": 5, "持之以恆": 5, "循序漸進": 5, "溫故知新": 5, "勤能補拙": 5, "不恥下問": 5, "學富五車": 5, "大公無私": 5, "見義勇為": 5, "禮尚往來": 5, "飲水思源": 5, "尊師重道": 5, "異想天開": 5, "得意忘形": 5, "班門弄斧": 5, "曇花一現": 5, "無可奈何": 5, "莫名其妙": 5};
