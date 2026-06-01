/**
 * Google Apps Script — Scrape live ICU bed data from MOHW for all 194 hospitals.
 *
 * SETUP (one time — no file import needed):
 *  1. Open https://sheets.google.com and create a NEW blank spreadsheet.
 *  2. Extensions → Apps Script → delete any existing code → paste this entire file → Save (Ctrl+S).
 *  3. Reload the spreadsheet.  A new "🏥 ICU Scraper" menu will appear.
 *
 * RUNNING (in order):
 *  1. 🏥 ICU Scraper → 0. Load hospital list
 *     - Fills the sheet with all 194 hospitals from the embedded data.
 *     - No internet access needed for this step.
 *  2. 🏥 ICU Scraper → 1. Scrape ICU data
 *     - Fetches each hospital's MOHW page (350ms between requests).
 *     - Auto-pauses after 5 min; re-run to resume where it left off.
 *     - Results saved to columns H (scraped_icu), I (scraped_burn), J (status).
 *  3. 🏥 ICU Scraper → 2. Email results as CSV
 *     - Sends icu_scraped.csv to your Google account email.
 */

"use strict";

// ── Config ─────────────────────────────────────────────────────────────────

const SHEET_NAME = "ICU";
const BASE_URL   = "https://openinfo.mohw.gov.tw/web/C02?type=1&basAgencyId={CODE}&zoneCode={ZONE}";

const KNOWN_ZONES = {
  "0101090517": {
    "仁愛院區": "A", "中興院區": "B", "忠孝院區": "C",
    "陽明院區": "D", "和平婦幼院區": "E",
    "松德院區": "F", "林森中醫昆明院區": "G",
  },
};

// Column indices (1-based)
const C_CODE   = 1;  // A
const C_CAMPUS = 2;  // B
const C_NAME   = 3;  // C
const C_ICU    = 4;  // D — July-2024 baseline
const C_BURN   = 5;  // E — July-2024 baseline
// Output columns:
const C_NEW_ICU  = 8;  // H
const C_NEW_BURN = 9;  // I
const C_STATUS   = 10; // J

// ── Embedded hospital list (194 hospitals, July-2024 baseline) ─────────────

const HOSPITAL_DATA = [["401180014","本院","國立台灣大學醫學院附設醫院",177,4],["501110514","本院","三軍總醫院附設民眾診療服務處",137,6],["601160016","本院","臺北榮民總醫院",207,6],["1101010012","本院","長庚醫療財團法人台北長庚紀念醫院",14,0],["1101020018","本院","國泰醫療財團法人國泰綜合醫院",56,4],["1101100011","本院","台灣基督長老教會馬偕醫療財團法人馬偕紀念醫院",65,8],["1101150011","本院","新光醫療財團法人新光吳火獅紀念醫院",62,3],["1301200010","本院","臺北市立萬芳醫院-委託財團法人臺北醫學大學辦理",50,0],["617060018","本院","臺中榮民總醫院",138,8],["1317040011","本院","中山醫學大學附設醫院",107,5],["1317050017","本院","中國醫藥大學附設醫院",151,10],["421040011","本院","國立成功大學醫學院附設醫院",126,10],["1141310019","本院","奇美醫療財團法人奇美醫院",109,8],["602030026","本院","高雄榮民總醫院",102,4],["1142100017","本院","長庚醫療財團法人高雄長庚紀念醫院",207,0],["1142120001","本院","義大醫療財團法人義大醫院",89,4],["1302050014","本院","財團法人私立高雄醫學大學附設中和紀念醫院",142,5],["1131010011","本院","醫療財團法人徐元智先生醫藥基金會亞東紀念醫院",102,0],["1131050515","本院","佛教慈濟醫療財團法人台北慈濟醫院",81,0],["1131100010","本院","台灣基督長老教會馬偕醫療財團法人淡水馬偕紀念醫院",61,0],["1331040513","本院","衛生福利部雙和醫院 (委託臺北醫學大學興建經營)",85,0],["1132070011","本院","長庚醫療財團法人林口長庚紀念醫院",299,17],["1137010024","本院","彰化基督教醫療財團法人彰化基督教醫院",149,0],["1145010010","本院","佛教慈濟醫療財團法人花蓮慈濟醫院",75,4],["401180023","本院","國立臺灣大學醫學院附設醫院兒童醫院",65,0],["1101100020","本院","台灣基督長老教會馬偕醫療財團法人馬偕兒童醫院",35,0],["1112010537","本院","新竹市立馬偕兒童醫院(委託台灣基督長老教會馬偕醫療財團法人興建經營)",14,0],["1137010042","本院","彰化基督教醫療財團法人彰化基督教兒童醫院",36,0],["1303260014","本院","中國醫藥大學兒童醫院",38,0],["101090517","仁愛院區","臺北市立聯合醫院-仁愛",30,0],["101090517","中興院區","臺北市立聯合醫院-中興",19,0],["101090517","忠孝院區","臺北市立聯合醫院-忠孝",24,0],["101090517","陽明院區","臺北市立聯合醫院-陽明",25,0],["101090517","和平婦幼院區","臺北市立聯合醫院-和平",30,0],["501010019","本院","三軍總醫院松山分院附設民眾診療服務處",16,0],["1101010021","本院","基督復臨安息日會醫療財團法人臺安醫院",21,0],["1101160017","本院","振興醫療財團法人振興醫院",78,6],["1101160026","本院","醫療財團法人辜公亮基金會和信治癌中心醫院",16,0],["1301170017","本院","臺北醫學大學附設醫院",59,0],["117030010","本院","衛生福利部臺中醫院",30,0],["136010010","本院","衛生福利部豐原醫院",53,0],["536190011","本院","國軍臺中總醫院附設民眾診療服務處",26,0],["917070029","本院","林新醫療社團法人林新醫院",35,0],["936030018","本院","李綜合醫療社團法人大甲李綜合醫院",21,0],["936050029","沙鹿總院","光田醫療社團法人光田綜合醫院-沙鹿",48,0],["936050029","大甲院區","光田醫療社團法人光田綜合醫院-大甲",37,0],["936060016","梧棲院區","童綜合醫療社團法人童綜合醫院-梧棲",90,10],["936060016","沙鹿院區","童綜合醫療社團法人童綜合醫院-沙鹿",0,0],["1136090519","本院","佛教慈濟醫療財團法人台中慈濟醫院",55,0],["1136200015","本院","仁愛醫療財團法人大里仁愛醫院",40,0],["1517011112","本院","澄清綜合醫院",30,0],["1517061032","本院","澄清綜合醫院中港分院",55,0],["121050011","本院","衛生福利部臺南醫院",32,0],["905320023","本院","台南市立醫院(委託秀傳醫療社團法人經營)",40,0],["1105040016","本院","台灣基督長老教會新樓醫療財團法人麻豆新樓醫院",22,0],["1121010018","本院","台灣基督長老教會新樓醫療財團法人台南新樓醫院",32,0],["1141090512","本院","奇美醫療財團法人柳營奇美醫院",67,0],["1305370013","本院","臺南市立安南醫院-委託中國醫藥大學興建經營",62,0],["102020011","本院","高雄市立聯合醫院",29,0],["102070020","本院","高雄市立大同醫院(委託財團法人私立高雄醫學大學附設中和紀念醫院經營)",45,0],["502030015","本院","國軍左營總醫院附設民眾診療服務處",32,8],["502080015","本院","國軍高雄總醫院附設民眾診療服務處",52,0],["902080013","本院","阮綜合醫療社團法人阮綜合醫院",53,0],["1102110011","本院","高雄市立小港醫院(委託財團法人私立高雄醫學大學經營)",20,0],["111070010","本院","衛生福利部基隆醫院",30,0],["1111060015","本院","長庚醫療財團法人基隆長庚紀念醫院",63,0],["412040012","本院","國立臺灣大學醫學院附設醫院新竹臺大分院新竹醫院",49,4],["1112010519","本院","台灣基督長老教會馬偕醫療財團法人新竹馬偕紀念醫院",32,0],["622020017","本院","臺中榮民總醫院嘉義分院",37,0],["1122010012","本院","戴德森醫療財團法人嘉義基督教醫院",85,0],["1122010021","本院","天主教中華聖母修女會醫療財團法人天主教聖馬爾定醫院",47,0],["131020016","本院","新北市立聯合醫院",27,0],["131060029","本院","衛生福利部臺北醫院",51,0],["1131090019","本院","行天宮醫療志業醫療財團法人恩主公醫院",30,0],["1131110516","本院","國泰醫療財團法人汐止國泰綜合醫院",46,0],["1231050017","本院","天主教耕莘醫療財團法人耕莘醫院",58,0],["132010014","本院","衛生福利部桃園醫院",68,2],["532090029","本院","國軍桃園總醫院附設民眾診療服務處",44,0],["632010014","本院","臺北榮民總醫院桃園分院",20,0],["932020025","本院","天成醫療社團法人天晟醫院",19,0],["1132010024","本院","沙爾德聖保祿修女會醫療財團法人聖保祿醫院",23,0],["1532011154","本院","敏盛綜合醫院",46,0],["1532100049","本院","聯新國際醫院",34,0],["933050018","本院","東元醫療社團法人東元綜合醫院",28,0],["434010518","本院","國立陽明交通大學附設醫院",41,0],["1134020019","本院","醫療財團法人羅許基金會羅東博愛醫院",46,0],["1134020028","本院","天主教靈醫會醫療財團法人羅東聖母醫院",45,0],["135010016","本院","衛生福利部苗栗醫院",26,0],["1135050020","本院","為恭醫療財團法人為恭紀念醫院",30,0],["137170515","本院","衛生福利部彰化醫院",16,0],["937010019","本院","秀傳醫療社團法人秀傳紀念醫院",70,0],["1137020511","本院","秀傳醫療財團法人彰濱秀傳紀念醫院",42,0],["138010027","本院","衛生福利部南投醫院",37,0],["1138020015","本院","埔基醫療財團法人埔里基督教醫院",14,0],["439010518","本院","國立臺灣大學醫學院附設醫院雲林分院",41,2],["1339060017","本院","中國醫藥大學北港附設醫院",40,0],["1140010510","本院","長庚醫療財團法人嘉義長庚紀念醫院",100,4],["1140030012","本院","佛教慈濟醫療財團法人大林慈濟醫院",59,3],["143010011","本院","衛生福利部屏東醫院",24,0],["943010017","本院","寶建醫療社團法人寶建醫院",30,0],["943030019","本院","安泰醫療社團法人安泰醫院",52,2],["1143010012","本院","屏基醫療財團法人屏東基督教醫院",57,0],["1343030018","本院","輔英科技大學附設醫院",18,0],["1543010109","本院","國仁醫院",15,0],["545040515","本院","國軍花蓮總醫院附設民眾診療服務處",10,0],["1145010038","本院","臺灣基督教門諾會醫療財團法人門諾醫院",39,0],["1146010014","本院","台灣基督長老教會馬偕醫療財團法人台東馬偕紀念醫院",49,0],["511040010","本院","三軍總醫院基隆分院附設民眾診療服務處",8,0],["512040014","本院","國軍桃園總醫院新竹分院附設民眾診療服務處",8,0],["517050010","本院","國軍臺中總醫院中清分院附設民眾診療服務處",6,0],["542020011","本院","國軍高雄總醫院岡山分院附設民眾診療服務處",7,0],["543010019","本院","國軍高雄總醫院屏東分院附設民眾診療服務處",9,0],["544010031","本院","三軍總醫院澎湖分院附設民眾診療服務處",15,0],["633030010","本院","臺北榮民總醫院新竹分院",10,0],["634030014","本院","臺北榮民總醫院蘇澳分院",9,0],["634070018","本院","臺北榮民總醫院員山分院",10,0],["638020014","本院","臺中榮民總醫院埔里分院",12,0],["640140012","灣橋院區","臺中榮民總醫院灣橋分院",8,0],["641310018","本院","高雄榮民總醫院臺南分院",20,0],["643130018","本院","屏東榮民總醫院龍泉分院",10,0],["645020015","本院","臺北榮民總醫院鳳林分院",4,0],["645030011","本院","臺北榮民總醫院玉里分院",8,0],["901190010","本院","西園醫療社團法人西園醫院",8,0],["903150014","本院","林新醫療社團法人烏日林新醫院",20,0],["935020027","本院","李綜合醫療社團法人苑裡李綜合醫院",15,0],["102080017","本院","高雄市立民生醫院",48,0],["122020517","本院","衛生福利部嘉義醫院",20,0],["131060010","本院","衛生福利部樂生療養院",14,0],["132110519","本院","衛生福利部桃園醫院新屋分院",6,0],["140010028","本院","衛生福利部朴子醫院",6,0],["141010013","本院","衛生福利部新營醫院",14,0],["141060513","本院","衛生福利部臺南醫院新化分院",6,0],["142030019","本院","衛生福利部旗山醫院",23,0],["143040019","本院","衛生福利部恆春旅遊醫院",6,0],["144010015","本院","衛生福利部澎湖醫院",8,0],["145010019","本院","衛生福利部花蓮醫院",8,0],["146010013","本院","衛生福利部臺東醫院",6,0],["190030516","本院","衛生福利部金門醫院",10,0],["291010010","本院","連江縣立醫院",1,0],["433050018","本院","國立臺灣大學醫學院附設醫院新竹臺大分院生醫醫院",25,0],["433050018","竹東院區","國立臺灣大學醫學院附設醫院新竹臺大分院生醫醫院",8,0],["439010527","本院","國立成功大學醫學院附設醫院斗六分院",15,0],["937050014","本院","員榮醫療社團法人員榮醫院",15,0],["938030016","本院","佑民醫療社團法人佑民醫院",18,0],["942020019","本院","高雄市立岡山醫院（委託秀傳醫療社團法人經營）",8,0],["943020013","本院","安泰醫療社團法人潮州安泰醫院",8,0],["943040015","本院","南門醫療社團法人南門醫院",8,0],["943160012","本院","枋寮醫療社團法人枋寮醫院",12,0],["1101110026","本院","康寧醫療財團法人康寧醫院",6,0],["1105050012","本院","奇美醫療財團法人佳里奇美醫院",21,0],["1107120017","本院","義大醫療財團法人義大癌治療醫院",16,0],["1107350015","本院","天主教聖功醫療財團法人聖功醫院",6,0],["1112010528","本院","國泰醫療財團法人新竹國泰綜合醫院",16,0],["1133060019","本院","天主教仁慈醫療財團法人仁慈醫院",11,0],["1137010051","本院","彰化基督教醫療財團法人漢銘基督教醫院",10,0],["1137020520","本院","彰化基督教醫療財團法人鹿港基督教醫院",14,0],["1137050019","本院","彰化基督教醫療財團法人員林基督教醫院",24,0],["1137080017","本院","彰化基督教醫療財團法人二林基督教醫院",18,0],["1138010019","本院","彰化基督教醫療財團法人南投基督教醫院",10,0],["1139030015","本院","天主教若瑟醫療財團法人若瑟醫院",20,0],["1139040011","本院","彰化基督教醫療財團法人雲林基督教醫院",24,0],["1139130010","本院","長庚醫療財團法人雲林長庚紀念醫院",6,0],["1142010518","本院","高雄市立鳳山醫院（委託長庚醫療財團法人經營）",4,0],["1145030012","本院","佛教慈濟醫療財團法人玉里慈濟醫院",3,0],["1146010032","本院","東基醫療財團法人台東基督教醫院",8,0],["1231030015","本院","天主教耕莘醫療財團法人永和耕莘醫院",21,0],["1303180011","本院","亞洲大學附屬醫院",40,0],["1331160010","本院","輔仁大學學校財團法人輔仁大學附設醫院",57,0],["1333050017","本院","中國醫藥大學新竹附設醫院",34,0],["1411030013","本院","醫療財團法人臺灣區煤礦業基金會臺灣礦工醫院",6,0],["1436020013","本院","東勢區農會附設農民醫院",6,0],["1501010010","本院","博仁綜合醫院",8,0],["1502040021","本院","健仁醫院",13,0],["1502081175","本院","邱外科醫院",10,0],["1503190020","本院","長安醫院",24,0],["1512011185","本院","南門綜合醫院",10,0],["1521031104","本院","郭綜合醫院",24,0],["1522011115","本院","陽明醫院",8,0],["1531010279","本院","板橋中興醫院",6,0],["1531071030","本院","仁愛醫院",6,0],["1532040039","天成院區","天成醫院",12,0],["1532040066","本院","怡仁綜合醫院",13,0],["1532101091","本院","陽明醫院",8,0],["1535010051","本院","大千綜合醫院",47,0],["1536100081","本院","清泉醫院",10,0],["1542010052","本院","大東醫院",8,0],["1542011282","本院","杏和醫院",8,0],["1542050056","本院","建佑醫院",14,0],["1131130018","本院","新北市立土城醫院(委託長庚醫療財團法人興建經營)",50,0],["1231050017","安康院區","天主教耕莘醫療財團法人耕莘醫院-安康",10,0],["401020013","本院","國立臺灣大學醫學院附設醫院癌醫中心分院",28,0],["643010011","本院","屏東榮民總醫院",26,0],["439010518","虎尾院區","國立臺灣大學醫學院附設醫院雲林分院",18,0],["1307020025","本院","財團法人私立高雄醫學大學附設高醫岡山醫院",10,0]];

// ── Step 0: Load embedded hospital list into sheet ─────────────────────────

function loadHospitalList() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  const header = [["code","campus","name","icu_jul2024","burn_jul2024","","","scraped_icu","scraped_burn","status"]];
  sheet.getRange(1, 1, 1, 10).setValues(header);

  const rows = HOSPITAL_DATA.map(h => [h[0], h[1], h[2], h[3], h[4]]);
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    `Loaded ${rows.length} hospitals.\n\nNext: 🏥 ICU Scraper → 1. Scrape ICU data`
  );
}

// ── Helpers ────────────────────────────────────────────────────────────────

function padCode_(code) {
  return String(code).padStart(10, "0");
}

function getZone_(code, campus) {
  if (!campus || campus === "本院") return "A";
  const zones = KNOWN_ZONES[padCode_(code)];
  if (zones && zones[campus]) return zones[campus];
  return "A";
}

function buildUrl_(code, campus) {
  return BASE_URL
    .replace("{CODE}", padCode_(code))
    .replace("{ZONE}", getZone_(code, campus));
}

function extractBeds_(html) {
  let icu = null;

  const m1 = html.match(/<t[hd][^>]*>[^<]*加護病床[^<]*(?!燒傷)[^<]*<\/t[hd]>\s*<t[hd][^>]*>\s*(\d+)/);
  if (m1) icu = parseInt(m1[1], 10);

  if (icu === null) {
    const m2 = html.match(/<dt[^>]*>[^<]*加護病床[^<]*(?!燒傷)[^<]*<\/dt>\s*<dd[^>]*>\s*(\d+)/);
    if (m2) icu = parseInt(m2[1], 10);
  }

  if (icu === null) {
    const m3 = html.match(/加護病床[^\d燒]{0,20}?(\d+)/);
    if (m3) icu = parseInt(m3[1], 10);
  }

  if (icu === null) return null;

  let burn = 0;
  const mb1 = html.match(/<t[hd][^>]*>[^<]*燒傷加護[^<]*<\/t[hd]>\s*<t[hd][^>]*>\s*(\d+)/);
  if (mb1) {
    burn = parseInt(mb1[1], 10);
  } else {
    const mb2 = html.match(/燒傷加護病床[^\d]{0,20}?(\d+)/);
    if (mb2) burn = parseInt(mb2[1], 10);
  }

  return { icu, burn };
}

// ── Step 1: Scrape ICU data ────────────────────────────────────────────────

function scrapeIcuData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + SHEET_NAME + '" not found.\nRun step 0 first.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data rows. Run step 0 first."); return; }

  const allData = sheet.getRange(2, 1, lastRow - 1, C_STATUS).getValues();

  let done = 0, failed = 0, skipped = 0;
  const startMs = Date.now();

  for (let i = 0; i < allData.length; i++) {
    const row    = allData[i];
    const code   = String(row[C_CODE   - 1] || "").trim();
    const campus = String(row[C_CAMPUS - 1] || "本院").trim();
    const status = String(row[C_STATUS - 1] || "").trim();

    if (!code) continue;
    if (status === "ok") { skipped++; continue; }

    const url = buildUrl_(code, campus);
    let newStatus = "fail";
    let newIcu = "", newBurn = 0;

    try {
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
          "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
          "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
      });

      const httpCode = resp.getResponseCode();
      if (httpCode === 200) {
        const result = extractBeds_(resp.getContentText("UTF-8"));
        if (result !== null) {
          newIcu = result.icu; newBurn = result.burn; newStatus = "ok"; done++;
        } else {
          newStatus = "no-data"; failed++;
        }
      } else if (httpCode === 429) {
        Utilities.sleep(10000); i--; continue;
      } else {
        newStatus = "http-" + httpCode; failed++;
      }
    } catch (e) {
      newStatus = "err:" + String(e).slice(0, 40); failed++;
    }

    sheet.getRange(i + 2, C_NEW_ICU, 1, 3).setValues([[newIcu, newBurn, newStatus]]);
    if ((i + 1) % 10 === 0) SpreadsheetApp.flush();
    Utilities.sleep(350);

    if (Date.now() - startMs > 295000) {
      SpreadsheetApp.getUi().alert(
        "Paused at row " + (i + 2) + " (5-min limit).\n\n" +
        done + " OK, " + failed + " failed, " + skipped + " skipped.\n\n" +
        "Re-run  🏥 ICU Scraper → 1. Scrape ICU data  to continue."
      );
      return;
    }
  }

  SpreadsheetApp.getUi().alert(
    "Scrape complete!\n\n✓ " + done + " succeeded  ✗ " + failed + " failed  ⏭ " + skipped + " already done\n\n" +
    "Run  🏥 ICU Scraper → 2. Email results as CSV"
  );
}

// ── Step 2: Email CSV ──────────────────────────────────────────────────────

function emailResultsCsv() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet "' + SHEET_NAME + '" not found.'); return; }

  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(2, 1, lastRow - 1, C_STATUS).getValues();

  const lines = ["code,campus,name,icu,burn_icu"];
  let count = 0;
  for (const row of allData) {
    const code   = String(row[C_CODE   - 1] || "").trim();
    const campus = String(row[C_CAMPUS - 1] || "").trim();
    const name   = String(row[C_NAME   - 1] || "").replace(/"/g, '""');
    const status = String(row[C_STATUS - 1] || "").trim();
    const icu    = row[C_NEW_ICU  - 1];
    const burn   = row[C_NEW_BURN - 1];
    if (status === "ok" && code) { lines.push(code + "," + campus + ',"' + name + '",' + icu + "," + burn); count++; }
  }

  if (count === 0) { SpreadsheetApp.getUi().alert("No scraped rows. Run step 1 first."); return; }

  const email = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(
    email,
    "ICU scrape results (" + count + " hospitals) — icu_scraped.csv",
    "Scraped " + count + " hospitals from openinfo.mohw.gov.tw.\n\nAttached: icu_scraped.csv",
    { attachments: [Utilities.newBlob(lines.join("\n"), "text/csv", "icu_scraped.csv")] }
  );
  SpreadsheetApp.getUi().alert("CSV with " + count + " rows emailed to " + email + ".");
}

// ── Menu ───────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏥 ICU Scraper")
    .addItem("0. Load hospital list", "loadHospitalList")
    .addSeparator()
    .addItem("1. Scrape ICU data", "scrapeIcuData")
    .addSeparator()
    .addItem("2. Email results as CSV", "emailResultsCsv")
    .addToUi();
}
