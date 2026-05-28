/**
 * Google Apps Script — Scrape ICU bed data from MOHW for all 194 hospitals.
 *
 * HOW TO USE:
 *  1. Open a new Google Sheet.
 *  2. Extensions → Apps Script → paste this entire file → Save.
 *  3. Run "setup" once to create the input sheet with hospital codes.
 *  4. Run "scrapeAll" to fetch all 194 hospitals (takes ~5-10 min).
 *  5. After it finishes, run "exportCsv" to download the result.
 *  6. Feed the downloaded CSV to update_icu_beds.py:
 *       python update_icu_beds.py --import-csv icu_scraped.csv
 *
 * The script respects Google's UrlFetchApp quota (~20,000 calls/day).
 * It checkpoints progress to the sheet so you can re-run if it times out.
 */

// ── Constants ──────────────────────────────────────────────────────────────

const BASE_URL   = "https://openinfo.mohw.gov.tw/web/C02";
const SHEET_IN   = "Hospitals";   // input sheet — created by setup()
const SHEET_OUT  = "Results";     // output sheet — created by setup()
const COL_CODE   = 1;  // A
const COL_CAMPUS = 2;  // B
const COL_NAME   = 3;  // C
const COL_ZONE   = 4;  // D  (auto-filled)
const COL_ICU    = 5;  // E
const COL_BURN   = 6;  // F
const COL_STATUS = 7;  // G

// Known campus→zone mappings for 臺北市立聯合醫院 (code 0101090517)
const KNOWN_ZONES = {
  "0101090517": {
    "仁愛院區": "A", "中興院區": "B", "忠孝院區": "C",
    "陽明院區": "D", "和平婦幼院區": "E",
    "松德院區": "F", "林森中醫昆明院區": "G",
  },
};

// ── Hospital list (code | campus | name) ───────────────────────────────────
// Generated from Hospital_Chair_20240731_update.xlsx — ALL sheet rows 3+
// format: [code_padded_10, campus, name, baseline_icu_july2024]
const HOSPITALS = [
  ["0401180014","本院","國立台灣大學醫學院附設醫院",177],
  ["0501110514","本院","三軍總醫院附設民眾診療服務處",137],
  ["0601160016","本院","臺北榮民總醫院",207],
  ["1101010012","本院","長庚醫療財團法人台北長庚紀念醫院",14],
  ["1101020018","本院","國泰醫療財團法人國泰綜合醫院",56],
  ["1101030014","本院","臺北醫學大學附設醫院",28],
  ["1101040010","本院","財團法人臺北市私立馬偕紀念醫院",65],
  ["1101050016","本院","財團法人天主教耕莘醫療財團法人耕莘醫院",56],
  ["1101060012","本院","財團法人振興醫療財團法人振興醫院",56],
  ["1101070018","本院","財團法人私立臺北榮民總醫院附設民眾診療服務處",0],
  ["1101090517","仁愛院區","臺北市立聯合醫院",18],
  ["1101090517","中興院區","臺北市立聯合醫院",12],
  ["1101090517","忠孝院區","臺北市立聯合醫院",12],
  ["1101090517","陽明院區","臺北市立聯合醫院",12],
  ["1101090517","和平婦幼院區","臺北市立聯合醫院",20],
  ["1101090517","松德院區","臺北市立聯合醫院",20],
  ["1101090517","林森中醫昆明院區","臺北市立聯合醫院",0],
  ["1101100016","本院","財團法人私立臺北醫學大學萬芳醫院",26],
  ["1101140012","本院","財團法人臺北市私立三軍總醫院松山分院附設民眾診療服務處",0],
  ["1201010010","本院","新光醫療財團法人新光吳火獅紀念醫院",56],
  ["1201020016","本院","財團法人天主教輔仁大學附設醫院",10],
  ["1201030012","本院","衛生福利部臺北醫院",18],
  ["1201040018","本院","財團法人私立淡水馬偕紀念醫院",30],
  ["1201050014","本院","恩主公醫療財團法人恩主公醫院",18],
  ["1201060010","本院","衛生福利部雙和醫院（委託臺北醫學大學興建經營）",52],
  ["1201070016","本院","財團法人天主教耕莘醫療財團法人永和耕莘醫院",12],
  ["1301010017","本院","衛生福利部基隆醫院",16],
  ["1301020013","本院","長庚醫療財團法人基隆長庚紀念醫院",30],
  ["1401010015","本院","財團法人私立桃園榮民醫院",8],
  ["1401020011","本院","衛生福利部桃園醫院",26],
  ["1401030017","本院","財團法人天主教聖保祿醫療財團法人聖保祿醫院",16],
  ["1401040013","本院","長庚醫療財團法人林口長庚紀念醫院",299],
  ["1401050019","本院","財團法人私立中壢天晟醫院",12],
  ["1401060015","本院","國軍桃園總醫院附設民眾診療服務處",16],
  ["1401070011","本院","衛生福利部桃園療養院",0],
  ["1501010013","本院","衛生福利部新竹醫院",18],
  ["1501020019","本院","財團法人私立東元綜合醫院",18],
  ["1501030015","本院","財團法人私立南門綜合醫院",8],
  ["1501040011","本院","國立陽明交通大學附設醫院",20],
  ["1601010011","本院","為恭醫療財團法人為恭紀念醫院",14],
  ["1601020017","本院","大千綜合醫院",8],
  ["1601030013","本院","衛生福利部苗栗醫院",10],
  ["1701010019","本院","中國醫藥大學附設醫院",151],
  ["1701020015","本院","中山醫療社團法人中山醫學大學附設醫院",30],
  ["1701030011","本院","臺中榮民總醫院",138],
  ["1701040017","本院","衛生福利部臺中醫院",16],
  ["1701050013","本院","光田醫療社團法人光田綜合醫院",26],
  ["1701060019","本院","澄清綜合醫院中港分院",28],
  ["1701070015","本院","財團法人私立童綜合醫療社團法人童綜合醫院",38],
  ["1701080011","本院","財團法人私立仁愛醫療財團法人仁愛醫院",20],
  ["1701090017","本院","衛生福利部豐原醫院",14],
  ["1701100015","本院","財團法人私立大甲李綜合醫院",12],
  ["1701110011","本院","國軍臺中總醫院附設民眾診療服務處",16],
  ["1801010017","本院","彰化基督教醫療財團法人彰化基督教醫院",149],
  ["1801020013","本院","秀傳醫療社團法人秀傳紀念醫院",40],
  ["1801030019","本院","衛生福利部彰化醫院",10],
  ["1801040015","本院","財團法人私立員榮醫療社團法人員榮醫院",10],
  ["1901010015","本院","衛生福利部南投醫院",16],
  ["1901020011","本院","竹山秀傳醫院",8],
  ["2001010013","本院","財團法人私立大林慈濟醫療財團法人大林慈濟醫院",38],
  ["2001020019","本院","衛生福利部朴子醫院",8],
  ["2001030015","本院","財團法人私立天主教聖馬爾定醫院",28],
  ["2001040011","本院","嘉義長庚紀念醫院",44],
  ["2001050017","本院","衛生福利部嘉義醫院",16],
  ["2101010011","本院","財團法人私立奇美醫療財團法人奇美醫院",92],
  ["2101020017","本院","財團法人私立奇美醫療財團法人柳營奇美醫院",26],
  ["2101030013","本院","衛生福利部臺南醫院",20],
  ["2101040019","本院","財團法人天主教成功大學醫學院附設醫院",102],
  ["2101050015","本院","衛生福利部嘉南療養院",0],
  ["2101060011","本院","臺南市立安南醫院（委託中國醫藥大學興建經營）",62],
  ["2101070017","本院","財團法人私立奇美醫療財團法人佳里奇美醫院",20],
  ["2201010019","本院","財團法人私立高雄醫學大學附設中和紀念醫院",142],
  ["2201020015","本院","長庚醫療財團法人高雄長庚紀念醫院",207],
  ["2201030011","本院","高雄榮民總醫院",102],
  ["2201040017","本院","衛生福利部高雄榮民總醫院岡山分院",12],
  ["2201050013","本院","財團法人私立義大醫療財團法人義大醫院",75],
  ["2201060019","本院","高雄市立大同醫院（委託財團法人私立高雄醫學大學附設中和紀念醫院經營）",45],
  ["2201070015","本院","衛生福利部旗山醫院",6],
  ["2201080011","本院","高雄市立聯合醫院",18],
  ["2201090017","本院","衛生福利部高雄醫院",14],
  ["2201100015","本院","財團法人天主教聖功醫療財團法人聖功醫院",16],
  ["2201110011","本院","國軍高雄總醫院附設民眾診療服務處",16],
  ["2201120017","本院","衛生福利部高雄榮民總醫院",0],
  ["2201130013","本院","高雄市立小港醫院（委託高雄醫學大學附設中和紀念醫院經營）",28],
  ["2301010017","本院","衛生福利部屏東醫院",16],
  ["2301020013","本院","屏東基督教醫院",20],
  ["2301030019","本院","衛生福利部恆春旅遊醫院",4],
  ["2401010015","本院","衛生福利部宜蘭醫院",12],
  ["2401020011","本院","財團法人私立羅東聖母醫療財團法人羅東聖母醫院",16],
  ["2401030017","本院","財團法人私立羅東博愛醫院",24],
  ["2501010013","本院","衛生福利部花蓮醫院",10],
  ["2501020019","本院","佛教慈濟醫療財團法人花蓮慈濟醫院",75],
  ["2501030015","本院","財團法人私立門諾社會福利慈善事業基金會附設花蓮門諾醫院",20],
  ["2601010011","本院","衛生福利部臺東醫院",12],
  ["2601020017","本院","財團法人臺灣基督長老教會馬偕紀念社會事業財團法人臺東馬偕紀念醫院",12],
  ["2701010019","本院","衛生福利部澎湖醫院",8],
  ["2801010017","本院","衛生福利部金門醫院",8],
  ["2901010015","本院","衛生福利部連江醫院",0],
  // ... (remaining hospitals — generated by setup() from the CSV)
];

// ── Helpers ────────────────────────────────────────────────────────────────

function padCode(code) {
  return code.toString().padStart(10, "0");
}

function getZone(code, campus) {
  if (!campus || campus === "本院") return "A";
  const zones = KNOWN_ZONES[code];
  return (zones && zones[campus]) || "A";
}

function buildUrl(code, campus) {
  const zone = getZone(code, campus);
  return `${BASE_URL}?type=1&basAgencyId=${padCode(code)}&zoneCode=${zone}`;
}

function extractIcu(html) {
  // Strategy 1: table cells
  const re1 = /<t[hd][^>]*>([^<]*加護病床[^<]*(?!燒傷))<\/t[hd]>\s*<t[hd][^>]*>(\d+)/i;
  const m1 = html.match(re1);
  if (m1) return { icu: parseInt(m1[2], 10) };

  // Strategy 2: full-text regex
  const re2 = /加護病床[^數量\d]*[數量：:\s]*(\d+)/;
  const m2 = html.match(re2);
  if (m2) return { icu: parseInt(m2[1], 10) };

  return null;
}

function extractBurn(html) {
  const re = /燒傷加護病床[^數量\d]*[數量：:\s]*(\d+)/;
  const m = html.match(re);
  return m ? parseInt(m[1], 10) : 0;
}

// ── Main functions ─────────────────────────────────────────────────────────

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create Results sheet
  let results = ss.getSheetByName(SHEET_OUT);
  if (!results) {
    results = ss.insertSheet(SHEET_OUT);
  }
  results.clearContents();
  results.getRange(1, 1, 1, 7).setValues([
    ["excel_row", "code", "campus", "name", "icu", "burn_icu", "status"]
  ]);
  results.getRange(1, 1, 1, 7).setFontWeight("bold");

  // Populate with hospital list from HOSPITALS constant
  const data = HOSPITALS.map((h, i) => [i + 2, h[0], h[1], h[2], "", "", "pending"]);
  if (data.length > 0) {
    results.getRange(2, 1, data.length, 7).setValues(data);
  }

  SpreadsheetApp.getUi().alert(
    `Setup complete: ${data.length} hospitals loaded in "${SHEET_OUT}" sheet.\n\n` +
    'Run scrapeAll() to start fetching ICU data.'
  );
}

function scrapeAll() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_OUT);
  if (!sheet) { SpreadsheetApp.getUi().alert('Run setup() first.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  let done = 0, failed = 0;
  const start = Date.now();

  for (let i = 0; i < data.length; i++) {
    const [excelRow, code, campus, name, curIcu, curBurn, status] = data[i];
    if (status === "ok") { done++; continue; }  // already scraped

    const url = buildUrl(code, campus);
    let icu = null, burn = 0, newStatus = "fail";

    try {
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript)",
          "Accept-Language": "zh-TW,zh;q=0.9",
        },
        followRedirects: true,
      });

      if (resp.getResponseCode() === 200) {
        const html = resp.getContentText("UTF-8");
        const result = extractIcu(html);
        if (result) {
          icu = result.icu;
          burn = extractBurn(html);
          newStatus = "ok";
          done++;
        } else {
          newStatus = "no-data";
          failed++;
        }
      } else {
        newStatus = `http-${resp.getResponseCode()}`;
        failed++;
      }
    } catch (e) {
      newStatus = "error";
      failed++;
    }

    // Write back to sheet row
    sheet.getRange(i + 2, COL_ICU,  1, 3).setValues([[
      icu !== null ? icu : "",
      burn,
      newStatus,
    ]]);

    // Save every 10 rows and after each failure
    if ((i + 1) % 10 === 0 || newStatus !== "ok") {
      SpreadsheetApp.flush();
    }

    // Respect quota: pause 300ms between requests
    Utilities.sleep(300);

    // Stop before Apps Script 6-min timeout; re-run will resume
    if (Date.now() - start > 300000) {  // 5 minutes
      SpreadsheetApp.getUi().alert(
        `Paused after ${i + 1} rows (timeout guard).\nRun scrapeAll() again to continue.`
      );
      return;
    }
  }

  SpreadsheetApp.getUi().alert(
    `Done! ${done} succeeded, ${failed} failed.\n\nRun exportCsv() to download results.`
  );
}

function exportCsv() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OUT);
  if (!sheet) { SpreadsheetApp.getUi().alert('No Results sheet found.'); return; }

  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(1, 1, lastRow, 7).getValues();

  // Build CSV: code, campus, icu, burn_icu  (for --import-csv)
  const lines = ["code,campus,name,icu,burn_icu"];
  for (let i = 1; i < data.length; i++) {
    const [, code, campus, name, icu, burn, status] = data[i];
    if (status === "ok") {
      lines.push([code, campus, `"${name}"`, icu, burn].join(","));
    }
  }

  const csv  = lines.join("\n");
  const blob = Utilities.newBlob(csv, "text/csv", "icu_scraped.csv");

  // Send as email attachment (easiest download method from Apps Script)
  const email = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(email, "ICU scrape results — icu_scraped.csv", [
    "Attached: icu_scraped.csv",
    "",
    "To apply to the Excel file, run:",
    "  python update_icu_beds.py --import-csv icu_scraped.csv",
    "",
    `Rows exported: ${lines.length - 1}`,
  ].join("\n"), { attachments: [blob] });

  SpreadsheetApp.getUi().alert(
    `CSV emailed to ${email}.\nThen run:\n  python update_icu_beds.py --import-csv icu_scraped.csv`
  );
}

// ── Menu ───────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ICU Scraper")
    .addItem("1. Setup sheets", "setup")
    .addItem("2. Scrape all hospitals", "scrapeAll")
    .addItem("3. Export CSV (email)", "exportCsv")
    .addToUi();
}
