/**
 * Google Apps Script — Scrape live ICU bed data from MOHW for all 194 hospitals.
 *
 * SETUP (one time):
 *  1. Open https://sheets.google.com and create a new spreadsheet.
 *  2. File → Import → upload icu_export.csv → "Replace current sheet" → Import.
 *  3. Rename that sheet to "ICU" (click the tab at the bottom).
 *  4. Extensions → Apps Script → delete any existing code → paste this entire file → Save.
 *  5. Reload the spreadsheet.  A new "🏥 ICU Scraper" menu will appear.
 *
 * RUNNING:
 *  1. Click  🏥 ICU Scraper → 1. Scrape ICU data
 *     - Fetches each row's MOHW page (300ms between requests).
 *     - Auto-pauses after 5 min to avoid Apps Script timeout; re-run to resume.
 *     - Progress saved to columns H (scraped_icu) and I (scraped_burn) and J (status).
 *  2. Click  🏥 ICU Scraper → 2. Email results as CSV
 *     - Sends icu_scraped.csv to your Google account email.
 *  3. Feed the CSV to the Python script:
 *       python update_icu_beds.py --import-csv icu_scraped.csv
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

// Column indices in the imported icu_export.csv (1-based)
// Header row: code, campus, name, icu, burn_icu
const C_CODE   = 1;  // A
const C_CAMPUS = 2;  // B
const C_NAME   = 3;  // C
const C_ICU    = 4;  // D — original July-2024 baseline
const C_BURN   = 5;  // E — original July-2024 baseline
// Output columns appended by this script:
const C_NEW_ICU  = 8;  // H
const C_NEW_BURN = 9;  // I
const C_STATUS   = 10; // J

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

/**
 * Extract ICU and burn ICU counts from raw HTML text.
 * Returns {icu: number, burn: number} or null if not found.
 */
function extractBeds_(html) {
  // --- ICU count ---
  let icu = null;

  // Strategy 1: table cell pattern <td>加護病床…</td><td>177</td>
  const m1 = html.match(/<t[hd][^>]*>[^<]*加護病床[^<]*(?!燒傷)[^<]*<\/t[hd]>\s*<t[hd][^>]*>\s*(\d+)/);
  if (m1) icu = parseInt(m1[1], 10);

  // Strategy 2: dl/dt+dd pattern
  if (icu === null) {
    const m2 = html.match(/<dt[^>]*>[^<]*加護病床[^<]*(?!燒傷)[^<]*<\/dt>\s*<dd[^>]*>\s*(\d+)/);
    if (m2) icu = parseInt(m2[1], 10);
  }

  // Strategy 3: raw text pattern
  if (icu === null) {
    const m3 = html.match(/加護病床[^\d燒]{0,20}?(\d+)/);
    if (m3) icu = parseInt(m3[1], 10);
  }

  if (icu === null) return null;

  // --- Burn ICU count ---
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

// ── Main scraper ───────────────────────────────────────────────────────────

function scrapeIcuData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      `Sheet "${SHEET_NAME}" not found.\n\n` +
      "Import icu_export.csv and rename the sheet to 'ICU', then try again."
    );
    return;
  }

  // Add output column headers if not present
  const headerRow = sheet.getRange(1, 1, 1, C_STATUS).getValues()[0];
  if (!headerRow[C_NEW_ICU - 1] || headerRow[C_NEW_ICU - 1] === "") {
    sheet.getRange(1, C_NEW_ICU).setValue("scraped_icu");
    sheet.getRange(1, C_NEW_BURN).setValue("scraped_burn");
    sheet.getRange(1, C_STATUS).setValue("status");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data rows found."); return; }

  const allData = sheet.getRange(2, 1, lastRow - 1, C_STATUS).getValues();

  let done = 0, failed = 0, skipped = 0;
  const startMs = Date.now();

  for (let i = 0; i < allData.length; i++) {
    const row    = allData[i];
    const code   = String(row[C_CODE   - 1] || "").trim();
    const campus = String(row[C_CAMPUS - 1] || "本院").trim();
    const status = String(row[C_STATUS - 1] || "").trim();

    if (!code) continue;

    // Skip rows already scraped successfully
    if (status === "ok") { skipped++; continue; }

    const url = buildUrl_(code, campus);
    let newStatus = "fail";
    let newIcu = "", newBurn = 0;

    try {
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript/1.0)",
          "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        },
      });

      const httpCode = resp.getResponseCode();
      if (httpCode === 200) {
        const html   = resp.getContentText("UTF-8");
        const result = extractBeds_(html);
        if (result !== null) {
          newIcu    = result.icu;
          newBurn   = result.burn;
          newStatus = "ok";
          done++;
        } else {
          newStatus = "no-data";
          failed++;
        }
      } else if (httpCode === 429) {
        // Rate limited — pause and retry this row
        Utilities.sleep(10000);
        i--;  // retry
        continue;
      } else {
        newStatus = `http-${httpCode}`;
        failed++;
      }
    } catch (e) {
      newStatus = `err:${String(e).slice(0, 40)}`;
      failed++;
    }

    // Write result back
    sheet.getRange(i + 2, C_NEW_ICU,  1, 3).setValues([[newIcu, newBurn, newStatus]]);

    // Flush every 10 rows
    if ((i + 1) % 10 === 0) SpreadsheetApp.flush();

    Utilities.sleep(350);  // ~350ms between requests

    // Auto-pause before 5-min Apps Script hard limit
    if (Date.now() - startMs > 295000) {
      SpreadsheetApp.getUi().alert(
        `Paused after processing row ${i + 2} (timeout guard).\n\n` +
        `Progress: ${done} OK, ${failed} failed, ${skipped} skipped.\n\n` +
        "Re-run 🏥 ICU Scraper → 1. Scrape ICU data to continue."
      );
      return;
    }
  }

  SpreadsheetApp.getUi().alert(
    `Scrape complete!\n\n` +
    `✓ ${done} succeeded  ✗ ${failed} failed  ⏭ ${skipped} already done\n\n` +
    "Run 🏥 ICU Scraper → 2. Email results as CSV when ready."
  );
}

// ── CSV export ─────────────────────────────────────────────────────────────

function emailResultsCsv() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found.`); return; }

  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(2, 1, lastRow - 1, C_STATUS).getValues();

  const lines = ["code,campus,name,icu,burn_icu"];
  let count = 0;
  for (const row of allData) {
    const code   = String(row[C_CODE   - 1] || "").trim();
    const campus = String(row[C_CAMPUS - 1] || "本院").trim();
    const name   = String(row[C_NAME   - 1] || "").replace(/"/g, '""');
    const status = String(row[C_STATUS - 1] || "").trim();
    const icu    = row[C_NEW_ICU  - 1];
    const burn   = row[C_NEW_BURN - 1];

    if (status === "ok" && code) {
      lines.push(`${code},${campus},"${name}",${icu},${burn}`);
      count++;
    }
  }

  if (count === 0) {
    SpreadsheetApp.getUi().alert("No successfully scraped rows found. Run the scraper first.");
    return;
  }

  const csv   = lines.join("\n");
  const blob  = Utilities.newBlob(csv, "text/csv", "icu_scraped.csv");
  const email = Session.getActiveUser().getEmail();

  GmailApp.sendEmail(
    email,
    `ICU scrape results (${count} hospitals) — icu_scraped.csv`,
    [
      `Scraped ${count} hospitals from openinfo.mohw.gov.tw.`,
      "",
      "To apply results to the Excel file, run:",
      "  python update_icu_beds.py --import-csv icu_scraped.csv",
      "",
      "To preview changes without writing:",
      "  python update_icu_beds.py --import-csv icu_scraped.csv --diff",
    ].join("\n"),
    { attachments: [blob] }
  );

  SpreadsheetApp.getUi().alert(`CSV with ${count} rows emailed to ${email}.`);
}

// ── Menu ───────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏥 ICU Scraper")
    .addItem("1. Scrape ICU data", "scrapeIcuData")
    .addSeparator()
    .addItem("2. Email results as CSV", "emailResultsCsv")
    .addToUi();
}
