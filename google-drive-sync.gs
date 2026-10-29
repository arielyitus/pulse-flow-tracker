/**
 * @OnlyCurrentDoc
 * ============================================================================
 * PULSE & FLOW — LOCKED-DOWN GOOGLE SHEET SYNC (ZERO GOOGLE DRIVE ACCESS!)
 * ============================================================================
 * WHY THIS IS SAFE:
 * The "@OnlyCurrentDoc" tag at the top of this file tells Google to JAIL this
 * script strictly to the ONE single Google Sheet you paste it into.
 * It does NOT use DriveApp and has ZERO permission to see, read, edit, or
 * delete any other files or folders in your Google Drive!
 *
 * SETUP INSTRUCTIONS (90 Seconds):
 * 1. Open a new blank Google Sheet by typing https://sheets.new in your browser
 *    and name it "Pulse & Flow Fitness Data".
 * 2. In that Google Sheet, click "Extensions" (top menu) -> "Apps Script".
 * 3. Delete the default 3 lines in Code.gs and paste this entire file.
 * 4. Click Save (disk icon), then click "Deploy" (top right) -> "New deployment".
 * 5. Click the gear icon -> choose "Web app":
 *      - Execute as: "Me"
 *      - Who has access: "Anyone" (protected by your Secret Sync PIN below)
 * 6. Click "Deploy" -> "Authorize access".
 *    Notice the permission screen now ONLY says:
 *    "View and manage spreadsheets that this application has been installed in"
 *    (It cannot see or touch any other file in your Google Drive!)
 * 7. Copy the Web app URL (ends in /exec) and paste it into the tracker app.
 * ============================================================================
 */

const STATE_TAB_NAME = '_AppStateJSON';
const LOG_TAB_NAME = 'Daily_Tracker_Log';
const CHUNK_SIZE = 40000; // Google Sheets max cell size is 50,000 chars; 40k is safely below

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readStoredWrapper(ss) {
  const sheet = ss.getSheetByName(STATE_TAB_NAME);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null;
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const fullJson = values.map(r => r[0] || '').join('');
  if (!fullJson) return null;
  return JSON.parse(fullJson);
}

function writeStoredWrapper(ss, wrapperObj) {
  const sheet = getOrCreateSheet(ss, STATE_TAB_NAME);
  sheet.clearContents();
  const fullJson = JSON.stringify(wrapperObj);
  const chunks = [];
  for (let i = 0; i < fullJson.length; i += CHUNK_SIZE) {
    chunks.push([fullJson.slice(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length > 0) {
    sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
  }
}

/**
 * Also writes a clean, human-readable summary table to the "Daily_Tracker_Log" tab
 * so you can view your weight, calories, protein, water, and workouts in Google Sheets!
 */
function updateReadableSummarySheet(ss, state) {
  try {
    const sheet = getOrCreateSheet(ss, LOG_TAB_NAME);
    sheet.clearContents();

    const headers = [['Date', 'Weight (lbs)', 'Calories (kcal)', 'Protein (g)', 'Carbs (g)', 'Fat (g)', 'Water (L)', 'Readiness', 'Headline / Notes']];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');

    const allDatesSet = {};
    if (state.weightLogs) Object.keys(state.weightLogs).forEach(d => allDatesSet[d] = true);
    if (state.nutritionLogs) Object.keys(state.nutritionLogs).forEach(d => allDatesSet[d] = true);
    if (state.journals) Object.keys(state.journals).forEach(d => allDatesSet[d] = true);

    const sortedDates = Object.keys(allDatesSet).sort().reverse();
    const rows = sortedDates.slice(0, 120).map(d => {
      const w = (state.weightLogs && state.weightLogs[d]) ? state.weightLogs[d] : '';
      const water = (state.waterLogs && state.waterLogs[d]) ? state.waterLogs[d] : 0;
      const nLog = (state.nutritionLogs && state.nutritionLogs[d]) ? state.nutritionLogs[d] : {};
      let cal = 0, prot = 0, carb = 0, fat = 0;
      ['breakfast', 'preworkout', 'lunch', 'dinner', 'snacks'].forEach(slot => {
        (nLog[slot] || []).forEach(item => {
          const m = Number(item.servings) || 1;
          cal += (Number(item.cal) || 0) * m;
          prot += (Number(item.prot) || 0) * m;
          carb += (Number(item.carb) || 0) * m;
          fat += (Number(item.fat) || 0) * m;
        });
      });
      const j = (state.journals && state.journals[d]) ? state.journals[d] : {};
      return [
        d,
        w,
        Math.round(cal),
        Math.round(prot * 10) / 10,
        Math.round(carb * 10) / 10,
        Math.round(fat * 10) / 10,
        water,
        j.trafficLight || '',
        j.headline || ''
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
  } catch (e) {
    // Non-fatal if readable table update fails
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const requestedPin = (params.pin || '').trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parsed = readStoredWrapper(ss);

    if (!parsed) {
      return jsonResponse({
        ok: true,
        exists: false,
        message: 'Sheet is ready! Push your first sync from the app.',
        state: null
      });
    }

    if (parsed._syncPin && parsed._syncPin !== requestedPin) {
      return jsonResponse({
        ok: false,
        error: 'Invalid Sync PIN. Check your Sync PIN in the Cloud Sync settings.'
      });
    }

    return jsonResponse({
      ok: true,
      exists: true,
      lastUpdated: parsed.updatedAt || 0,
      state: parsed.state || parsed
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'Empty POST body' });
    }

    const payload = JSON.parse(e.postData.contents);
    const incomingPin = (payload.pin || '').trim();
    const newState = payload.state;

    if (!newState || typeof newState !== 'object') {
      return jsonResponse({ ok: false, error: 'Missing state object in payload' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = readStoredWrapper(ss);
    if (existing && existing._syncPin && existing._syncPin !== incomingPin) {
      return jsonResponse({
        ok: false,
        error: 'Invalid Sync PIN. Cannot overwrite Google Sheet state.'
      });
    }

    const wrapper = {
      _syncPin: incomingPin,
      savedAtIso: new Date().toISOString(),
      updatedAt: newState.updatedAt || Date.now(),
      state: newState
    };

    writeStoredWrapper(ss, wrapper);
    updateReadableSummarySheet(ss, newState);

    return jsonResponse({
      ok: true,
      folderName: 'Single-Sheet Locked (@OnlyCurrentDoc)',
      fileName: ss.getName(),
      updatedAt: wrapper.updatedAt
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
