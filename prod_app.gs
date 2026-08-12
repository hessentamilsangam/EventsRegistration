/************************************************************
 * 1. Existing doPost — DO NOT MODIFY (kept exactly as given)
 ************************************************************/
function doPost(e) {
  try {
    Logger.log("Incoming request:");
    Logger.log(e.postData ? e.postData.contents : "NO POST DATA RECEIVED");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sheet1");

    const data = JSON.parse(e.postData.contents);

    const row = [
      data.timestamp || "",
      data.eventTitle || "",
      data.registrationId || "",
      data.fullName || "",
      data.email || "",
      data.phone || "",
      data.isMember || "",
      data.htsMemberId || "",
      data.foodPreference || "",
      data.adults || "",
      data.kids || "",
      data.attendeeNames || "",
      "", // attendstatus column initially empty
      "" // paymentStatus column initially empty
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/************************************************************
 * 2. doGet — routing for QR verification
 ************************************************************/
function doGet(e) {
  if (e.parameter.mode === "verify") {
    return verifyRegistrationId(e);
  } else if (e.parameter.mode === "dashboard") {
    return getDashboardData();
  }
  return ContentService.createTextOutput("INVALID_MODE");
}

/************************************************************
 * 3. Secure verification endpoint (for QR scan webpage)
 ************************************************************/
function verifyRegistrationId(e) {
  // NOTE: password check removed to match client-side (no password)
  const regId = e.parameter.registrationId;

  if (!regId) {
    return jsonOutput({ status: "MISSING_ID" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");

  const data = sheet.getDataRange().getValues();
  const header = data[0];

  const regIdCol = header.indexOf("registrationId");
  const fullNameCol = header.indexOf("fullName");
  const attendeeCol = header.indexOf("attendeeNames");
  const statusCol = header.indexOf("attendStatus");

  let found = false;
  let statusMessage = "";
  let fullName = "";
  let attendeeNames = "";

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][regIdCol]) === regId) {

      found = true;

      fullName = data[i][fullNameCol] || "";
      attendeeNames = data[i][attendeeCol] || "";

      if (String(data[i][statusCol]).toUpperCase() === "OK") {
        statusMessage = "Already updated";
      } else {
        sheet.getRange(i + 1, statusCol + 1).setValue("OK");
        statusMessage = "Updated successfully";
      }

      break;
    }
  }

  if (!found) {
    return jsonOutput({
      status: "Invalid registrationId",
      fullName: "",
      attendeeNames: ""
    });
  }

  return jsonOutput({
    status: statusMessage,
    fullName: fullName,
    attendeeNames: attendeeNames
  });
}

/************************************************************
 * 4. Helper — Return JSON properly
 ************************************************************/
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function dashboardResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1');
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  // header map (lowercased) -> column index
  const header = data[0].map(h => String(h || '').trim());
  const idx = {};
  header.forEach((h, i) => { idx[h.toLowerCase()] = i; });

  // helper to find a value by trying multiple possible header names
  function getCell(row, variants) {
    for (let v of variants) {
      if (!v) continue;
      const i = idx[String(v).toLowerCase()];
      if (i !== undefined && i < row.length) return row[i];
    }
    return '';
  }

  function toNumber(v) {
    const n = Number(v);
    return Number.isFinite(n) ? Math.round(n) : 0;
  }

  function normalizeYesNo(v) {
    const s = String(v || '').trim().toLowerCase();
    return (s === 'yes' || s === 'y' || s === 'true' || s === '1') ? 'Yes' : 'No';
  }

  function normalizeOk(v) {
    const s = String(v || '').trim().toLowerCase();
    return (s === 'ok' || s === 'yes' || s === 'attended' || s === 'paid' || s === 'true') ? 'ok' : s;
  }

  const result = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    // Try common header name variants for each expected property
    const regId = getCell(row, ['registrationid']);
    const fullName = getCell(row, ['fullName']);
    const eventTitle = getCell(row, ['eventTitle']);
    const isMemberRaw = getCell(row, ['isMember']);
    const foodPref = getCell(row, ['foodPreference']);
    const adultsRaw = getCell(row, ['adults']);
    const kidsRaw = getCell(row, ['kids']);
    const attendRaw = getCell(row, ['attendStatus']);
    const paymentRaw = getCell(row, ['paymentStatus']);

    const record = {
      registrationId: String(regId || '').trim(),
      fullName: String(fullName || '').trim(),
      eventTitle: String(eventTitle || '').trim(),

      // normalized fields used by the frontend
      isMember: normalizeYesNo(isMemberRaw),           // "Yes" or "No" (frontend checks === 'yes' case-insensitive)
      foodPreference: String(foodPref || '').trim(),   // e.g. "Veg", "NonVeg", "Veg - Jain" etc.
      adults: toNumber(adultsRaw),
      kids: toNumber(kidsRaw),

      // normalize to "ok" when matched; otherwise pass through lowercase-ish string
      attendStatus: normalizeOk(attendRaw),            // frontend checks === 'ok'
      paymentStatus: normalizeOk(paymentRaw)           // frontend checks === 'ok'
    };

    // Add any other raw columns if you want them passed through:
    // header.forEach(h => { record[h] = row[ header.indexOf(h) ]; });

    // Only include non-empty rows (useful to skip blank trailing lines)
    const anyVal = Object.values(record).some(v => v !== '' && v !== 0 && v !== 'No');
    if (anyVal) result.push(record);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}