// ═══════════════════════════════════════════════════════
//  Squad Gym — Google Apps Script Backend
// ═══════════════════════════════════════════════════════
//
//  POSTUP:
//  1. Otevři script.google.com → klikni "Nový projekt"
//  2. Celý tento soubor vlož místo obsahu (Code.gs)
//  3. Klikni: Nasadit → Nové nasazení
//     - Typ: Webová aplikace
//     - Spustit jako: Já
//     - Kdo má přístup: Kdokoli
//  4. Zkopíruj "URL nasazení" a vlož do index.html jako SCRIPT_URL
//  5. Při změně kódu vždy: Nasadit → Spravovat nasazení → upravit verzi
// ═══════════════════════════════════════════════════════

function doGet(e) {
  try {
    const gc = (e.parameter && e.parameter.gc) || 'sg_default';
    return json(getAllData(gc));
  } catch (err) {
    return json({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case 'upsertCheckin':   return json(upsertCheckin(d));
      case 'upsertPR':        return json(upsertPR(d));
      case 'upsertProfile':   return json(upsertProfile(d));
      case 'insertChallenge': return json(insertChallenge(d));
      case 'upsertProgress':  return json(upsertProgress(d));
      case 'upsertSetting':   return json(upsertSetting(d));
      case 'deletePR':        return json(deletePR(d));
      case 'deactivateChallenge': return json(deactivateChallenge(d));
      default:                return json({ error: 'Neznámá akce: ' + d.action });
    }
  } catch (err) {
    return json({ error: err.toString() });
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheets setup ──────────────────────────────────────
const SCHEMA = {
  checkins:   ['group_code','member','date','status','note','ts'],
  prs:        ['group_code','member','exercise','value','unit','date','ts'],
  profiles:   ['group_code','member','avatar_data','stats_json','updated_at'],
  challenges: ['id','group_code','title','description','type','target_value','end_date','created_by','created_at','active'],
  progress:   ['id','challenge_id','group_code','member','value','updated_at'],
  settings:   ['group_code','key','value','updated_at']
};

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = SCHEMA[name];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#f1f5f9');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Google Sheets může auto-konvertovat date stringy na Date objekty.
// Tato funkce je vždy převede zpět na 'YYYY-MM-DD' string.
function normVal(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' +
           String(v.getMonth() + 1).padStart(2, '0') + '-' +
           String(v.getDate()).padStart(2, '0');
  }
  return v;
}

function sheetRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      if (v === '' || v === null) { obj[h] = null; }
      else if (v instanceof Date) { obj[h] = normVal(v); }
      else { obj[h] = v; }
    });
    return obj;
  });
}

// ── GET all data ──────────────────────────────────────
function getAllData(gc) {
  // Ensure all sheets exist
  Object.keys(SCHEMA).forEach(name => getSheet(name));

  let checkins   = sheetRows(getSheet('checkins')).filter(r => r.group_code === gc);
  let prs        = sheetRows(getSheet('prs')).filter(r => r.group_code === gc);
  let profiles   = sheetRows(getSheet('profiles')).filter(r => r.group_code === gc);
  let challenges = sheetRows(getSheet('challenges')).filter(r => r.group_code === gc);
  let progress   = sheetRows(getSheet('progress')).filter(r => r.group_code === gc);
  let settings   = sheetRows(getSheet('settings')).filter(r => r.group_code === gc);

  // Typová korekce
  prs.forEach(r => { r.value = r.value !== null ? parseFloat(r.value) : null; });

  challenges.forEach(r => {
    r.target_value = r.target_value !== null ? parseFloat(r.target_value) || null : null;
    r.active = (r.active === true || r.active === 'TRUE' || r.active === 'true');
  });
  challenges = challenges.filter(r => r.active);

  progress.forEach(r => { r.value = r.value !== null ? parseFloat(r.value) : 0; });

  return { checkins, prs, profiles, challenges, progress, settings };
}

// ── Upsert helper ─────────────────────────────────────
function upsertByKeys(sheetName, keyFields, data) {
  const sheet = getSheet(sheetName);
  const headers = SCHEMA[sheetName];
  const rows = sheet.getDataRange().getValues();
  const sheetHeaders = rows[0];

  let foundRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (keyFields.every(k => {
      const idx = sheetHeaders.indexOf(k);
      return idx !== -1 && String(normVal(rows[i][idx])) === String(normVal(data[k]));
    })) {
      foundRow = i;
      break;
    }
  }

  const newRow = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : '');

  if (foundRow !== -1) {
    sheet.getRange(foundRow + 1, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  return { ok: true };
}

// ── Operace ───────────────────────────────────────────
function upsertCheckin(d) {
  return upsertByKeys('checkins', ['group_code','member','date'], d);
}

function upsertPR(d) {
  return upsertByKeys('prs', ['group_code','member','exercise'], d);
}

function upsertProfile(d) {
  return upsertByKeys('profiles', ['group_code','member'], d);
}

function insertChallenge(d) {
  const sheet = getSheet('challenges');
  const headers = SCHEMA['challenges'];
  const id = d.id || Utilities.getUuid();
  const row = headers.map(h => {
    if (h === 'id') return id;
    if (h === 'active') return true;
    return (d[h] !== undefined && d[h] !== null) ? d[h] : '';
  });
  sheet.appendRow(row);
  return { ok: true, id };
}

function upsertSetting(d) {
  return upsertByKeys('settings', ['group_code','key'], d);
}

function deletePR(d) {
  const sheet = getSheet('prs');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: false };
  const h = rows[0];
  const gi = h.indexOf('group_code'), mi = h.indexOf('member'), ei = h.indexOf('exercise');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][gi]) === String(d.group_code) &&
        String(rows[i][mi]) === String(d.member) &&
        String(rows[i][ei]) === String(d.exercise)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

function deactivateChallenge(d) {
  const sheet = getSheet('challenges');
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: false };
  const h = rows[0];
  const ii = h.indexOf('id'), ai = h.indexOf('active');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][ii]) === String(d.id)) {
      if (ai !== -1) sheet.getRange(i + 1, ai + 1).setValue(false);
      return { ok: true };
    }
  }
  return { ok: false };
}

function upsertProgress(d) {
  const sheet = getSheet('progress');
  const rows = sheet.getDataRange().getValues();
  if (rows.length > 1) {
    const h = rows[0];
    const ci = h.indexOf('challenge_id'), mi = h.indexOf('member');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][ci]) === String(d.challenge_id) && String(rows[i][mi]) === String(d.member)) {
        const vi = h.indexOf('value'), ui = h.indexOf('updated_at');
        if (vi !== -1) sheet.getRange(i+1, vi+1).setValue(d.value);
        if (ui !== -1) sheet.getRange(i+1, ui+1).setValue(d.updated_at || Date.now());
        return { ok: true };
      }
    }
  }
  sheet.appendRow([Utilities.getUuid(), d.challenge_id, d.group_code, d.member, d.value, d.updated_at || Date.now()]);
  return { ok: true };
}
