const SPREADSHEET_ID = '1otC4NLLEZqTYxvlu9q1uAIWRxPNs909m8B3sh7JxM60';
const MAIN_SHEET_NAME = 'Sh2';
const MATERIAL_SHEET_NAME = 'Sh1';
const HEADER_ROW = 1;
const SH1_HEADER_ROW = 1;
const ADMIN_PIN = '5555';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard งานก่อสร้าง (WBS) - กฟภ.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getTargetSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  throw new Error('ไม่พบ Spreadsheet กรุณาระบุ SPREADSHEET_ID ในไฟล์ Code.gs');
}

function isOwnerSession_() {
  try {
    const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    return Boolean(active && effective && active === effective);
  } catch (e) {
    return false;
  }
}

function verifyAdminPin(pin) {
  if (isOwnerSession_()) return { ok: true };
  if (String(pin || '').trim() === String(ADMIN_PIN)) return { ok: true };
  return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
}

function parseNum_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const cleaned = String(v || '').replace(/[,฿%\s]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function findColIndexByHints_(headers, hints, exactHints) {
  if (exactHints && exactHints.length) {
    for (let e = 0; e < exactHints.length; e++) {
      const ex = exactHints[e].toLowerCase();
      const idx = headers.findIndex(col => String(col || '').trim().toLowerCase() === ex);
      if (idx >= 0) return idx;
    }
  }
  for (let h = 0; h < hints.length; h++) {
    const hint = hints[h].toLowerCase();
    const idx = headers.findIndex(col => String(col || '').toLowerCase().includes(hint));
    if (idx >= 0) return idx;
  }
  return -1;
}

function resolveSh1Map_(headers, clientSh1Map) {
  let map = clientSh1Map ? Object.assign({}, clientSh1Map) : {};
  if (map.wbs === undefined || Number(map.wbs) < 0) {
    const propsRaw = PropertiesService.getScriptProperties().getProperty('WBS_DASHBOARD_CONFIG');
    if (propsRaw) {
      try {
        const parsed = JSON.parse(propsRaw);
        if (parsed && parsed.sh1ColMap) map = Object.assign({}, parsed.sh1ColMap, map);
      } catch (e) {}
    }
  }
  const pick = (key, hints, exactHints, preferIdx) => {
    if (map[key] !== undefined && Number(map[key]) >= 0) return Number(map[key]);
    if (preferIdx !== undefined && headers.length > preferIdx) return preferIdx;
    return findColIndexByHints_(headers, hints, exactHints);
  };
  return {
    wbsIdx: pick('wbs', ['wbs', 'องค์ประกอบ', 'รหัสงาน']),
    matCodeIdx: pick('matCode', ['เลขพัสดุ', 'รหัสพัสดุ', 'รหัสวัสดุ', 'material'], ['วัสดุ']),
    matIdx: pick('matText', ['ข้อความวัสดุ', 'ชื่อวัสดุ', 'รายการวัสดุ', 'รายละเอียดวัสดุ']),
    moveIdx: pick('moveType', ['ประเภทการเคลื่อนย้าย', 'การเคลื่อนย้าย', 'เคลื่อนย้าย', 'movement'], null, 4),
    colVIdx: pick('colVVal', ['ต้องการ'], null, 21),
    mb52Idx: pick('mb52Val', ['ยอด mb52', 'mb52']),
    matMoveIdx: pick('matMoveStatus', ['การเคลื่อนย้ายพัสดุ', 'เคลื่อนย้ายพัสดุ'])
  };
}

function isValidPendingMaterialRow_(valV, moveStr, matCodeStr, matTextStr) {
  if (!(valV > 0)) return false;
  const mt = String(moveStr || '').trim();
  if (mt === '581' || (mt !== '' && !isNaN(Number(mt)) && Number(mt) === 581)) return false;
  const code = String(matCodeStr || '').trim();
  if (code.startsWith('9')) return false;
  const txt = String(matTextStr || '').toUpperCase();
  if (txt.includes('METER')) return false;
  return true;
}

function getSheetDashboardData() {
  try {
    const ss = getTargetSpreadsheet_();
    const sheet = MAIN_SHEET_NAME ? ss.getSheetByName(MAIN_SHEET_NAME) : ss.getSheets()[0];
    if (!sheet) {
      const available = ss.getSheets().map(s => '"' + s.getName() + '"').join(', ');
      throw new Error('ไม่พบแผ่นงาน "' + MAIN_SHEET_NAME + '" (แผ่นงานที่มีคือ: ' + available + ')');
    }
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < HEADER_ROW + 1 || lastCol < 1) throw new Error('แผ่นงาน "' + sheet.getName() + '" ยังไม่มีข้อมูลเพียงพอ');

    const allValues = sheet.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol).getDisplayValues();
    let sh1Headers = [];
    const sh1 = ss.getSheetByName(MATERIAL_SHEET_NAME);
    if (sh1 && sh1.getLastRow() >= SH1_HEADER_ROW && sh1.getLastColumn() >= 1) {
      sh1Headers = sh1.getRange(SH1_HEADER_ROW, 1, 1, sh1.getLastColumn()).getDisplayValues()[0];
    }

    const savedConfigRaw = PropertiesService.getScriptProperties().getProperty('WBS_DASHBOARD_CONFIG');
    let savedConfig = null;
    if (savedConfigRaw) { try { savedConfig = JSON.parse(savedConfigRaw); } catch (e) {} }

    return {
      ok: true, isOwner: isOwnerSession_(), sheetTitle: ss.getName(), tabName: sheet.getName(),
      headers: allValues[0], rows: allValues.slice(1), sh1Headers: sh1Headers,
      sh1Found: Boolean(sh1), savedConfig: savedConfig,
      fetchedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function getMaterialsForWbs(wbsCode, clientSh1Map) {
  try {
    const ss = getTargetSpreadsheet_();
    const sh1 = ss.getSheetByName(MATERIAL_SHEET_NAME);
    if (!sh1) throw new Error('ไม่พบแผ่นงาน "' + MATERIAL_SHEET_NAME + '"');
    const lastRow = sh1.getLastRow(), lastCol = sh1.getLastColumn();
    if (lastRow < SH1_HEADER_ROW + 1 || lastCol < 1) return { ok: true, items: [], totalRowsForWbs: 0 };

    const allValues = sh1.getRange(SH1_HEADER_ROW, 1, lastRow - SH1_HEADER_ROW + 1, lastCol).getDisplayValues();
    const headers = allValues[0], rows = allValues.slice(1);
    const idx = resolveSh1Map_(headers, clientSh1Map);
    if (idx.wbsIdx < 0) throw new Error('ไม่พบคอลัมน์ WBS ในชีต ' + MATERIAL_SHEET_NAME);

    const target = String(wbsCode || '').trim().toUpperCase();
    let totalRowsForWbs = 0;
    const matchedRows = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rowWbs = (idx.wbsIdx < r.length) ? String(r[idx.wbsIdx] || '').trim().toUpperCase() : '';
      if (rowWbs !== target) continue;
      totalRowsForWbs++;

      const moveStr = (idx.moveIdx >= 0 && idx.moveIdx < r.length) ? String(r[idx.moveIdx] || '').trim() : '';
      const valV = (idx.colVIdx >= 0 && idx.colVIdx < r.length) ? parseNum_(r[idx.colVIdx]) : 0;
      const matCode = (idx.matCodeIdx >= 0 && idx.matCodeIdx < r.length && String(r[idx.matCodeIdx]).trim() !== '') ? String(r[idx.matCodeIdx]).trim() : '-';
      const matText = (idx.matIdx >= 0 && idx.matIdx < r.length && String(r[idx.matIdx]).trim() !== '') ? String(r[idx.matIdx]).trim() : '-';

      if (isValidPendingMaterialRow_(valV, moveStr, matCode, matText)) {
        matchedRows.push({
          matCode: matCode,
          matText: matText,
          moveType: moveStr || '-',
          colVVal: valV,
          mb52Val: (idx.mb52Idx >= 0 && idx.mb52Idx < r.length && String(r[idx.mb52Idx]).trim() !== '') ? String(r[idx.mb52Idx]).trim() : '-',
          matMoveStatus: (idx.matMoveIdx >= 0 && idx.matMoveIdx < r.length && String(r[idx.matMoveIdx]).trim() !== '') ? String(r[idx.matMoveIdx]).trim() : '-'
        });
      }
    }

    return { ok: true, sheetName: sh1.getName(), totalRowsForWbs: totalRowsForWbs, items: matchedRows };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function getAggregatedMaterialsForWbsList(wbsList, clientSh1Map) {
  try {
    const ss = getTargetSpreadsheet_();
    const sh1 = ss.getSheetByName(MATERIAL_SHEET_NAME);
    if (!sh1) throw new Error('ไม่พบแผ่นงาน "' + MATERIAL_SHEET_NAME + '"');
    const lastRow = sh1.getLastRow(), lastCol = sh1.getLastColumn();
    if (lastRow < SH1_HEADER_ROW + 1 || lastCol < 1) return { ok: true, rows: [] };

    const allValues = sh1.getRange(SH1_HEADER_ROW, 1, lastRow - SH1_HEADER_ROW + 1, lastCol).getDisplayValues();
    const headers = allValues[0], rows = allValues.slice(1);
    const idx = resolveSh1Map_(headers, clientSh1Map);
    if (idx.wbsIdx < 0) throw new Error('ไม่พบคอลัมน์ WBS ในชีต ' + MATERIAL_SHEET_NAME);

    const targetSet = {};
    (wbsList || []).forEach(w => {
      const k = String(w || '').trim().toUpperCase();
      if (k) targetSet[k] = true;
    });

    const matched = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rawWbs = (idx.wbsIdx < r.length) ? String(r[idx.wbsIdx] || '').trim() : '';
      if (!rawWbs || !targetSet[rawWbs.toUpperCase()]) continue;

      const moveStr = (idx.moveIdx >= 0 && idx.moveIdx < r.length) ? String(r[idx.moveIdx] || '').trim() : '';
      const valV = (idx.colVIdx >= 0 && idx.colVIdx < r.length) ? parseNum_(r[idx.colVIdx]) : 0;
      const matCode = (idx.matCodeIdx >= 0 && idx.matCodeIdx < r.length && String(r[idx.matCodeIdx]).trim() !== '') ? String(r[idx.matCodeIdx]).trim() : '-';
      const matText = (idx.matIdx >= 0 && idx.matIdx < r.length && String(r[idx.matIdx]).trim() !== '') ? String(r[idx.matIdx]).trim() : '-';

      if (isValidPendingMaterialRow_(valV, moveStr, matCode, matText)) {
        matched.push({
          wbs: rawWbs,
          matCode: matCode,
          matText: matText,
          moveType: moveStr || '-',
          colVVal: valV,
          mb52Val: (idx.mb52Idx >= 0 && idx.mb52Idx < r.length && String(r[idx.mb52Idx]).trim() !== '') ? String(r[idx.mb52Idx]).trim() : '-',
          matMoveStatus: (idx.matMoveIdx >= 0 && idx.matMoveIdx < r.length && String(r[idx.matMoveIdx]).trim() !== '') ? String(r[idx.matMoveIdx]).trim() : '-'
        });
      }
    }
    return { ok: true, sheetName: sh1.getName(), rows: matched };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function getWbsWithPendingMaterials(wbsList, clientSh1Map) {
  try {
    if (!Array.isArray(wbsList) || wbsList.length === 0) {
      return { ok: true, pendingWbsList: [], checkedCount: 0 };
    }
    const ss = getTargetSpreadsheet_();
    const sh1 = ss.getSheetByName(MATERIAL_SHEET_NAME);
    if (!sh1) throw new Error('ไม่พบแผ่นงาน "' + MATERIAL_SHEET_NAME + '"');
    const lastRow = sh1.getLastRow(), lastCol = sh1.getLastColumn();
    if (lastRow < SH1_HEADER_ROW + 1 || lastCol < 1) return { ok: true, pendingWbsList: [], checkedCount: 0 };

    const targetSet = {};
    let targetCount = 0;
    for (let k = 0; k < wbsList.length; k++) {
      const code = String(wbsList[k] || '').trim().toUpperCase();
      if (code && !targetSet[code]) {
        targetSet[code] = true;
        targetCount++;
      }
    }
    if (targetCount === 0) return { ok: true, pendingWbsList: [], checkedCount: 0 };

    const allValues = sh1.getRange(SH1_HEADER_ROW, 1, lastRow - SH1_HEADER_ROW + 1, lastCol).getDisplayValues();
    const headers = allValues[0], rows = allValues.slice(1);
    const idx = resolveSh1Map_(headers, clientSh1Map);
    if (idx.wbsIdx < 0) throw new Error('ไม่พบคอลัมน์ WBS ในชีต ' + MATERIAL_SHEET_NAME);

    const pendingMap = {};
    let pendingCount = 0;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const wbsUpper = (idx.wbsIdx < r.length) ? String(r[idx.wbsIdx] || '').trim().toUpperCase() : '';
      if (!wbsUpper || !targetSet[wbsUpper] || pendingMap[wbsUpper]) continue;

      const moveStr = (idx.moveIdx >= 0 && idx.moveIdx < r.length) ? String(r[idx.moveIdx] || '').trim() : '';
      const valV = (idx.colVIdx >= 0 && idx.colVIdx < r.length) ? parseNum_(r[idx.colVIdx]) : 0;
      const matCode = (idx.matCodeIdx >= 0 && idx.matCodeIdx < r.length && String(r[idx.matCodeIdx]).trim() !== '') ? String(r[idx.matCodeIdx]).trim() : '-';
      const matText = (idx.matIdx >= 0 && idx.matIdx < r.length && String(r[idx.matIdx]).trim() !== '') ? String(r[idx.matIdx]).trim() : '-';

      if (isValidPendingMaterialRow_(valV, moveStr, matCode, matText)) {
        pendingMap[wbsUpper] = true;
        pendingCount++;
        if (pendingCount >= targetCount) break;
      }
    }
    return { ok: true, pendingWbsList: Object.keys(pendingMap), checkedCount: targetCount };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function saveDashboardConfig(configJson, pin) {
  try {
    if (!isOwnerSession_() && String(pin || '').trim() !== String(ADMIN_PIN)) {
      return { ok: false, error: 'ไม่มีสิทธิ์บันทึกการตั้งค่า (เฉพาะเจ้าของระบบเท่านั้น)' };
    }
    PropertiesService.getScriptProperties().setProperty('WBS_DASHBOARD_CONFIG', configJson);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}
