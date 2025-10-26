const VEHICLE_IN_USE_CACHE_KEY = 'vehicle_in_use_payload_v3';
const VEHICLE_IN_USE_PROP_KEY = 'vehicle_in_use_payload_v3_json';
const VEHICLE_RELEASED_CACHE_KEY = 'vehicle_released_dropdown_payload_v1';
const VEHICLE_RELEASED_PROP_KEY = 'vehicle_released_dropdown_payload_v1_json';
const VEHICLE_RELEASED_VERSION_PROP_KEY = 'vehicle_released_dropdown_payload_v1_version';
const VEHICLE_RELEASED_CACHE_TTL_SECONDS = 120;
const VEHICLE_SUMMARY_HEADER = [
  'Ref','Date and time of entry','Project','Team','R.Beneficiary','Vehicle Number',
  'Make','Model','Category','Usage Type','Owner','Status','Last Users remarks','Ratings','Submitter username','R.Ben Time','R. Ben'
];

function upsertVehicleSummaryRow(sheetName, rowData, keyType) {
  try {
    if (!rowData || typeof rowData !== 'object') return;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
    }
    const currentHeader = sh.getRange(1, 1, 1, VEHICLE_SUMMARY_HEADER.length).getValues()[0];
    const needsHeaderUpdate = VEHICLE_SUMMARY_HEADER.some(function(expected, idx){
      return String(currentHeader[idx] || '').trim() !== expected;
    });
    if (needsHeaderUpdate) {
      sh.getRange(1, 1, 1, VEHICLE_SUMMARY_HEADER.length).setValues([VEHICLE_SUMMARY_HEADER]);
      sh.setFrozenRows(1);
    }

    const rawKey = keyType === 'vehicle'
      ? String(
          rowData['Vehicle Number'] ||
          rowData.carNumber ||
          ''
        ).trim().toUpperCase()
      : String(
          rowData['R.Beneficiary'] ||
          rowData['R. Ben'] ||
          rowData.responsibleBeneficiary ||
          ''
        ).trim().toLowerCase();

    if (!rawKey) return;

    const keyColumnName = keyType === 'vehicle' ? 'Vehicle Number' : 'R.Beneficiary';
    const keyColumnIndex = VEHICLE_SUMMARY_HEADER.indexOf(keyColumnName);
    if (keyColumnIndex < 0) return;

    const lastRow = sh.getLastRow();
    let targetRow = -1;
    if (lastRow > 1) {
      const keyRange = sh.getRange(2, keyColumnIndex + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keyRange.length; i++) {
        const cellValue = String(keyRange[i][0] || '').trim();
        if (!cellValue) continue;
        const cellKey = keyType === 'vehicle' ? cellValue.toUpperCase() : cellValue.toLowerCase();
        if (cellKey === rawKey) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const rowValues = [
      rowData.Ref || '',
      rowData['Date and time of entry'] || rowData.Timestamp || '',
      rowData.Project || '',
      rowData.Team || '',
      rowData['R.Beneficiary'] || rowData['R. Ben'] || rowData.responsibleBeneficiary || '',
      rowData['Vehicle Number'] || rowData.carNumber || '',
      rowData.Make || rowData.make || '',
      rowData.Model || rowData.model || '',
      rowData.Category || rowData.category || '',
      rowData['Usage Type'] || rowData.usageType || '',
      rowData.Owner || rowData.owner || '',
      rowData.Status || rowData.status || '',
      rowData['Last Users remarks'] || rowData.remarks || '',
      rowData.Ratings || rowData.rating || rowData.stars || '',
      rowData['Submitter username'] || rowData.submitter || rowData.Submitter || '',
      rowData['R.Ben Time'] || rowData.rBenTime || rowData.responsibleBeneficiaryTime || rowData.responsibleTime || '',
      rowData['R. Ben'] || rowData.rBenShort || rowData.responsibleBeneficiary || ''
    ];

    if (targetRow > 0) {
      sh.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sh.appendRow(rowValues);
    }
  } catch (err) {
    console.error('upsertVehicleSummaryRow failed:', err, { sheetName, keyType });
  }
}

function invalidateVehicleInUseCache() {
  try {
    CacheService.getScriptCache().remove(VEHICLE_IN_USE_CACHE_KEY);
  } catch (_cacheErr) {
    // ignore cache purge failures
  }
  try {
    PropertiesService.getScriptProperties().deleteProperty(VEHICLE_IN_USE_PROP_KEY);
  } catch (_propErr) {
    // ignore property purge failures
  }
}

function invalidateVehicleReleasedCache(reason) {
  try {
    CacheService.getScriptCache().remove(VEHICLE_RELEASED_CACHE_KEY);
  } catch (_cacheErr) {
    // ignore cache purge failures
  }
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(VEHICLE_RELEASED_PROP_KEY);
    props.deleteProperty(VEHICLE_RELEASED_VERSION_PROP_KEY);
  } catch (_propErr) {
    // ignore property purge failures
  }
  if (reason) {
    try {
      console.log('[CACHE] Vehicle_Released cache invalidated:', reason);
    } catch (_logErr) {
      // logging optional
    }
  }
}

function invalidateVehicleCache(reason) {
  try {
    CacheService.getScriptCache().remove(VEHICLE_CACHE_KEY);
  } catch (_cacheErr) {
    // ignore cache purge failures
  }
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(VEHICLE_PROP_KEY);
    props.deleteProperty(VEHICLE_VERSION_PROP_KEY);
  } catch (_propErr) {
    // ignore property purge failures
  }
  if (reason) {
    try {
      console.log('[CACHE] Vehicle cache invalidated:', reason);
    } catch (_logErr) {
      // logging optional
    }
  }
}

function invalidateVehicleSheetCache(sheetName, reason) {
  try {
    const cacheKey = 'VEHICLE_' + sheetName.toUpperCase() + '_CACHE_KEY';
    const propKey = 'VEHICLE_' + sheetName.toUpperCase() + '_PROP_KEY';
    const versionPropKey = 'VEHICLE_' + sheetName.toUpperCase() + '_VERSION_PROP_KEY';

    CacheService.getScriptCache().remove(cacheKey);
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty(propKey);
    props.deleteProperty(versionPropKey);

    if (reason) {
      console.log('[CACHE] Vehicle sheet cache invalidated for', sheetName + ':', reason);
    }
  } catch (err) {
    console.error('invalidateVehicleSheetCache failed:', err);
  }
}

/**
 * Gets vehicle assignment data from Vehicle_InUse tab
 * Returns latest IN USE vehicle mapped per beneficiary so UI can prefill rows.
 */
function getVehicleInUseData() {
  try {
    const nowMs = Date.now();
    const cacheFreshMsWithAssignments = 5 * 60 * 1000; // 5 minutes
    const cacheFreshMsEmpty = 60 * 1000;               // 1 minute for empty payloads

    function shouldUseVehicleInUseSnapshot(payload) {
      if (!payload || payload.ok === false) return false;
      const assignments = Array.isArray(payload.assignments) ? payload.assignments : [];
      const tsCandidate = payload.checkedAt || payload.generatedAt || payload.updatedAt;
      const tsValue = _parseTs_(tsCandidate);
      if (!tsValue) return false;
      const ageMs = nowMs - tsValue;
      if (isNaN(ageMs) || ageMs < 0) return false;
      const freshnessLimit = assignments.length ? cacheFreshMsWithAssignments : cacheFreshMsEmpty;
      return ageMs <= freshnessLimit;
    }

    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    let props = null;
    try {
      props = PropertiesService.getScriptProperties();
    } catch (_propErr) {
      props = null;
    }
    if (cache) {
      const cached = cache.get(VEHICLE_IN_USE_CACHE_KEY);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          if (parsed && parsed.ok && shouldUseVehicleInUseSnapshot(parsed)) {
            parsed.cached = true;
            const referenceTs = parsed.checkedAt || parsed.generatedAt || parsed.updatedAt;
            const refValue = _parseTs_(referenceTs);
            if (refValue) parsed.cacheAgeMs = nowMs - refValue;
            return parsed;
          }
        } catch (_err) {
          cache.remove(VEHICLE_IN_USE_CACHE_KEY);
        }
      }
    }

    _maybeAutoRefreshCarTPSummaries_(5);

    if (props) {
      const stored = props.getProperty(VEHICLE_IN_USE_PROP_KEY);
      if (stored) {
        try {
          const parsed = JSON.parse(stored);
          if (parsed && parsed.ok && shouldUseVehicleInUseSnapshot(parsed)) {
            parsed.cached = true;
            parsed.fromProperties = true;
            const referenceTs = parsed.checkedAt || parsed.generatedAt || parsed.updatedAt;
            const refValue = _parseTs_(referenceTs);
            if (refValue) parsed.cacheAgeMs = nowMs - refValue;
            return parsed;
          }
        } catch (_propParseErr) {
          try { props.deleteProperty(VEHICLE_IN_USE_PROP_KEY); } catch (_){ /* ignore */ }
        }
      }
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Vehicle_InUse');
    let refreshed = false;

    if (!sheet) {
      console.warn('[BACKEND] Vehicle_InUse sheet missing, triggering refresh');
      try { refreshVehicleStatusSheets(); refreshed = true; } catch (refreshErr) {
        console.error('[BACKEND] refreshVehicleStatusSheets() failed:', refreshErr);
      }
      sheet = ss.getSheetByName('Vehicle_InUse');
    }

    if (!sheet) {
      return { ok: false, source: 'Vehicle_InUse', error: 'Vehicle_InUse sheet not found' };
    }

    let lastRow = sheet.getLastRow();
    if (lastRow <= 1 && !refreshed) {
      console.log('[BACKEND] Vehicle_InUse empty, attempting refresh');
      try { refreshVehicleStatusSheets(); refreshed = true; } catch (refreshErr) {
        console.error('[BACKEND] refreshVehicleStatusSheets() failed on empty sheet:', refreshErr);
      }
      lastRow = sheet.getLastRow();
    }

    if (lastRow <= 1) {
      const emptyTs = new Date().toISOString();
      const emptyPayload = {
        ok: true,
        source: 'Vehicle_InUse',
        assignments: [],
        updatedAt: '',
        generatedAt: emptyTs,
        checkedAt: emptyTs,
        message: 'No assignments found'
      };
      if (cache) cache.put(VEHICLE_IN_USE_CACHE_KEY, JSON.stringify(emptyPayload), 15);
      if (props) {
        try { props.setProperty(VEHICLE_IN_USE_PROP_KEY, JSON.stringify(emptyPayload)); } catch (_err) { /* ignore */ }
      }
      return emptyPayload;
    }

    const summaryData = getVehicleInUseSummary();
    if (!summaryData.ok) {
      const message = summaryData.error || 'Vehicle_InUse summary unavailable';
      return { ok: false, source: 'Vehicle_InUse', error: message };
    }

    const assignments = (summaryData.assignments || []).map(function(entry, idx) {
      const beneficiary = String(entry.beneficiary || entry.responsibleBeneficiary || '').trim();
      const vehicleNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
      if (!beneficiary || !vehicleNumber) return null;
      const timestampValue = entry.latestTimestamp || entry.entryDate || '';
      const ts = _parseTs_(timestampValue);
      return {
        beneficiary: beneficiary,
        responsibleBeneficiary: beneficiary,
        vehicleNumber: vehicleNumber,
        project: entry.project || '',
        team: entry.team || '',
        status: entry.assignmentStatus || entry.status || '',
        entryDate: timestampValue || '',
        entryTimestamp: ts,
        rowNumber: entry.rowNumber || (idx + 2),
        sheet: 'Vehicle_InUse',
        ref: entry.ref || '',
        make: entry.make || '',
        model: entry.model || '',
        category: entry.category || '',
        usageType: entry.usageType || '',
        owner: entry.owner || '',
        remarks: entry.remarks || '',
        ratings: entry.stars || entry.ratings || '',
        submitter: entry.submitter || ''
      };
    }).filter(Boolean);

    assignments.sort(function(a, b) {
      const aTs = typeof a.entryTimestamp === 'number' ? a.entryTimestamp : 0;
      const bTs = typeof b.entryTimestamp === 'number' ? b.entryTimestamp : 0;
      if (aTs !== bTs) return bTs - aTs;
      return (b.rowNumber || 0) - (a.rowNumber || 0);
    });

    const newestTs = assignments.length ? assignments.reduce(function(max, item){
      return item.entryTimestamp && item.entryTimestamp > max ? item.entryTimestamp : max;
    }, 0) : 0;

    const updatedAt = newestTs > 0 ? new Date(newestTs).toISOString() : (summaryData.updatedAt || new Date().toISOString());
    const generatedIso = new Date().toISOString();
    const payload = {
      ok: true,
      source: 'Vehicle_InUse',
      assignments: assignments,
      updatedAt: updatedAt,
      generatedAt: generatedIso,
      checkedAt: generatedIso
    };

    if (cache) {
      try { cache.put(VEHICLE_IN_USE_CACHE_KEY, JSON.stringify(payload), 15); } catch (_err) { /* ignore */ }
    }
    if (props) {
      try { props.setProperty(VEHICLE_IN_USE_PROP_KEY, JSON.stringify(payload)); } catch (_err) { /* ignore */ }
    }

    return payload;
  } catch (error) {
    console.error('[BACKEND] getVehicleInUseData failed:', error);
    return { ok: false, source: 'Vehicle_InUse', error: String(error) };
  }
}

// Read summary sheets that mirror Vehicle_InUse or Vehicle_Released snapshots.
function getVehicleSummaryRows(sheetName) {
  const tried = [];
  const notes = [];
  const errors = [];
  const candidates = [];

  try {
    if (typeof SHEET_ID === 'string' && SHEET_ID) {
      candidates.push({ id: SHEET_ID, label: 'SHEET_ID' });
    }
  } catch (_){ /* ignore */ }

  try {
    if (typeof CAR_SHEET_ID === 'string' && CAR_SHEET_ID && CAR_SHEET_ID !== SHEET_ID) {
      candidates.push({ id: CAR_SHEET_ID, label: 'CAR_SHEET_ID' });
    }
  } catch (_){ /* ignore */ }

  if (!candidates.length) {
    console.warn(`[BACKEND] getVehicleSummaryRows(${sheetName}) has no sheet id candidates.`);
    return { rows: [], updatedAt: '', rowsFetched: 0, error: 'No spreadsheet IDs available for summary lookup' };
  }

  for (var c = 0; c < candidates.length; c++) {
    var candidate = candidates[c];
    tried.push(candidate.label + ':' + candidate.id);
    try {
      const ss = SpreadsheetApp.openById(candidate.id);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        notes.push(`${candidate.label}:${candidate.id} missing sheet ${sheetName}`);
        continue;
      }

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) {
        notes.push(`${candidate.label}:${candidate.id} has no data in ${sheetName}`);
        continue;
      }

      const range = sheet.getRange(1, 1, lastRow, lastCol);
      const values = range.getValues();
      const display = range.getDisplayValues();
      const headers = display[0] && display[0].some(Boolean) ? display[0] : values[0];
      const IX = _headerIndex_(headers);

      const rows = [];
      for (let r = 1; r < values.length; r++) {
        const valueRow = values[r];
        const displayRow = display[r];
        let hasData = false;
        const merged = valueRow.map(function(cell, idx){
          let val = cell;
          if (val === '' || val === null) val = displayRow[idx];
          if (!hasData && val != null && String(val).trim() !== '') hasData = true;
          return val;
        });
        if (hasData) rows.push(merged);
      }

      if (!rows.length) {
        notes.push(`${candidate.label}:${candidate.id} only has empty rows in ${sheetName}`);
        continue;
      }

      let latestTs = 0;
      let tsIndex = -1;
      try {
        tsIndex = IX.get(['Date and time of entry', 'Date and time', 'Timestamp', 'Latest Timestamp']);
      } catch (_err) {
        tsIndex = -1;
      }
      if (tsIndex >= 0) {
        rows.forEach(function(row) {
          const ts = _parseTs_(row[tsIndex]);
          if (ts > latestTs) latestTs = ts;
        });
      }

      const updatedAt = latestTs > 0 ? new Date(latestTs).toISOString() : new Date().toISOString();

      return {
        rows: rows,
        headerIndex: IX,
        headerRow: headers,
        updatedAt: updatedAt,
        rowsFetched: rows.length,
        sheetId: candidate.id,
        sheetLabel: candidate.label,
        notes: notes.length ? notes.slice() : undefined,
        tried: tried.slice()
      };
    } catch (error) {
      console.error(`[BACKEND] Failed summary lookup on ${candidate.label}:${candidate.id} for ${sheetName}:`, error);
      errors.push(`${candidate.label}:${candidate.id}: ${error}`);
    }
  }

  const result = { rows: [], updatedAt: '', rowsFetched: 0 };
  if (notes.length) result.notes = notes.slice();
  if (errors.length) result.error = errors.join('; ');
  if (tried.length) result.tried = tried.slice();
  return result;
}

function getVehicleInUseSummary() {
  const summary = getVehicleSummaryRows('Vehicle_InUse');
  if (summary.error) {
    return { ok: false, source: 'Vehicle_InUse', assignments: [], updatedAt: '', error: summary.error };
  }

  if (!summary.rows.length) {
    if (summary.notes || summary.tried) {
      console.warn('[BACKEND] Vehicle_InUse summary empty', { notes: summary.notes, tried: summary.tried });
    }
    return { ok: true, source: 'Vehicle_InUse', assignments: [], updatedAt: summary.updatedAt || '', message: 'Vehicle_InUse summary empty' };
  }

  const IX = summary.headerIndex;
  const idx = function(labels) {
    if (!IX) return -1;
    try { return IX.get(labels); } catch (_err) { return -1; }
  };
  const beneficiaryIdx = idx(['R.Beneficiary', 'Beneficiary', 'Responsible Beneficiary','R. Ben','R Ben']);
  let vehicleIdx = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car', 'Vehicle']);
  if (vehicleIdx < 0 && Array.isArray(summary.headerRow)) {
    vehicleIdx = _findCarNumberColumn_(summary.headerRow, summary.rows);
    if (vehicleIdx >= 0) {
      try {
        console.log('[BACKEND] Vehicle_Released header fallback matched car column at index', vehicleIdx, {
          header: String(summary.headerRow[vehicleIdx] || '')
        });
      } catch (_logErr) {
        // logging optional
      }
    }
  }
  if (vehicleIdx < 0) {
    try {
      console.error('[BACKEND] Vehicle_Released summary missing vehicle column', {
        headers: Array.isArray(summary.headerRow) ? summary.headerRow : null,
        sheetId: summary.sheetId || null,
        sheetLabel: summary.sheetLabel || null
      });
    } catch (_logErr) {
      // best-effort logging
    }
  }
  const assignmentIdx = idx(['Status', 'Assignment Status']);
  const tsIdx = idx(['Date and time of entry', 'Timestamp', 'Latest Timestamp']);
  const projectIdx = idx(['Project', 'Project Name', 'ProjectName']);
  const teamIdx = idx(['Team', 'Team Name', 'TeamName']);
  const remarksIdx = idx(['Last Users remarks', 'Remarks']);
  const ownerIdx = idx(['Owner']);
  const categoryIdx = idx(['Category']);
  const usageIdx = idx(['Usage Type']);
  const rowNumberIdx = idx(['RowNumber', 'Row Number']);
  const makeIdx = idx(['Make']);
  const modelIdx = idx(['Model']);
  const starsIdx = idx(['Ratings', 'Stars', 'Rating']);
  const rBenTimeIdx = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);

  const assignments = summary.rows.map(function(row) {
    return {
      beneficiary: beneficiaryIdx >= 0 ? row[beneficiaryIdx] : '',
      vehicleNumber: vehicleIdx >= 0 ? row[vehicleIdx] : '',
      assignmentStatus: assignmentIdx >= 0 ? row[assignmentIdx] : '',
      status: assignmentIdx >= 0 ? row[assignmentIdx] : '',
      latestTimestamp: tsIdx >= 0 ? row[tsIdx] : '',
      project: projectIdx >= 0 ? row[projectIdx] : '',
      team: teamIdx >= 0 ? row[teamIdx] : '',
      remarks: remarksIdx >= 0 ? row[remarksIdx] : '',
      owner: ownerIdx >= 0 ? row[ownerIdx] : '',
      category: categoryIdx >= 0 ? row[categoryIdx] : '',
      usageType: usageIdx >= 0 ? row[usageIdx] : '',
      make: makeIdx >= 0 ? row[makeIdx] : '',
      model: modelIdx >= 0 ? row[modelIdx] : '',
      stars: starsIdx >= 0 ? Number(row[starsIdx] || 0) || 0 : 0,
      rowNumber: rowNumberIdx >= 0 ? row[rowNumberIdx] : '',
      rBenTime: rBenTimeIdx >= 0 ? row[rBenTimeIdx] : ''
    };
  }).filter(function(entry) {
    return String(entry.vehicleNumber || '').trim() !== '';
  });

  assignments.sort(function(a, b) {
    var tsA = a.latestTimestamp ? new Date(a.latestTimestamp).getTime() : 0;
    var tsB = b.latestTimestamp ? new Date(b.latestTimestamp).getTime() : 0;
    if (tsA !== tsB) return tsB - tsA;
    return String(a.beneficiary || '').localeCompare(String(b.beneficiary || ''));
  });

  return {
    ok: true,
    source: 'Vehicle_InUse',
    assignments: assignments,
    updatedAt: summary.updatedAt || new Date().toISOString()
  };
}

function getVehicleReleasedSummary() {
  const summary = getVehicleSummaryRows('Vehicle_Released');
  if (summary.error) {
    return { ok: false, source: 'Vehicle_Released', vehicles: [], updatedAt: '', error: summary.error };
  }

  if (!summary.rows.length) {
    if (summary.notes || summary.tried) {
      console.warn('[BACKEND] Vehicle_Released summary empty', { notes: summary.notes, tried: summary.tried });
    }
    return { ok: true, source: 'Vehicle_Released', vehicles: [], updatedAt: summary.updatedAt || '', message: 'Vehicle_Released summary empty' };
  }

  const IX = summary.headerIndex;
  const idx = function(labels) {
    if (!IX) return -1;
    try { return IX.get(labels); } catch (_err) { return -1; }
  };
  const vehicleIdx = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car', 'Vehicle']);
  const releaseTsIdx = idx(['Date and time of entry', 'Timestamp', 'Latest Timestamp']);
  const statusIdx = idx(['Status', 'Release Status']);
  const projectIdx = idx(['Project', 'Project Name', 'ProjectName']);
  const teamIdx = idx(['Team', 'Team Name', 'TeamName']);
  const remarksIdx = idx(['Last Users remarks', 'Remarks']);
  const ownerIdx = idx(['Owner']);
  const categoryIdx = idx(['Category']);
  const usageIdx = idx(['Usage Type']);
  const makeIdx = idx(['Make']);
  const modelIdx = idx(['Model']);
  const beneficiaryIdx = idx(['R.Beneficiary', 'Responsible Beneficiary', 'R. Ben', 'R Ben']);
  const starsIdx = idx(['Ratings', 'Stars', 'Rating']);
  const rBenTimeIdx = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);

  var vehicles = summary.rows.map(function(row) {
    return {
      vehicleNumber: vehicleIdx >= 0 ? row[vehicleIdx] : '',
      latestRelease: releaseTsIdx >= 0 ? row[releaseTsIdx] : '',
      status: statusIdx >= 0 ? row[statusIdx] : 'Released',
      project: projectIdx >= 0 ? row[projectIdx] : '',
      team: teamIdx >= 0 ? row[teamIdx] : '',
      remarks: remarksIdx >= 0 ? row[remarksIdx] : '',
      owner: ownerIdx >= 0 ? row[ownerIdx] : '',
      category: categoryIdx >= 0 ? row[categoryIdx] : '',
      usageType: usageIdx >= 0 ? row[usageIdx] : '',
      make: makeIdx >= 0 ? row[makeIdx] : '',
      model: modelIdx >= 0 ? row[modelIdx] : '',
      responsibleBeneficiary: beneficiaryIdx >= 0 ? row[beneficiaryIdx] : '',
      'R. Ben': beneficiaryIdx >= 0 ? row[beneficiaryIdx] : '',
      stars: starsIdx >= 0 ? Number(row[starsIdx] || 0) || 0 : 0,
      rBenTime: rBenTimeIdx >= 0 ? row[rBenTimeIdx] : ''
    };
  }).filter(function(entry) {
    return String(entry.vehicleNumber || '').trim() !== '';
  });

  vehicles.sort(function(a, b) {
    var tsA = a.latestRelease ? new Date(a.latestRelease).getTime() : 0;
    var tsB = b.latestRelease ? new Date(b.latestRelease).getTime() : 0;
    if (tsA !== tsB) return tsB - tsA;
    return String(a.vehicleNumber || '').localeCompare(String(b.vehicleNumber || ''));
  });

  return {
    ok: true,
    source: 'Vehicle_Released',
    vehicles: vehicles,
    updatedAt: summary.updatedAt || new Date().toISOString()
  };
}

/**
 * Simple test function to check CarT_P sheet access
 */
function testCarTPAccess() {
  try {
    console.log('Testing CarT_P sheet access...');
    const sh = _openCarTP_();
    if (!sh) {
      return { ok: false, error: 'CarT_P sheet not found' };
    }
    
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    
    if (lastRow <= 1) {
      return { ok: false, error: 'CarT_P sheet has no data', lastRow, lastCol };
    }
    
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const sampleRow = lastRow > 1 ? sh.getRange(2, 1, 1, lastCol).getValues()[0] : null;
    
    return { 
      ok: true, 
      sheetName: sh.getName(),
      lastRow, 
      lastCol, 
      header: header,
      sampleRow: sampleRow
    };
    
  } catch (error) {
    return { ok: false, error: error.toString() };
  }
}

/**
 * SIMPLE test functions that bypass _openCarTP_ and test directly
 */
function testDirectReleaseHistory() {
  try {
    console.log('testDirectReleaseHistory: Testing with direct return...');
    return [['Test Message', 'This is a test release history response']];
  } catch (error) {
    console.error('testDirectReleaseHistory error:', error);
    return [['Error in testDirectReleaseHistory', error.toString()]];
  }
}

/**
 * Test function specifically for IN USE history - call this directly from script editor
 */
function testInUseHistoryFunction() {
  console.log('=== Testing IN USE History Function ===');
  
  try {
    const testCarNumber = 'TEST-VEHICLE-GENERIC';
    console.log('Testing with car number:', testCarNumber);
    
    const result = getInUseHistoryForcedWorking(testCarNumber, '');
    console.log('Function returned:', result);
    console.log('Result type:', typeof result);
    console.log('Is array?:', Array.isArray(result));
    
    if (Array.isArray(result)) {
      console.log('Array length:', result.length);
      result.forEach((row, index) => {
        console.log(`Row ${index}:`, row);
      });
    } else {
      console.error('ERROR: Function did not return an array!');
    }
    
    return result;
  } catch (error) {
    console.error('Test failed with error:', error);
    return [['Test Error', error.toString()]];
  }
}

/**
 * ULTIMATE TEST: This function is absolutely bulletproof and cannot return null
 */
function testSimpleReturn() {
  return [['Test', 'Working'], ['Status', 'SUCCESS']];
}

/**
 * DEBUGGING: Check if the issue is with our main functions
 */
function debugGetInUseHistory(carNumber) {
  console.log('debugGetInUseHistory called with:', carNumber);
  
  // Test 1: Can we return a simple array?
  try {
    const simpleResult = [['DEBUG', 'STEP1'], ['Simple return', 'works']];
    console.log('Step 1 - Simple return test passed');
    
    // Test 2: Can we access CAR_SHEET_ID?
    const sheetId = CAR_SHEET_ID;
    console.log('Step 2 - CAR_SHEET_ID:', sheetId);
    
    // Test 3: Can we call _openCarTP_?
    const sheet = _openCarTP_();
    console.log('Step 3 - _openCarTP_ result:', sheet ? 'Found sheet' : 'No sheet');
    
    if (!sheet) {
      return [['DEBUG', 'SHEET_ACCESS'], ['Error', 'Cannot access CarT_P sheet'], ['SheetID', sheetId || 'UNDEFINED']];
    }
    
    // Test 4: Can we get basic sheet info?
    const lastRow = sheet.getLastRow();
    const sheetName = sheet.getName();
    console.log('Step 4 - Sheet info - Name:', sheetName, 'LastRow:', lastRow);
    
    return [
      ['DEBUG', 'SHEET_INFO'],
      ['SheetName', sheetName],
      ['LastRow', lastRow.toString()],
      ['CarNumber', carNumber || 'NO_CAR'],
      ['Timestamp', new Date().toISOString()]
    ];
    
  } catch (error) {
    console.error('debugGetInUseHistory error:', error);
    return [['DEBUG', 'ERROR'], ['Message', error.toString()], ['CarNumber', carNumber || 'NO_CAR']];
  }
}

/**
 * PRODUCTION: Real IN USE history from CarT_P sheet with bulletproof error handling
 */
function getInUseHistoryForcedWorking(carNumber, teamName) {
  try {
    console.log('getInUseHistoryForcedWorking called with car:', carNumber, 'team:', teamName);

    const normalizedCar = String(carNumber || '').trim();
    const normalizedTeam = String(teamName || '').trim().toLowerCase();

    if (!normalizedCar) {
      console.log('No car number provided');
      return [['Error', 'No car number provided']];
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Vehicle_InUse');
    if (!sh) {
      console.log('Vehicle_InUse sheet not found');
      return [['Info', 'Vehicle_InUse sheet not found. Please refresh vehicle status summaries.']];
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol < 12) {
      console.log('Vehicle_InUse sheet has insufficient data');
      return [['Info', 'Vehicle_InUse sheet has no history data to display']];
    }

    const columnCount = Math.min(11, lastCol - 1); // columns B-L inclusive (11 columns)
    if (columnCount <= 0) {
      return [['Info', 'Vehicle_InUse sheet missing required columns']];
    }

    const range = sh.getRange(1, 2, lastRow, columnCount);
    const values = range.getDisplayValues();
    const rawValues = range.getValues();
    const header = values[0];
    const dataRows = values.slice(1);
    const rawRows = rawValues.slice(1);

    const loweredHeader = header.map(function(cell){ return String(cell || '').trim().toLowerCase(); });

    function findIndex(aliases, fallback){
      const list = Array.isArray(aliases) ? aliases : [];
      for (let col = 0; col < loweredHeader.length; col++) {
        const headerCell = loweredHeader[col];
        if (!headerCell) continue;
        for (let i = 0; i < list.length; i++) {
          const needle = String(list[i] || '').trim().toLowerCase();
          if (!needle) continue;
          if (headerCell === needle) return col;
          if (headerCell.indexOf(needle) !== -1) return col;
          if (needle.indexOf(headerCell) !== -1 && headerCell.length > 2) return col;
        }
      }
      const fallbackIndex = typeof fallback === 'number' ? fallback : -1;
      return fallbackIndex >= 0 && fallbackIndex < loweredHeader.length ? fallbackIndex : -1;
    }

    const CAR_INDEX = findIndex(['vehicle number', 'car number', 'vehicle no', 'car no', 'car #', 'car'], 4);
    const TEAM_INDEX = findIndex(['team', 'team name'], 2);
    const STATUS_INDEX = findIndex(['status', 'in use/ release', 'in use', 'in use / release'], 10);
    const DATE_INDEX = findIndex(['date and time of entry', 'date and time', 'timestamp', 'date'], 5);
    const BENEFICIARY_INDEX = findIndex(['r.beneficiary', 'responsible beneficiary', 'name of responsible beneficiary', 'member', 'beneficiary'], -1);

    if (CAR_INDEX < 0 || STATUS_INDEX < 0 || DATE_INDEX < 0) {
      console.log('Required columns missing in Vehicle_InUse history extract');
      return [['Info', 'Vehicle_InUse sheet missing required columns for history view']];
    }

    function pickCell(displayRow, rawRow, index){
      if (!displayRow || index < 0 || index >= displayRow.length) return '';
      if (rawRow && index < rawRow.length) {
        const raw = rawRow[index];
        if (raw instanceof Date) return raw;
        if (raw !== null && raw !== '' && raw !== undefined) return raw;
      }
      return displayRow[index];
    }

    function splitNames(value){
      if (!value && value !== 0) return [];
      if (Array.isArray(value)) {
        return value
          .map(function(entry){ return String(entry || '').trim(); })
          .filter(Boolean);
      }
      const text = String(value || '').trim();
      if (!text) return [];
      return text
        .split(/[,;\n\|]+/)
        .map(function(part){ return String(part || '').trim(); })
        .filter(Boolean);
    }

    const targetCar = normalizedCar.toLowerCase();
    const entries = [];

    for (let i = 0; i < dataRows.length; i++) {
      const displayRow = dataRows[i];
      const rawRow = rawRows[i];
      const carValue = String(pickCell(displayRow, rawRow, CAR_INDEX) || '').trim().toLowerCase();
      if (!carValue || carValue !== targetCar) continue;

      const teamValue = TEAM_INDEX >= 0
        ? String(pickCell(displayRow, rawRow, TEAM_INDEX) || '').trim().toLowerCase()
        : '';
      if (normalizedTeam && teamValue !== normalizedTeam) continue;

      const statusValue = String(pickCell(displayRow, rawRow, STATUS_INDEX) || '').trim();
      if (!statusValue) continue;
      const normalizedStatus = typeof _normStatus_ === 'function'
        ? _normStatus_(statusValue)
        : statusValue.toUpperCase();
      if (normalizedStatus !== 'IN USE') continue;

      const dateValue = pickCell(displayRow, rawRow, DATE_INDEX);
      const timestamp = typeof _parseDateTimeFlexible_ === 'function'
        ? (_parseDateTimeFlexible_(dateValue) || _parseDateTimeFlexible_(displayRow[DATE_INDEX]))
        : new Date(dateValue).getTime();
      const displayDate = displayRow[DATE_INDEX];

      const beneficiarySource = BENEFICIARY_INDEX >= 0 ? pickCell(displayRow, rawRow, BENEFICIARY_INDEX) : '';
      const names = typeof _splitBeneficiaryNames_ === 'function'
        ? _splitBeneficiaryNames_(beneficiarySource)
        : splitNames(beneficiarySource);

      entries.push({
        displayRow: displayRow.slice(),
        rawRow: rawRow,
        timestamp: timestamp || 0,
        displayDate: displayDate,
        names: Array.isArray(names) && names.length ? names : [''],
        rowIndex: i
      });
    }

    console.log(`Filtered ${entries.length} rows for car ${normalizedCar} and team ${normalizedTeam || '(any)'}`);

    if (!entries.length) {
      const teamMessage = normalizedTeam ? ` and team: ${teamName}` : '';
      return [['Info', `No IN USE history found for vehicle: ${normalizedCar}${teamMessage}`]];
    }

    const insertIndex = STATUS_INDEX + 1;
    const headerWithDays = header.slice();
    headerWithDays.splice(insertIndex, 0, 'Days in Use');

    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const msPerDay = 24 * 60 * 60 * 1000;

    function computeDaySpan(ts){
      if (!ts || !isFinite(ts)) return '';
      const date = new Date(ts);
      if (isNaN(date.getTime())) return '';
      const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const diff = todayStart.getTime() - start.getTime();
      if (diff < 0) return '0';
      return String(Math.floor(diff / msPerDay) + 1);
    }

    const latestByBeneficiary = new Map();

    entries.forEach(function(entry){
      const ts = entry.timestamp || 0;
      entry.names.forEach(function(name, idx){
        const trimmed = String(name || '').trim();
        const key = trimmed ? trimmed.toLowerCase() : `__row_${entry.rowIndex}_${idx}`;
        const existing = latestByBeneficiary.get(key);
        if (!existing || ts > (existing.timestamp || 0)) {
          latestByBeneficiary.set(key, {
            name: trimmed,
            timestamp: ts,
            displayRow: entry.displayRow.slice()
          });
        }
      });
    });

    const rowsWithDays = [];
    latestByBeneficiary.forEach(function(entry){
      const row = entry.displayRow.slice();
      if (BENEFICIARY_INDEX >= 0 && entry.name) {
        row[BENEFICIARY_INDEX] = entry.name;
      }
      const daysValue = computeDaySpan(entry.timestamp);
      row.splice(insertIndex, 0, daysValue);
      rowsWithDays.push({ row: row, timestamp: entry.timestamp || 0 });
    });

    rowsWithDays.sort(function(a, b){
      return (b.timestamp || 0) - (a.timestamp || 0);
    });

    const finalRows = rowsWithDays.map(function(entry){ return entry.row; });
    const result = [headerWithDays].concat(finalRows);
    console.log('Returning IN USE history result with', result.length, 'rows');
    return result;

  } catch (error) {
    console.error('Error in getInUseHistoryForcedWorking:', error);
    return [['Error', 'Error retrieving IN USE history: ' + error.toString()]];
  }
}

/**
 * PRODUCTION: Real RELEASE history from CarT_P sheet with bulletproof error handling
 */
function getReleaseHistoryForcedWorking(carNumber) {
  try {
    console.log('getReleaseHistoryForcedWorking called with car:', carNumber);

    const inputNumber = String(carNumber || '').trim();
    const targetKey = _vehicleKey_(inputNumber);
    if (!targetKey) {
      return [['Error', 'Vehicle number required']];
    }

    const sh = _openCarTP_();
    if (!sh) {
      console.log('CarT_P sheet not found - returning info message');
      return [['Info', 'CarT_P sheet not accessible - please check sheet permissions']];
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) {
      console.log('CarT_P sheet has no data');
      return [['Info', 'CarT_P sheet contains no data']];
    }

    const headerDisplay = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    let IX;
    try {
      IX = _headerIndex_(headerDisplay);
    } catch (err) {
      console.error('getReleaseHistoryForcedWorking header error:', err);
      return [['Error', 'Required columns missing in CarT_P: ' + err.message]];
    }

    const idxVehicle = IX.get(['Vehicle Number', 'Car Number', 'Vehicle', 'Vehicle No', 'Car No', 'Car']);
    const idxStatus = IX.get(['Status', 'In Use / release', 'In Use', 'Release Status']);
    const idxDate = IX.get(['Date and time of entry', 'Date and time', 'Timestamp', 'Date']);
  const idxProject = IX.get(['Project']);
  const idxTeam = IX.get(['Team', 'Team Name']);
  const idxBeneficiary = IX.get(['R.Beneficiary', 'Responsible Beneficiary', 'Responsible beneficiary', 'Name of Responsible beneficiary','R. Ben','R Ben']);
  const idxRemarks = IX.get(['Last Users remarks', 'Remarks', 'User Remarks', 'Last 3 Remarks', 'Last Users remark']);
  const idxRatings = IX.get(['Ratings', 'Stars', 'Rating']);

    const valuesRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const sheetValues = valuesRange.getValues();
    const sheetDisplay = valuesRange.getDisplayValues();

    const inUseTimestamps = new Map();
    const releaseEntries = [];

    for (let r = 0; r < sheetValues.length; r++) {
      const rowValues = sheetValues[r];
      const rowDisplay = sheetDisplay[r];

      const rawCarValue = (rowValues[idxVehicle] != null && rowValues[idxVehicle] !== '')
        ? rowValues[idxVehicle]
        : rowDisplay[idxVehicle];
      const rowKey = _vehicleKey_(rawCarValue);
      if (!rowKey || rowKey !== targetKey) {
        continue;
      }

      const rawStatusValue = (rowValues[idxStatus] != null && rowValues[idxStatus] !== '')
        ? rowValues[idxStatus]
        : rowDisplay[idxStatus];
      const status = _normStatus_(rawStatusValue);
      if (!status) {
        continue;
      }

  const rawDateValue = rowValues[idxDate];
  const displayDateValue = rowDisplay[idxDate];
  const timestamp = _parseDateTimeFlexible_(rawDateValue) || _parseDateTimeFlexible_(displayDateValue);

      const rawBeneficiary = rowValues[idxBeneficiary];
      const displayBeneficiary = String(rowDisplay[idxBeneficiary] || rawBeneficiary || '').trim();
      const beneficiaryNames = _splitBeneficiaryNames_(rawBeneficiary || displayBeneficiary);
      const beneficiaryKeys = beneficiaryNames.map(function(name){ return _beneficiaryKey_(name); }).filter(Boolean);

      if (status === 'IN USE') {
        if (beneficiaryKeys.length && timestamp) {
          beneficiaryKeys.forEach(function(key){
            if (!inUseTimestamps.has(key)) inUseTimestamps.set(key, []);
            inUseTimestamps.get(key).push(timestamp);
          });
        }
        continue;
      }

      if (status !== 'RELEASE') {
        continue;
      }

  const projectValue = String(rowDisplay[idxProject] || rowValues[idxProject] || '').trim();
  const teamValue = String(rowDisplay[idxTeam] || rowValues[idxTeam] || '').trim();

      const remarksRaw = idxRemarks >= 0 ? (rowValues[idxRemarks] != null && rowValues[idxRemarks] !== '' ? rowValues[idxRemarks] : rowDisplay[idxRemarks]) : '';
      const ratingsRaw = idxRatings >= 0 ? (rowValues[idxRatings] != null && rowValues[idxRatings] !== '' ? rowValues[idxRatings] : rowDisplay[idxRatings]) : '';

      releaseEntries.push({
        ts: timestamp,
        rawDate: rawDateValue,
  displayDate: displayDateValue || (timestamp ? Utilities.formatDate(new Date(timestamp), TZ(), 'dd-MMM-yyyy HH:mm') : ''),
        project: projectValue,
        team: teamValue,
        beneficiaryDisplay: displayBeneficiary,
        beneficiaryKeys: beneficiaryKeys,
        status: status,
        remarks: String(remarksRaw || '').trim(),
        ratings: String(ratingsRaw || '').trim()
      });
    }

    if (!releaseEntries.length) {
      return [['Info', `No RELEASE history found for vehicle: ${inputNumber || carNumber}`]];
    }

    inUseTimestamps.forEach(function(list, key){
      list.sort(function(a, b){ return (a || 0) - (b || 0); });
      inUseTimestamps.set(key, list);
    });

    releaseEntries.sort(function(a, b){
      const tsA = a.ts || 0;
      const tsB = b.ts || 0;
      if (tsA === tsB) {
        return 0;
      }
      return tsB - tsA;
    });

    const limit = 10;
    const limitedEntries = releaseEntries.slice(0, limit);

    function formatDateValue(entry) {
      if (entry.displayDate) {
        return entry.displayDate;
      }
      if (entry.ts) {
        try {
          return Utilities.formatDate(new Date(entry.ts), TZ(), 'dd-MMM-yyyy HH:mm');
        } catch (_formatErr) {
          /* ignore */
        }
      }
      if (entry.rawDate instanceof Date) {
        try {
          return Utilities.formatDate(entry.rawDate, TZ(), 'dd-MMM-yyyy HH:mm');
        } catch (_dateErr) {
          /* ignore */
        }
      }
      return '';
    }

    function computeDaysUsed(entry) {
      if (!entry || !entry.beneficiaryKeys || !entry.beneficiaryKeys.length) {
        return '0';
      }
      const releaseTs = entry.ts;
      if (!releaseTs) {
        return '0';
      }

      const dayMs = 24 * 60 * 60 * 1000;
      let bestDiff = null;

      entry.beneficiaryKeys.forEach(function(key){
        const tsList = inUseTimestamps.get(key);
        if (!tsList || !tsList.length) return;
        for (let i = tsList.length - 1; i >= 0; i--) {
          const candidateTs = tsList[i];
          if (!candidateTs || candidateTs > releaseTs) {
            continue;
          }
          const diff = releaseTs - candidateTs;
          if (diff < 0) {
            continue;
          }
          if (bestDiff === null || diff < bestDiff) {
            bestDiff = diff;
          }
          break; // nearest previous entry found
        }
      });

      if (bestDiff === null) {
        return '0';
      }

      const diffDays = bestDiff / dayMs;
      if (!isFinite(diffDays) || diffDays < 0) {
        return '0';
      }
      if (diffDays < 1) {
        return diffDays.toFixed(2);
      }
      return String(Math.round(diffDays));
    }

    const header = ['Date and time of entry', 'Project', 'Team', 'R.Beneficiary', 'Status', 'Days used', 'Ratings', 'Last users remarks'];
    const rows = limitedEntries.map(function(entry){
      return [
        formatDateValue(entry) || '',
        entry.project || '',
        entry.team || '',
        entry.beneficiaryDisplay || '',
        entry.status || 'RELEASE',
        computeDaysUsed(entry),
        entry.ratings || '',
        entry.remarks || ''
      ];
    });

    console.log('Returning computed RELEASE history rows:', rows.length);
    return [header].concat(rows);
  } catch (error) {
    console.error('Error in getReleaseHistoryForcedWorking:', error);
    return [['Error', 'Error retrieving RELEASE history: ' + error.toString()]];
  }
}

/**
 * Test the CarT_P sheet connection specifically
 */
function testCarTPConnection() {
  try {
    console.log('testCarTPConnection called');
    console.log('CAR_SHEET_ID:', CAR_SHEET_ID);
    
    const sh = _openCarTP_();
    if (!sh) {
      return {
        status: 'error',
        message: 'CarT_P sheet not found',
        sheetId: CAR_SHEET_ID,
        availableSheets: []
      };
    }
    
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    
    return {
      status: 'success',
      message: 'CarT_P sheet found',
      sheetName: sh.getName(),
      sheetId: CAR_SHEET_ID,
      lastRow: lastRow,
      lastCol: lastCol,
      hasData: lastRow > 1
    };
  } catch (error) {
    return {
      status: 'error',
      message: 'Error testing CarT_P connection: ' + error.toString(),
      sheetId: CAR_SHEET_ID
    };
  }
}

/**
 * TEMPORARY: Force return test data for debugging - bypasses all sheet access
 */
/**
 * TEMPORARY: Force return test data for debugging - bypasses all sheet access
 */
function getVehicleReleaseHistoryWrapper(carNumber) {
  console.log('getVehicleReleaseHistoryWrapper called with:', carNumber);
  console.log('WRAPPER: Forcing test data return to debug null issue');
  
  // FORCE RETURN TEST DATA - This should never be null
  const testResult = [
    ['Status', 'Car Number', 'Date', 'Message'],
    ['DEBUG', carNumber || 'NO_CAR', new Date().toLocaleDateString(), 'Wrapper function is working - sheet access bypassed for debugging']
  ];
  
  console.log('WRAPPER: Returning test result:', testResult);
  return testResult;
}

/**
 * Simple test function to check CarT_P data and test history functions
 */
function _vehicleSheetReleaseVehicles(){
  try {
    const summary = getVehicleReleasedSummary();
    if (!summary.ok) {
      console.warn('_vehicleSheetReleaseVehicles: Vehicle_Released summary unavailable');
      return [];
    }

    const out = summary.vehicles.map(function(entry){
      const carNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
      if (!carNumber) return null;
      const status = _normStatus_(entry.status || 'RELEASE');
      if (status !== 'RELEASE') return null;
      return {
        carNumber: carNumber,
        project: entry.project || '',
        team: entry.team || '',
        make: entry.make || '',
        model: entry.model || '',
        usageType: entry.usageType || '',
        contractType: entry.contractType || '',
        owner: entry.owner || '',
        category: entry.category || '',
        status: status,
        remarks: entry.remarks || '',
        stars: entry.stars || 0,
        dateTime: entry.latestRelease || summary.updatedAt || '',
        responsibleBeneficiary: entry.responsibleBeneficiary || '',
        'R.Beneficiary': entry.responsibleBeneficiary || ''
      };
    }).filter(Boolean);

    out.sort(function(a,b){ return String(a.carNumber||'').localeCompare(String(b.carNumber||'')); });
    return out;
  } catch (err) {
    console.error('vehicleSheetReleaseVehicles error:', err);
    return [];
  }
}

function getVehicleReleaseSnapshots(limitPerVehicle) {
  try {
    const maxPerVehicle = Math.max(1, Math.min(10, Number(limitPerVehicle) || 3));
    const summary = getVehicleSummaryRows('Vehicle_History');
    if (!summary.rows || !summary.rows.length) {
      return { ok: true, limit: maxPerVehicle, vehicles: {} };
    }

    const IX = summary.headerIndex;
    const idx = function(labels, required) {
      if (!IX) return required ? -1 : -1;
      try { return IX.get(labels); } catch (_err) { return required ? -1 : -1; }
    };

    const vehicleIdx = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car'], true);
    if (vehicleIdx < 0) {
      console.warn('getVehicleReleaseSnapshots: Vehicle number column missing in Vehicle_History summary');
      return { ok: true, limit: maxPerVehicle, vehicles: {} };
    }
    const statusIdx = idx(['Status', 'In Use/Release', 'In Use / release', 'In Use'], false);
    const ratingIdx = idx(['Ratings', 'Stars', 'Rating'], false);
    const remarkIdx = idx(['Last Users remarks', 'Remarks', 'Feedback'], false);
    const dateIdx = idx(['Date and time of entry', 'Date and time', 'Timestamp', 'Date'], false);

    const grouped = new Map();
    summary.rows.forEach(function(row) {
      if (!row) return;
      const rawVehicleNumber = String(row[vehicleIdx] || '').trim();
      if (!rawVehicleNumber) return;
  const status = statusIdx >= 0 ? _normStatus_(row[statusIdx]) : '';
  if (status && status !== 'RELEASE' && status !== 'IN USE') return;
      const canonicalKey = _vehicleKey_(rawVehicleNumber);
      if (!canonicalKey) return;
      const aliasKey = rawVehicleNumber.toUpperCase();
      const targetKeys = canonicalKey === aliasKey ? [canonicalKey] : [canonicalKey, aliasKey];

      const dateValue = dateIdx >= 0 ? row[dateIdx] : '';
      const timestamp = dateIdx >= 0 ? _parseTs_(dateValue) : 0;
      const entry = {
        vehicleNumber: rawVehicleNumber,
  status: status || '',
  rating: ratingIdx >= 0 ? row[ratingIdx] || '' : '',
  remark: remarkIdx >= 0 ? row[remarkIdx] || '' : '',
        timestamp: timestamp || 0,
        date: dateValue || ''
      };

      targetKeys.forEach(function(key){
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(entry);
      });
    });

    const vehicles = {};
    grouped.forEach(function(list, key) {
      list.sort(function(a, b) {
        return (b.timestamp || 0) - (a.timestamp || 0);
      });
      vehicles[key] = list.slice(0, maxPerVehicle).map(function(entry) {
        return {
          status: entry.status || '',
          rating: entry.rating,
          remark: entry.remark,
          timestamp: entry.timestamp || null,
          date: entry.date || ''
        };
      });
    });
    return { ok: true, limit: maxPerVehicle, vehicles: vehicles };
  } catch (err) {
    console.error('getVehicleReleaseSnapshots failed:', err);
    return { ok: false, error: String(err) };
  }
}

function getCarTPVehicleMeta(limitPerVehicle) {
  try {
    const maxPerVehicle = Math.max(1, Math.min(10, Number(limitPerVehicle) || 3));
    const summary = getVehicleSummaryRows('CarT_P');
    if (summary.error) {
      return { ok: false, source: 'CarT_P', error: summary.error };
    }
    if (!summary.rows.length) {
      return {
        ok: true,
        source: 'CarT_P',
        vehicles: {},
        limit: maxPerVehicle,
        updatedAt: summary.updatedAt || ''
      };
    }

    const IX = summary.headerIndex;
    const idx = function(labels, fallbackIndex) {
      if (IX) {
        try {
          const found = IX.get(labels);
          if (typeof found === 'number' && found >= 0) {
            return found;
          }
        } catch (_err) {
          // fall through
        }
      }
      return typeof fallbackIndex === 'number' ? fallbackIndex : -1;
    };

    const vehicleIdx = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car', 'Vehicle'], 5);
    if (vehicleIdx < 0) {
      console.warn('getCarTPVehicleMeta: Vehicle number column missing in CarT_P summary');
      return {
        ok: true,
        source: 'CarT_P',
        vehicles: {},
        limit: maxPerVehicle,
        updatedAt: summary.updatedAt || '',
        message: 'Vehicle number column missing in CarT_P summary'
      };
    }
    const remarksIdx = idx(['Last 3 User Remarks', 'Last Users remarks', 'Remarks', 'User Remarks', 'Last 3 Remarks'], 12);
    const ratingsIdx = idx(['Last 3 Ratings', 'Ratings', 'Stars', 'Rating'], 13);

    const vehicles = {};

    function metaList(value) {
      if (value == null) return [];
      if (Array.isArray(value)) {
        return value.map(function(entry){ return String(entry || '').trim(); }).filter(Boolean);
      }
      const text = String(value || '').trim();
      if (!text) return [];
      return text
        .split(/[\r\n•;\|]+/)
        .map(function(part){ return String(part || '').trim(); })
        .filter(Boolean);
    }

    function appendUniqueLimited(target, entries) {
      if (!entries || !entries.length) return;
      entries.forEach(function(entry){
        const value = String(entry || '').trim();
        if (!value) return;
        const lower = value.toLowerCase();
        const exists = target.some(function(existing){
          return String(existing || '').trim().toLowerCase() === lower;
        });
        if (!exists) {
          target.push(value);
        }
      });
      if (target.length > maxPerVehicle) {
        target.length = maxPerVehicle;
      }
    }

    summary.rows.forEach(function(row){
      if (!row) return;
      const rawVehicle = String(row[vehicleIdx] || '').trim();
      if (!rawVehicle) return;
      const canonicalKey = _vehicleKey_(rawVehicle);
      if (!canonicalKey) return;
      const aliasKey = rawVehicle.toUpperCase();
      let bucket = vehicles[canonicalKey];
      if (!bucket) {
        bucket = { ratings: [], remarks: [] };
        vehicles[canonicalKey] = bucket;
      }
      if (aliasKey && !vehicles[aliasKey]) {
        vehicles[aliasKey] = bucket;
      }
      if (remarksIdx >= 0) {
        appendUniqueLimited(bucket.remarks, metaList(row[remarksIdx]));
      }
      if (ratingsIdx >= 0) {
        appendUniqueLimited(bucket.ratings, metaList(row[ratingsIdx]));
      }
    });

    return {
      ok: true,
      source: 'CarT_P',
      limit: maxPerVehicle,
      vehicles: vehicles,
      updatedAt: summary.updatedAt || ''
    };
  } catch (err) {
    console.error('getCarTPVehicleMeta failed:', err);
    return { ok: false, source: 'CarT_P', error: String(err) };
  }
}

/**
 * Helper function to filter vehicle history data by car number
 */
function filterVehicleHistory(data, carNumber) {
  if (!data || data.length < 2) return [];
  
  const carIdx = data[0].indexOf('Vehicle Number');
  if (carIdx < 0) {
    console.log('Vehicle Number column not found in history data');
    return [];
  }
  
  const filteredRows = data.slice(1).filter(row => {
    const rowCarNumber = String(row[carIdx] || '').trim();
    const searchCarNumber = String(carNumber || '').trim();
    return rowCarNumber === searchCarNumber;
  });
  
  console.log(`Found ${filteredRows.length} history entries for car ${carNumber}`);
  
  if (filteredRows.length === 0) return [];
  
  return [data[0]].concat(filteredRows);
}

/**
 * Debug function to test vehicle history functionality
 */
function testVehicleHistory() {
  try {
    console.log('Testing vehicle history functionality...');
    
    // First, check if CarT_P sheet has data
    const carTPData = _readCarTP_objects_();
    console.log('CarT_P data rows:', carTPData.length);
    
    if (carTPData.length === 0) {
      console.log('No data in CarT_P, adding sample data...');
      const sampleResult = addSampleCarData();
      console.log('Sample data result:', sampleResult);
      
      // Re-read after adding sample data
      const newCarTPData = _readCarTP_objects_();
      console.log('CarT_P data rows after adding sample:', newCarTPData.length);
    }
    
    // Show available car numbers
    const availableCars = carTPData.map(row => row['Vehicle Number']).filter(Boolean);
    console.log('Available car numbers:', availableCars);
    
    // Refresh the summary sheets to ensure they have current data
    console.log('Refreshing vehicle status sheets...');
    const refreshResult = refreshVehicleStatusSheets();
    console.log('Refresh result:', refreshResult);
    
    // Use the first available car number for testing
    let testCarNumber = 'TEST-VEHICLE-DEFAULT'; // Default test value
    if (availableCars.length > 0) {
      testCarNumber = availableCars[0];
    }
    
    console.log(`Testing RELEASE history for car: ${testCarNumber}`);
    const releaseHistory = getVehicleReleaseHistory(testCarNumber);
    console.log('Release history result length:', releaseHistory.length);
    console.log('Release history result:', releaseHistory);
    
    console.log(`Testing IN USE history for car: ${testCarNumber}`);
    const inUseHistory = getInUseHistoryForcedWorking(testCarNumber, '');
    console.log('In use history result length:', inUseHistory.length);
    console.log('In use history result:', inUseHistory);
    
    // Check summary sheets directly
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const releasedSheet = ss.getSheetByName('Vehicle_Released');
    const inUseSheet = ss.getSheetByName('Vehicle_InUse');
    const historySheet = ss.getSheetByName('Vehicle_History');
    
    return {
      ok: true,
      carTPDataRows: carTPData.length,
      availableCars: availableCars,
      refreshResult: refreshResult,
      releaseHistory: releaseHistory,
      inUseHistory: inUseHistory,
      testCarNumber: testCarNumber,
      summarySheets: {
        released: releasedSheet ? releasedSheet.getLastRow() : 0,
        inUse: inUseSheet ? inUseSheet.getLastRow() : 0,
        history: historySheet ? historySheet.getLastRow() : 0
      }
    };
  } catch (error) {
    console.error('Error in testVehicleHistory:', error);
    return {
      ok: false,
      error: String(error)
    };
  }
}

/**
 * Force refresh of vehicle summary sheets and test history functions
 */
function forceRefreshAndTest() {
  try {
    console.log('Force refreshing vehicle summary sheets...');
    
    // Clear existing summary sheets first
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetNames = ['Vehicle_Released', 'Vehicle_InUse', 'Vehicle_History'];
    
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        ss.deleteSheet(sheet);
        console.log(`Deleted existing ${name} sheet`);
      }
    });
    
    // Now refresh to recreate
    const refreshResult = refreshVehicleStatusSheets();
    console.log('Refresh result:', refreshResult);
    
    // Test with first available car
    const carTPData = _readCarTP_objects_();
    const availableCars = carTPData.map(row => row['Vehicle Number']).filter(Boolean);
    
    if (availableCars.length > 0) {
      const testCar = availableCars[0];
      console.log(`Testing history for car: ${testCar}`);
      
      const releaseHistory = getVehicleReleaseHistory(testCar);
      const inUseHistory = getInUseHistoryForcedWorking(testCar, '');
      
      return {
        ok: true,
        refreshResult: refreshResult,
        testCar: testCar,
        releaseHistoryLength: releaseHistory.length,
        inUseHistoryLength: inUseHistory.length,
        availableCars: availableCars
      };
    } else {
      return {
        ok: false,
        error: 'No cars available in CarT_P sheet'
      };
    }
  } catch (error) {
    console.error('Error in forceRefreshAndTest:', error);
    return {
      ok: false,
      error: String(error)
    };
  }
}

/**
 * Refreshes summary sheets for vehicle status:
 * - Vehicle_InUse: latest IN USE row per beneficiary
 * - Vehicle_Released: latest RELEASE row per beneficiary
 * - Vehicle_History: full CarT_P history (unchanged)
 * Call this after any change (assignment, release, beneficiary change).
 */
function _sanitizeResponsibleName(value) {
  if (!value && value !== 0) return '';
  let text = String(value).trim();
  if (!text) return '';
  text = text.replace(/^name of (responsible )?beneficiary\s*:?/i, '');
  text = text.replace(/^responsible beneficiary\s*:?/i, '');
  text = text.replace(/^r\.?\s*ben\s*:?/i, '');
  text = text.replace(/^r\.?\s*beneficiary\s*:?/i, '');
  text = text.replace(/^beneficiary\s*:?/i, '');
  text = text.replace(/^name\s*:?/i, '');
  text = text.replace(/^[\s:,-]+/, '');
  text = text.replace(/\s+/g, ' ').trim();
  if (!text) return '';
  if (/^(name of (responsible )?beneficiary|responsible beneficiary|beneficiary|name)$/i.test(text)) {
    return '';
  }
  return text;
}

function _extractResponsibleName(row) {
  if (!row || typeof row !== 'object') return '';
  const sources = [
    row['R. Ben'],
    row.rBenShort,
    row.responsibleBeneficiary,
    row['Responsible Beneficiary'],
    row['Name of Responsible beneficiary'],
    row['R.Beneficiary']
  ];
  for (let i = 0; i < sources.length; i++) {
    const source = sources[i];
    if (!source && source !== 0) continue;
    const pieces = _splitBeneficiaryNames_(source)
      .map(_sanitizeResponsibleName)
      .filter(Boolean);
    if (pieces.length) {
      return pieces[0];
    }
    const single = _sanitizeResponsibleName(source);
    if (single) return single;
  }
  return '';
}

function refreshVehicleStatusSheets() {
  const allCarRows = _readCarTP_objects_();
  if (!allCarRows.length) {
    console.log('No CarT_P data found.');
    invalidateVehicleReleasedCache('CarT_P refresh encountered no data');
    return { ok: false, error: 'No CarT_P data' };
  }

  const latestByBeneficiary = new Map();
  const latestByVehicle = new Map();

  for (let idx = 0; idx < allCarRows.length; idx++) {
    const row = allCarRows[idx];

    const tsRaw = typeof row._ts === 'number' ? row._ts : _parseTs_(row['Date and time of entry']);
    row._ts = (typeof tsRaw === 'number' && !isNaN(tsRaw)) ? tsRaw : 0;
    row._rowIndex = (typeof row._rowIndex === 'number') ? row._rowIndex : (idx + 2);

    const names = _beneficiaryNamesFromRow_(row);
    const responsibleName = _extractResponsibleName(row);
    const responsibleKey = _beneficiaryKey_(responsibleName);
    if (names.length) {
      names.forEach(function(name){
        const cleaned = _sanitizeResponsibleName(name) || _norm(name);
        if (!cleaned) return;
        const key = _beneficiaryKey_(cleaned);
        if (!key) return;
        const prev = latestByBeneficiary.get(key);
        if (!prev || row._ts > prev._ts || (row._ts === prev._ts && row._rowIndex >= prev._rowIndex)) {
          const clone = Object.assign({}, row);
          clone['R.Beneficiary'] = cleaned;
          clone['R. Ben'] = responsibleName || '';
          clone.responsibleBeneficiary = responsibleName;
          clone.__beneficiaryKey = key;
          latestByBeneficiary.set(key, clone);
        }
      });
    } else {
      const beneficiary = _sanitizeResponsibleName(String(
        row['R.Beneficiary'] ||
        row['R. Ben'] ||
        row.responsibleBeneficiary ||
        row['Responsible Beneficiary'] ||
        row['Name of Responsible beneficiary'] ||
        ''
      ));
      if (beneficiary) {
        const key = _beneficiaryKey_(beneficiary);
        if (!key) {
          continue;
        }
        const prev = latestByBeneficiary.get(key);
        if (!prev || row._ts > prev._ts || (row._ts === prev._ts && row._rowIndex >= prev._rowIndex)) {
          const clone = Object.assign({}, row);
          clone['R.Beneficiary'] = beneficiary;
          clone['R. Ben'] = responsibleName || '';
          clone.responsibleBeneficiary = responsibleName;
          clone.__beneficiaryKey = key;
          latestByBeneficiary.set(key, clone);
        }
      }
    }

    const vehicle = String(row['Vehicle Number'] || '').trim().toUpperCase();
    if (vehicle) {
      const prevVeh = latestByVehicle.get(vehicle);
      if (!prevVeh || row._ts > prevVeh._ts || (row._ts === prevVeh._ts && row._rowIndex >= prevVeh._rowIndex)) {
        latestByVehicle.set(vehicle, row);
      }
    }
  }

  const finalInUseSummaries = [];
  latestByBeneficiary.forEach(function(row){
    if (_normStatus_(row.Status) === 'IN USE') {
      if (!row['R. Ben']) {
        const responsible = _extractResponsibleName(row);
        row['R. Ben'] = responsible;
        row.responsibleBeneficiary = responsible;
      }
      finalInUseSummaries.push(row);
    }
  });

  const finalReleasedSummaries = [];
  latestByVehicle.forEach(function(row){
    if (_normStatus_(row.Status) === 'RELEASE') {
      finalReleasedSummaries.push(row);
    }
  });

  const sortByLatestEntry = (a, b) => {
    const tsA = typeof a._ts === 'number' ? a._ts : 0;
    const tsB = typeof b._ts === 'number' ? b._ts : 0;
    if (tsA !== tsB) return tsB - tsA;
    const idxA = typeof a._rowIndex === 'number' ? a._rowIndex : 0;
    const idxB = typeof b._rowIndex === 'number' ? b._rowIndex : 0;
    return idxB - idxA;
  };

  finalInUseSummaries.sort(sortByLatestEntry);
  finalReleasedSummaries.sort(sortByLatestEntry);

  console.log(`Writing ${finalInUseSummaries.length} rows to Vehicle_InUse sheet.`);
  writeVehicleSummarySheet('Vehicle_InUse', finalInUseSummaries);

  console.log(`Writing ${finalReleasedSummaries.length} rows to Vehicle_Released sheet.`);
  writeVehicleSummarySheet('Vehicle_Released', finalReleasedSummaries);

  writeVehicleSummarySheet('Vehicle_History', allCarRows);

  return {
    ok: true,
    inUse: finalInUseSummaries.length,
    released: finalReleasedSummaries.length
  };
}

/**
 * Helper to write summary data to a named sheet.
 */
function writeVehicleSummarySheet(sheetName, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clearContents();
  sh.getRange(1,1,1,VEHICLE_SUMMARY_HEADER.length).setValues([VEHICLE_SUMMARY_HEADER]);
  sh.setFrozenRows(1);
  if (rows.length) {
    function selectResponsible(row) {
      const sources = [
        row['R. Ben'],
        row.rBenShort,
        row.responsibleBeneficiary,
        row['R.Beneficiary'],
        row['Responsible Beneficiary'],
        row['Name of Responsible beneficiary']
      ];
      function sanitize(value) {
        if (!value && value !== 0) return '';
        let text = String(value).trim();
        if (!text) return '';
        text = text.replace(/^name of (responsible )?beneficiary\s*:?/i, '');
        text = text.replace(/^responsible beneficiary\s*:?/i, '');
        text = text.replace(/^r\.?\s*ben\s*:?/i, '');
        text = text.replace(/^r\.?\s*beneficiary\s*:?/i, '');
        text = text.replace(/^beneficiary\s*:?/i, '');
        text = text.replace(/^name\s*:?/i, '');
        text = text.replace(/^[\s:,-]+/, '');
        text = text.replace(/\s+/g, ' ').trim();
        if (!text) return '';
        if (/^(name of (responsible )?beneficiary|responsible beneficiary|beneficiary|name)$/i.test(text)) {
          return '';
        }
        return text;
      }
      function pickFrom(source) {
        const cleanedList = _splitBeneficiaryNames_(source)
          .map(sanitize)
          .filter(Boolean);
        if (cleanedList.length) {
          return cleanedList[0];
        }
        const single = sanitize(source);
        return single || '';
      }
      for (let i = 0; i < sources.length; i++) {
        const candidate = pickFrom(sources[i]);
        if (candidate) {
          return candidate;
        }
      }
      return '';
    }
    const values = rows.map(r => [
      r.Ref || r['Reference Number'] || '',
      r['Date and time of entry'] || '',
      r.Project || '',
      r.Team || '',
      r['R.Beneficiary'] || r['R. Ben'] || r.responsibleBeneficiary || '',
      r['Vehicle Number'] || '',
      r.Make || '',
      r.Model || '',
      r.Category || '',
      r['Usage Type'] || '',
      r.Owner || '',
      _normStatus_(r.Status) || (r.Status || ''),
      r['Last Users remarks'] || '',
      r.Ratings || '',
      r['Submitter username'] || '',
      r['R.Ben Time'] || r.rBenTime || r.responsibleBeneficiaryTime || r.responsibleTime || '',
      selectResponsible(r)
    ]);
    sh.getRange(2,1,values.length,VEHICLE_SUMMARY_HEADER.length).setValues(values);
  }
  try { sh.autoResizeColumns(1, VEHICLE_SUMMARY_HEADER.length); } catch(_){}
  if (sheetName === 'Vehicle_Released') {
    invalidateVehicleReleasedCache('Vehicle_Released sheet rewritten');
  }
}

function _sheetMatchesVehicleReleased_(sheet) {
  if (!sheet || typeof sheet.getName !== 'function') return false;
  const name = String(sheet.getName() || '').trim();
  if (name === 'Vehicle_Released') return true;
  const normalized = name.replace(/\s+/g, '').toLowerCase();
  return normalized === 'vehicle_released' || normalized === 'vehiclereleased';
}

function _shouldProcessVehicleReleasedEvent_(e) {
  try {
    if (e && e.range && typeof e.range.getSheet === 'function') {
      if (!_sheetMatchesVehicleReleased_(e.range.getSheet())) {
        return false;
      }
    } else if (e && e.source && typeof e.source.getActiveSheet === 'function') {
      if (!_sheetMatchesVehicleReleased_(e.source.getActiveSheet())) {
        return false;
      }
    }
    if (e && e.source && typeof e.source.getId === 'function' && SHEET_ID) {
      if (e.source.getId() !== SHEET_ID) {
        return false;
      }
    }
    return true;
  } catch (err) {
    console.warn('_shouldProcessVehicleReleasedEvent_ error', err);
  }
  return false;
}

function onVehicleReleasedChange(e) {
  const changeType = e && e.changeType ? e.changeType : 'UNKNOWN';
  try {
    if (!_shouldProcessVehicleReleasedEvent_(e)) {
      return { ok: false, skipped: true, changeType };
    }
    invalidateVehicleReleasedCache(`Vehicle_Released change (${changeType})`);
    return { ok: true, invalidated: true, changeType };
  } catch (err) {
    console.error('onVehicleReleasedChange error', err, 'changeType:', changeType);
    return { ok: false, error: String(err), changeType };
  }
}

/**
 * Synchronisation helpers to keep summary sheets fresh when CarT_P changes.
 */
function _sheetMatchesCarTP_(sheet) {
  if (!sheet || typeof sheet.getName !== 'function') return false;
  if (sheet.getName() !== 'CarT_P') return false;
  if (!CAR_SHEET_ID) return true;
  try {
    const parent = typeof sheet.getParent === 'function' ? sheet.getParent() : null;
    if (parent && typeof parent.getId === 'function') {
      return parent.getId() === CAR_SHEET_ID;
    }
  } catch (err) {
    console.warn('_sheetMatchesCarTP_ parent check failed', err);
  }
  return true;
}

function _shouldProcessCarTPEvent_(e, options) {
  options = options || {};
  try {
    if (e && e.range && typeof e.range.getSheet === 'function') {
      return _sheetMatchesCarTP_(e.range.getSheet());
    }
    if (options.fallbackToActiveSheet && e && e.source && typeof e.source.getActiveSheet === 'function') {
      if (CAR_SHEET_ID && typeof e.source.getId === 'function' && e.source.getId() !== CAR_SHEET_ID) {
        return false;
      }
      return _sheetMatchesCarTP_(e.source.getActiveSheet());
    }
    if (e && e.source && typeof e.source.getId === 'function' && CAR_SHEET_ID && e.source.getId() === CAR_SHEET_ID) {
      // Installable change events often omit the edited range. If the event
      // originated from the Car spreadsheet itself, treat it as a CarT_P
      // update so downstream refresh logic still runs.
      return true;
    }
  } catch (err) {
    console.warn('_shouldProcessCarTPEvent_ error', err);
  }
  return false;
}

function _runCarTPSummaryRefresh_(source, meta) {
  const started = Date.now();
  try {
    const refreshResult = refreshVehicleStatusSheets();
    const syncResult = syncVehicleSheetFromCarTP();
    const durationMs = Date.now() - started;
    const context = meta ? Object.assign({}, meta) : {};
    context.source = source || 'manual';
    context.durationMs = durationMs;
    context.refreshResult = refreshResult;
    context.syncResult = syncResult;
    console.log('CarT_P summaries refreshed', context);
    invalidateVehicleReleasedCache(`CarT_P summary refresh (${context.source || 'unknown'})`);
    return { ok: true, durationMs, refreshResult, syncResult };
  } catch (err) {
    console.error('CarT_P summary refresh failed', err, { source, meta });
    return { ok: false, error: String(err) };
  }
}

function _maybeAutoRefreshCarTPSummaries_(ttlSec) {
  try {
    const props = PropertiesService.getScriptProperties();
    const ttlMs = Math.max(3, Number(ttlSec) || 5) * 1000;
    const now = Date.now();
    const last = Number(props.getProperty('CAR_TP_LAST_AUTO_REFRESH') || '0');
    if (now - last < ttlMs) {
      return { ok: true, skipped: true, reason: 'ttl' };
    }
    const refreshResult = refreshVehicleStatusSheets();
    const syncResult = syncVehicleSheetFromCarTP();
    props.setProperty('CAR_TP_LAST_AUTO_REFRESH', String(now));
    invalidateVehicleReleasedCache('Auto refresh trigger');
    return { ok: true, refreshResult: refreshResult, syncResult: syncResult };
  } catch (error) {
    console.error('_maybeAutoRefreshCarTPSummaries_ error:', error);
    try { PropertiesService.getScriptProperties().deleteProperty('CAR_TP_LAST_AUTO_REFRESH'); } catch (_e) { /* ignore */ }
    return { ok: false, error: String(error) };
  }
}

/**
 * Simple trigger: refresh vehicle summary sheets whenever CarT_P is edited.
 */
function onEdit(e) {
  try {
    if (!_shouldProcessCarTPEvent_(e)) {
      return;
    }
    const meta = {};
    try {
      if (e && e.range && typeof e.range.getA1Notation === 'function') {
        meta.range = e.range.getA1Notation();
      }
    } catch (_ignored) {
      // ignore errors while gathering metadata
    }
    return _runCarTPSummaryRefresh_('onEdit', meta);
  } catch (err) {
    console.error('onEdit handler failed', err);
  }
}

/**
 * Installable trigger handler to keep vehicle summary sheets in sync when CarT_P changes.
 */
function onCarTPChange(e) {
  const changeType = e && e.changeType ? e.changeType : 'UNKNOWN';
  try {
    if (!_shouldProcessCarTPEvent_(e, { fallbackToActiveSheet: true })) {
      return { ok: false, skipped: true, changeType, reason: 'Not a CarT_P change' };
    }
    const sameSpreadsheet = !CAR_SHEET_ID || CAR_SHEET_ID === SHEET_ID;
    if (changeType === 'EDIT' && sameSpreadsheet) {
      // When the script is bound to the same spreadsheet as CarT_P,
      // the simple onEdit trigger will already refresh the summaries.
      // Skip here to avoid running the refresh twice for the same edit.
      return { ok: true, skipped: true, changeType, reason: 'Handled by onEdit' };
    }
    return _runCarTPSummaryRefresh_('onCarTPChange', { changeType });
  } catch (err) {
    console.error('onCarTPChange error', err, 'changeType:', changeType);
    return { ok: false, error: String(err), changeType };
  }
}

/**
 * Quick diagnostic + setup for realtime Vehicle_InUse updates.
 * - Verifies CarT_P access (via debugCarTPOverview)
 * - Ensures the summary tab name exists exactly as 'Vehicle_InUse' in SHEET_ID
 * - Ensures the installable onChange trigger is attached to the CarT_P spreadsheet
 * - Forces one refresh and sync
 * Returns a summary object for visibility.
 */
function checkCarTPRealtimeSync() {
  const summary = { ok: true, steps: [] };

  // 1) Verify CarT_P access
  try {
    const car = debugCarTPOverview();
    if (!car || !car.ok) {
      summary.ok = false;
      summary.steps.push('CarT_P not accessible. Check CAR_SHEET_ID and sharing/permissions.');
      return summary;
    }
    summary.steps.push(`CarT_P found: ${car.sheetName} (rows=${car.lastRow}, cols=${car.lastCol})`);
  } catch (e) {
    summary.ok = false;
    summary.steps.push('CarT_P overview failed: ' + String(e));
    return summary;
  }

  // 2) Ensure Vehicle_InUse tab exists in target SHEET_ID
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Vehicle_InUse');
    if (!sh) {
      sh = ss.insertSheet('Vehicle_InUse');
      summary.steps.push('Created Vehicle_InUse tab (exact name).');
    } else {
      summary.steps.push('Vehicle_InUse tab exists.');
    }
  } catch (e) {
    summary.ok = false;
    summary.steps.push('Failed to verify/create Vehicle_InUse tab: ' + String(e));
    return summary;
  }

  // 3) Ensure installable onChange trigger for CarT_P
  try {
    const trigRes = createCarTPOnChangeTrigger();
    if (trigRes && trigRes.ok) {
      summary.steps.push('onChange trigger ensured for CarT_P.');
    } else {
      summary.steps.push('Trigger creation attempted; check triggers in editor if issues persist.');
    }
  } catch (e) {
    summary.steps.push('Trigger setup error: ' + String(e));
  }

  // 4) Force one refresh + sync
  try {
    const res = _runCarTPSummaryRefresh_('manual-check', {});
    const inUse = res && res.refreshResult ? res.refreshResult.inUse : undefined;
    const released = res && res.refreshResult ? res.refreshResult.released : undefined;
    summary.steps.push(`Refresh done. InUse=${inUse}, Released=${released}`);
  } catch (e) {
    summary.ok = false;
    summary.steps.push('Refresh failed: ' + String(e));
  }

  return summary;
}

/**
 * Create an installable onChange trigger for the CarT_P spreadsheet.
 */
function createCarTPOnChangeTrigger() {
  try {
    removeCarTPOnChangeTrigger();
    const carSpreadsheet = SpreadsheetApp.openById(CAR_SHEET_ID);
    ScriptApp.newTrigger('onCarTPChange')
      .forSpreadsheet(carSpreadsheet)
      .onChange()
      .create();
    return { ok: true, message: 'CarT_P onChange trigger created' };
  } catch (err) {
    return { ok: false, message: String(err) };
  }
}

/**
 * Remove any existing installable triggers that call onCarTPChange.
 */
function removeCarTPOnChangeTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'onCarTPChange') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    return { ok: true, message: 'Removed CarT_P onChange triggers' };
  } catch (err) {
    return { ok: false, message: String(err) };
  }
}
/**
 * Get responsible beneficiary and teammate usage breakdown for a car from CarT_P.
 */
function getCarReleaseDetails(carNumber) {
  const targetCar = String(carNumber || '').trim();
  if (!targetCar) {
    return { ok: false, error: 'Car number required' };
  }

  const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
  const sheet = ss.getSheetByName('CarT_P');
  if (!sheet) return { ok: false, error: 'CarT_P sheet not found' };

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    return { ok: true, responsible: '', teamMembersUsing: [], teamMembersNotUsing: [], teamMembers: [] };
  }

  const header = data[0];
  const carIdx = header.indexOf('Vehicle Number');
  const teamIdx = header.indexOf('Team');
  const statusIdx = header.indexOf('Status');
  let rbenIdx = header.indexOf('R.Beneficiary');
  if (rbenIdx < 0) {
    rbenIdx = header.indexOf('R. Ben');
  }
  const dateIdx = header.indexOf('Date and time of entry');

  if (carIdx < 0 || teamIdx < 0 || statusIdx < 0 || rbenIdx < 0 || dateIdx < 0) {
    return { ok: false, error: 'Required columns missing in CarT_P' };
  }

  const normalizeHeader = function(value){
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
  };

  const normalizedHead = header.map(normalizeHeader);

  const findIndex = function(labels){
    const list = (Array.isArray(labels) ? labels : [labels])
      .map(normalizeHeader)
      .filter(Boolean);
    if (!list.length) return -1;
    for (let i = 0; i < normalizedHead.length; i++) {
      if (!normalizedHead[i]) continue;
      if (list.indexOf(normalizedHead[i]) !== -1) {
        return i;
      }
    }
    return -1;
  };

  const projectIdx = findIndex(['Project', 'Project Name']);
  const makeIdx = findIndex(['Make', 'Car Make', 'Brand']);
  const modelIdx = findIndex(['Model', 'Car Model']);
  const categoryIdx = findIndex(['Category', 'Vehicle Category', 'Category Name', 'Cat']);
  const usageIdx = findIndex(['Usage Type', 'Usage', 'Use Type']);
  const ownerIdx = findIndex(['Owner', 'Owner Name', 'Owner Info']);
  const remarksIdx = findIndex(['Last Users remarks', 'Remarks', 'Feedback']);
  const starsIdx = findIndex(['Stars', 'Ratings', 'Rating']);
  const submitIdx = findIndex(['Submitter username', 'Submitter', 'User']);
  const respShortIdx = findIndex(['R. Ben', 'R Ben']);
  const respFullIdx = findIndex(['Responsible Beneficiary', 'Name of Responsible beneficiary']);

  const splitNames = function(value) {
    if (!value) return [];
    return String(value)
      .split(/[,;\n]+/)
      .map(function(name){ return String(name || '').trim(); })
      .filter(Boolean);
  };

  const rowsForCar = [];
  let responsible = '';
  let teamName = '';
  let teamKey = '';
  let latestInUseTs = -1;
  let detailRow = null;
  let detailRowTs = -1;
  let fallbackRow = null;
  let fallbackTs = -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[carIdx]).trim() !== targetCar) continue;
    rowsForCar.push(row);
    const status = _normStatus_(row[statusIdx]);
    const ts = _parseTs_(row[dateIdx]);
    if (ts >= fallbackTs) {
      fallbackTs = ts;
      fallbackRow = row;
    }
    const candidateTeam = String(row[teamIdx] || '').trim();
    const names = splitNames(row[rbenIdx]);
    const responsibleShortRaw = respShortIdx >= 0 ? row[respShortIdx] : '';
    const responsibleFullRaw = respFullIdx >= 0 ? row[respFullIdx] : '';
    const responsibleCandidate = _sanitizeResponsibleName(
      responsibleShortRaw ||
      responsibleFullRaw ||
      (names.length ? names[0] : '')
    );
    if (status === 'IN USE' && names.length) {
      if (ts >= latestInUseTs) {
        latestInUseTs = ts;
        responsible = responsibleCandidate || names[0];
        teamName = candidateTeam;
        teamKey = _normTeamKey(candidateTeam);
      }
      if (ts >= detailRowTs) {
        detailRow = row;
        detailRowTs = ts;
      }
    } else if (!detailRow || ts >= detailRowTs) {
      detailRow = row;
      detailRowTs = ts;
    }
  }

  if (!rowsForCar.length) {
    return { ok: true, responsible: '', teamMembersUsing: [], teamMembersNotUsing: [], teamMembers: [] };
  }

  if (!detailRow) {
    detailRow = fallbackRow;
  }

  if (!teamName) {
    const fallbackTeam = detailRow ? String(detailRow[teamIdx] || '').trim() : String(rowsForCar[0][teamIdx] || '').trim();
    teamName = fallbackTeam;
    teamKey = _normTeamKey(fallbackTeam);
    if (!responsible) {
      const fallbackShort = respShortIdx >= 0 ? (detailRow ? detailRow[respShortIdx] : rowsForCar[0][respShortIdx]) : '';
      const fallbackFull = respFullIdx >= 0 ? (detailRow ? detailRow[respFullIdx] : rowsForCar[0][respFullIdx]) : '';
      const fallbackNames = splitNames(detailRow ? detailRow[rbenIdx] : rowsForCar[0][rbenIdx]);
      const fallbackCandidate = _sanitizeResponsibleName(
        fallbackShort ||
        fallbackFull ||
        (fallbackNames.length ? fallbackNames[0] : '')
      );
      if (fallbackCandidate) {
        responsible = fallbackCandidate;
      } else if (fallbackNames.length) {
        responsible = fallbackNames[0];
      }
    }
  }

  const memberStatusMap = new Map();
  rowsForCar.forEach(function(row) {
    const rowTeam = String(row[teamIdx] || '').trim();
    if (teamKey && _normTeamKey(rowTeam) !== teamKey) return;
    const status = _normStatus_(row[statusIdx]);
    const ts = _parseTs_(row[dateIdx]);
    const names = splitNames(row[rbenIdx]);
    names.forEach(function(name){
      if (!name) return;
      const key = _beneficiaryKey_(name);
      const existing = memberStatusMap.get(key);
      if (!existing || ts >= existing.ts) {
        memberStatusMap.set(key, {
          name: name,
          status: status || '',
          ts: ts || 0
        });
      }
    });
  });

  const responsibleKey = _beneficiaryKey_(responsible);
  const teamMembersUsing = [];
  const teamMembersNotUsing = [];
  const usingKeySet = new Set();
  const notUsingKeySet = new Set();
  if (responsibleKey) {
    usingKeySet.add(responsibleKey);
  }

  memberStatusMap.forEach(function(entry, key){
    if (responsibleKey && key === responsibleKey) return;
    if (entry.status === 'IN USE') {
      if (!usingKeySet.has(key)) {
        usingKeySet.add(key);
        teamMembersUsing.push(entry.name);
      }
    } else if (key) {
      if (!notUsingKeySet.has(key)) {
        notUsingKeySet.add(key);
        teamMembersNotUsing.push(entry.name);
      }
    }
  });

  if (teamKey) {
    try {
      const ddRows = _readDD_compact_();
      if (Array.isArray(ddRows) && ddRows.length) {
        const latestByMember = new Map();
        ddRows.forEach(function(row){
          if (!row || row.teamKey !== teamKey) return;
          const benKey = row.beneficiaryKey;
          if (!benKey) return;
          const composite = row.teamKey + '|' + benKey;
          const ts = Number(row.timestamp || 0);
          const existing = latestByMember.get(composite);
          if (!existing || ts >= existing.ts) {
            latestByMember.set(composite, {
              name: row.beneficiary,
              key: benKey,
              ts: ts
            });
          }
        });
        latestByMember.forEach(function(info){
          const key = info && info.key;
          if (!key) return;
          if (usingKeySet.has(key)) return;
          if (notUsingKeySet.has(key)) return;
          notUsingKeySet.add(key);
          teamMembersNotUsing.push(info.name);
        });
      }
    } catch (ddErr) {
      console.warn('getCarReleaseDetails: DD reconciliation failed for team', teamName, ddErr);
    }
  }

  const sortNames = function(list) {
    return list.sort(function(a, b){
      return a.localeCompare(b, undefined, { sensitivity: 'base' });
    });
  };

  sortNames(teamMembersUsing);
  sortNames(teamMembersNotUsing);

  const readString = function(idx){
    if (idx < 0 || !detailRow) return '';
    const value = detailRow[idx];
    if (value == null) return '';
    if (value instanceof Date) return Utilities.formatDate(value, TZ(), 'yyyy-MM-dd HH:mm');
    return String(value).trim();
  };

  const starValue = function(idx){
    if (idx < 0 || !detailRow) return 0;
    const raw = detailRow[idx];
    const num = Number(raw);
    return Number.isFinite(num) ? num : 0;
  };

  const projectValue = readString(projectIdx);
  const teamValue = teamName || readString(teamIdx);
  const statusValue = detailRow && statusIdx >= 0 ? _normStatus_(detailRow[statusIdx]) || '' : '';

  const carData = {
    project: projectValue,
    projectName: projectValue,
    team: teamValue,
    teamName: teamValue,
    carNumber: targetCar,
    category: readString(categoryIdx),
    usageType: readString(usageIdx),
    owner: readString(ownerIdx),
    make: readString(makeIdx),
    model: readString(modelIdx),
    status: statusValue,
    lastUsers: readString(rbenIdx),
    remarks: readString(remarksIdx),
    stars: starValue(starsIdx),
    submitter: readString(submitIdx)
  };

  return {
    ok: true,
    responsible: responsible || '',
    team: teamValue || '',
    project: projectValue || '',
    teamMembersUsing: teamMembersUsing,
    teamMembersNotUsing: teamMembersNotUsing,
    teamMembers: teamMembersUsing.slice(),
    carData: carData
  };
}
/** 
 * Web App backend for "test fund req"
 * Sheet ID: 107RBhtlS9c5iHhmDZXqOOWJ52cllcmmjuSBCVCMe8uA
 * Saves rows into "Sheet1".
 * Project/Teams/Beneficiaries come from "DD".
 *
 * Speed upgrade:
 *  - 10 min cache with versioned keys (manual bust supported)
 *  - Faster DD reader (header via getDisplayValues, body via getValues, narrowed range)
 *  - Project→Teams index to serve Projects & Teams with one cached build
 *  - NEW: Client-preload of compact DD rows (for instant client-side filtering)
 */

// Google Apps Script globals are available by default
// HtmlService, SpreadsheetApp, LockService, PropertiesService, CacheService, Logger

// Top-level configuration: use guarded assignments so redeclaration across
// multiple script files doesn't throw a SyntaxError in Apps Script.
var SHEET_ID = (typeof SHEET_ID !== 'undefined') ? SHEET_ID : '107RBhtlS9c5iHhmDZXqOOWJ52cllcmmjuSBCVCMe8uA';
// If your CarT_P sheet lives in a different Spreadsheet, set this to that ID.
// Otherwise, it defaults to SHEET_ID.
var CAR_SHEET_ID = (typeof CAR_SHEET_ID !== 'undefined') ? CAR_SHEET_ID : SHEET_ID;
var SHEET_NAME = (typeof SHEET_NAME !== 'undefined') ? SHEET_NAME : 'submissions';
var SUBMISSIONS_SHEET_NAME = (typeof SUBMISSIONS_SHEET_NAME !== 'undefined') ? SUBMISSIONS_SHEET_NAME : 'submissions';
var DATA_SHEET_NAME = (typeof DATA_SHEET_NAME !== 'undefined') ? DATA_SHEET_NAME : 'DD';
var CACHE_TTL_SEC = (typeof CACHE_TTL_SEC !== 'undefined') ? CACHE_TTL_SEC : 600; // 10 minutes
// Default RAG folder (for Drive-wide docs/sheets grounding)
var RAG_DEFAULT_FOLDER_ID = (typeof RAG_DEFAULT_FOLDER_ID !== 'undefined') ? RAG_DEFAULT_FOLDER_ID : '1CsFgq7ocPT_X8RsMK_qknAH_ftpsz0Rm';

/**
 * Timezone helper. Prefer the Spreadsheet timezone when available, otherwise default to Africa/Dar_es_Salaam.
 * Use as TZ() throughout the code to avoid ReferenceError when TZ is missing.
 */
function TZ(){
  try{
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = ss.getSpreadsheetTimeZone();
    if (tz && typeof tz === 'string' && tz.trim()) return tz;
  }catch(e){ /* ignore */ }
  return 'Africa/Dar_es_Salaam';
}

/** Ensure the spreadsheet timezone is at least set to East Africa Time (EAT) if not configured. */
function ensureSpreadsheetTZ(){
  try{
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = ss.getSpreadsheetTimeZone();
    if (!tz || tz.indexOf('Dar_es_Salaam') === -1) {
      try { ss.setSpreadsheetTimeZone('Africa/Dar_es_Salaam'); } catch(e) { /* may not have permission */ }
    }
  }catch(e){ /* ignore */ }
}

/**
 * Parse an amount value robustly from a spreadsheet cell.
 * Accepts numbers or strings like "Tsh 115", "115", "1,234.50" and returns a Number.
 * Returns 0 for non-numeric values.
 */
function parseAmount(v){
  try{
    if (v === null || typeof v === 'undefined' || v === '') return 0;
    if (typeof v === 'number') return v;
    var s = String(v || '').replace(/\u00A0/g,' ').trim(); // normalize nbsp
    // find first numeric token (allows commas and dots)
    var m = s.match(/-?\d[\d,\.\s]*/);
    if (!m) return 0;
    var numStr = m[0].replace(/[ ,\s]+/g,'');
    var n = parseFloat(numStr);
    return isNaN(n) ? 0 : n;
  }catch(e){ return 0; }
}

// Use the existing `_escHtml` helper (defined later) for HTML escaping. Removed duplicate `_escHtml_`.

/* -------------------------- templating + UI -------------------------- */

function include(filename) {
  try {
    if (filename === undefined || filename === null || filename === '') {
      console.error('Include called with invalid filename: ' + filename);
      return '<!-- Error: Invalid filename passed to include function -->';
    }
    console.log('Including file: ' + filename);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error('Error including file ' + filename + ': ' + error.toString());
    return '<!-- Error loading ' + filename + ': ' + error.toString() + ' -->';
  }
}

/**
 * Injects:
 *  - PTI: Project→Teams index
 *  - DDC: Compact DD client pack for instant filter (arrays for minimal size)
 */
function doGet(e) {
  // Handle chat requests
  if (e && e.parameter && e.parameter.action === 'chat') {
    try {
      console.log('Chat request received:', e.parameter.message);
      var message = e.parameter.message || '';
      var response = (typeof Chat !== 'undefined' && Chat && typeof Chat.processChatQuery === 'function')
        ? Chat.processChatQuery(message)
        : processChatQuery(message);
      console.log('Chat response:', response);
      return ContentService
        .createTextOutput(JSON.stringify({response: response}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      console.error('doGet chat error:', error);
      return ContentService
        .createTextOutput(JSON.stringify({error: 'Chat system error: ' + error.message}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Typeahead suggestions endpoint
  if (e && e.parameter && e.parameter.action === 'suggest') {
    try {
      var prefix = (e.parameter.prefix || '').toString();
      var limit = Math.max(1, Math.min(50, Number(e.parameter.limit || 12)));
      var types = (e.parameter.types || '').toString().trim();
      var typeList = types ? types.split(',').map(function(t){ return String(t||'').trim().toLowerCase(); }).filter(Boolean) : [];
      var res = getTypeaheadSuggestions(prefix, limit, typeList);
      return ContentService.createTextOutput(JSON.stringify({ ok:true, suggestions: res })).setMimeType(ContentService.MimeType.JSON);
    } catch (errS) {
      return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(errS) })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // LLM connectivity ping
  if (e && e.parameter && e.parameter.action === 'llmPing') {
    try {
      var provider = (e.parameter.provider || '').toLowerCase();
      var result;
      if (provider === 'deepseek') {
        result = (typeof debugDeepSeekPing === 'function') ? debugDeepSeekPing() : { ok:false, error:'debugDeepSeekPing not available' };
      } else if (provider === 'openrouter') {
        result = (typeof debugOpenRouterPing === 'function') ? debugOpenRouterPing() : { ok:false, error:'debugOpenRouterPing not available' };
      } else {
        // Auto-select based on configured provider
        try {
          var p = (typeof _llmProvider_ === 'function') ? _llmProvider_() : 'openrouter';
          if (p === 'deepseek') {
            result = (typeof debugDeepSeekPing === 'function') ? debugDeepSeekPing() : { ok:false, error:'debugDeepSeekPing not available' };
          } else if (p === 'openrouter') {
            result = (typeof debugOpenRouterPing === 'function') ? debugOpenRouterPing() : { ok:false, error:'debugOpenRouterPing not available' };
          } else {
            result = { ok:false, error:'No LLM configured' };
          }
        } catch(_autoErr){
          // Fallback
          result = (typeof debugDeepSeekPing === 'function') ? debugDeepSeekPing() : (
            (typeof debugOpenRouterPing === 'function') ? debugOpenRouterPing() : { ok:false, error:'No LLM ping available' }
          );
        }
      }
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err) })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // Ensure spreadsheet timezone is Tanzania (EAT)
  try { ensureSpreadsheetTZ(); } catch(_tz) {}

  var tpl = HtmlService.createTemplateFromFile('Index.html');
  var out = tpl.evaluate()
    .setTitle('Fund Request — Split-Flap Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  try {
    var pti = getProjectTeamIndex(); // warm path (cached)
    var ddc = getDDClientPack();     // compact rows (cached)

    var html = out.getContent();
    function safeScript(id, obj){
      var json = JSON.stringify(obj).replace(/<\/script>/ig,'<\\/script>');
      return '<script type="application/json" id="'+ id +'">'+ json +'</script>';
    }
    var tags = safeScript('PTI', pti) + '\n' + safeScript('DDC', ddc);

    if (html.indexOf('</body>') !== -1) {
      html = html.replace('</body>', tags + '\n</body>');
    } else {
      html += tags;
    }
    out.setContent(html);
  } catch (e) {
    console.error('Injection failed: ' + e);
  }

  return out;
}

/* Utility to stream HTML fragments */
function getFragment(name){
  Logger.log('getFragment called with: ' + name);
  
  // Validate that name is defined and not empty
  if (!name || name === undefined || name === null || name === '') {
    Logger.log('ERROR: getFragment called with invalid name: ' + name);
    return '<!-- ERROR: getFragment called with invalid name: ' + name + ' -->';
  }
  
  return include(name);
}

/* -------------------------- ensure Sheet1 header -------------------------- */

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const header = [
    'Timestamp',
    'Beneficiary',
    'Account Holder',
    'Row Total',
    'Designation',
    'Fuel From','Fuel To','Fuel Amt',
  'DA From','DA To','DA Amt',
    'Car From','Car To','Car Amt',
    'Airtime From','Airtime To','Airtime Amt',
    'Transport From','Transport To','Transport Amt',
    'Misc From','Misc To','Misc Amt',
    'Mob No',
    'Display Name',
    'W/H Charges',
    'Remarks'
  ];

  const rng = sh.getRange(1, 1, 1, header.length);
  const values = rng.getValues()[0];
  const needsHeader = values.every(v => v === '' || v === null);
  if (needsHeader) {
    rng.setValues([header]);
    sh.setFrozenRows(1);
  } else {
    if (sh.getLastColumn() < header.length) {
      sh.insertColumnsAfter(sh.getLastColumn(), header.length - sh.getLastColumn());
    }
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sh;
}

/* -------------------------- save (unchanged behavior) -------------------------- */

function saveBatch(rows) {
  if (!Array.isArray(rows)) return { ok: false, error: 'Invalid payload' };
  const validMob = v => typeof v === 'string' && /^0\d{9}$/.test(v);
  
  // Helper function to convert DD/MM/YYYY string to Date object for calculations
     function parseDate(dateStr) {
       if (!dateStr || dateStr.trim() === '') return '';
       const parts = dateStr.split('/');
       if (parts.length === 3) {
         const day = parseInt(parts[0], 10);
         const month = parseInt(parts[1], 10);
         const year = parseInt(parts[2], 10);
         if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
           // Return Date object for Google Sheets calculations
           return new Date(year, month - 1, day); // month is 0-indexed in Date constructor
         }
       }
       return dateStr; // Return original string if parsing fails
     }

  const lock = LockService.getDocumentLock(); // Keeping this as-is per your choice. (ScriptLock is an optional later tweak)
  lock.waitLock(30000);
  try {
    const sh = getOrCreateSheet_();
    // Create Tanzania timezone timestamp as Date object for proper formatting
    // Use proper timezone conversion for East Africa Time (UTC+3)
    const now = new Date();
    const utcTime = now.getTime() + (now.getTimezoneOffset() * 60000);
    const tanzaniaTime = new Date(utcTime + (3 * 60 * 60 * 1000));
    const timestamp = tanzaniaTime; // Store as Date object to preserve correct formatting
    
    const data = rows.map(r => ([
      timestamp,
      (r.beneficiary || '').toString().trim(),
      (r.accountHolder || '').toString().trim(),
      Number(r.total || 0),
      (r.designation || '').toString().trim(),
      parseDate(r.fuel?.from || ''),   parseDate(r.fuel?.to || ''),   Number(r.fuel?.amount || 0),
  parseDate(r.da?.from || ''),   parseDate(r.da?.to || ''),   Number(r.da?.amount || 0),
      parseDate(r.car?.from  || ''),   parseDate(r.car?.to  || ''),   Number(r.car?.amount  || 0),
      parseDate(r.air?.from  || ''),   parseDate(r.air?.to  || ''),   Number(r.air?.amount  || 0),
      parseDate(r.transport?.from || ''), parseDate(r.transport?.to || ''), Number(r.transport?.amount || 0),
      parseDate(r.misc?.from || ''),   parseDate(r.misc?.to || ''),   Number(r.misc?.amount || 0),
      (validMob(r.mob) ? r.mob : (r.mob || '')),
      (r.displayName || '').toString().trim(),
      Number(r.whCharges || 0),
      (r.remarks || '').toString().trim()
    ]));
    if (data.length === 0) return { ok: false, error: 'No rows to save' };
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    return { ok: true, added: data.length, lastRow: startRow + data.length - 1 };
  } catch (e) {
    return { ok: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Submit data to submissions tab with validation
 * Only submits if there are no violations
 * Submits all teams regardless of filtering
 */
function submitToSubmissions(submissionData) {
  if (!submissionData || !Array.isArray(submissionData.rows)) {
    return { ok: false, error: 'Invalid submission data' };
  }

  // Check for violations
  if (submissionData.hasViolations) {
    return { ok: false, error: 'Cannot submit: There are validation violations that must be resolved first' };
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
    
    // Create submissions sheet if it doesn't exist
    if (!sh) {
      sh = ss.insertSheet(SUBMISSIONS_SHEET_NAME);
      // Set up header for submissions sheet
      const header = [
        'Timestamp',
        'Submission ID',
        'Beneficiary',
        'Account Holder',
        'Team Name',
        'Project Name',
        'Row Total',
        'Designation',
        'Fuel From','Fuel To','Fuel Amt',
  'DA From','DA To','DA Amt',
        'Car From','Car To','Car Amt',
        'Airtime From','Airtime To','Airtime Amt',
        'Transport From','Transport To','Transport Amt',
        'Misc From','Misc To','Misc Amt',
        'Mob No',
        'Display Name',
        'W/H Charges',
        'Remarks'
      ];
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
    }

    // Generate unique submission ID
    const submissionId = 'SUB_' + new Date().getTime();
    // Set timestamp to Tanzania timezone (EAT - UTC+3) as Date object
    // Use proper timezone conversion for East Africa Time (UTC+3)
    const now = new Date();
    const utcTime = now.getTime() + (now.getTimezoneOffset() * 60000);
    const tanzaniaTime = new Date(utcTime + (3 * 60 * 60 * 1000));
    const timestamp = tanzaniaTime; // Store as Date object to preserve correct formatting
    
    // Helper function to convert DD/MM/YYYY string to Date object for calculations
     function parseDate(dateStr) {
       if (!dateStr || dateStr.trim() === '') return '';
       const parts = dateStr.split('/');
       if (parts.length === 3) {
         const day = parseInt(parts[0], 10);
         const month = parseInt(parts[1], 10);
         const year = parseInt(parts[2], 10);
         if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
           // Return Date object for Google Sheets calculations
           return new Date(year, month - 1, day); // month is 0-indexed in Date constructor
         }
       }
       return dateStr; // Return original string if parsing fails
     }
    
    // Prepare data for insertion
    const data = submissionData.rows.map(r => ([
      timestamp,
      submissionId,
      (r.beneficiary || '').toString().trim(),
      (r.accountHolder || '').toString().trim(),
      (r.teamName || r.team || '').toString().trim(),
      (r.projectName || r.project || '').toString().trim(),
      Number(r.total || 0),
      (r.designation || '').toString().trim(),
      parseDate(r.fuel?.from || ''),   parseDate(r.fuel?.to || ''),   Number(r.fuel?.amount || 0),
  parseDate(r.da?.from || ''),   parseDate(r.da?.to || ''),   Number(r.da?.amount || 0),
      parseDate(r.car?.from  || ''),   parseDate(r.car?.to  || ''),   Number(r.car?.amount  || 0),
      parseDate(r.air?.from  || ''),   parseDate(r.air?.to  || ''),   Number(r.air?.amount  || 0),
      parseDate(r.transport?.from || ''), parseDate(r.transport?.to || ''), Number(r.transport?.amount || 0),
      parseDate(r.misc?.from || ''),   parseDate(r.misc?.to || ''),   Number(r.misc?.amount || 0),
      (r.mob || '').toString().trim(),
      (r.displayName || '').toString().trim(),
      Number(r.whCharges || 0),
      (r.remarks || '').toString().trim()
    ]));

    if (data.length === 0) {
      return { ok: false, error: 'No rows to submit' };
    }

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    
    return { 
      ok: true, 
      submissionId: submissionId,
      submitted: data.length, 
      lastRow: startRow + data.length - 1 
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

/**
  function createVehicleReleasedOnChangeTrigger() {
    try {
      removeVehicleReleasedOnChangeTrigger();
      ScriptApp.newTrigger('onVehicleReleasedChange')
        .forSpreadsheet(SHEET_ID)
        .onChange()
        .create();
      return { ok: true, message: 'Vehicle_Released onChange trigger created' };
    } catch (error) {
      console.error('Failed to create Vehicle_Released onChange trigger:', error);
      return { ok: false, error: String(error) };
    }
  }

  function removeVehicleReleasedOnChangeTrigger() {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'onVehicleReleasedChange') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      return { ok: true, message: 'Removed Vehicle_Released onChange triggers' };
    } catch (error) {
      console.error('Failed to remove Vehicle_Released onChange triggers:', error);
      return { ok: false, error: String(error) };
    }
  }
 * Build a latest Ops (Ops_P + Ops_T) mapping keyed by account holder.
 * Internal helper; returns JSON-safe data for caching/serialization.
 */
function _buildAllLatestOpsMap_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const state = Object.create(null);
  const stats = {
    permanentRows: 0,
    temporaryRows: 0,
    sources: []
  };

  function ingest(sheetName, changeType) {
    let sh;
    try {
      sh = ss.getSheetByName(sheetName);
    } catch (err) {
      console.warn('ingest ops sheet failed:', sheetName, err);
      return;
    }
    if (!sh) return;
    if (stats.sources.indexOf(sheetName) === -1) stats.sources.push(sheetName);

    const values = sh.getDataRange().getValues();
    if (!values || values.length <= 1) return;

    const rowCount = values.length - 1;
    if (changeType === 'Permanent') {
      stats.permanentRows += rowCount;
    } else if (changeType === 'Temporary') {
      stats.temporaryRows += rowCount;
    }

    const header = values[0].map(function(h){
      return String(h || '').trim().toLowerCase();
    });

    function findIndex(patterns) {
      for (var i = 0; i < header.length; i++) {
        var col = header[i];
        if (!col) continue;
        for (var j = 0; j < patterns.length; j++) {
          if (col.indexOf(patterns[j]) !== -1) {
            return i;
          }
        }
      }
      return -1;
    }

    const idx = {
      timestamp: findIndex(['timestamp', 'date']),
      accountHolder: findIndex(['account holder', 'acc holder', 'account_holder']),
      mobNo: findIndex(['mobile number', 'mobile no', 'mob no', 'mob num', 'mobile']),
      displayName: findIndex(['display name', 'displayname']),
      operator: findIndex(['operator']),
      submitter: findIndex(['submitter', 'user', 'email', 'entered by'])
    };

    const rows = values.slice(1);
    rows.forEach(function(row){
      var acc = idx.accountHolder >= 0 ? String(row[idx.accountHolder] || '').trim() : '';
      if (!acc) return;
      var key = acc.toLowerCase();

      var tsVal = idx.timestamp >= 0 ? row[idx.timestamp] : null;
      var ts = 0;
      if (tsVal instanceof Date) {
        ts = tsVal.getTime();
      } else if (tsVal) {
        var parsed = new Date(tsVal);
        if (!isNaN(parsed.getTime())) ts = parsed.getTime();
      }

      var entry = state[key];
      if (!entry || ts >= entry.timestampMs) {
        state[key] = {
          accountHolder: acc,
          mobNo: idx.mobNo >= 0 ? String(row[idx.mobNo] || '').trim() : '',
          displayName: idx.displayName >= 0 ? String(row[idx.displayName] || '').trim() : '',
          operator: idx.operator >= 0 ? String(row[idx.operator] || '').trim() : '',
          submitter: idx.submitter >= 0 ? String(row[idx.submitter] || '').trim() : '',
          changeType: changeType,
          sheet: sheetName,
          timestampMs: ts
        };
      }
    });
  }

  ingest('Ops_P', 'Permanent');
  ingest('Ops_T', 'Temporary');

  const map = {};
  const list = [];
  Object.keys(state).forEach(function(key){
    const entry = state[key];
    if (!entry) return;
    const iso = entry.timestampMs ? new Date(entry.timestampMs).toISOString() : '';
    const payload = {
      accountHolder: entry.accountHolder,
      mobNo: entry.mobNo,
      displayName: entry.displayName,
      operator: entry.operator,
      submitter: entry.submitter,
      changeType: entry.changeType,
      sourceSheet: entry.sheet,
      timestamp: iso
    };
    map[entry.accountHolder] = {
      mobNo: entry.mobNo,
      displayName: entry.displayName,
      operator: entry.operator,
      submitter: entry.submitter,
      changeType: entry.changeType,
      sourceSheet: entry.sheet,
      timestamp: iso
    };
    list.push(payload);
  });

  list.sort(function(a, b){
    return a.accountHolder.localeCompare(b.accountHolder);
  });

  return {
    ok: true,
    map: map,
    list: list,
    stats: stats,
    generated: new Date().toISOString(),
    total: list.length
  };
}

/**
 * Public entry-point consumed by the frontend. Wraps the Ops map builder in script cache logic.
 */
function getAllLatestOpsMapCached(options) {
  try {
    var forceRefresh = false;
    if (options === true) {
      forceRefresh = true;
    } else if (options && typeof options === 'object') {
      forceRefresh = !!(options.force || options.refresh || options.forceRefresh);
    }

    var cacheKey = 'ops:latest-map:v1';
    var cache;
    try {
      cache = CacheService.getScriptCache();
    } catch (cacheErr) {
      console.warn('CacheService unavailable for getAllLatestOpsMapCached:', cacheErr);
    }

    if (!forceRefresh && cache) {
      var cached = cache.get(cacheKey);
      if (cached) {
        try {
          var parsed = JSON.parse(cached);
          parsed.cached = true;
          parsed.refreshed = false;
          return parsed;
        } catch (parseErr) {
          console.warn('Failed to parse cached ops map:', parseErr);
        }
      }
    }

    var result = _buildAllLatestOpsMap_();
    if (cache) {
      try {
        cache.put(cacheKey, JSON.stringify(result), 300); // cache for 5 minutes
      } catch (putErr) {
        console.warn('Failed to cache ops map:', putErr);
      }
    }
    result.cached = false;
    result.refreshed = true;
    return result;
  } catch (err) {
    console.error('getAllLatestOpsMapCached error:', err);
    return { ok: false, error: String(err) };
  }
}

/**
 * Return unique operator names from Ops_P column K (K1, K2, ...).
 * If the sheet or column is empty, returns an empty array.
 */
function getOpsOperators() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('Ops_P');
    if (!sh) return [];
    var lastRow = sh.getLastRow();
    if (lastRow < 1) return [];
    // Column K is index 11
    var vals = sh.getRange(1, 11, lastRow, 1).getDisplayValues()
      .map(function(r){ return String(r[0] || '').trim(); })
      .filter(function(v){ return v && v.length; });
    // Unique preserve order
    var seen = Object.create(null);
    var out = [];
    vals.forEach(function(v){ if(!seen[v]){ seen[v]=true; out.push(v); } });
    return out;
  } catch (e) {
    return [];
  }
}

/**
 * Get current user identity details for display/logging.
 */
function getCurrentUser() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  var username = '';
  try {
    username = email ? String(email).split('@')[0] : '';
  } catch(_e) {}
  return { email: email, username: username };
}

/**
 * Save a Mobile Money Number change request into Ops_P (permanent) or Ops_T (temporary).
 * Payload: { type: 'Permanent'|'Temporary', accountHolder: string, mobNo: string, displayName: string, operator: string }
 */
function saveOpsChange(payload) {
  try {
    if (!payload || typeof payload !== 'object') return { ok:false, error:'Invalid payload' };
    var type = String(payload.type || '').toLowerCase();
    var sheetName = (type === 'temporary' || type === 'temp') ? 'Ops_T' : 'Ops_P';

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.appendRow(['Timestamp', 'Account Holder', 'Mobile Number', 'Display Name', 'Operator', 'Submitter']);
      sh.setFrozenRows(1);
    }

    // Tanzania time as Date object
    var now = new Date();
    var utcTime = now.getTime() + (now.getTimezoneOffset() * 60000);
    var tzDate = new Date(utcTime + (3 * 60 * 60 * 1000)); // UTC+3

    var userEmail = '';
    try { userEmail = Session.getActiveUser().getEmail() || ''; } catch(e) {}

    var accHolder = (payload.accountHolder || '').toString().trim();
    var mobNo = (payload.mobNo || '').toString().trim();
    var displayName = (payload.displayName || '').toString().trim();
    var operator = (payload.operator || '').toString().trim();

    sh.appendRow([tzDate, accHolder, mobNo, displayName, operator, userEmail]);
    return { ok:true };
  } catch (e) {
    return { ok:false, error:String(e) };
  }
}

/**
 * Submit car release data to CarT_P sheet
 * Enhanced version with better validation and error handling
 */
function submitCarRelease(releaseData) {
  // Input validation
  if (!releaseData) {
    return { ok: false, error: 'No release data provided' };
  }
  
  // Validate required fields
  const requiredFields = ['carNumber', 'remarks', 'stars'];
  for (const field of requiredFields) {
    if (!releaseData[field] || (field === 'remarks' && releaseData[field].toString().trim().length < 10)) {
      return { ok: false, error: `Invalid or missing ${field}` };
    }
  }
  
  // Validate stars rating
  const stars = Number(releaseData.stars);
  if (isNaN(stars) || stars < 1 || stars > 5) {
    return { ok: false, error: 'Stars rating must be between 1 and 5' };
  }

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'System is busy, please try again' };
    }
    
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    let sh = ss.getSheetByName('CarT_P');
    
    // Create CarT_P sheet if it doesn't exist
    if (!sh) {
      sh = ss.insertSheet('CarT_P');
      // Set up header for CarT_P sheet (matches new structure with Responsible Beneficiary column)
      const header = [
        'Ref',
        'Date and time of entry',
        'Project',
        'Team',
        'R.Beneficiary',
        'Vehicle Number',
        'Make',
        'Model',
        'Category',
        'Usage Type',
        'Owner',
        'Status',
        'Last Users remarks',
        'Ratings',
        'Submitter username',
        'R.Ben Time'
      ];
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);

      // Format header row
      const headerRange = sh.getRange(1, 1, 1, header.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
    }

    // Generate unique reference number with better format
    const timestamp = Date.now();
    const randomSuffix = Math.random().toString(36).substr(2, 5).toUpperCase();
    const refNumber = `CAR-${timestamp}-${randomSuffix}`;
    
    // Preserve spreadsheet timezone (ensureSpreadsheetTZ handles defaults)
    const tanzaniaTime = new Date();
    
    // Get current user email for better tracking
    let submitter = releaseData.submitter || '';
    try {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        submitter = userEmail;
      }
    } catch (e) {
      // Fallback to provided submitter or 'System'
      submitter = submitter || 'System';
    }

    // Build a unique, ordered list of team members involved in this release
    const teamMemberSet = new Set();
    const teamMembersList = [];
    const addTeamMember = (name) => {
      const trimmed = String(name || '').trim();
      if (!trimmed) return;
      const key = trimmed.toLowerCase();
      if (teamMemberSet.has(key)) return;
      teamMemberSet.add(key);
      teamMembersList.push(trimmed);
    };

    addTeamMember(releaseData.responsibleBeneficiary);
    if (Array.isArray(releaseData.teamMembersUsing) && releaseData.teamMembersUsing.length) {
      releaseData.teamMembersUsing.forEach(addTeamMember);
    } else if (Array.isArray(releaseData.teamMembers) && releaseData.teamMembers.length) {
      releaseData.teamMembers.forEach(addTeamMember);
    } else if (Array.isArray(releaseData.lastUsers) && releaseData.lastUsers.length) {
      releaseData.lastUsers.forEach(addTeamMember);
    } else if (typeof releaseData.lastUsers === 'string' && releaseData.lastUsers) {
      releaseData.lastUsers.split(/[,;\n]+/).forEach(addTeamMember);
    }

    const teamMembersString = teamMembersList.join(', ');
    const primaryBeneficiary = teamMembersList.length
      ? teamMembersList[0]
      : String(releaseData.responsibleBeneficiary || '').trim();
    const shortBeneficiaryValue = primaryBeneficiary ? primaryBeneficiary : '';
    const beneficiaryRows = teamMembersList.length
      ? teamMembersList.slice()
      : (primaryBeneficiary ? [primaryBeneficiary] : []);
    const defaultBeneficiaryText = teamMembersList.length
      ? teamMembersString
      : primaryBeneficiary;
    releaseData.responsibleBeneficiary = primaryBeneficiary;
    releaseData.teamMembers = beneficiaryRows;

    // Ensure Responsible Beneficiary column exists even on legacy layouts
    (function ensureResponsibleColumn(){
      try {
        const existingHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
        const hasColumn = existingHead.some(function(h){
          const norm = String(h||'').trim().toLowerCase().replace(/[^a-z]/g,'');
          return norm === 'rbeneficiary' || norm === 'responsiblebeneficiary';
        });
        if (hasColumn) return;
        const teamIdx = existingHead.findIndex(function(h){
          const norm = String(h||'').trim().toLowerCase();
          return norm === 'team' || norm === 'team name';
        });
        if (teamIdx >= 0) {
          sh.insertColumnAfter(teamIdx + 1);
          sh.getRange(1, teamIdx + 2).setValue('R.Beneficiary');
        } else {
          const lastCol = sh.getLastColumn();
          sh.insertColumnAfter(lastCol);
          sh.getRange(1, lastCol + 1).setValue('R.Beneficiary');
        }
      } catch (err) {
        console.warn('submitCarRelease: ensure R.Beneficiary column failed', err);
      }
    })();

    (function ensureShortResponsibleColumn(){
      try {
        const existingHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
        const normalized = existingHead.map(function(h){
          return String(h || '').trim().toLowerCase().replace(/[^a-z]/g, '');
        });
        if (normalized.indexOf('rben') >= 0) {
          return;
        }
        const timeIdx = normalized.indexOf('rbentime');
        if (timeIdx >= 0) {
          sh.insertColumnAfter(timeIdx + 1);
          sh.getRange(1, timeIdx + 2).setValue('R. Ben');
          return;
        }
        const respIdx = normalized.indexOf('rbeneficiary');
        if (respIdx >= 0) {
          sh.insertColumnAfter(respIdx + 1);
          sh.getRange(1, respIdx + 2).setValue('R. Ben');
          return;
        }
        const lastCol = sh.getLastColumn();
        sh.insertColumnAfter(lastCol);
        sh.getRange(1, lastCol + 1).setValue('R. Ben');
      } catch (err) {
        console.warn('submitCarRelease: ensure R. Ben column failed', err);
      }
    })();

    // Prepare data for insertion using header-aware mapping
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    const rowValues = new Array(header.length).fill('');

    const rawTeamIn = String((releaseData.teamName || releaseData.team || '')).trim();
    const rawProjectIn = String((releaseData.projectName || releaseData.project || '')).trim();
    const carUpper = String(releaseData.carNumber || '').trim().toUpperCase();

    let cachedValues = null;
    let cachedLastRow = null;

    function safeIdx(labels) {
      try { return IX.get(labels); } catch(_e) { return -1; }
    }

    function resolveContext(rawValue, labelSet) {
      const cleaned = String(rawValue || '').trim();
      if (cleaned && !/^unknown\b/i.test(cleaned)) return cleaned;
      const carIdx = safeIdx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car']);
      const targetIdx = safeIdx(labelSet);
      if (carIdx < 0 || targetIdx < 0 || !carUpper) return cleaned;
      const lastRow = cachedLastRow != null ? cachedLastRow : sh.getLastRow();
      cachedLastRow = lastRow;
      if (lastRow <= 1) return cleaned;
      try {
        if (!cachedValues) {
          cachedValues = sh.getRange(2, 1, lastRow - 1, header.length).getDisplayValues();
        }
        const values = cachedValues;
        for (let i = values.length - 1; i >= 0; i--) {
          const carCell = String(values[i][carIdx] || '').trim().toUpperCase();
          if (!carCell || carCell !== carUpper) continue;
          const cellValue = String(values[i][targetIdx] || '').trim();
          if (cellValue && !/^unknown\b/i.test(cellValue)) return cellValue;
        }
      } catch (lookupErr) {
        console.warn('resolveContext fallback failed:', lookupErr);
      }
      return /^unknown\b/i.test(cleaned) ? '' : cleaned;
    }

    const resolvedProject = resolveContext(rawProjectIn, ['Project','Project Name']);
    const resolvedTeam = resolveContext(rawTeamIn, ['Team','Team Name']);
    const finalProject = resolvedProject || (/^unknown\b/i.test(rawProjectIn) ? '' : rawProjectIn);
    const finalTeam = resolvedTeam || (/^unknown\b/i.test(rawTeamIn) ? '' : rawTeamIn);
    const finalTeamLower = finalTeam ? finalTeam.toLowerCase() : '';

    const remarksValue = (releaseData.remarks || '').toString().trim();

    function assign(labels, value, required) {
      try {
        const idx = IX.get(labels);
        rowValues[idx] = value;
      } catch (err) {
        if (required) throw err;
        // optional column missing – ignore
      }
    }

    assign(['Reference Number','Ref','Ref Number'], refNumber, true);
    assign(['Date and time of entry','Date and time','Timestamp','Date'], tanzaniaTime, true);
    assign(['Project'], finalProject, false);
    assign(['Team','Team Name'], finalTeam, false);
    assign(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], defaultBeneficiaryText, false);
    assign(['R. Ben','R Ben'], shortBeneficiaryValue, false);
    assign(['Vehicle Number','Car Number','Vehicle No','Car No','Car #','Car'], (releaseData.carNumber || '').toString().trim().toUpperCase(), true);
    assign(['Make','Car Make','Brand'], (releaseData.make || '').toString().trim(), false);
    assign(['Model','Car Model'], (releaseData.model || '').toString().trim(), false);
    assign(['Category','Vehicle Category','Cat','Category Name'], (releaseData.category || '').toString().trim(), false);
    assign(['Usage Type','Usage','Use Type'], (releaseData.usageType || '').toString().trim(), false);
    assign(['Contract Type','Contract','Agreement Type'], (releaseData.contractType || '').toString().trim(), false);
    assign(['Owner','Owner Name','Owner Info'], (releaseData.owner || '').toString().trim(), false);
    assign(['In Use/Release','In Use / release','In Use','Status'], 'RELEASE', false);
    assign(['Last Users remarks','Remarks','Feedback'], remarksValue, false);
    assign(['Stars','Ratings','Rating'], stars, false);
    assign(['Submitter username','Submitter','User'], submitter.toString().trim(), false);
    assign(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], tanzaniaTime, false);

    // Update existing IN USE rows for this vehicle/team to mark all team members as released
    const idxRefCol = safeIdx(['Reference Number','Ref','Ref Number']);
    const idxVehicleCol = safeIdx(['Vehicle Number','Car Number','Vehicle No','Car No','Car #','Car']);
    const idxStatusCol = safeIdx(['In Use/Release','In Use / release','In Use','Status']);
    const idxTeamCol = safeIdx(['Team','Team Name']);
    const idxRemarksCol = safeIdx(['Last Users remarks','Remarks','Feedback']);
    const idxStarsCol = safeIdx(['Stars','Ratings','Rating']);
    const idxSubmitCol = safeIdx(['Submitter username','Submitter','User']);
    const idxRespTimeCol = safeIdx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time']);
    const idxRBCol = safeIdx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible']);
    const idxRBenShortCol = safeIdx(['R. Ben','R Ben']);
    const shortBeneficiaryLower = shortBeneficiaryValue ? shortBeneficiaryValue.toLowerCase() : '';

    const lastRowExisting = sh.getLastRow();
    if (lastRowExisting > 1 && idxVehicleCol >= 0 && idxStatusCol >= 0) {
      const existingCount = lastRowExisting - 1;
      const dataRange = sh.getRange(2, 1, existingCount, header.length);
      const existingValues = dataRange.getValues();
      const existingDisplay = dataRange.getDisplayValues();
      const updates = [];

      for (let r = 0; r < existingValues.length; r++) {
        const rowValue = existingValues[r];
        const displayValue = existingDisplay[r];
        const rowCar = String(rowValue[idxVehicleCol] || displayValue[idxVehicleCol] || '').trim().toUpperCase();
        if (!rowCar || rowCar !== carUpper) continue;

        const statusRaw = rowValue[idxStatusCol] || displayValue[idxStatusCol] || '';
        const normalizedStatus = _normStatus_(statusRaw);
        if (normalizedStatus !== 'IN USE') continue;

        if (idxTeamCol >= 0 && finalTeamLower) {
          const rowTeam = String(rowValue[idxTeamCol] || displayValue[idxTeamCol] || '').trim().toLowerCase();
          if (rowTeam && rowTeam !== finalTeamLower) continue;
        }

        const existingRB = idxRBCol >= 0
          ? String(rowValue[idxRBCol] || displayValue[idxRBCol] || '').trim()
          : '';
        const existingShort = idxRBenShortCol >= 0
          ? String(rowValue[idxRBenShortCol] || displayValue[idxRBenShortCol] || '').trim()
          : '';
        const isResponsibleRow = !!shortBeneficiaryLower && (
          (existingRB && existingRB.toLowerCase() === shortBeneficiaryLower) ||
          (existingShort && existingShort.toLowerCase() === shortBeneficiaryLower)
        );

        const updated = rowValue.slice();
        updated[idxStatusCol] = 'RELEASE';
        if (idxRBCol >= 0 && defaultBeneficiaryText) {
          if (!existingRB) {
            updated[idxRBCol] = defaultBeneficiaryText;
          }
        }
        if (idxRBenShortCol >= 0) {
          if (isResponsibleRow) {
            if (!existingShort && shortBeneficiaryValue) {
              updated[idxRBenShortCol] = shortBeneficiaryValue;
            }
          } else if (shortBeneficiaryLower) {
            updated[idxRBenShortCol] = '';
          }
        }
        if (idxRemarksCol >= 0 && remarksValue) {
          updated[idxRemarksCol] = remarksValue;
        }
        if (idxStarsCol >= 0) {
          updated[idxStarsCol] = stars;
        }
        if (idxSubmitCol >= 0 && submitter) {
          updated[idxSubmitCol] = submitter;
        }
        if (idxRespTimeCol >= 0) {
          updated[idxRespTimeCol] = tanzaniaTime;
        }
        updates.push({ rowNumber: r + 2, values: updated });
      }

      for (let i = 0; i < updates.length; i++) {
        const update = updates[i];
        sh.getRange(update.rowNumber, 1, 1, header.length).setValues([update.values]);
      }
    }

    const beneficiaryTargets = beneficiaryRows.length
      ? beneficiaryRows
      : (defaultBeneficiaryText ? [defaultBeneficiaryText] : []);
    const rowsToInsert = beneficiaryTargets.map(function(name, index) {
      const rowCopy = rowValues.slice();
      if (idxRBCol >= 0) {
        if (index === 0) {
          rowCopy[idxRBCol] = defaultBeneficiaryText || name || '';
        } else {
          rowCopy[idxRBCol] = name || defaultBeneficiaryText || '';
        }
      }
      if (idxRBenShortCol >= 0) {
        rowCopy[idxRBenShortCol] = index === 0 ? (shortBeneficiaryValue || defaultBeneficiaryText || '') : '';
      }
      if (idxRefCol >= 0 && index > 0) {
        rowCopy[idxRefCol] = refNumber + '-' + (index + 1);
      }
      if (idxRespTimeCol >= 0 && index > 0) {
        rowCopy[idxRespTimeCol] = '';
      }
      return rowCopy;
    });
    if (!rowsToInsert.length) {
      rowsToInsert.push(rowValues.slice());
    }

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rowsToInsert.length, rowValues.length).setValues(rowsToInsert);

    try {
      const summarySource = rowsToInsert[0] || rowValues;
      const summaryObj = {};
      for (let i = 0; i < VEHICLE_SUMMARY_HEADER.length; i++) {
        summaryObj[VEHICLE_SUMMARY_HEADER[i]] = summarySource[i] || '';
      }
      summaryObj.status = summaryObj.Status || 'RELEASE';
      upsertVehicleSummaryRow('Vehicle_Released', summaryObj, 'vehicle');
    } catch (summaryErr) {
      console.warn('submitCarRelease: Vehicle_Released summary update failed', summaryErr);
    }

    // Auto-resize columns for better readability
    sh.autoResizeColumns(1, rowValues.length);

    try {
      refreshVehicleStatusSheets();
    } catch (refreshErr) {
      console.error('Vehicle summary refresh (release) failed:', refreshErr);
    }

    try {
      syncVehicleSheetFromCarTP();
    } catch (syncErr) {
      console.error('Vehicle sheet sync (release) failed:', syncErr);
    }

    try { CacheService.getScriptCache().remove('VEH_PICKER_V1'); } catch (_cacheClearErr) { /* cache purge best-effort */ }
    invalidateVehicleInUseCache();
    invalidateVehicleReleasedCache('Vehicle release persisted to CarT_P');

    return {
      ok: true,
      refNumber: refNumber,
      submitted: rowsToInsert.length,
      lastRow: startRow,
      timestamp: tanzaniaTime.toISOString()
    };
  } catch (e) {
    console.error('Car release submission error:', e);
    return { 
      ok: false, 
      error: `Submission failed: ${e.message || String(e)}` 
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn('Failed to release lock:', e);
    }
  }
}

/**
 * Release a single user from an in-use vehicle entry.
 * When this was the last active user, falls back to full vehicle release to trigger feedback capture.
 */
function releaseCarUser(payload) {
  if (!payload) {
    return { ok: false, error: 'No release data provided' };
  }

  const carNumber = String(payload.carNumber || '').trim().toUpperCase();
  if (!carNumber) {
    return { ok: false, error: 'Missing car number' };
  }

  const requestedUsersRaw = [];
  if (Array.isArray(payload.users)) {
    payload.users.forEach((user) => requestedUsersRaw.push(user));
  }
  if (payload.user || payload.releasedUser) {
    requestedUsersRaw.push(payload.user || payload.releasedUser);
  }

  const requestedUsers = requestedUsersRaw
    .map((name) => String(name || '').trim())
    .filter(Boolean);

  if (!requestedUsers.length) {
    return { ok: false, error: 'Missing user to release' };
  }

  const uniqueTargets = new Map();
  requestedUsers.forEach((name) => {
    const key = name.toLowerCase();
    if (!key) return;
    if (!uniqueTargets.has(key)) {
      uniqueTargets.set(key, name);
    }
  });

  if (!uniqueTargets.size) {
    return { ok: false, error: 'Missing user to release' };
  }

  const lock = LockService.getDocumentLock();
  let lockHeld = false;
  try {
    if (!lock.tryLock(30000)) {
      return { ok: false, error: 'System is busy, please try again' };
    }
    lockHeld = true;

    const sh = _openCarTP_();
    if (!sh) {
      return { ok: false, error: 'CarT_P sheet not found' };
    }

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      return { ok: false, error: 'CarT_P sheet has no data' };
    }

    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    const idx = (labels, required) => {
      try {
        return IX.get(labels);
      } catch (e) {
        if (required) throw e;
        return -1;
      }
    };

    const iCar = idx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car'], true);
    const iStatus = idx(['In Use/Release','In Use / release','In Use','Status'], true);
    const iBeneficiary = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], true);
    const iBeneficiaryShort = idx(['R. Ben','R Ben'], false);
    const iDate = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    const iRemarks = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars = idx(['Ratings','Stars','Rating'], false);
    const iSubmit = idx(['Submitter username','Submitter','User'], false);
    const iRespTime = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);
    const iRef = idx(['Reference Number','Ref','Ref Number'], false);
    const iProject = idx(['Project'], false);
    const iTeam = idx(['Team','Team Name'], false);
    const iMake = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iCategory = idx(['Category','Vehicle Category','Cat','Category Name'], false);
    const iUsage = idx(['Usage Type','Usage','Use Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);

    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const rows = dataRange.getValues();
    const display = dataRange.getDisplayValues();

    const activeUsers = new Map();
    for (let r = rows.length - 1; r >= 0; r--) {
      const row = rows[r];
      const carValue = String(row[iCar] || display[r][iCar] || '').trim().toUpperCase();
      if (!carValue || carValue !== carNumber) continue;

      const statusValue = String(row[iStatus] || display[r][iStatus] || '').trim().toUpperCase();
      const beneficiaryRaw = row[iBeneficiary] || display[r][iBeneficiary] || '';
      const beneficiary = String(beneficiaryRaw || '').trim();
      const beneficiaryKey = beneficiary.toLowerCase();
      if (!beneficiaryKey) continue;

      if (statusValue === 'IN USE') {
        if (!activeUsers.has(beneficiaryKey)) {
          activeUsers.set(beneficiaryKey, { name: beneficiary || uniqueTargets.get(beneficiaryKey) || '', indices: [] });
        }
        activeUsers.get(beneficiaryKey).indices.push(r);
      }
    }

    const releaseTargets = [];
    const missingUsers = [];
    uniqueTargets.forEach((name, key) => {
      const info = activeUsers.get(key);
      if (info && Array.isArray(info.indices) && info.indices.length) {
        releaseTargets.push({
          key,
          requestedName: name,
          displayName: info.name || name,
          indices: info.indices
        });
      } else {
        missingUsers.push(name);
      }
    });

    if (!releaseTargets.length) {
      return { ok: false, error: 'The selected user(s) are not currently assigned to this vehicle' };
    }

    const remainingUsersKeys = Array.from(activeUsers.keys()).filter((key) => !uniqueTargets.has(key));
    const requireFeedback = !!payload.requireFeedback;
    const remarksValue = String(payload.remarks || '').trim();
    if (requireFeedback) {
      if (remarksValue.length < 10) {
        return { ok: false, error: 'Please provide at least 10 characters of remarks' };
      }
      const rating = Number(payload.stars);
      if (isNaN(rating) || rating < 1 || rating > 5) {
        return { ok: false, error: 'Stars rating must be between 1 and 5' };
      }
    }
    const starsValue = Number(payload.stars || 0);

    const operationTimestamp = new Date();
    let submitter = String(payload.submitter || '').trim();
    try {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        submitter = userEmail;
      }
    } catch (_e) {
      // ignore
    }
    if (!submitter) {
      submitter = 'System';
    }

    const generateRef = () => {
      const timestamp = Date.now();
      const randomSuffix = Math.random().toString(36).substr(2, 5).toUpperCase();
      return `USR-${timestamp}-${randomSuffix}`;
    };

    const newRows = [];
    const summaryEntries = [];
    const releasedUsers = [];

    releaseTargets.forEach((target) => {
      const rowIndex = target.indices[target.indices.length - 1];
      const currentRow = rows[rowIndex].slice();
      const currentDisplay = display[rowIndex];
      const releaseDate = new Date(operationTimestamp.getTime());
      const releaseNote = remarksValue
        ? remarksValue
        : `Released ${target.displayName || target.requestedName} on ${Utilities.formatDate(releaseDate, TZ(), 'dd-MMM-yyyy HH:mm')}`;

      const releaseRow = currentRow.slice();
      if (iStatus >= 0) releaseRow[iStatus] = 'RELEASE';
      if (iBeneficiary >= 0) releaseRow[iBeneficiary] = target.displayName || target.requestedName;
      if (iBeneficiaryShort >= 0) releaseRow[iBeneficiaryShort] = target.displayName || target.requestedName;
      if (iDate >= 0) releaseRow[iDate] = releaseDate;
      if (iRemarks >= 0) releaseRow[iRemarks] = releaseNote;
      if (iStars >= 0) releaseRow[iStars] = starsValue;
      if (iSubmit >= 0) releaseRow[iSubmit] = submitter;
      if (iRespTime >= 0) releaseRow[iRespTime] = releaseDate;
      const referenceNumber = generateRef();
      if (iRef >= 0) {
        releaseRow[iRef] = referenceNumber;
      }
      newRows.push(releaseRow);

      const summaryObj = {
        Ref: referenceNumber,
        'Date and time of entry': releaseDate,
        Project: iProject >= 0 ? (releaseRow[iProject] || currentDisplay[iProject] || '') : '',
        Team: iTeam >= 0 ? (releaseRow[iTeam] || currentDisplay[iTeam] || '') : '',
        'R.Beneficiary': target.displayName || target.requestedName,
        'R. Ben': target.displayName || target.requestedName,
        'Vehicle Number': releaseRow[iCar] || currentDisplay[iCar] || carNumber,
        Make: iMake >= 0 ? (releaseRow[iMake] || currentDisplay[iMake] || '') : '',
        Model: iModel >= 0 ? (releaseRow[iModel] || currentDisplay[iModel] || '') : '',
        Category: iCategory >= 0 ? (releaseRow[iCategory] || currentDisplay[iCategory] || '') : '',
        'Usage Type': iUsage >= 0 ? (releaseRow[iUsage] || currentDisplay[iUsage] || '') : '',
        Owner: iOwner >= 0 ? (releaseRow[iOwner] || currentDisplay[iOwner] || '') : '',
        Status: 'RELEASE',
        'Last Users remarks': releaseNote,
        Ratings: iStars >= 0 ? (releaseRow[iStars] || '') : '',
        'Submitter username': submitter,
        'R.Ben Time': iRespTime >= 0 ? (releaseRow[iRespTime] || '') : ''
      };
      summaryEntries.push(summaryObj);
      releasedUsers.push(target.displayName || target.requestedName);
    });

    if (newRows.length) {
      const appendStart = sh.getLastRow() + 1;
      sh.getRange(appendStart, 1, newRows.length, lastCol).setValues(newRows);
    }

    summaryEntries.forEach((entry) => {
      try {
        upsertVehicleSummaryRow('Vehicle_Released', entry, 'vehicle');
      } catch (summaryErr) {
        console.warn('Failed to append Vehicle_Released summary for user release', summaryErr);
      }
    });

    try { refreshVehicleStatusSheets(); } catch (refreshErr) { console.error('Partial user release refresh failed:', refreshErr); }
    try { syncVehicleSheetFromCarTP(); } catch (syncErr) { console.error('Partial user release sync failed:', syncErr); }
    try { CacheService.getScriptCache().remove('VEH_PICKER_V1'); } catch (_cacheErr) { /* ignore */ }
    invalidateVehicleInUseCache();
    invalidateVehicleReleasedCache('User release updated assignments');

    return {
      ok: true,
      partial: remainingUsersKeys.length > 0,
      releasedUsers: releasedUsers,
      missingUsers: missingUsers
    };
  } catch (error) {
    console.error('releaseCarUser error:', error);
    return { ok: false, error: String(error) };
  } finally {
    if (lockHeld) {
      lock.releaseLock();
    }
  }
}

/**
 * Switch the responsible beneficiary for an IN USE vehicle without altering other active users.
 * Creates a fresh IN USE row for the selected beneficiary and refreshes cached summaries.
 */
function changeVehicleResponsibleBeneficiary(carNumber, beneficiaryName, options) {
  try {
    const targetCarRaw = String(carNumber || '').trim();
    const targetCar = _vehicleKey_(targetCarRaw);
    const candidateName = _sanitizeResponsibleName(beneficiaryName) || String(beneficiaryName || '').trim();

    if (!targetCar) {
      return { ok: false, error: 'Vehicle number required' };
    }
    if (!candidateName) {
      return { ok: false, error: 'Beneficiary name required' };
    }

    const sh = _openCarTP_();
    if (!sh) {
      return { ok: false, error: 'CarT_P sheet not found' };
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) {
      return { ok: false, error: 'CarT_P sheet has no data' };
    }

    const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    function idx(labels, required) {
      try {
        return IX.get(labels);
      } catch (err) {
        if (required) throw err;
        return -1;
      }
    }

    const iRef   = idx(['Reference Number', 'Ref', 'Ref Number'], false);
    const iDate  = idx(['Date and time of entry', 'Date and time', 'Timestamp', 'Date'], false);
    const iProj  = idx(['Project'], false);
    const iTeam  = idx(['Team', 'Team Name'], false);
    const iCar   = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car'], true);
    const iMake  = idx(['Make', 'Car Make', 'Brand'], false);
    const iModel = idx(['Model', 'Car Model'], false);
    const iCat   = idx(['Category', 'Vehicle Category', 'Cat'], false);
    const iUse   = idx(['Usage Type', 'Usage', 'Use Type'], false);
    const iOwner = idx(['Owner', 'Owner Name', 'Owner Info'], false);
    const iResp  = idx(['R.Beneficiary', 'Responsible Beneficiary', 'R Beneficiary', 'Responsible', 'R. Ben', 'R Ben'], false);
    const iRespShort = idx(['R. Ben', 'R Ben'], false);
    const iRespFull  = idx(['Responsible Beneficiary', 'Name of Responsible beneficiary'], false);
    const iStatus    = idx(['Status', 'In Use/Release', 'In Use / release', 'In Use'], false);
    const iRemarks   = idx(['Last Users remarks', 'Remarks', 'Feedback'], false);
    const iRatings   = idx(['Ratings', 'Stars', 'Rating'], false);
    const iSubmit    = idx(['Submitter username', 'Submitter', 'User'], false);
    const iRespTime  = idx(['R.Ben Time', 'R.Ben timestamp', 'Responsible Beneficiary Time'], false);

    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    const display = dataRange.getDisplayValues();

    let latestInUseRow = null;
    let latestInUseDisplay = null;
    let latestInUseTs = -1;
    let latestInUseIndex = -1;
    let latestAnyRow = null;
    let latestAnyDisplay = null;
    let latestAnyTs = -1;
    let latestAnyIndex = -1;

    const beneficiaryMap = new Map();

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const dispRow = display[r];

      const rawCar = row[iCar] != null && row[iCar] !== '' ? row[iCar] : dispRow[iCar];
      const rowKey = _vehicleKey_(rawCar);
      if (!rowKey || rowKey !== targetCar) continue;

      const ts = _parseTs_(row[iDate] != null && row[iDate] !== '' ? row[iDate] : dispRow[iDate]);
      if (ts >= latestAnyTs) {
        latestAnyTs = ts;
        latestAnyRow = row.slice();
        latestAnyDisplay = dispRow.slice();
        latestAnyIndex = r;
      }

      const status = _normStatus_(row[iStatus] != null && row[iStatus] !== '' ? row[iStatus] : dispRow[iStatus]);
      if (status === 'IN USE') {
        if (ts >= latestInUseTs) {
          latestInUseTs = ts;
          latestInUseRow = row.slice();
          latestInUseDisplay = dispRow.slice();
          latestInUseIndex = r;
        }
      }

      const names = _splitBeneficiaryNames_(row[iResp] != null && row[iResp] !== '' ? row[iResp] : dispRow[iResp]);
      names.forEach((name) => {
        const normalized = _beneficiaryKey_(name);
        if (!normalized) return;
        const existing = beneficiaryMap.get(normalized);
        if (!existing || ts >= existing.ts) {
          beneficiaryMap.set(normalized, {
            name: _sanitizeResponsibleName(name) || name,
            status: status || '',
            ts: ts || 0
          });
        }
      });
    }

    if (!latestInUseRow) {
      // No IN USE entry exists, fall back to latest row but still enforce validation
      latestInUseRow = latestAnyRow ? latestAnyRow.slice() : null;
      latestInUseDisplay = latestAnyDisplay ? latestAnyDisplay.slice() : null;
      latestInUseTs = latestAnyTs;
      latestInUseIndex = latestAnyIndex;
    }

    if (!latestInUseRow) {
      return { ok: false, error: 'No matching vehicle entries found.' };
    }

    const normalizedCandidate = _beneficiaryKey_(candidateName);
    const candidateEntry = beneficiaryMap.get(normalizedCandidate);
    if (!candidateEntry || candidateEntry.status !== 'IN USE') {
      return { ok: false, error: `${candidateName} is not currently marked as IN USE for this vehicle.` };
    }

    const currentNames = _splitBeneficiaryNames_(latestInUseRow[iResp] != null && latestInUseRow[iResp] !== '' ? latestInUseRow[iResp] : latestInUseDisplay[iResp]);
    const currentResponsible = currentNames.length ? currentNames[0] : '';
    if (currentResponsible && _beneficiaryKey_(currentResponsible) === normalizedCandidate) {
      return {
        ok: true,
        unchanged: true,
        previousResponsible: currentResponsible,
        newResponsible: candidateEntry.name,
        carDetails: getCarReleaseDetails(targetCarRaw)
      };
    }

    const baseRow = latestInUseRow ? latestInUseRow.slice() : new Array(lastCol).fill('');
    const now = new Date();
    const nowIso = now instanceof Date ? now : new Date(now);
    const sanitizedCandidate = _sanitizeResponsibleName(candidateEntry.name) || candidateEntry.name || candidateName;

    function setValue(index, value) {
      if (index >= 0 && index < baseRow.length) {
        baseRow[index] = value;
      }
    }

    setValue(iDate, now);
    setValue(iStatus, 'IN USE');
    setValue(iResp, sanitizedCandidate);
    setValue(iRespShort, sanitizedCandidate);
    setValue(iRespFull, sanitizedCandidate);
    setValue(iRespTime, now);
    setValue(iSubmit, (function(){
      try {
        const email = Session.getActiveUser().getEmail();
        return email || 'System';
      } catch (_e) {
        return 'System';
      }
    })());
    setValue(iRatings, 0);
    setValue(iRemarks, '');

    if (iRef >= 0) {
      const ref = `RESP-${Date.now()}-${Math.random().toString(36).slice(2, 6).toUpperCase()}`;
      setValue(iRef, ref);
    }

    const ensureField = (index, fallback) => {
      if (index < 0) return;
      if (baseRow[index] == null || baseRow[index] === '') {
        setValue(index, fallback);
      }
    };

    const fallbackRow = latestAnyRow || latestInUseRow;
    if (fallbackRow) {
      ensureField(iProj, fallbackRow[iProj]);
      ensureField(iTeam, fallbackRow[iTeam]);
      ensureField(iMake, fallbackRow[iMake]);
      ensureField(iModel, fallbackRow[iModel]);
      ensureField(iCat, fallbackRow[iCat]);
      ensureField(iUse, fallbackRow[iUse]);
      ensureField(iOwner, fallbackRow[iOwner]);
      ensureField(iCar, fallbackRow[iCar]);
    }

    sh.getRange(lastRow + 1, 1, 1, lastCol).setValues([baseRow]);

    try {
      invalidateVehicleInUseCache();
      invalidateVehicleReleasedCache('Responsible beneficiary updated');
    } catch (_cacheErr) {
      // swallow cache errors
    }

    try {
      refreshVehicleStatusSheets();
    } catch (refreshErr) {
      console.warn('refreshVehicleStatusSheets failed after responsible change:', refreshErr);
    }

    try {
      syncVehicleSheetFromCarTP();
    } catch (syncErr) {
      console.warn('syncVehicleSheetFromCarTP failed after responsible change:', syncErr);
    }

    const updatedDetails = getCarReleaseDetails(targetCarRaw);

    return {
      ok: true,
      previousResponsible: currentResponsible || '',
      newResponsible: sanitizedCandidate,
      carNumber: targetCarRaw,
      carDetails: updatedDetails
    };
  } catch (error) {
    console.error('changeVehicleResponsibleBeneficiary error:', error);
    return { ok: false, error: String(error && error.message ? error.message : error) };
  }
}

function addVehicleSecondaryBeneficiary(carNumber, beneficiaryName, options) {
  try {
    const targetCarRaw = String(carNumber || '').trim();
    const targetCar = _vehicleKey_(targetCarRaw);
    const rawCandidate = String(beneficiaryName || '').trim();
    const candidateName = _sanitizeResponsibleName(rawCandidate) || rawCandidate;

    if (!targetCar) {
      return { ok: false, error: 'Vehicle number required' };
    }
    if (!candidateName) {
      return { ok: false, error: 'Beneficiary name required' };
    }

    const sh = _openCarTP_();
    if (!sh) return { ok: false, error: 'CarT_P sheet not found' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) {
      return { ok: false, error: 'CarT_P sheet has no data' };
    }

    const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    function idx(labels, required) {
      try {
        return IX.get(labels);
      } catch (err) {
        if (required) throw err;
        return -1;
      }
    }

    const iRef   = idx(['Reference Number', 'Ref', 'Ref Number'], false);
    const iDate  = idx(['Date and time of entry', 'Date and time', 'Timestamp', 'Date'], false);
    const iProj  = idx(['Project'], false);
    const iTeam  = idx(['Team', 'Team Name'], false);
    const iCar   = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car'], true);
    const iMake  = idx(['Make', 'Car Make', 'Brand'], false);
    const iModel = idx(['Model', 'Car Model'], false);
    const iCat   = idx(['Category', 'Vehicle Category', 'Cat'], false);
    const iUse   = idx(['Usage Type', 'Usage', 'Use Type'], false);
    const iOwner = idx(['Owner', 'Owner Name', 'Owner Info'], false);
    const iResp  = idx(['R.Beneficiary', 'Responsible Beneficiary', 'R Beneficiary', 'Responsible', 'R. Ben', 'R Ben'], false);
    const iRespShort = idx(['R. Ben', 'R Ben'], false);
    const iRespFull  = idx(['Responsible Beneficiary', 'Name of Responsible beneficiary'], false);
    const iStatus    = idx(['Status', 'In Use/Release', 'In Use / release', 'In Use'], false);
    const iRemarks   = idx(['Last Users remarks', 'Remarks', 'Feedback'], false);
    const iRatings   = idx(['Ratings', 'Stars', 'Rating'], false);
    const iSubmit    = idx(['Submitter username', 'Submitter', 'User'], false);
    const iRespTime  = idx(['R.Ben Time', 'R.Ben timestamp', 'Responsible Beneficiary Time'], false);

    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    const display = dataRange.getDisplayValues();

    let latestInUseRow = null;
    let latestInUseDisplay = null;
    let latestInUseTs = -1;
    let latestAnyRow = null;
    let latestAnyDisplay = null;
    let latestAnyTs = -1;

    const beneficiaryMap = new Map();

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const dispRow = display[r];
      const rawCar = row[iCar] != null && row[iCar] !== '' ? row[iCar] : dispRow[iCar];
      const rowKey = _vehicleKey_(rawCar);
      if (!rowKey || rowKey !== targetCar) continue;

      const ts = _parseTs_(row[iDate] != null && row[iDate] !== '' ? row[iDate] : dispRow[iDate]);
      if (ts >= latestAnyTs) {
        latestAnyTs = ts;
        latestAnyRow = row.slice();
        latestAnyDisplay = dispRow.slice();
      }
      const status = _normStatus_(row[iStatus] != null && row[iStatus] !== '' ? row[iStatus] : dispRow[iStatus]);
      if (status === 'IN USE' && ts >= latestInUseTs) {
        latestInUseTs = ts;
        latestInUseRow = row.slice();
        latestInUseDisplay = dispRow.slice();
      }

      const names = _splitBeneficiaryNames_(row[iResp] != null && row[iResp] !== '' ? row[iResp] : dispRow[iResp]);
      names.forEach((name) => {
        const normalized = _beneficiaryKey_(name);
        if (!normalized) return;
        const existing = beneficiaryMap.get(normalized);
        if (!existing || ts >= existing.ts) {
          beneficiaryMap.set(normalized, { name, status: status || '', ts: ts || 0 });
        }
      });
    }

    const normalizedCandidate = _beneficiaryKey_(candidateName);
    const existingEntry = beneficiaryMap.get(normalizedCandidate);
    if (existingEntry && existingEntry.status === 'IN USE') {
      return { ok: false, error: `${candidateName} is already marked as IN USE for this vehicle.` };
    }

    if (!latestInUseRow) {
      latestInUseRow = latestAnyRow ? latestAnyRow.slice() : null;
      latestInUseDisplay = latestAnyDisplay ? latestAnyDisplay.slice() : null;
    }
    if (!latestInUseRow) {
      return { ok: false, error: 'No baseline IN USE entry found for this vehicle.' };
    }

    const currentResponsible =
      _sanitizeResponsibleName(
        (iRespShort >= 0 ? latestInUseRow[iRespShort] : '') ||
        (iRespFull >= 0 ? latestInUseRow[iRespFull] : '')
      ) ||
      _sanitizeResponsibleName(
        (iRespShort >= 0 && latestInUseDisplay ? latestInUseDisplay[iRespShort] : '') ||
        (iRespFull >= 0 && latestInUseDisplay ? latestInUseDisplay[iRespFull] : '')
      );

    const baseRow = latestInUseRow.slice();
    const now = new Date();
    const responsibleDisplay = currentResponsible || candidateName;
    const existingRespTime = iRespTime >= 0
      ? (latestInUseRow[iRespTime] || (latestInUseDisplay ? latestInUseDisplay[iRespTime] : ''))
      : '';

    if (iResp >= 0) baseRow[iResp] = candidateName;
    if (iRespShort >= 0) baseRow[iRespShort] = responsibleDisplay;
    if (iRespFull >= 0) baseRow[iRespFull] = responsibleDisplay;
    if (iDate >= 0) baseRow[iDate] = now;
    if (iStatus >= 0) baseRow[iStatus] = 'IN USE';
    if (iRemarks >= 0) baseRow[iRemarks] = '';
    if (iRatings >= 0) baseRow[iRatings] = 0;
    if (iRespTime >= 0) baseRow[iRespTime] = existingRespTime || now;
    if (iSubmit >= 0) {
      try {
        baseRow[iSubmit] = Session.getActiveUser().getEmail() || 'System';
      } catch (_e) {
        baseRow[iSubmit] = 'System';
      }
    }
    if (iRef >= 0) {
      baseRow[iRef] = `SEC-${Date.now()}-${Math.random().toString(36).slice(2, 6).toUpperCase()}`;
    }

    const fallbackRow = latestAnyRow || latestInUseRow;
    const ensureField = (index, fallback) => {
      if (index < 0) return;
      if (baseRow[index] == null || baseRow[index] === '') {
        baseRow[index] = fallback;
      }
    };
    if (fallbackRow) {
      ensureField(iProj, fallbackRow[iProj]);
      ensureField(iTeam, fallbackRow[iTeam]);
      ensureField(iMake, fallbackRow[iMake]);
      ensureField(iModel, fallbackRow[iModel]);
      ensureField(iCat, fallbackRow[iCat]);
      ensureField(iUse, fallbackRow[iUse]);
      ensureField(iOwner, fallbackRow[iOwner]);
      ensureField(iCar, fallbackRow[iCar]);
    }

    sh.getRange(lastRow + 1, 1, 1, lastCol).setValues([baseRow]);

    try { invalidateVehicleInUseCache(); } catch (_e) {}
    try { invalidateVehicleReleasedCache('Vehicle beneficiary added'); } catch (_e) {}
    try { refreshVehicleStatusSheets(); } catch (refreshErr) { console.warn('refreshVehicleStatusSheets failed after secondary beneficiary add:', refreshErr); }
    try { syncVehicleSheetFromCarTP(); } catch (syncErr) { console.warn('syncVehicleSheetFromCarTP failed after secondary beneficiary add:', syncErr); }

    const updatedDetails = getCarReleaseDetails(targetCarRaw);
    return {
      ok: true,
      newBeneficiary: candidateName,
      carNumber: targetCarRaw,
      carDetails: updatedDetails
    };
  } catch (error) {
    console.error('addVehicleSecondaryBeneficiary error:', error);
    return { ok: false, error: String(error && error.message ? error.message : error) };
  }
}

function assignCarToTeamWithReturn(payload) {
  const result = assignCarToTeam(payload);
  if (!result || result.ok === false) {
    return result;
  }
  const beneficiaries = Array.isArray(payload?.beneficiaries)
    ? payload.beneficiaries.map((name) => String(name || '').trim()).filter(Boolean)
    : [];
  let primary = beneficiaries.length ? beneficiaries[0] : String(payload?.responsibleBeneficiary || '').trim();
  if (!primary && typeof payload?.responsibleBeneficiary === 'string') {
    primary = payload.responsibleBeneficiary.trim();
  }
  return Object.assign({}, result, {
    assignedBeneficiaries: beneficiaries,
    responsibleBeneficiary: primary || ''
  });
}

/**
 * Get cars from CarT_P sheet.
 * Returns the latest entry for each vehicle number (across all statuses by default).
 * Pass includeOnlyRelease=true to filter down to cars whose latest status is RELEASE.
 */
function getAvailableCars(includeOnlyRelease) {
  try {
    // Open CarT_P sheet with a tolerant finder
    const sh = _openCarTP_();
    if (!sh) { console.log('CarT_P sheet not found'); return []; }
    
    const lastRow = sh.getLastRow();
    console.log(`CarT_P sheet found, last row: ${lastRow}`);
    
    if (lastRow <= 1) {
      console.log('No data in CarT_P sheet');
      return [];
    }
    
    // Read header and data (header-driven indices to avoid column order issues)
    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);

    // Tolerant indices
    function idx(labels, required){
      try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; }
    }
    const iRef      = idx(['Reference Number','Ref','Ref Number'], false);
    // Date can be optional; fall back to row order if missing
    const iDate     = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    const iProject  = idx(['Project'], false);
    const iTeam     = idx(['Team'], false);
    // Car Number: allow tolerant fallback if standard aliases fail
    let iCarNo = -1;
    try { iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); } catch(_){ iCarNo = -1; }
    if (iCarNo < 0) {
      iCarNo = _findCarNumberColumn_(head);
    }
    if (iCarNo < 0) throw new Error('Car Number column not found');
    const iMake     = idx(['Make','Car Make','Brand'], false);
    const iModel    = idx(['Model','Car Model'], false);
    const iCategory = idx(['Category','Vehicle Category','Cat'], false);
    const iUsage    = idx(['Usage Type','Usage','Use Type'], false);
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner    = idx(['Owner','Owner Name','Owner Info'], false);
    const iResp     = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible','R. Ben','R Ben'], false);
    // Status optional to avoid hard dependency on header naming
    const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iRemarks  = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars    = idx(['Stars','Rating'], false);
    const iSubmit   = idx(['Submitter username','Submitter','User'], false);

    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const disp  = dataRange.getDisplayValues();
    console.log(`Retrieved ${data.length} rows from CarT_P sheet (header-driven)`);
    
    // Group by car number and find latest entry for each
    const carMap = new Map();
    
    data.forEach((row, index) => {
      // Prefer raw values; fallback to display values if raw is empty
      const rawCar = row[iCarNo];
      const dispCar = disp[index][iCarNo];
      const carNumber = String((rawCar == null || String(rawCar).trim()==='') ? dispCar : rawCar).trim();
      if (!carNumber) return;
      const status = String(iStatus>=0 ? (row[iStatus] || disp[index][iStatus]) : '').trim().toUpperCase();
      const dateTime = (iDate>=0 ? (row[iDate] || disp[index][iDate] || null) : null);
      // If date column missing or unparsable, use row order as monotonic timestamp
      const tsFromDate = (dateTime instanceof Date) ? dateTime.getTime() : (dateTime ? (new Date(dateTime)).getTime() : NaN);
      const ts = isNaN(tsFromDate) ? (index + 1) : tsFromDate;

      const carData = {
        refNumber: iRef>=0 ? row[iRef] : '',
        dateTime: dateTime,
        ts: ts,
        project: iProject>=0 ? row[iProject] : '',
        team:    iTeam>=0    ? row[iTeam]    : '',
        carNumber: carNumber,
        make: iMake>=0 ? row[iMake] : '',
        model: iModel>=0 ? row[iModel] : '',
        category: iCategory>=0 ? (row[iCategory] || disp[index][iCategory] || '') : '',
        usageType: iUsage>=0 ? row[iUsage] : '',
        contractType: iContract>=0 ? row[iContract] : '',
        owner: iOwner>=0 ? row[iOwner] : (iOwner>=0 ? disp[index][iOwner] : ''),
        responsibleBeneficiary: iResp>=0 ? (row[iResp] || disp[index][iResp] || '') : '',
        'R.Beneficiary': iResp>=0 ? (row[iResp] || disp[index][iResp] || '') : '',
        'R. Ben': iResp>=0 ? (row[iResp] || disp[index][iResp] || '') : '',
        status: status,
        remarks: iRemarks>=0 ? row[iRemarks] : '',
        stars: iStars>=0 ? row[iStars] : 0,
        submitter: iSubmit>=0 ? row[iSubmit] : '',
        rowIndex: index + 2
      };

      // Keep only the latest entry for each car number (by timestamp, fallback to row order)
      if (!carMap.has(carNumber)) {
        carMap.set(carNumber, carData);
      } else {
        const prev = carMap.get(carNumber);
        const prevTs = (prev && typeof prev.ts==='number') ? prev.ts : 0;
        if (ts >= prevTs) carMap.set(carNumber, carData);
      }
    });
    
    console.log(`Processed ${carMap.size} unique cars`);

    // Build latest-per-vehicle list
    const allCars = [];
    carMap.forEach((carData) => {
      const resp = carData.responsibleBeneficiary || '';
      allCars.push({
        carNumber: carData.carNumber,
        make: carData.make,
        model: carData.model,
        usageType: carData.usageType,
        contractType: carData.contractType,
        owner: carData.owner,
        remarks: carData.remarks,
        project: carData.project,
        team: carData.team,
        status: carData.status,
        stars: carData.stars,
        dateTime: carData.dateTime,
        category: carData.category,
        responsibleBeneficiary: resp,
        'R.Beneficiary': resp
      });
    });

    // Optional filter for only RELEASE if explicitly requested
    if (includeOnlyRelease === true) {
      const onlyRelease = allCars.filter(c => {
        const s = String(c.status || '').trim().toUpperCase();
        return s === 'RELEASE' || s.startsWith('RELEASE'); // allow RELEASE/RELEASED variants
      });
      console.log(`Returning ${onlyRelease.length} vehicles with latest status RELEASE (explicit filter)`);
      return onlyRelease;
    }

    console.log(`Returning ${allCars.length} vehicles (latest per car, all statuses)`);
    return allCars;
    
  } catch (error) {
    console.error('Error getting available cars:', error);
    return [];
  }
}


/**
 * Return one entry per unique Vehicle Number from CarT_P, ignoring status and recency.
 * Chooses the last-seen row (bottom-most) for metadata if duplicates exist.
 */
function getAllUniqueCars() {
  try {
    const map = new Map();

    const inUse = getVehicleInUseSummary();
    if (inUse.ok && Array.isArray(inUse.assignments)) {
      inUse.assignments.forEach(function(entry){
        const carNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
        if (!carNumber) return;
        const key = carNumber.toUpperCase();
        if (map.has(key)) return;
        map.set(key, {
          carNumber: carNumber,
          make: entry.make || '',
          model: entry.model || '',
          category: entry.category || '',
          usageType: entry.usageType || '',
          contractType: entry.contractType || '',
          owner: entry.owner || '',
          remarks: entry.remarks || '',
          project: entry.project || '',
          team: entry.team || '',
          status: entry.assignmentStatus || 'IN USE',
          stars: entry.stars || 0,
          responsibleBeneficiary: entry.beneficiary || entry.responsibleBeneficiary || '',
          'R.Beneficiary': entry.beneficiary || entry.responsibleBeneficiary || ''
        });
      });
    }

    const released = getVehicleReleasedSummary();
    if (released.ok && Array.isArray(released.vehicles)) {
      released.vehicles.forEach(function(entry){
        const carNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
        if (!carNumber) return;
        const key = carNumber.toUpperCase();
        if (!map.has(key)) {
          map.set(key, {
            carNumber: carNumber,
            make: entry.make || '',
            model: entry.model || '',
            category: entry.category || '',
            usageType: entry.usageType || '',
            contractType: entry.contractType || '',
            owner: entry.owner || '',
            remarks: entry.remarks || '',
            project: entry.project || '',
            team: entry.team || '',
            status: entry.status || 'RELEASE',
            stars: entry.stars || 0,
            responsibleBeneficiary: entry.responsibleBeneficiary || '',
            'R.Beneficiary': entry.responsibleBeneficiary || ''
          });
        }
      });
    }

    const out = Array.from(map.values()).sort(function(a,b){
      return String(a.carNumber||'').localeCompare(String(b.carNumber||''));
    });
    console.log(`getAllUniqueCars -> ${out.length} unique vehicles from summary sheets`);
    return out;
  } catch (e) {
    console.error('getAllUniqueCars error:', e);
    return [];
  }
}

/**
 * Return one entry per unique Vehicle Number from CarT_P, but only from rows
 * where Status indicates RELEASE (same-row filter). Iterates bottom→top so the
 * last RELEASE occurrence per car wins.
 */
function getUniqueReleaseCars() {
  try {
    const list = _vehicleSheetReleaseVehicles();
    console.log(`getUniqueReleaseCars -> ${list.length} vehicles via Vehicle_Released summary`);
    return list;
  } catch (e) {
    console.error('getUniqueReleaseCars error:', e);
    return [];
  }
}

/**
 * Return unique vehicles where the latest entry by Date/Time (column like 'Date and time of entry')
 * has Status = RELEASE. Ignores older entries regardless of their statuses.
 */
function getLatestReleaseCars() {
  try {
    const summary = getVehicleReleasedSummary();
    if (!summary.ok || !Array.isArray(summary.vehicles)) {
      console.warn('getLatestReleaseCars: Vehicle_Released summary unavailable or empty');
      return [];
    }

    const out = summary.vehicles.map(function(entry) {
      const carNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
      if (!carNumber) return null;
      const statusRaw = String(entry.status || '').trim().toUpperCase();
      const status = statusRaw ? statusRaw : 'RELEASE';
      const compactStatus = status.replace(/[\s_\-/]+/g,'');
      if (compactStatus !== 'RELEASE' && compactStatus !== 'RELEASED') return null;

      return {
        carNumber: carNumber,
        make: entry.make || '',
        model: entry.model || '',
        category: entry.category || '',
        usageType: entry.usageType || '',
        contractType: entry.contractType || '',
        owner: entry.owner || '',
        responsibleBeneficiary: entry.responsibleBeneficiary || '',
        'R.Beneficiary': entry.responsibleBeneficiary || '',
        remarks: entry.remarks || '',
        project: entry.project || '',
        team: entry.team || '',
        status: status,
        dateTime: entry.latestRelease || summary.updatedAt || ''
      };
    }).filter(Boolean);

    out.sort(function(a, b){ return String(a.carNumber||'').localeCompare(String(b.carNumber||'')); });
    console.log(`getLatestReleaseCars -> ${out.length} vehicles sourced from Vehicle_Released summary`);
    return out;
  } catch (e) {
    console.error('getLatestReleaseCars error:', e);
    return [];
  }
}

/**
 * Get car history for a specific car number
 * Returns users, feedback, and ratings from previous assignments
 */
function getCarHistory(carNumber) {
  try {
    if (!carNumber) {
      console.log('No car number provided for history lookup');
      return { users: [], feedback: [], ratings: [] };
    }
    const sh = _openCarTP_();
    if (!sh) { console.log('CarT_P sheet not found'); return { users: [], feedback: [], ratings: [] }; }
    
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      console.log('No data in CarT_P sheet');
      return { users: [], feedback: [], ratings: [] };
    }
    // Header-driven lookup
    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iDate     = idx(['Date and time of entry','Date and time','Timestamp','Date'], true);
    const iCarNo    = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], true);
    const iRemarks  = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars    = idx(['Stars','Rating'], false);
    const iSubmit   = idx(['Submitter username','Submitter','User'], false);
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Filter entries for this specific car number
    const carHistory = data.filter(row => String(row[iCarNo]||'').trim() === carNumber);
    
    // Sort by date (most recent first)
    carHistory.sort((a, b) => {
      const dateA = a[iDate] || new Date(0);
      const dateB = b[iDate] || new Date(0);
      const ta = (dateA instanceof Date) ? dateA.getTime() : (new Date(dateA)).getTime() || 0;
      const tb = (dateB instanceof Date) ? dateB.getTime() : (new Date(dateB)).getTime() || 0;
      return tb - ta;
    });
    
    // Extract history data (limit to last 3 entries)
    const recentHistory = carHistory.slice(0, 3);
    
    const users = [];
    const feedback = [];
    const ratings = [];
    
    recentHistory.forEach(row => {
      const submitter = iSubmit>=0 ? (row[iSubmit] || 'Unknown User') : 'Unknown User';
      const remarks = iRemarks>=0 ? (row[iRemarks] || '') : '';
      const stars = iStars>=0 ? (row[iStars] || 0) : 0;
      const dateTime = row[iDate];
      
      // Format user with date
      if (submitter && submitter !== 'Unknown User') {
    const dateStr = dateTime ? Utilities.formatDate(new Date(dateTime), TZ(), 'dd-MMM-yy') : 'Unknown date';
        users.push(`${submitter} (${dateStr})`);
      }
      
      // Add feedback if available
      if (remarks && remarks.trim()) {
        feedback.push(remarks.trim());
      }
      
      // Add rating if available
      if (stars && stars > 0) {
        ratings.push(Number(stars));
      }
    });
    
    console.log(`Found ${carHistory.length} history entries for car ${carNumber}`);
    
    return {
      users: users,
      feedback: feedback,
      ratings: ratings
    };
    
  } catch (error) {
    console.error('Error getting car history:', error);
    return { users: [], feedback: [], ratings: [] };
  }
}

/**
 * Legacy wrapper conserved for backwards compatibility.
 * Delegates to the primary testCarConnection implementation below.
 */
function testCarConnectionLegacy() {
  try {
    return testCarConnection();
  } catch (e) {
    return { success: false, error: String(e) };
  }
}

/** Debug helper: return CarT_P overview (sheet id, names, header, size) */
function debugCarTPOverview(){
  try{
    console.log('debugCarTPOverview called - checking CAR_SHEET_ID:', CAR_SHEET_ID);
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    const sheets = ss.getSheets().map(s=>s.getName());
    console.log('Available sheets:', sheets);
    const sh = _openCarTP_();
    if (!sh){
      console.log('CarT_P sheet not found in sheets:', sheets);
      return { ok:false, reason:'CarT_P sheet not found', sheetId: CAR_SHEET_ID, sheets };
    }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    console.log(`Found CarT_P sheet: ${sh.getName()}, rows: ${lastRow}, cols: ${lastCol}`);
    const headDisp = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const headNorm = headDisp.map(h => (String(h||'').trim().toLowerCase()));
    console.log('Header:', headDisp);
    return { ok:true, sheetId: CAR_SHEET_ID, sheetName: sh.getName(), sheets, lastRow, lastCol, header: headDisp, headerNorm: headNorm };
  }catch(e){
    console.error('debugCarTPOverview error:', e);
    return { ok:false, error: String(e), sheetId: CAR_SHEET_ID };
  }
}

/** Simple test function to verify basic connectivity */
function testCarConnection(){
  try{
    console.log('testCarConnection called');
    const result = debugCarTPOverview();
    console.log('debugCarTPOverview result:', result);
    if (result.ok) {
      // Test both with and without RELEASE filter
      const allCars = getAvailableCars(false);
      const releaseCars = getAvailableCars(true);
      console.log('getAvailableCars(false) result:', allCars);
      console.log('getAvailableCars(true) result:', releaseCars);
      
      return { 
        success: true, 
        debug: result, 
        cars: releaseCars, 
        carCount: releaseCars ? releaseCars.length : 0,
        allCarsCount: allCars ? allCars.length : 0,
        allCars: allCars
      };
    } else {
      return { success: false, debug: result, error: 'Sheet not found' };
    }
  } catch(e) {
    console.error('testCarConnection error:', e);
    return { success: false, error: String(e) };
  }
}

/**
 * Force clear CarT_P sheet and add fresh sample data for testing
 */
function forceClearAndAddSampleData() {
  try {
    console.log('forceClearAndAddSampleData called');
    
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    let sh = ss.getSheetByName('CarT_P');
    
    if (sh) {
      // Clear all content first
      sh.clear();
      console.log('Cleared existing CarT_P sheet');
    } else {
      // Create sheet if it doesn't exist
      sh = ss.insertSheet('CarT_P');
      console.log('Created new CarT_P sheet');
    }
    
    // Set up header for CarT_P sheet
    const header = [
      'Reference Number',
      'Date and time of entry',
      'Project',
      'Team',
      'R.Beneficiary',
      'Vehicle Number',
      'Make',
      'Model',
      'Category',
      'Usage Type',
      'Owner',
      'Status',
      'Last Users remarks',
      'Ratings',
      'Submitter username',
      'R.Ben Time'
    ];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    console.log('Added header to CarT_P sheet');
    
    // Add sample car data - mix of RELEASE and IN USE entries (anonymized vehicle numbers)
    const sampleData = [
      ['REF-001', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'RELEASE', 'This car was good for city driving', 4, 'Jane Smith', new Date()],
      ['REF-002', new Date(), 'Project Falcon', 'Beta Team', 'Jane Smith', 'TEST-VEHICLE-B', 'Honda', 'Civic', 'Sedan', 'Rental', 'Jane Smith', 'IN USE', 'Currently assigned to team', 5, 'John Doe', new Date()],
      ['REF-003', new Date(), 'Project Eagle', 'Gamma Squad', 'Mike Johnson', 'TEST-VEHICLE-C', 'Nissan', 'Sentra', 'Sedan', 'Rental', 'Mike Johnson', 'RELEASE', 'Excellent vehicle for long trips', 5, 'Sarah Wilson', new Date()],
      ['REF-004', new Date(), 'Project Falcon', 'Delta Force', 'Sarah Wilson', 'TEST-VEHICLE-D', 'Toyota', 'Camry', 'Sedan', 'Rental', 'Sarah Wilson', 'IN USE', 'Comfortable and reliable vehicle', 4, 'Mike Johnson', new Date()],
      ['REF-005', new Date(), 'Project Eagle', 'Echo Team', 'Alex Brown', 'TEST-VEHICLE-E', 'Honda', 'Accord', 'Sedan', 'Rental', 'Alex Brown', 'RELEASE', 'Great car for team transportation', 5, 'Lisa Davis', new Date()],
      ['REF-006', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'IN USE', 'Reassigned to same team', 4, 'Jane Smith', new Date()]
    ];
    
    // Add the sample data starting from row 2
    const range = sh.getRange(2, 1, sampleData.length, sampleData[0].length);
    range.setValues(sampleData);
    
    console.log(`Added ${sampleData.length} sample car records to CarT_P sheet`);
    
    return {
      success: true,
      message: `Successfully added ${sampleData.length} sample car records with mixed statuses`,
      recordsAdded: sampleData.length,
      header: header
    };
    
  } catch (error) {
    console.error('Error in forceClearAndAddSampleData:', error);
    return {
      success: false,
      error: String(error)
    };
  }
}

/**
 * Add sample car data to CarT_P sheet for testing
 * This function adds some sample RELEASE status cars if the sheet is empty
 */
function addSampleCarData() {
  try {
    console.log('addSampleCarData called');
    
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    let sh = ss.getSheetByName('CarT_P');
    
    // Create CarT_P sheet if it doesn't exist
    if (!sh) {
      console.log('CarT_P sheet not found, creating it');
      sh = ss.insertSheet('CarT_P');
      // Set up header for CarT_P sheet
      const header = [
        'Reference Number',
        'Date and time of entry',
        'Project',
        'Team',
        'R.Beneficiary',
        'Vehicle Number',
        'Make',
        'Model',
        'Category',
        'Usage Type',
        'Owner',
        'In Use/Release',
        'Last Users remarks',
        'Stars',
        'Submitter username',
        'R.Ben Time'
      ];
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
      console.log('Created CarT_P sheet with header');
    }
    
    const lastRow = sh.getLastRow();
    console.log('Current last row in CarT_P:', lastRow);
    
    // Check if there's already data (beyond header)
    if (lastRow > 1) {
      console.log('CarT_P sheet already has data');
      return { success: false, error: 'CarT_P sheet already has data. Clear it first if you want to add sample data.' };
    }
    
    // Add sample car data - mix of RELEASE and IN USE entries (anonymized vehicle numbers)
    const sampleData = [
      ['REF-001', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'RELEASE', 'This car was good for city driving', 4, 'Jane Smith', new Date()],
      ['REF-002', new Date(), 'Project Falcon', 'Beta Team', 'Jane Smith', 'TEST-VEHICLE-B', 'Honda', 'Civic', 'Sedan', 'Rental', 'Jane Smith', 'IN USE', 'Currently assigned to team', 5, 'John Doe', new Date()],
      ['REF-003', new Date(), 'Project Eagle', 'Gamma Squad', 'Mike Johnson', 'TEST-VEHICLE-C', 'Nissan', 'Sentra', 'Sedan', 'Rental', 'Mike Johnson', 'RELEASE', 'Excellent vehicle for long trips', 5, 'Sarah Wilson', new Date()],
      ['REF-004', new Date(), 'Project Falcon', 'Delta Force', 'Sarah Wilson', 'TEST-VEHICLE-D', 'Toyota', 'Camry', 'Sedan', 'Rental', 'Sarah Wilson', 'IN USE', 'Comfortable and reliable vehicle', 4, 'Mike Johnson', new Date()],
      ['REF-005', new Date(), 'Project Eagle', 'Echo Team', 'Alex Brown', 'TEST-VEHICLE-E', 'Honda', 'Accord', 'Sedan', 'Rental', 'Alex Brown', 'RELEASE', 'Great car for team transportation', 5, 'Lisa Davis', new Date()],
      ['REF-006', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'IN USE', 'Reassigned to same team', 4, 'Jane Smith', new Date()]
    ];
    
    // Add the sample data starting from row 2
    const range = sh.getRange(2, 1, sampleData.length, sampleData[0].length);
    range.setValues(sampleData);
    
    console.log(`Added ${sampleData.length} sample car records to CarT_P sheet`);
    
    return {
      success: true,
      message: `Successfully added ${sampleData.length} sample car records with RELEASE status`,
      recordsAdded: sampleData.length
    };
    
  } catch (error) {
    console.error('Error adding sample car data:', error);
    return {
      success: false,
      error: String(error)
    };
  }
}

/** Debug helper: return Vehicle sheet overview (sheet id, names, header, size) */
function debugVehicleOverview(){
  try{
    console.log('debugVehicleOverview called - checking SHEET_ID:', SHEET_ID);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = ss.getSheets().map(s=>s.getName());
    console.log('Available sheets:', sheets);
    const sh = _openVehicleSheet_();
    if (!sh){
      console.log('Vehicle sheet not found in sheets:', sheets);
      return { ok:false, reason:'Vehicle sheet not found', sheetId: SHEET_ID, sheets };
    }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    console.log(`Found Vehicle sheet: ${sh.getName()}, rows: ${lastRow}, cols: ${lastCol}`);
    const headDisp = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const headNorm = headDisp.map(h => (String(h||'').trim().toLowerCase()));
    console.log('Header:', headDisp);
    return { ok:true, sheetId: SHEET_ID, sheetName: sh.getName(), sheets, lastRow, lastCol, header: headDisp, headerNorm: headNorm };
  }catch(e){
    console.error('debugVehicleOverview error:', e);
    return { ok:false, error: String(e), sheetId: SHEET_ID };
  }
}

/** Get all Vehicle data for display in popup modal */
function getAllCarTPData() {
  try {
    console.log('getAllCarTPData called - fetching from Vehicle sheet');
    
    // Open Vehicle sheet instead of CarT_P
    const sh = _openVehicleSheet_();
    if (!sh) {
      console.log('Vehicle sheet not found');
      return { success: false, error: 'Vehicle sheet not found' };
    }
    
    const lastRow = sh.getLastRow();
    console.log(`Vehicle sheet found, last row: ${lastRow}`);
    
    if (lastRow <= 1) {
      console.log('No data in Vehicle sheet');
      return { success: true, records: [], message: 'No data found in Vehicle sheet' };
    }
    
    // Read header and all data
    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    
    // Get column indices with tolerant matching
    function idx(labels, required) {
      try { return IX.get(labels); } catch(e) { if(required) throw e; return -1; }
    }
    
    const iRef = idx(['Reference Number','Ref','Ref Number'], false);
    const iDate = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    const iProject = idx(['Project'], false);
    const iTeam = idx(['Team'], false);
    
    // Car Number with fallback
    let iCarNo = -1;
    try { iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); } catch(_){ iCarNo = -1; }
    if (iCarNo < 0) {
      iCarNo = _findCarNumberColumn_(head);
    }
    if (iCarNo < 0) throw new Error('Car Number column not found');
    
    const iMake = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iUsage = idx(['Usage Type','Usage','Use Type'], false);
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);
    const iStatus = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iRemarks = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars = idx(['Stars','Rating'], false);
    const iSubmit = idx(['Submitter username','Submitter','User'], false);
    
    // Get all data
    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const disp = dataRange.getDisplayValues();
    
    console.log(`Retrieved ${data.length} rows from CarT_P sheet`);
    
    // Process all records (not just latest per car)
    const records = [];
    
    data.forEach((row, index) => {
      // Get car number with fallback to display values
      const rawCar = row[iCarNo];
      const dispCar = disp[index][iCarNo];
      const carNumber = String((rawCar == null || String(rawCar).trim() === '') ? dispCar : rawCar).trim();
      
      if (!carNumber) return; // Skip rows without car number
      
      const record = {
        refNumber: iRef >= 0 ? (row[iRef] || '') : '',
        dateTime: iDate >= 0 ? (row[iDate] || disp[index][iDate] || '') : '',
        project: iProject >= 0 ? (row[iProject] || '') : '',
        team: iTeam >= 0 ? (row[iTeam] || '') : '',
        carNumber: carNumber,
        make: iMake >= 0 ? (row[iMake] || '') : '',
        model: iModel >= 0 ? (row[iModel] || '') : '',
        usageType: iUsage >= 0 ? (row[iUsage] || '') : '',
        contractType: iContract >= 0 ? (row[iContract] || '') : '',
        owner: iOwner >= 0 ? (row[iOwner] || '') : '',
        status: iStatus >= 0 ? String(row[iStatus] || disp[index][iStatus] || '').trim().toUpperCase() : '',
        remarks: iRemarks >= 0 ? (row[iRemarks] || '') : '',
        stars: iStars >= 0 ? (row[iStars] || 0) : 0,
        submitter: iSubmit >= 0 ? (row[iSubmit] || '') : '',
        rowIndex: index + 2
      };
      
      records.push(record);
    });
    
    // Sort by date (newest first) or by row index if no date
    records.sort((a, b) => {
      const dateA = a.dateTime ? new Date(a.dateTime) : new Date(0);
      const dateB = b.dateTime ? new Date(b.dateTime) : new Date(0);
      
      if (dateA.getTime() !== dateB.getTime()) {
        return dateB.getTime() - dateA.getTime(); // Newest first
      }
      
      return b.rowIndex - a.rowIndex; // Fallback to row order (newest first)
    });
    
    console.log(`Returning ${records.length} Vehicle records`);
    
    return {
      success: true,
      records: records,
      totalRecords: records.length,
      sheetName: sh.getName()
    };
    
  } catch (error) {
    console.error('Error getting all Vehicle data:', error);
    return {
      success: false,
      error: String(error),
      message: 'Failed to retrieve Vehicle data'
    };
  }
}

/**
 * Sync the Vehicle sheet so it contains one row per unique vehicle number with
 * the most recent status and metadata from CarT_P. Rewrites only the managed
 * columns (matching CarT_P header) while preserving other sheet content.
 */
function syncVehicleSheetFromCarTP(){
  try {
    const carRows = _readCarTP_objects_();
    if (!carRows.length) {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sh = _openVehicleSheet_();
      if (!sh) {
        sh = ss.insertSheet('Vehicle');
      }
      const header = ['Ref','Date and time of entry','Project','Team','R.Beneficiary','Vehicle Number','Make','Model','Category','Usage Type','Owner','Status','Last Users remarks','Ratings','Submitter username'];
      if (sh.getMaxColumns() < header.length) {
        sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
      }
      sh.getRange(1,1,1,header.length).setValues([header]);
      sh.setFrozenRows(1);
      const maxRows = sh.getMaxRows();
      if (maxRows > 1) {
        sh.getRange(2,1,maxRows-1,header.length).clearContent();
      }
      return { ok:true, updated:0 };
    }

    const latestByCar = new Map();
    for (let i = 0; i < carRows.length; i++) {
      const row = carRows[i];
      const carNumber = String(row['Vehicle Number'] || '').trim();
      if (!carNumber) continue;
      const key = carNumber.toUpperCase();
      const ts = (typeof row._ts === 'number' && !isNaN(row._ts)) ? row._ts : 0;
      const seq = i + 1;
      const prev = latestByCar.get(key);
      if (!prev || ts > prev.ts || (ts === prev.ts && seq > prev.seq)) {
        latestByCar.set(key, { ts, seq, row });
      }
    }

    const records = Array.from(latestByCar.values())
      .map(entry => entry.row)
      .sort((a,b) => String(a['Vehicle Number'] || '').localeCompare(String(b['Vehicle Number'] || '')));

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = _openVehicleSheet_();
    if (!sh) {
      sh = ss.insertSheet('Vehicle');
    }

    const header = ['Ref','Date and time of entry','Project','Team','R.Beneficiary','Vehicle Number','Make','Model','Category','Usage Type','Owner','Status','Last Users remarks','Ratings','Submitter username'];
    if (sh.getMaxColumns() < header.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
    }

    const neededRows = records.length + 1;
    if (sh.getMaxRows() < neededRows) {
      sh.insertRowsAfter(sh.getMaxRows(), neededRows - sh.getMaxRows());
    }

    sh.getRange(1,1,1,header.length).setValues([header]);
    sh.setFrozenRows(1);

    // Clear managed columns before writing to avoid stale values
    const maxRows = sh.getMaxRows();
    if (maxRows > 1) {
      sh.getRange(2,1,maxRows-1,header.length).clearContent();
    }

    if (records.length) {
      const values = records.map(r => {
        const dateVal = r['Date and time of entry'];
        const isDate = dateVal instanceof Date;
        const responsible = String(r['R.Beneficiary'] || r['R. Ben'] || r.responsibleBeneficiary || '').trim();
        return [
          r.Ref || r['Reference Number'] || '',
          isDate ? dateVal : (dateVal || ''),
          r.Project || '',
          r.Team || '',
          responsible,
          r['Vehicle Number'] || '',
          r.Make || '',
          r.Model || '',
          r.Category || '',
          r['Usage Type'] || '',
          r.Owner || '',
          _normStatus_(r.Status),
          r['Last Users remarks'] || '',
          r.Ratings || '',
          r['Submitter username'] || ''
        ];
      });
      sh.getRange(2,1,values.length,header.length).setValues(values);
    }

    try { sh.autoResizeColumns(1, header.length); } catch (_autoErr) { /* optional */ }

    return { ok:true, updated: records.length };
  } catch (e) {
    console.error('syncVehicleSheetFromCarTP error:', e);
    return { ok:false, error: String(e) };
  }
}

/** Locate CarT_P sheet with tolerant matching */
function _openCarTP_(){
  try{
    console.log('[OPEN_CARTP] Attempting to open CarT_P sheet...');
    console.log('[OPEN_CARTP] Using CAR_SHEET_ID:', CAR_SHEET_ID);
    
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    console.log('[OPEN_CARTP] Spreadsheet opened successfully:', {
      spreadsheetId: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl()
    });
    
    let sh = ss.getSheetByName('CarT_P');
    if (sh) {
      console.log('[OPEN_CARTP] ✅ Found CarT_P sheet directly');
      return sh;
    }
    
    console.log('[OPEN_CARTP] CarT_P not found, searching with tolerant matching...');
    const all = ss.getSheets();
    console.log('[OPEN_CARTP] Available sheets:', all.map(sheet => sheet.getName()));
    
    for (let i=0;i<all.length;i++){
      const n = (all[i].getName()||'').toLowerCase().replace(/\s+/g,'');
      console.log('[OPEN_CARTP] Checking sheet:', all[i].getName(), 'normalized:', n);
      
      if (n === 'cart_p' || n === 'cartp' || (n.includes('car') && (n.includes('t_p')||n.includes('tp')||n.includes('release')))){
        console.log('[OPEN_CARTP] ✅ Found matching sheet:', all[i].getName());
        return all[i];
      }
    }
    
    console.error('[OPEN_CARTP] ❌ No CarT_P sheet found in spreadsheet');
    return null;
  }catch(e){
    console.error('[OPEN_CARTP] ❌ Error opening CAR sheet:', e);
    console.error('[OPEN_CARTP] Error details:', {
      message: e.message,
      stack: e.stack,
      name: e.name
    });
    return null;
  }
}

/**
 * Test function to verify CarT_P sheet access and data writing
 * Call this from frontend to debug the issue
 */
function testCarTPAccess() {
  try {
    console.log('[TEST_CARTP] Starting CarT_P access test...');
    console.log('[TEST_CARTP] CAR_SHEET_ID:', CAR_SHEET_ID);
    console.log('[TEST_CARTP] SHEET_ID:', SHEET_ID);
    
    // Test 1: Can we open the spreadsheet?
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    console.log('[TEST_CARTP] ✅ Spreadsheet opened:', {
      id: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl()
    });
    
    // Test 2: Can we find CarT_P sheet?
    const sh = _openCarTP_();
    if (!sh) {
      console.error('[TEST_CARTP] ❌ CarT_P sheet not found');
      return { ok: false, error: 'CarT_P sheet not found' };
    }
    
    console.log('[TEST_CARTP] ✅ CarT_P sheet found:', {
      name: sh.getName(),
      lastRow: sh.getLastRow(),
      lastColumn: sh.getLastColumn()
    });
    
    // Test 3: Can we write a test row?
    const testRow = [
      'TEST-' + Date.now(),
      new Date(),
      'TEST_PROJECT',
      'TEST_TEAM',
      'TEST_BENEFICIARY',
      'TEST-VEHICLE-123',
      'Test Make',
      'Test Model',
      'Test Category',
      'Test Usage',
      'Test Owner',
      'IN USE',
      'Test Remarks',
      5,
      'test@example.com',
      new Date()
    ];
    
    const startRow = sh.getLastRow() + 1;
    console.log('[TEST_CARTP] Attempting to write test row at row:', startRow);
    
    sh.getRange(startRow, 1, 1, testRow.length).setValues([testRow]);
    
    console.log('[TEST_CARTP] ✅ Test row written successfully');
    console.log('[TEST_CARTP] New last row:', sh.getLastRow());
    
    // Test 4: Can we read back the data?
    const writtenData = sh.getRange(startRow, 1, 1, testRow.length).getValues()[0];
    console.log('[TEST_CARTP] ✅ Data read back:', writtenData);
    
    return {
      ok: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      sheetName: sh.getName(),
      lastRow: sh.getLastRow(),
      testRowWritten: startRow,
      testData: writtenData
    };
    
  } catch (error) {
    console.error('[TEST_CARTP] ❌ Test failed:', error);
    return {
      ok: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Wrapper function for assignCarToTeam to ensure proper parameter handling
 */
function assignCarToTeamWrapper(payloadJson) {
  try {
    console.log('[WRAPPER] assignCarToTeamWrapper called');
    console.log('[WRAPPER] Received payloadJson type:', typeof payloadJson);
    console.log('[WRAPPER] Received payloadJson:', payloadJson);
    
    let payload;
    if (typeof payloadJson === 'string') {
      console.log('[WRAPPER] Parsing JSON string...');
      payload = JSON.parse(payloadJson);
    } else if (typeof payloadJson === 'object' && payloadJson !== null) {
      console.log('[WRAPPER] Using object directly...');
      payload = payloadJson;
    } else {
      console.error('[WRAPPER] Invalid payload type:', typeof payloadJson);
      return { ok: false, error: 'Invalid payload type: ' + typeof payloadJson };
    }
    
    console.log('[WRAPPER] Parsed payload:', JSON.stringify(payload, null, 2));
    console.log('[WRAPPER] Calling assignCarToTeam...');
    
    return assignCarToTeam(payload);
    
  } catch (error) {
    console.error('[WRAPPER] Error in assignCarToTeamWrapper:', error);
    return { ok: false, error: 'Wrapper error: ' + error.message };
  }
}

/** Locate Vehicle sheet with tolerant matching */
function _openVehicleSheet_(){
  try{
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Vehicle');
    if (sh) return sh;
    const all = ss.getSheets();
    for (let i=0;i<all.length;i++){
      const n = (all[i].getName()||'').toLowerCase().replace(/\s+/g,'');
      if (n === 'vehicle' || n === 'vehicles' || (n.includes('vehicle'))){
        return all[i];
      }
    }
    return null;
  }catch(e){
    console.error('Open Vehicle sheet failed:', e);
    return null;
  }
}

/** Locate Vehicle_Released summary sheet */
function _openVehicleReleasedSheet_(){
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Vehicle_Released');
    if (sh) return sh;
    const all = ss.getSheets();
    for (let i = 0; i < all.length; i++) {
      const name = String(all[i].getName() || '').toLowerCase().replace(/\s+/g, '');
      if (name === 'vehicle_released' || name === 'vehiclereleased') {
        return all[i];
      }
    }
    return null;
  } catch (e) {
    console.error('Open Vehicle_Released sheet failed:', e);
    return null;
  }
}

/** Debug helper: return CarT_P overview (sheet id, names, header, size) */
function debugCarTPOverview(){
  try{
    const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
    const sheets = ss.getSheets().map(s=>s.getName());
    const sh = _openCarTP_();
    if (!sh){
      console.log('CarT_P sheet not found in sheets:', sheets);
      return { ok:false, reason:'CarT_P sheet not found', sheetId: CAR_SHEET_ID, sheets };
    }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const headDisp = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const headNorm = headDisp.map(h => (String(h||'').trim().toLowerCase()));
    return { ok:true, sheetId: CAR_SHEET_ID, sheetName: sh.getName(), sheets, lastRow, lastCol, header: headDisp, headerNorm: headNorm };
  }catch(e){
    return { ok:false, error: String(e), sheetId: CAR_SHEET_ID };
  }
}

/** Debug: extract raw rows as parsed by header mapping (no RELEASE filter, no grouping) */
function debugExtractCars(){
  const sh = _openCarTP_();
  if (!sh) return { ok:false, reason:'CarT_P not found' };
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const IX = _headerIndex_(head);
  function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
  const iDate     = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
  // Car Number with tolerant fallback
  let iCarNo = -1; try{ iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); }catch(_){ iCarNo = -1; }
  if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
  if (iCarNo < 0) throw new Error('Car Number column not found');
  const iMake     = idx(['Make','Car Make','Brand'], false);
  const iModel    = idx(['Model','Car Model'], false);
  const iUsage    = idx(['Usage Type','Usage','Use Type'], false);
  const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
  const iOwner    = idx(['Owner','Owner Name','Owner Info'], false);
  const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
  const rng  = sh.getRange(2,1,Math.max(0,lastRow-1), lastCol);
  const data = rng.getValues();
  const disp = rng.getDisplayValues();
  const out = [];
  for (let r=0;r<data.length;r++){
    const row = data[r];
    const car = String((row[iCarNo]==null || String(row[iCarNo]).trim()==='') ? disp[r][iCarNo] : row[iCarNo]).trim();
    const rec = {
      row: r+2,
      carNumber: car,
      status: iStatus>=0? (row[iStatus] || disp[r][iStatus]) : '',
      make: iMake>=0? (row[iMake] || disp[r][iMake]) : '',
      model: iModel>=0? (row[iModel] || disp[r][iModel]) : '',
      usageType: iUsage>=0? (row[iUsage] || disp[r][iUsage]) : '',
      contractType: iContract>=0? (row[iContract] || disp[r][iContract]) : '',
      owner: iOwner>=0? (row[iOwner] || disp[r][iOwner]) : '',
      dateTime: iDate>=0? (row[iDate] || disp[r][iDate]) : ''
    };
    if (car || rec.make || rec.model || rec.owner || rec.status) out.push(rec);
  }
  return { ok:true, count: out.length, rows: out.slice(0,50) };
}

/** Debug: show how columns are detected + sample car numbers */
function debugCarTPMapping(){
  try{
    const sh = _openCarTP_();
    if (!sh) return { ok:false, reason:'CarT_P not found' };
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iRef      = idx(['Reference Number','Ref','Ref Number'], false);
    const iDate     = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    let iCarNo = -1; try{ iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); }catch(_){ iCarNo = -1; }
    const iCarNoHeur = _findCarNumberColumn_(head);
    if (iCarNo < 0 && iCarNoHeur >= 0) iCarNo = iCarNoHeur;
    const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iMake     = idx(['Make','Car Make','Brand'], false);
    const iModel    = idx(['Model','Car Model'], false);

    const data = sh.getRange(2,1,Math.max(0,lastRow-1), lastCol).getValues();
    const sample = [];
    const seen = new Set();
    if (iCarNo >= 0){
      for (let r=0;r<data.length && sample.length<10;r++){
        const row = data[r];
        const car = String(row[iCarNo]||'').trim();
        if (!car) continue;
        if (!seen.has(car)){
          seen.add(car);
          sample.push({
            row: r+2,
            carNumber: car,
            status: iStatus>=0 ? row[iStatus] : '',
            make: iMake>=0 ? row[iMake] : '',
            model: iModel>=0 ? row[iModel] : ''
          });
        }
      }
    }

    // Also compute grouping size using same logic as getAvailableCars
    let groups = 0;
    if (iCarNo >= 0){
      const map = new Map();
      for (let r=0;r<data.length;r++){
        const row = data[r];
        const car = String(row[iCarNo]||'').trim();
        if (!car) continue;
        const date = iDate>=0 ? row[iDate] : null;
        const tsFromDate = (date instanceof Date) ? date.getTime() : (date ? (new Date(date)).getTime() : NaN);
        const ts = isNaN(tsFromDate) ? (r+1) : tsFromDate;
        const rec = { ts: ts, row: r+2 };
        const prev = map.get(car);
        if (!prev || ts >= prev.ts) map.set(car, rec);
      }
      groups = map.size;
    }

    return {
      ok:true,
      lastRow, lastCol,
      header: head,
      carIndex: iCarNo,
      carHeader: iCarNo>=0 ? head[iCarNo] : '',
      carIndexHeuristic: iCarNoHeur,
      sampleCount: sample.length,
      sample,
      groupedVehicles: groups
    };
  }catch(e){
    return { ok:false, error:String(e) };
  }
}

/** Debug: simple write probe — writes a test string into first blank cell in column A (Ref) */
function debugWriteTest(value){
  try{
    const sh = _openCarTP_();
    if (!sh) return { ok:false, error:'CarT_P not found' };
    const lastRow = sh.getLastRow();
    // scan for first blank in column A starting from row 2
    let targetRow = 2;
    for (; targetRow <= lastRow; targetRow++){
      const v = sh.getRange(targetRow, 1).getDisplayValue();
      if (!String(v||'').trim()) break; // found blank
    }
    if (targetRow > lastRow) targetRow = lastRow + 1; // append if no internal blank
    const text = (value == null || String(value).trim() === '') ? ('TEST-'+ new Date().getTime()) : String(value);
    sh.getRange(targetRow, 1).setValue(text);
    return { ok:true, row: targetRow, col: 1, value: text };
  }catch(e){
    return { ok:false, error: String(e) };
  }
}

function _ensureCarTPSchemaAndHeader_(sh){
  if (!sh) return [];
  const hasRows = sh.getLastRow() > 0;
  const lastCol = sh.getLastColumn();
  if (!hasRows || lastCol === 0){
    const header = [
      'Ref',
      'Date and time of entry',
      'Project',
      'Team',
      'R.Beneficiary',
      'Vehicle Number',
      'Make',
      'Model',
      'Category',
      'Usage Type',
      'Owner',
      'Status',
      'Last Users remarks',
      'Ratings',
      'Submitter username'
    ];
    sh.getRange(1,1,1,header.length).setValues([header]);
  }

  const ensureResponsibleColumn = () => {
    try {
      const currentHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      const needsColumn = !currentHead.some(function(h){
        const norm = String(h||'').trim().toLowerCase().replace(/[^a-z]/g,'');
        return norm === 'rbeneficiary' || norm === 'responsiblebeneficiary';
      });
      if (!needsColumn) return;
      const teamIdx = currentHead.findIndex(function(h){
        const norm = String(h||'').trim().toLowerCase();
        return norm === 'team' || norm === 'team name';
      });
      if (teamIdx >= 0) {
        sh.insertColumnAfter(teamIdx + 1);
        sh.getRange(1, teamIdx + 2).setValue('R.Beneficiary');
      } else {
        const col = sh.getLastColumn();
        sh.insertColumnAfter(col);
        sh.getRange(1, col + 1).setValue('R.Beneficiary');
      }
    } catch (err) {
      console.warn('Unable to ensure R.Beneficiary column', err);
    }
  };

  const ensureShortResponsibleColumn = () => {
    try {
      const currentHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      const normalized = currentHead.map(function(h){
        return String(h || '').trim().toLowerCase().replace(/[^a-z]/g, '');
      });
      if (normalized.indexOf('rben') >= 0) return;
      const timeIdx = normalized.indexOf('rbentime');
      if (timeIdx >= 0) {
        sh.insertColumnAfter(timeIdx + 1);
        sh.getRange(1, timeIdx + 2).setValue('R. Ben');
        return;
      }
      const respIdx = normalized.indexOf('rbeneficiary');
      if (respIdx >= 0) {
        sh.insertColumnAfter(respIdx + 1);
        sh.getRange(1, respIdx + 2).setValue('R. Ben');
        return;
      }
      const col = sh.getLastColumn();
      sh.insertColumnAfter(col);
      sh.getRange(1, col + 1).setValue('R. Ben');
    } catch (err) {
      console.warn('Unable to ensure R. Ben column', err);
    }
  };

  const ensureFullResponsibleColumn = () => {
    try {
      const currentHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      const normalized = currentHead.map(function(h){
        return String(h || '').trim().toLowerCase().replace(/[^a-z]/g, '');
      });
      if (normalized.indexOf('responsiblebeneficiary') >= 0 || normalized.indexOf('nameofresponsiblebeneficiary') >= 0) {
        return;
      }
      const respIdx = normalized.indexOf('rbeneficiary');
      if (respIdx >= 0) {
        sh.insertColumnAfter(respIdx + 1);
        sh.getRange(1, respIdx + 2).setValue('Responsible Beneficiary');
        return;
      }
      const col = sh.getLastColumn();
      sh.insertColumnAfter(col);
      sh.getRange(1, col + 1).setValue('Responsible Beneficiary');
    } catch (err) {
      console.warn('Unable to ensure Responsible Beneficiary column', err);
    }
  };

  const ensureRBTimeColumn = () => {
    try {
      const currentHead = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
      const normalized = currentHead.map((h) => String(h || '').trim().toLowerCase().replace(/[^a-z]/g, ''));
      const timeIdx = normalized.indexOf('rbentime');
      const submitIdx = normalized.indexOf('submitterusername');
      if (timeIdx >= 0) {
        if (submitIdx >= 0 && timeIdx !== submitIdx + 1) {
          const currentCol = timeIdx + 1;
          const destination = submitIdx + 2;
          if (currentCol !== destination) {
            try {
              sh.moveColumns(sh.getRange(1, currentCol, sh.getMaxRows(), 1), destination);
            } catch (moveErr) {
              console.warn('Unable to reposition R.Ben Time column', moveErr);
            }
          }
        }
        return;
      }
      if (submitIdx >= 0) {
        sh.insertColumnAfter(submitIdx + 1);
        sh.getRange(1, submitIdx + 2).setValue('R.Ben Time');
      } else {
        const col = sh.getLastColumn();
        sh.insertColumnAfter(col);
        sh.getRange(1, col + 1).setValue('R.Ben Time');
      }
    } catch (err) {
      console.warn('Unable to ensure R.Ben Time column', err);
    }
  };

  ensureResponsibleColumn();
  ensureShortResponsibleColumn();
  ensureFullResponsibleColumn();
  ensureRBTimeColumn();

  const width = Math.max(15, sh.getLastColumn());
  return sh.getRange(1,1,1,width).getDisplayValues()[0];
}

/**
 * Assign a vehicle to a team (and optionally multiple beneficiaries) and log entries in CarT_P.
 * - Does not overwrite existing UI values; frontend only passes beneficiaries that were empty.
 * - Writes one row per beneficiary with Status='IN USE' to indicate assignment.
 * - Uses tolerant header mapping; creates CarT_P with expected headers if missing.
 */
function assignCarToTeam(payload){
  try{
    console.log('[ASSIGN_CAR] Function called with payload:', JSON.stringify(payload, null, 2));
    
    if(!payload) {
      console.error('[ASSIGN_CAR] No payload provided');
      return { ok:false, error:'No payload' };
    }
    
    const project = _norm(payload.project);
    const team    = _norm(payload.team);
    const carNum  = _norm(payload.carNumber);
    const rawBeneficiaries = Array.isArray(payload.beneficiaries) ? payload.beneficiaries : [];
    const bens = [];
    const seenBeneficiaries = new Set();
    rawBeneficiaries.forEach(name => {
      const trimmed = _norm(name);
      if (!trimmed) return;
      const key = trimmed.toLowerCase();
      if (seenBeneficiaries.has(key)) return;
      seenBeneficiaries.add(key);
      bens.push(trimmed);
    });
    const car     = payload.car || {};
    
    console.log('[ASSIGN_CAR] Parsed data:', { project, team, carNum, beneficiaries: bens });
    
    if(!project || !team || !carNum) {
      console.error('[ASSIGN_CAR] Missing required fields:', { project, team, carNum });
      return { ok:false, error:'Missing project/team/carNumber' };
    }

    console.log('[ASSIGN_CAR] Opening CarT_P sheet...');
    const sh = _openCarTP_();
    if(!sh) {
      console.error('[ASSIGN_CAR] CarT_P sheet not found');
      return { ok:false, error:'CarT_P not found' };
    }
    
    console.log('[ASSIGN_CAR] CarT_P sheet found:', {
      sheetName: sh.getName(),
      spreadsheetId: sh.getParent().getId(),
      lastRow: sh.getLastRow(),
      lastColumn: sh.getLastColumn()
    });

    const head = _ensureCarTPSchemaAndHeader_(sh);
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iRef   = idx(['Ref','Reference Number','Ref Number'], false);
    const iDate  = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    let iCarNo   = -1; try{ iCarNo = idx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car'], false);}catch(_){ iCarNo=-1; }
    if(iCarNo<0) iCarNo = _findCarNumberColumn_(head);
    const iProj  = idx(['Project'], false);
    const iTeam  = idx(['Team','Team Name'], false);
    const iResp  = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible','R. Ben','R Ben'], false);
    const iRespShort = idx(['R. Ben','R Ben'], false);
    const iRespFull = idx(['Responsible Beneficiary','Name of Responsible beneficiary'], false);
    const iMake  = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iCat   = idx(['Category','Vehicle Category','Cat'], false);
    const iUse   = idx(['Usage Type','Usage','Use Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);
    const iStat  = idx(['Status','In Use/Release','In Use / release','In Use'], false);
    const iRem   = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iRate  = idx(['Ratings','Stars','Rating'], false);
    const iSub   = idx(['Submitter username','Submitter','User'], false);
    const iRespTime = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);

    // Generate a base ref prefix
    const baseTs = Date.now();
    const userEmail = (function(){ try{ return Session.getActiveUser().getEmail()||''; }catch(e){ return ''; } })();
    const submitter = userEmail || 'System';

    // Build rows
    const now = new Date();
    const rows = [];
    const make  = _norm(car.make);
    const model = _norm(car.model);
    const cat   = _norm(car.category);
    const use   = _norm(car.usageType);
    const owner = _norm(car.owner);

    const responsible = _norm(payload.responsibleBeneficiary) || (bens.length ? bens[0] : '');
    const responsibleKey = responsible.toLowerCase();

    const targetBeneficiaries = [];
    const targetSeen = new Set();
    if (responsible) {
      targetBeneficiaries.push(responsible);
      targetSeen.add(responsibleKey);
    }
    bens.forEach(name => {
      const key = name.toLowerCase();
      if (targetSeen.has(key)) return;
      targetSeen.add(key);
      targetBeneficiaries.push(name);
    });
    if (!targetBeneficiaries.length) {
      console.warn('[ASSIGN_CAR] No beneficiaries supplied; creating placeholder row.');
      targetBeneficiaries.push(responsible || '');
    }

    const toRow = (beneficiaryName, isResponsibleRow)=>{
      const columnCount = Math.max(head.length, sh.getLastColumn());
      const arr = new Array(columnCount).fill('');
      if(iRef >=0)  arr[iRef]  = 'REF-'+baseTs+'-'+Math.random().toString(36).slice(2,6).toUpperCase();
      if(iDate>=0)  arr[iDate] = now;
      if(iProj>=0)  arr[iProj] = project;
      if(iTeam>=0)  arr[iTeam] = team;
      if(iResp>=0){
        if(isResponsibleRow){
          arr[iResp] = responsible || _norm(beneficiaryName);
        } else {
          arr[iResp] = _norm(beneficiaryName) || responsible;
        }
      }
      if (iRespShort >= 0) {
        if (responsible) {
          arr[iRespShort] = responsible;
        } else {
          arr[iRespShort] = _norm(beneficiaryName);
        }
      }
      if(iCarNo>=0) arr[iCarNo]= carNum;
      if(iMake>=0)  arr[iMake] = make;
      if(iModel>=0) arr[iModel]= model;
      if(iCat>=0)   arr[iCat]  = cat;
      if(iUse>=0)   arr[iUse]  = use;
      if(iOwner>=0) arr[iOwner]= owner;
      if(iStat>=0)  arr[iStat] = 'IN USE';
      if(iRem>=0)   arr[iRem]  = '';
      if(iRate>=0)  arr[iRate] = 0;
      if(iSub>=0)   arr[iSub]  = submitter;
      if (iRespFull >= 0) {
        arr[iRespFull] = responsible || '';
      }
      if (iRespTime >= 0) {
        arr[iRespTime] = now;
      }
      return arr;
    };

    targetBeneficiaries.forEach(name => {
      const key = _norm(name).toLowerCase();
      const isResponsible = !!responsible && key === responsibleKey;
      rows.push(toRow(name, isResponsible));
    });

    if(rows.length){
      const start = sh.getLastRow() + 1;
      console.log('[ASSIGN_CAR] Writing data to CarT_P:', {
        startRow: start,
        rowCount: rows.length,
        columnCount: rows[0].length,
        sheetName: sh.getName(),
        spreadsheetId: sh.getParent().getId()
      });
      
      console.log('[ASSIGN_CAR] Data being written:', rows);
      
      sh.getRange(start,1,rows.length,rows[0].length).setValues(rows);
      
      console.log('[ASSIGN_CAR] ✅ Data successfully written to CarT_P sheet');
      console.log('[ASSIGN_CAR] New last row after write:', sh.getLastRow());

      try {
        rows.forEach(function(rowValues){
          const summaryObj = {};
          for (let i = 0; i < VEHICLE_SUMMARY_HEADER.length; i++) {
            summaryObj[VEHICLE_SUMMARY_HEADER[i]] = rowValues[i] || '';
          }
          summaryObj.status = summaryObj.Status || 'IN USE';
          upsertVehicleSummaryRow('Vehicle_InUse', summaryObj, 'beneficiary');
        });
      } catch (summaryErr) {
        console.warn('[ASSIGN_CAR] Vehicle_InUse summary update failed:', summaryErr);
      }
    } else {
      console.warn('[ASSIGN_CAR] No rows to write - this should not happen');
    }

    try {
      refreshVehicleStatusSheets();
    } catch (refreshErr) {
      console.error('Vehicle summary refresh (assign) failed:', refreshErr);
    }

    try {
      syncVehicleSheetFromCarTP();
    } catch (syncErr) {
      console.error('Vehicle sheet sync (assign) failed:', syncErr);
    }

    try { CacheService.getScriptCache().remove('VEH_PICKER_V1'); } catch (_cacheDropErr) { /* ignore */ }

    const result = { ok:true, written: rows.length };
    console.log('[ASSIGN_CAR] ✅ COMPLETED SUCCESSFULLY - Returning result:', result);
    return result;
  }catch(e){
    console.error('[ASSIGN_CAR] ❌ ERROR occurred:', e);
    console.error('[ASSIGN_CAR] Error details:', {
      message: e.message,
      stack: e.stack,
      name: e.name
    });
    return { ok:false, error:String(e) };
  }
}

function _vehicleNumberExists_(vehicleNumber) {
  const result = { exists: false, addedBy: '', source: '', status: '' };
  const target = _norm(vehicleNumber).toUpperCase();
  if (!target) return result;
  try {
    const vehicleSheet = _openVehicleSheet_();
    if (vehicleSheet) {
      const vehicleCol = _columnLetterToIndex_('AE');
      const addedByCol = _columnLetterToIndex_('AF');
      if (vehicleCol > 0) {
        const lastRow = vehicleSheet.getLastRow();
        if (lastRow > 1) {
          const rowCount = lastRow - 1;
          const vehicleValues = vehicleSheet.getRange(2, vehicleCol, rowCount, 1).getDisplayValues();
          let addedByValues = null;
          if (addedByCol > 0) {
            addedByValues = vehicleSheet.getRange(2, addedByCol, rowCount, 1).getDisplayValues();
          }
          for (let i = 0; i < vehicleValues.length; i++) {
            const candidate = _norm(vehicleValues[i][0]).toUpperCase();
            if (!candidate) continue;
            if (candidate === target) {
              result.exists = true;
              result.source = 'Vehicle';
              result.addedBy = addedByValues && addedByValues[i] ? _norm(addedByValues[i][0]) : '';
              return result;
            }
          }
        }
      }
    }
  } catch (sheetErr) {
    console.warn('[NEW_VEHICLE] Vehicle sheet duplicate check failed:', sheetErr);
  }

  try {
    const sh = _openCarTP_();
    if (!sh) return result;
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return result;
    const head = _ensureCarTPSchemaAndHeader_(sh);
    const IX = _headerIndex_(head);
    let iCarNo = -1;
    try { iCarNo = IX.get(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car']); } catch (_err) { iCarNo = -1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    let columnIndex = iCarNo >= 0 ? iCarNo + 1 : 7; // default to column G when headers missing
    if (columnIndex < 1) return result;
    let iStatus = -1;
    try { iStatus = IX.get(['Status','In Use/Release','In Use / release','In Use']); } catch (_err) { iStatus = -1; }
    const rowCount = lastRow - 1;
    const vehicleValues = sh.getRange(2, columnIndex, rowCount, 1).getDisplayValues();
    let statusValues = null;
    if (iStatus >= 0) {
      statusValues = sh.getRange(2, iStatus + 1, rowCount, 1).getDisplayValues();
    }
    for (let i = 0; i < vehicleValues.length; i++) {
      const candidate = _norm(vehicleValues[i][0]).toUpperCase();
      if (!candidate) continue;
      if (candidate === target) {
        result.exists = true;
        result.source = 'CarT_P';
        result.status = statusValues && statusValues[i] ? _norm(statusValues[i][0]) : '';
        return result;
      }
    }
  } catch (carErr) {
    console.warn('[NEW_VEHICLE] CarT_P duplicate check failed:', carErr);
  }

  return result;
}

function _upsertVehicleCatalogRow_(entry) {
  try {
    const sh = _openVehicleSheet_();
    if (!sh) {
      console.warn('[NEW_VEHICLE] Vehicle sheet unavailable; skipping catalog update');
      return;
    }

    const makeCol = _columnLetterToIndex_('Z');
    const modelCol = _columnLetterToIndex_('AA');
    const categoryCol = _columnLetterToIndex_('AB');
    const usageCol = _columnLetterToIndex_('AC');
    const ownerCol = _columnLetterToIndex_('AD');
    const vehicleCol = _columnLetterToIndex_('AE');
    const addedByCol = _columnLetterToIndex_('AF');

    if (makeCol < 1 || vehicleCol < 1 || addedByCol < 1) {
      console.warn('[NEW_VEHICLE] Catalog column mapping failed; skipping.');
      return;
    }

    const startCol = makeCol;
    const colCount = addedByCol - startCol + 1;
    if (colCount <= 0) return;

    const maxCols = sh.getMaxColumns();
    if (maxCols < addedByCol) {
      sh.insertColumnsAfter(maxCols, addedByCol - maxCols);
    }

    const headerRange = sh.getRange(1, startCol, 1, colCount);
    const headerCurrent = headerRange.getDisplayValues()[0];
    const headerTemplate = ['Make','Model','Category','Usage Type','Owner','Vehicle Number','Added By'];
    let headerNeedsUpdate = false;
    const headerWrite = headerTemplate.map((label, idx) => {
      const current = headerCurrent[idx];
      if (typeof current === 'string' && current.trim()) {
        return current;
      }
      headerNeedsUpdate = true;
      return label;
    });
    if (headerNeedsUpdate) {
      headerRange.setValues([headerWrite]);
    }

    const vehicleNumber = _norm(entry.vehicleNumber).toUpperCase();
    if (!vehicleNumber) {
      console.warn('[NEW_VEHICLE] No vehicle number provided for catalog update');
      return;
    }

    const make = _norm(entry.make);
    const model = _norm(entry.model);
    const category = _norm(entry.category);
    const usageType = _norm(entry.usageType);
    const owner = _norm(entry.owner);
    const addedBy = _norm(entry.addedBy);

    const lastRow = sh.getLastRow();
    const bodyRows = Math.max(0, lastRow - 1);
    let targetRow = -1;
    if (bodyRows > 0) {
      const vehicleValues = sh.getRange(2, vehicleCol, bodyRows, 1).getDisplayValues();
      for (let i = 0; i < vehicleValues.length; i++) {
        const value = _norm(vehicleValues[i][0]).toUpperCase();
        if (value && value === vehicleNumber) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const ensureRow = targetRow > 0 ? targetRow : Math.max(lastRow + 1, 2);
    const neededRows = Math.max(ensureRow, 2);
    const maxRows = sh.getMaxRows();
    if (maxRows < neededRows) {
      sh.insertRowsAfter(maxRows, neededRows - maxRows);
    }

    let rowValues = [make, model, category, usageType, owner, vehicleNumber, addedBy];

    if (targetRow > 0) {
      const existing = sh.getRange(targetRow, startCol, 1, colCount).getDisplayValues()[0];
      const merged = [
        make || existing[0] || '',
        model || existing[1] || '',
        category || existing[2] || '',
        usageType || existing[3] || '',
        owner || existing[4] || '',
        vehicleNumber || existing[5] || '',
        addedBy || existing[6] || ''
      ];
      rowValues = merged;
    }

    sh.getRange(ensureRow, startCol, 1, colCount).setValues([rowValues]);
    try {
      console.log('[NEW_VEHICLE] Vehicle catalog updated', {
        row: ensureRow,
        vehicleNumber: vehicleNumber,
        make: rowValues[0],
        model: rowValues[1]
      });
    } catch (_logErr) {
      // logging best effort
    }
  } catch (err) {
    console.warn('[NEW_VEHICLE] Failed to upsert vehicle catalog row', err);
  }
}

function submitNewVehicleRelease(payload){
  try {
    console.log('[NEW_VEHICLE] submitNewVehicleRelease called with payload:', JSON.stringify(payload || {}, null, 2));
    if (!payload || typeof payload !== 'object') {
      return { ok: false, error: 'No payload' };
    }

    const vehicleNumberInput = _norm(payload.vehicleNumber);
    if (!vehicleNumberInput) {
      return { ok: false, error: 'Vehicle number required' };
    }
    const vehicleNumber = vehicleNumberInput.toUpperCase();

    const make = _norm(payload.make);
    const model = _norm(payload.model);
    const category = _norm(payload.category);
    const usageType = _norm(payload.usageType);
    const owner = _norm(payload.owner);
    const project = _norm(payload.project);
    const team = _norm(payload.team);
    const responsibleBeneficiary = _norm(payload.responsibleBeneficiary);

    const duplicateCheck = _vehicleNumberExists_(vehicleNumber);
    if (duplicateCheck.exists) {
      const origin = duplicateCheck.source ? ` (${duplicateCheck.source})` : '';
      const addedByNote = duplicateCheck.addedBy ? ` — added by ${duplicateCheck.addedBy}` : '';
      return { ok: false, error: `Vehicle number ${vehicleNumber} already exists${origin}${addedByNote}` };
    }

    const sh = _openCarTP_();
    if (!sh) {
      return { ok: false, error: 'CarT_P not found' };
    }

    const head = _ensureCarTPSchemaAndHeader_(sh);
    const IX = _headerIndex_(head);
    const idx = (labels, required) => {
      try { return IX.get(labels); }
      catch (err) { if (required) throw err; return -1; }
    };

    const iRef   = idx(['Ref','Reference Number','Ref Number'], false);
    const iDate  = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    let iCarNo   = -1; try{ iCarNo = idx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car'], false);}catch(_){ iCarNo=-1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    const iProj  = idx(['Project'], false);
    const iTeam  = idx(['Team','Team Name'], false);
    const iResp  = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible','R. Ben','R Ben'], false);
    const iRespShort = idx(['R. Ben','R Ben'], false);
    const iRespFull = idx(['Responsible Beneficiary','Name of Responsible beneficiary'], false);
    const iMake  = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iCat   = idx(['Category','Vehicle Category','Cat'], false);
    const iUse   = idx(['Usage Type','Usage','Use Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);
    const iStat  = idx(['Status','In Use/Release','In Use / release','In Use'], false);
    const iRem   = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iRate  = idx(['Ratings','Stars','Rating'], false);
    const iSub   = idx(['Submitter username','Submitter','User'], false);
    const iRespTime = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);

    const now = new Date();
    const submitter = (function(){ try{ return Session.getActiveUser().getEmail() || ''; }catch(_){ return ''; } })() || 'System';
    const baseRef = 'NEW-' + Date.now() + '-' + Math.random().toString(36).slice(2,6).toUpperCase();

    const columnCount = Math.max(head.length, sh.getLastColumn());
    const row = new Array(columnCount).fill('');
    if (iRef >= 0) row[iRef] = baseRef;
    if (iDate >= 0) row[iDate] = now;
    if (iProj >= 0) row[iProj] = project;
    if (iTeam >= 0) row[iTeam] = team;
    if (iResp >= 0) row[iResp] = responsibleBeneficiary;
    if (iRespShort >= 0) row[iRespShort] = responsibleBeneficiary;
    if (iRespFull >= 0) row[iRespFull] = responsibleBeneficiary;
    if (iCarNo >= 0) row[iCarNo] = vehicleNumber;
    if (iMake >= 0) row[iMake] = make;
    if (iModel >= 0) row[iModel] = model;
    if (iCat >= 0) row[iCat] = category;
    if (iUse >= 0) row[iUse] = usageType;
    if (iOwner >= 0) row[iOwner] = owner;
    if (iStat >= 0) row[iStat] = 'RELEASE';
    if (iRem >= 0) row[iRem] = _norm(payload.remarks);
    if (iRate >= 0) row[iRate] = Number(payload.rating) || 0;
    if (iSub >= 0) row[iSub] = submitter;
    if (iRespTime >= 0) row[iRespTime] = now;

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, 1, columnCount).setValues([row]);
    console.log('[NEW_VEHICLE] Row appended at', startRow);

    try {
      _upsertVehicleCatalogRow_({
        vehicleNumber: vehicleNumber,
        make: make,
        model: model,
        category: category,
        usageType: usageType,
        owner: owner,
        addedBy: submitter
      });
    } catch (catalogErr) {
      console.warn('[NEW_VEHICLE] Catalog update failed:', catalogErr);
    }

    try {
      const summaryObj = {};
      for (let i = 0; i < VEHICLE_SUMMARY_HEADER.length; i++) {
        summaryObj[VEHICLE_SUMMARY_HEADER[i]] = row[i] || '';
      }
      summaryObj.Status = 'RELEASE';
      upsertVehicleSummaryRow('Vehicle_Released', summaryObj, 'vehicle');
    } catch (summaryErr) {
      console.warn('[NEW_VEHICLE] Vehicle_Released summary update failed:', summaryErr);
    }

    try {
      invalidateVehicleReleasedCache('new_vehicle_release');
    } catch (cacheErr) {
      console.warn('[NEW_VEHICLE] Cache invalidation failed:', cacheErr);
    }

    try {
      syncVehicleSheetFromCarTP();
    } catch (syncErr) {
      console.warn('[NEW_VEHICLE] Vehicle sheet sync failed:', syncErr);
    }

    return { ok: true, status: 'RELEASE', ref: row[iRef] || baseRef };
  } catch (error) {
    console.error('[NEW_VEHICLE] submitNewVehicleRelease error:', error);
    return { ok: false, error: String(error) };
  }
}

function getNewVehicleOptions(){
  try {
    const empty = {
      make: [],
      model: [],
      category: [],
      usageType: [],
      owner: [],
      vehicleNumber: [],
      catalog: []
    };
    const sh = _openVehicleSheet_();
    if (!sh) {
      return { ok: false, error: 'Vehicle sheet not found', options: empty };
    }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return {
        ok: true,
        options: empty,
        rowCount: 0,
        updatedAt: new Date().toISOString()
      };
    }

    const columnLetterToIndex = function(letter){
      if (!letter) return -1;
      let idx = 0;
      const cleaned = String(letter).toUpperCase().replace(/[^A-Z]/g, '');
      for (let i = 0; i < cleaned.length; i++) {
        idx *= 26;
        idx += (cleaned.charCodeAt(i) - 64); // 'A' => 1
      }
      return idx > 0 ? idx : -1; // 1-based
    };

    const letterMap = {
      make: 'Z',
      model: 'AA',
      category: 'AB',
      usageType: 'AC',
      owner: 'AD',
      vehicleNumber: 'AE'
    };

    const columnIndexes = {};
    Object.keys(letterMap).forEach(function(key){
      const col = columnLetterToIndex(letterMap[key]);
      columnIndexes[key] = col > 0 ? (col - 1) : -1; // convert to zero-based
    });

    const range = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    if (!range.length) {
      return {
        ok: true,
        options: empty,
        rowCount: 0,
        updatedAt: new Date().toISOString()
      };
    }

    const uniqueMaps = {
      make: new Map(),
      model: new Map(),
      category: new Map(),
      usageType: new Map(),
      owner: new Map()
    };
    const catalogMap = new Map();

    function recordUnique(map, rawValue){
      if (!rawValue && rawValue !== 0) return;
      const value = String(rawValue).trim();
      if (!value) return;
      const key = value.toLowerCase();
      const current = map.get(key);
      if (current) {
        current.count += 1;
      } else {
        map.set(key, { value: value, count: 1 });
      }
    }

    range.forEach(function(row){
      const make = columnIndexes.make >= 0 ? String(row[columnIndexes.make] || '').trim() : '';
      const model = columnIndexes.model >= 0 ? String(row[columnIndexes.model] || '').trim() : '';
      const category = columnIndexes.category >= 0 ? String(row[columnIndexes.category] || '').trim() : '';
      const usageType = columnIndexes.usageType >= 0 ? String(row[columnIndexes.usageType] || '').trim() : '';
      const owner = columnIndexes.owner >= 0 ? String(row[columnIndexes.owner] || '').trim() : '';
      const vehicleNumberRaw = columnIndexes.vehicleNumber >= 0 ? String(row[columnIndexes.vehicleNumber] || '').trim() : '';
      if (!vehicleNumberRaw) {
        // Still record meta lists so dropdowns remain populated even if some rows miss vehicle numbers
        recordUnique(uniqueMaps.make, make);
        recordUnique(uniqueMaps.model, model);
        recordUnique(uniqueMaps.category, category);
        recordUnique(uniqueMaps.usageType, usageType);
        recordUnique(uniqueMaps.owner, owner);
        return;
      }
      const vehicleNumber = vehicleNumberRaw.toUpperCase();
      recordUnique(uniqueMaps.make, make);
      recordUnique(uniqueMaps.model, model);
      recordUnique(uniqueMaps.category, category);
      recordUnique(uniqueMaps.usageType, usageType);
      recordUnique(uniqueMaps.owner, owner);

      if (!catalogMap.has(vehicleNumber)) {
        catalogMap.set(vehicleNumber, {
          vehicleNumber: vehicleNumber,
          make: make,
          model: model,
          category: category,
          usageType: usageType,
          owner: owner
        });
      }
    });

    try {
      const cartpSheet = _openCarTP_();
      if (cartpSheet) {
        const cartpLastRow = cartpSheet.getLastRow();
        if (cartpLastRow > 1) {
          const vehicleColIndex = _columnLetterToIndex_('G');
          if (vehicleColIndex > 0) {
            const cartpVehicles = cartpSheet.getRange(2, vehicleColIndex, cartpLastRow - 1, 1).getDisplayValues();
            cartpVehicles.forEach(function(rowValue){
              const raw = rowValue && rowValue.length ? rowValue[0] : rowValue;
              const vehicleNumber = String(raw || '').trim().toUpperCase();
              if (!vehicleNumber) return;
              if (!catalogMap.has(vehicleNumber)) {
                catalogMap.set(vehicleNumber, {
                  vehicleNumber: vehicleNumber,
                  make: '',
                  model: '',
                  category: '',
                  usageType: '',
                  owner: ''
                });
              }
            });
          }
        }
      }
    } catch (cartpErr) {
      console.warn('[NEW_VEHICLE] Failed to hydrate catalog with CarT_P column G', cartpErr);
    }

    function mapToOptionList(map){
      return Array.from(map.values())
        .sort(function(a, b){
          return a.value.localeCompare(b.value, undefined, { sensitivity: 'base' });
        })
        .map(function(entry){
          const label = entry.count > 1 ? entry.value + ' [' + entry.count + ']' : entry.value;
          return { value: entry.value, label: label };
        });
    }

    const catalog = Array.from(catalogMap.values())
      .sort(function(a, b){
        return a.vehicleNumber.localeCompare(b.vehicleNumber, undefined, { sensitivity: 'base' });
      });

    const vehicleNumberOptions = catalog.map(function(entry){
      const descriptor = [entry.make, entry.model].filter(Boolean).join(' ');
      return {
        value: entry.vehicleNumber,
        label: descriptor ? (entry.vehicleNumber + ' — ' + descriptor) : entry.vehicleNumber,
        make: entry.make,
        model: entry.model,
        category: entry.category,
        usageType: entry.usageType,
        owner: entry.owner
      };
    });

    const options = {
      make: mapToOptionList(uniqueMaps.make),
      model: mapToOptionList(uniqueMaps.model),
      category: mapToOptionList(uniqueMaps.category),
      usageType: mapToOptionList(uniqueMaps.usageType),
      owner: mapToOptionList(uniqueMaps.owner),
      vehicleNumber: vehicleNumberOptions,
      catalog: catalog
    };

    return {
      ok: true,
      updatedAt: new Date().toISOString(),
      rowCount: catalog.length,
      options: options
    };
  } catch (error) {
    console.error('getNewVehicleOptions error:', error);
    return { ok: false, error: String(error) };
  }
}


/* =====================  PROJECT / TEAM / BENEFICIARIES (DD)  ===================== */

function _openDataSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sh) throw new Error('Data sheet "DD" not found.');
  return sh;
}

function _norm(s) { return String(s == null ? '' : s).trim(); }
function _up(s)   { return _norm(s).toUpperCase(); }
function _normProjectKey(s) {
  try {
    return String(s == null ? '' : s)
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .replace(/\s*\/\s*/g, '/')
      .replace(/\s*[-]+\s*/g, '-')
      .replace(/\s*\|\s*/g, '|');
  } catch(_e) {
    return '';
  }
}
function _normTeamKey(s) {
  try {
    return String(s == null ? '' : s)
      .trim()
      .toLowerCase()
      .replace(/[\s_\-]+/g, ' ')
      .replace(/\s+/g, ' ');
  } catch(_e) {
    return '';
  }
}
function _toNum(v){
  return parseAmount(v);
}

/** tolerant header map: lowercase, trimmed, and alias-aware */
function _headerIndex_(headRow) {
  const idxNorm = {};   // normalized (trim + lower)
  const idxSan  = {};   // sanitized (trim + lower + remove non-alnum)

  const norm = (s) => _norm(s).toLowerCase();
  const san  = (s) => norm(s).replace(/[^a-z0-9]+/g,'');

  headRow.forEach((h, i) => {
    const n = norm(h);
    const z = san(h);
    if (!(n in idxNorm)) idxNorm[n] = i;
    if (z && !(z in idxSan)) idxSan[z] = i;
  });

  function find(keys) {
    for (let k of (Array.isArray(keys) ? keys : [keys])) {
      const nk = norm(k);
      const zk = san(k);
      if (nk in idxNorm) return idxNorm[nk];
      if (zk && (zk in idxSan)) return idxSan[zk];
      // loose contains fallback on both maps
      for (let seen in idxNorm) { if (seen.indexOf(nk) !== -1) return idxNorm[seen]; }
      for (let seen in idxSan) { if (zk && seen.indexOf(zk) !== -1) return idxSan[seen]; }
    }
    return -1;
  }

  return {
    get: (labels) => {
      const i = find(labels);
      if (i === -1) throw new Error('Missing header: ' + JSON.stringify(labels));
      return i;
    }
  };
}

/** Fallback: aggressively detect the Car Number column by heuristics */
function _findCarNumberColumn_(headRow, rows){
  try{
    // headRow is array of header display strings
    for (var i=0;i<headRow.length;i++){
      var h = String(headRow[i]||'').toLowerCase().replace(/\s+/g,' ').trim();
      // must mention car/vehicle and number/no/# or registration/plate
      var hasVehicle = /(car|vehicle)/.test(h);
      var hasNumberish = /(number|no\b|#|registration|reg|plate)/.test(h);
      if (hasVehicle && hasNumberish) return i;
    }
    if (Array.isArray(rows) && rows.length) {
      var sampleCount = Math.min(rows.length, 25);
      var colScores = new Map();
      for (var rowIdx = 0; rowIdx < sampleCount; rowIdx++) {
        var row = rows[rowIdx];
        if (!Array.isArray(row)) continue;
        for (var col = 0; col < row.length; col++) {
          var raw = row[col];
          var text = raw == null ? '' : String(raw).trim();
          if (!text) continue;
          // prefer values that look like vehicle identifiers (letters+numbers, allow dashes/slashes/spaces)
          if (!/[0-9]/.test(text)) continue;
          if (!/[A-Za-z]/.test(text)) continue;
          if (text.length < 3) continue;
          if (/(beneficiary|team|project)/i.test(String(headRow[col]||''))) continue;
          var score = colScores.get(col) || 0;
          score += /^[A-Z0-9\-\/\s]+$/.test(text) ? 2 : 1;
          if (/vehicle|car|reg|plate|number/i.test(String(headRow[col]||''))) score += 2;
          colScores.set(col, score);
        }
      }
      var bestCol = -1;
      var bestScore = 0;
      colScores.forEach(function(score, col){
        if (score > bestScore) {
          bestScore = score;
          bestCol = col;
        }
      });
      if (bestCol >= 0 && bestScore >= 3) return bestCol;
    }
  }catch(e){ /* ignore */ }
  return -1;
}

function _columnLetterToIndex_(letter) {
  if (!letter) return -1;
  let idx = 0;
  const cleaned = String(letter).toUpperCase().replace(/[^A-Z]/g, '');
  if (!cleaned) return -1;
  for (let i = 0; i < cleaned.length; i++) {
    idx = idx * 26 + (cleaned.charCodeAt(i) - 64);
  }
  return idx > 0 ? idx : -1;
}

/* -------------------------- versioned cache helpers -------------------------- */

function _cacheVersion_() {
  const sp = PropertiesService.getScriptProperties();
  let v = Number(sp.getProperty('DD_VER') || '1');
  if (!v || isNaN(v)) { v = 1; sp.setProperty('DD_VER', String(v)); }
  return v;
}

function _ptCacheVersion_() {
  const sp = PropertiesService.getScriptProperties();
  let v = Number(sp.getProperty('PT_VER') || '1');
  if (!v || isNaN(v)) { v = 1; sp.setProperty('PT_VER', String(v)); }
  return v;
}

function bustDDCache() {
  const sp = PropertiesService.getScriptProperties();
  let v = Number(sp.getProperty('DD_VER') || '1');
  v = (isNaN(v) ? 1 : v) + 1;
  sp.setProperty('DD_VER', String(v));
  // No direct delete in CacheService; version bump invalidates keys.
  return { ok:true, version: v, ttlSec: CACHE_TTL_SEC };
}

/**
 * Bust the project-team index cache to force regeneration with new logic
 */
function bustPTCache() {
  const sp = PropertiesService.getScriptProperties();
  let v = Number(sp.getProperty('PT_VER') || '1');
  v = (isNaN(v) ? 1 : v) + 1;
  sp.setProperty('PT_VER', String(v));
  // This will invalidate the PT_INDEX_V2_ cache key
  return { ok:true, version: v, message: 'Project-Team cache busted' };
}

/**
 * Debug function to identify projects without teams
 */
function debugProjectTeams() {
  const idx = getProjectTeamIndex();
  const projects = getProjects();
  
  const results = {
    totalProjects: projects.length,
    projectsWithTeams: 0,
    projectsWithoutTeams: [],
    teamsByProject: idx.teamsByProject || {},
    sampleTeamData: {}
  };
  
  projects.forEach(proj => {
    const teams = idx.teamsByProject[proj.id] || [];
    if (teams.length > 0) {
      results.projectsWithTeams++;
      if (Object.keys(results.sampleTeamData).length < 3) {
        results.sampleTeamData[proj.id] = { name: proj.name, teams: teams };
      }
    } else {
      results.projectsWithoutTeams.push({ id: proj.id, name: proj.name });
    }
  });
  
  return results;
}

/**
 * Read DD sheet -> compact rows used by the app, with a short cache.
 */
function _readDD_compact_() {
  const cache = CacheService.getScriptCache();
  const key = 'DD_COMPACT_V3_' + _cacheVersion_();
  try {
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* ignore cache errors */ }

  const sh = _openDataSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  let IX;
  try { IX = _headerIndex_(head); }
  catch (e) {
    console.error('Header parsing failed', e);
    Logger.log('Header parsing failed: %s', e);
    return [];
  }

  const iName   = IX.get(['Name of Beneficiary','Beneficiary']);
  const iDesig  = IX.get(['Resource Designation','Designation']);
  const iDA     = IX.get(['Default DA Amt','Default DA','DA Default']);
  const iProj   = IX.get(['Project']);
  const iTeam   = IX.get(['Team','Team Name']);
  const iAcct   = IX.get(['Account Holder Name','Account Holder']);
  const iUse    = IX.get(['In Use / release','In Use','In Use/Release']);
  let iDateTime = -1;
  try {
    iDateTime = IX.get(['Date and Time','Date & Time','Timestamp','Updated At','Date']);
  } catch (_missingDateCol) {
    iDateTime = -1;
  }

  const indices = [iName, iDesig, iDA, iProj, iTeam, iAcct, iUse];
  if (iDateTime >= 0) indices.push(iDateTime);
  const minIdx = Math.min.apply(null, indices);
  const maxIdx = Math.max.apply(null, indices);
  const width  = maxIdx - minIdx + 1;

  const shBody = sh.getRange(2, minIdx + 1, lastRow - 1, width).getValues();

  const out = new Array(shBody.length);
  for (let r = 0; r < shBody.length; r++) {
    const row = shBody[r];
    const get = (abs) => row[abs - minIdx];
    const beneficiaryValue = _norm(get(iName));
    const teamValue = _norm(get(iTeam));
    const dateRaw = iDateTime >= 0 ? get(iDateTime) : '';
    const ts = iDateTime >= 0 ? _parseTs_(dateRaw) : 0;
    out[r] = {
      beneficiary: beneficiaryValue,
      designation: _norm(get(iDesig)),
      defaultDa:   _toNum(get(iDA)),
      project:     _norm(get(iProj)),
      team:        teamValue,
      account:     _norm(get(iAcct)),
      inuse:       _norm(get(iUse)),
      dateTime:    dateRaw,
      timestamp:   ts,
      teamKey:     _normTeamKey(teamValue),
      beneficiaryKey: _beneficiaryKey_(beneficiaryValue)
    };
  }

  try {
    const s = JSON.stringify(out);
    if (s.length < 90000) cache.put(key, s, CACHE_TTL_SEC);
  } catch (e) { /* ignore cache errors */ }

  return out;
}

/**
 * Get historical submission date ranges for a beneficiary for a given field
 * field can be: 'da', 'fuel', 'car', 'air', 'transport', 'misc'
 * Returns an array of { from: 'YYYY-MM-DD', to: 'YYYY-MM-DD' }
 */
function getSubmissionRangesForBeneficiary(beneficiary, field) {
  try {
    const name = (beneficiary || '').toString().trim();
    if (!name) return [];
    const fieldKey = (field || '').toString().trim().toLowerCase();

    // Map field key to submissions header base
    const MAP = {
      fuel: 'Fuel',
  da: 'DA',
      car: 'Car',
      air: 'Airtime',
      transport: 'Transport',
      misc: 'Misc'
    };
  const base = MAP[fieldKey] || 'DA'; // default to DA to satisfy current need

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
    if (!sh) return [];

    const rng = sh.getDataRange();
    const values = rng.getValues();
    if (!values || values.length < 2) return [];

    const header = values[0].map(v => (v || '').toString());
    const idxBeneficiary = header.indexOf('Beneficiary');
    const idxFrom = header.indexOf(base + ' From');
    const idxTo   = header.indexOf(base + ' To');
    if (idxBeneficiary === -1 || idxFrom === -1 || idxTo === -1) return [];

    // Normalize incoming values to YYYY-MM-DD
    const toYMD = (val) => {
      if (!val) return '';
      if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) {
        // Format with East Africa Time to avoid off-by-one issues
        return Utilities.formatDate(val, TZ(), 'yyyy-MM-dd');
      }
      // Accept strings like DD/MM/YYYY or YYYY-MM-DD
      const s = String(val).trim();
      if (!s) return '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) {
        const d = ('0' + Number(m[1])).slice(-2);
        const mo = ('0' + Number(m[2])).slice(-2);
        const y = m[3];
        return `${y}-${mo}-${d}`;
      }
      return '';
    };

    const nameLc = name.toLowerCase();
    const out = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const ben = (row[idxBeneficiary] || '').toString().trim().toLowerCase();
      if (ben !== nameLc) continue;
      const from = toYMD(row[idxFrom]);
      const to = toYMD(row[idxTo]);
      if (from && to) out.push({ from, to });
    }
    return out;
  } catch (e) {
    console.error('getSubmissionRangesForBeneficiary error: ' + e);
    return [];
  }
}

/** Build and cache Project→Teams index for fast lookups */
function getProjectTeamIndex() {
  const cache = CacheService.getScriptCache();
  const key = 'PT_INDEX_V2_' + _ptCacheVersion_(); // Updated version to bust cache
  try {
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* ignore cache errors */ }

  const rows = _readDD_compact_().filter(r => _up(r.inuse) === 'IN USE' && r.project && r.beneficiary);
  const projSet = new Map(); // project -> true
  const projTeamBen = new Map(); // project -> Map(team -> Set(beneficiary-lc))

  for (const r of rows) {
    const proj = r.project;
    const team = r.team || 'Unassigned'; // Use 'Unassigned' for empty teams instead of empty string
    const ben  = r.beneficiary.toLowerCase();
    projSet.set(proj, true);
    if (!projTeamBen.has(proj)) projTeamBen.set(proj, new Map());
    const tmap = projTeamBen.get(proj);
    if (!tmap.has(team)) tmap.set(team, new Set());
    tmap.get(team).add(ben);
  }

  const projects = Array.from(projSet.keys())
    .map(p => ({ id: p, name: p }))
    .sort((a,b)=> a.name.localeCompare(b.name, undefined, {sensitivity:'base'}));

  const teamsByProject = {};
  projTeamBen.forEach((tmap, proj) => {
    const arr = [];
    tmap.forEach((set, team) => {
      // Always include teams if they have beneficiaries, even if team name is empty/unassigned
      if (set.size > 0) {
        const displayName = team === 'Unassigned' ? 'Unassigned Team' : team;
        arr.push({ id: team, name: displayName, count: set.size });
      }
    });
    arr.sort((a,b)=> a.name.localeCompare(b.name, undefined, {sensitivity:'base'}));
    teamsByProject[proj] = arr;
  });

  const index = { projects, teamsByProject };

  try {
    const s = JSON.stringify(index);
    if (s.length < 90000) cache.put(key, s, CACHE_TTL_SEC);
  } catch (e) { /* ignore cache errors */ }

  return index;
}

/** NEW: Compact client pack of "IN USE" DD rows for instant on-page filtering */
function getDDClientPack(){
  const cache = CacheService.getScriptCache();
  const key = 'DDC_PACK_V1_' + _cacheVersion_();
  try{
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  }catch(e){}

  const rows = _readDD_compact_().filter(r => _up(r.inuse) === 'IN USE' && r.beneficiary && r.project);
  // Pack as arrays to reduce size: [ben, acct, desig, da, proj, team]
  const packed = rows.map(r => [
    r.beneficiary, r.account, r.designation, Number(r.defaultDa||0), r.project, (r.team||'')
  ]);

  const obj = { v:_cacheVersion_(), packed:true, rows: packed };

  try {
    const s = JSON.stringify(obj);
    // No strict need to cache, but do it for parity
    if (s.length < 90000) cache.put(key, s, CACHE_TTL_SEC);
  } catch(e){}

  return obj;
}

/** ping */
function health(){ return {ok:true, ts: new Date().toISOString()}; }

/**
 * Installable trigger helper: create an onChange trigger for the `submissions` sheet.
 * Run from Apps Script editor to set up automatic cache busting when submissions change.
 */
function createSubmissionsOnChangeTrigger(){
  // Creates a project trigger that calls onSubmissionsChange when the spreadsheet changes
  try{
    // Remove any existing triggers first to avoid duplicates
    removeSubmissionsOnChangeTrigger();
    ScriptApp.newTrigger('onSubmissionsChange')
      .forSpreadsheet(SpreadsheetApp.openById(SHEET_ID))
      .onChange()
      .create();
    return { ok:true, message: 'Trigger created' };
  }catch(e){ return { ok:false, message: String(e) }; }
}

/**
 * Remove existing onSubmissionsChange triggers for this project.
 */
function removeSubmissionsOnChangeTrigger(){
  try{
    const ts = ScriptApp.getProjectTriggers();
    for (const t of ts){
      if (t.getHandlerFunction && t.getHandlerFunction() === 'onSubmissionsChange') ScriptApp.deleteTrigger(t);
    }
    return { ok:true, message: 'Removed triggers' };
  }catch(e){ return { ok:false, message: String(e) }; }
}

/**
 * Handler called by the installable trigger when the spreadsheet changes.
 * Busts script caches used by the app and records the last submissions update timestamp.
 */
function onSubmissionsChange(e){
  try{
    // Only act on EDIT/CHANGE events related to the submissions spreadsheet
    // We optimistically bust caches on any change event; keep logic minimal for reliability.
    bustDDCache();
    bustPTCache();
    bustTypeaheadCache();
    // Record last update time for quick inspection/use by clients
    const sp = PropertiesService.getScriptProperties();
    sp.setProperty('SUBMISSIONS_LAST_UPDATE', new Date().toISOString());
    return { ok:true, ts: new Date().toISOString() };
  }catch(e){ console.error('onSubmissionsChange error', e); return { ok:false, message: String(e) }; }
}

/**
 * Debug helper: list DA rows with parsed From/To/Amount and overlap status for an input range.
 * Usage (Apps Script editor): debugListDARows('01-Sep-2025','30-Sep-2025', 200)
 */
function debugListDARows(startStr, endStr, limit){
  try{
    limit = typeof limit === 'number' ? limit : 200;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
    if (!sh) return { ok:false, message:'no_submissions_sheet' };
    const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn();
    if (lastRow <= 1) return { ok:true, rows:[] };
    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => String(h||'').trim());
    // Candidate header names
    const fromCandidates = ['DA From','DA From Date','DA FromDate','DA Date From','DA Date','DA From (Date)','DA From'];
    const toCandidates = ['DA To','DA To Date','DA ToDate','DA Date To','DA To (Date)','DA To'];
    const amtCandidates = ['DA Amount','DA Amt','DA','ER DA Amt','ER DA Amount','ER DA','ERDA','er da amt','erda amt'];

    function findIdx(cands){
      for (const c of cands){ const i = head.indexOf(c); if (i >= 0) return i; }
      return -1;
    }

    // Try header-based first
    let idxFrom = findIdx(fromCandidates);
    let idxTo = findIdx(toCandidates);
    let idxAmt = findIdx(amtCandidates);
    // Fallback to fixed letters (L/M for DA) if header not found
    function colLetterToIndex(letter){ if (!letter) return -1; const s = String(letter||'').toUpperCase().trim(); let idx=0; for (let i=0;i<s.length;i++){ idx = idx*26 + (s.charCodeAt(i)-64); } return idx>0 ? idx-1 : -1; }
    if (idxFrom < 0) idxFrom = colLetterToIndex('L');
    if (idxTo < 0) idxTo = colLetterToIndex('M');
    if (idxAmt < 0) {
      // search common amount columns by heuristics
      const amtHints = ['DA Amount','DA Amt','DA','ER DA Amt','ER DA Amount','ER DA','ERDA'];
      for (let i=0;i<head.length;i++){ const h = String(head[i]||'').toLowerCase(); if (amtHints.some(a=>h.indexOf(a.toLowerCase())>=0)){ idxAmt = i; break; } }
    }

    const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    const disp = sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues();

    function parseDateFlexible(v){
      if (!v && v !== 0) return null;
      if (v instanceof Date) return v;
      const s = String(v||'').trim(); if (!s) return null;
      const d = parseDateTokenFlexible(s); if (d) return d;
      const d2 = new Date(s); if (!isNaN(d2.getTime())) return d2; return null;
    }

    const inStart = (startStr instanceof Date) ? startStr : parseDateFlexible(startStr);
    const inEnd = (endStr instanceof Date) ? endStr : parseDateFlexible(endStr);
    const out = [];
    for (let r=0; r<values.length && out.length < limit; r++){
      const row = values[r]; const drow = disp[r];
      const fromRaw = (idxFrom>=0) ? (row[idxFrom] || drow[idxFrom]) : null;
      const toRaw = (idxTo>=0) ? (row[idxTo] || drow[idxTo]) : null;
      const amtRaw = (idxAmt>=0) ? (row[idxAmt] || drow[idxAmt]) : null;
      const from = parseDateFlexible(fromRaw);
      const to = parseDateFlexible(toRaw);
      const amt = (function(a){ if (a==null || a==='') return 0; try{ return parseAmount(a); }catch(_){ return 0; } })(amtRaw);
      const rowNum = r + 2;
      // Determine overlap if both parsed
      let overlaps = false;
      if (from && to && inStart && inEnd){
        const fmt = (d)=> Utilities.formatDate(new Date(d), TZ(), 'yyyy-MM-dd');
        const a = fmt(from), b = fmt(to), x = fmt(inStart), y = fmt(inEnd);
        overlaps = (x <= b) && (a <= y);
      }
      out.push({ row: rowNum, fromRaw: String(fromRaw||''), toRaw: String(toRaw||''), from: from?Utilities.formatDate(new Date(from),TZ(),'dd-MMM-yyyy'):null, to: to?Utilities.formatDate(new Date(to),TZ(),'dd-MMM-yyyy'):null, amount: amt, overlaps });
    }
    return { ok:true, head, idxFrom, idxTo, idxAmt, rows: out, inputRange: { start: inStart?Utilities.formatDate(inStart,TZ(),'dd-MMM-yyyy'):null, end: inEnd?Utilities.formatDate(inEnd,TZ(),'dd-MMM-yyyy'):null } };
  }catch(e){ console.error('debugListDARows error', e); return { ok:false, message: String(e) }; }
}

/**
 * Wrapper: run debugListDARows and log the JSON result to Logger so it appears in Execution logs.
 * Usage: debugListDARowsLog('01-Sep-2025','30-Sep-2025',200)
 */
function debugListDARowsLog(startStr, endStr, limit){
  try{
    const res = debugListDARows(startStr, endStr, limit);
    try{ Logger.log(JSON.stringify(res, null, 2)); }catch(_e){ Logger.log(String(res)); }
    return res;
  }catch(e){ Logger.log('debugListDARowsLog error: ' + String(e)); return { ok:false, message: String(e) }; }
}

/**
 * Write debugListDARows output to a sheet called 'DEBUG_DA_ROWS' for easy inspection.
 * Usage: debugListDARowsToSheet('01-Sep-2025','30-Sep-2025',200)
 */
function debugListDARowsToSheet(startStr, endStr, limit){
  try{
    const res = debugListDARows(startStr, endStr, limit);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('DEBUG_DA_ROWS');
    if (!sh) sh = ss.insertSheet('DEBUG_DA_ROWS');
    // Clear sheet
    sh.clearContents();
    // Write metadata
    sh.getRange(1,1).setValue('inputStart');
    sh.getRange(1,2).setValue(res.inputRange && res.inputRange.start || '');
    sh.getRange(2,1).setValue('inputEnd');
    sh.getRange(2,2).setValue(res.inputRange && res.inputRange.end || '');
    sh.getRange(3,1).setValue('idxFrom'); sh.getRange(3,2).setValue(res.idxFrom);
    sh.getRange(4,1).setValue('idxTo'); sh.getRange(4,2).setValue(res.idxTo);
    sh.getRange(5,1).setValue('idxAmt'); sh.getRange(5,2).setValue(res.idxAmt);
    // Headers
    const headers = ['row','fromRaw','toRaw','fromParsed','toParsed','amount','overlaps'];
    sh.getRange(7,1,1,headers.length).setValues([headers]);
    // Rows
    const rows = (res.rows || []).map(r => [r.row, r.fromRaw, r.toRaw, r.from || '', r.to || '', r.amount, r.overlaps ? 'YES' : 'NO']);
    if (rows.length) sh.getRange(8,1,rows.length, headers.length).setValues(rows);
    return { ok:true, written: rows.length, sheet: 'DEBUG_DA_ROWS' };
  }catch(e){ console.error('debugListDARowsToSheet error', e); return { ok:false, message: String(e) }; }
}

/**
 * Test wrapper so you can run from the Apps Script function picker
 * Run this from the editor to execute the debug for 01-Sep-2025 -> 30-Sep-2025
 */
function debugListDARows_TestSep2025(){
  return debugListDARowsLog('01-Sep-2025','30-Sep-2025',200);
}

/* -------------------------- Typeahead (Predictive Text) -------------------------- */

function _taCacheVersion_(){
  var sp = PropertiesService.getScriptProperties();
  var v = Number(sp.getProperty('TA_VER') || '1');
  if (!v || isNaN(v)) { v = 1; sp.setProperty('TA_VER', String(v)); }
  return v;
}

function bustTypeaheadCache(){
  var sp = PropertiesService.getScriptProperties();
  var v = Number(sp.getProperty('TA_VER') || '1');
  v = (isNaN(v) ? 1 : v) + 1;
  sp.setProperty('TA_VER', String(v));
  return { ok:true, version: v };
}

/** Build and cache a compact typeahead index across relevant sheets */
function _buildTypeaheadIndex_(){
  var cache = CacheService.getScriptCache();
  var key = 'TA_IDX_V1_' + _taCacheVersion_();
  try{ var hit = cache.get(key); if (hit) return JSON.parse(hit); }catch(_e){}

  var seen = new Map(); // key: type+'\u0000'+valueLower -> { t, v, w }
  function add(t, v, w){
    v = String(v||'').trim(); if (!v) return;
    var k = t + '\u0000' + v.toLowerCase();
    var cur = seen.get(k);
    if (!cur) seen.set(k, { t:t, v:v, w: Number(w)||0 });
    else cur.w += Number(w)||0;
  }

  // 1) Expense keywords (and common phrases)
  var expenseTerms = [
  'Fuel','DA','Vehicle Rent','Airtime','Transport','Misc','Total','Total Expense','Expenses','Cost','Spending'
  ];
  expenseTerms.forEach(function(x){ add('expense', x, 50); add('expense', x.toLowerCase(), 15); });

  // 2) Dates (months + relative)
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  months.forEach(function(m){ add('date', m, 35); add('date', m.slice(0,3), 20); });
  ['This month','Last month','Next month','Today','Yesterday'].forEach(function(d){ add('date', d, 30); });

  // 3) DD data (beneficiaries, projects, teams, accounts)
  try {
    var dd = _readDD_compact_();
    dd.forEach(function(r){
      if (r.beneficiary) add('beneficiary', r.beneficiary, 45);
      if (r.project) add('project', r.project, 40);
      if (r.team) add('team', r.team, 38);
      if (r.account) add('account', r.account, 20);
      if (r.designation) add('designation', r.designation, 18);
    });
  } catch(_eDD){}

  // 4) CarT_P data (vehicle numbers, make/model, status)
  try {
    // Also index CarT_P headers for predictive text like "Vehicle Number"
    var shCar = _openCarTP_();
    if (shCar){
      var hc = shCar.getLastColumn();
      if (hc>0){
        var hcar = shCar.getRange(1,1,1,hc).getDisplayValues()[0];
        hcar.forEach(function(hv){ if (hv) add('header', hv, 28); });
      }
    }
    var carRows = _readCarTP_objects_();
    carRows.forEach(function(r){
      if (r['Vehicle Number']) add('vehicle', r['Vehicle Number'], 45);
      if (r.Make) add('make', r.Make, 20);
      if (r.Model) add('model', r.Model, 18);
      if (r.Owner) add('owner', r.Owner, 12);
      if (r.Status) add('status', _normStatus_(r.Status), 25);
      if (r.Team) add('team', r.Team, 14);
      if (r.Category) add('category', r.Category, 10);
      if (r['Usage Type']) add('usage', r['Usage Type'], 10);
    });
  } catch(_eCar){}

  // 5) Vehicle master (enrich vehicle/make/model/owner)
  try {
    var vm = _readVehicleMasterIndex_();
    Object.keys(vm).forEach(function(k){
      add('vehicle', k, 40);
      var m = vm[k];
      if (m.make) add('make', m.make, 15);
      if (m.model) add('model', m.model, 15);
      if (m.owner) add('owner', m.owner, 10);
      if (m.category) add('category', m.category, 10);
      if (m.usageType) add('usage', m.usageType, 10);
    });
  } catch(_eVM){}

  // 6) Submissions: headers + unique values (beneficiary, project, team)
  try {
    var sh = getSheet('submissions');
    if (sh) {
      var lastRow = sh.getLastRow(); var lastCol = sh.getLastColumn();
      if (lastCol > 0) {
        var head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
        head.forEach(function(h){ if (h) add('header', h, 30); });
      }
      if (lastRow > 1 && lastCol > 0){
        var vals = sh.getRange(2,1,Math.min(1000,lastRow-1), lastCol).getDisplayValues(); // cap to 1000 rows
        var h = (function(){ try{ return sh.getRange(1,1,1,lastCol).getDisplayValues()[0]; }catch(_){ return []; }})();
        var iBen = h.indexOf('Beneficiary');
        var iProj= h.indexOf('Project'); if (iProj<0) iProj = h.indexOf('Project Name');
        var iTeam= h.indexOf('Team'); if (iTeam<0) iTeam = h.indexOf('Team Name');
        var seenBen={},seenProj={},seenTeam={};
        vals.forEach(function(r){
          if (iBen>=0){ var v=r[iBen]; if (v && !seenBen[v]){ add('beneficiary', v, 22); seenBen[v]=1; } }
          if (iProj>=0){ var p=r[iProj]; if (p && !seenProj[p]){ add('project', p, 18); seenProj[p]=1; } }
          if (iTeam>=0){ var t=r[iTeam]; if (t && !seenTeam[t]){ add('team', t, 16); seenTeam[t]=1; } }
        });
      }
    }
  } catch(_eSub){}

  // Assemble and cap size
  var arr = Array.from(seen.values());
  // Prefer higher weight, then shorter label, then alpha
  arr.sort(function(a,b){ if (b.w!==a.w) return b.w-a.w; if (a.v.length!==b.v.length) return a.v.length-b.v.length; return String(a.v||'').localeCompare(String(b.v||'')); });
  // Cap to keep cache payload light (~100KB limit)
  var MAX = 6000;
  if (arr.length > MAX) arr = arr.slice(0, MAX);
  var idx = { v:_taCacheVersion_(), tokens: arr };
  try { var s = JSON.stringify(idx); if (s.length < 90000) CacheService.getScriptCache().put(key, s, CACHE_TTL_SEC); } catch(_ePut){}
  return idx;
}

/** Public: return suggestions for a prefix */
function getTypeaheadSuggestions(prefix, limit, types, offset){
  try{
    prefix = String(prefix||'').trim();
    offset = Number(offset || 0) || 0;
    var wantTypes = Array.isArray(types) ? types.map(function(t){return String(t||'').toLowerCase();}) : [];
    var idx = _buildTypeaheadIndex_();
    var tokens = idx && idx.tokens ? idx.tokens : [];
    // Filter by types if provided
    if (wantTypes && wantTypes.length){
      tokens = tokens.filter(function(t){ return wantTypes.indexOf(String(t.t||'').toLowerCase()) !== -1; });
    }
    // No prefix: return top-N curated defaults
    if (!prefix){
      var preferred = tokens.filter(function(t){ return t.t==='expense' || t.t==='date' || t.t==='status'; });
      if (preferred.length < (limit||10)) preferred = tokens.slice(0, limit||10);
      return preferred.slice(offset, offset + (limit||10)).map(function(t, i){ return { id: (t.i||('ta_'+(offset+i))), label:t.v, type:t.t, meta:{ weight:t.w } }; });
    }
    var pl = prefix.toLowerCase();
    // Match by startsWith on value or any token part
    var out = [];
    for (var i=0;i<tokens.length;i++){
      var t = tokens[i];
      var vl = String(t.v||'').toLowerCase();
      if (vl.indexOf(pl) === 0) { out.push(t); continue; }
      var parts = vl.split(/\s+/);
      for (var j=0;j<parts.length;j++){ if (parts[j].indexOf(pl) === 0){ out.push(t); break; } }
      if (out.length >= Math.max(200, (limit||10)*10)) break; // stop excessive scanning
    }
    // Rank results: stronger prefix match first (startsWith > part-start), then weight
    out.sort(function(a,b){
      function rank(t){ var vl=String(t.v||'').toLowerCase(); if (vl.indexOf(pl)===0) return 2; var parts=vl.split(/\s+/); for (var k=0;k<parts.length;k++){ if (parts[k].indexOf(pl)===0) return 1; } return 0; }
      var ra = rank(a), rb = rank(b); if (rb!==ra) return rb-ra; if (b.w!==a.w) return b.w-a.w; return String(a.v||'').localeCompare(String(b.v||''));
    });
    // Slice using offset/limit for paging
    var slice = out.slice(offset, offset + Math.max(1, Math.min(50, Number(limit||10))));
    return slice.map(function(t, i){ return { id: (t.i||('ta_'+(offset+i))), label:t.v, type:t.t, meta:{ weight:t.w } }; });
  }catch(e){
    console.error('getTypeaheadSuggestions error:', e);
    return [];
  }
}

/** Build CSV and HTML table from rows (rows = array of objects or array of arrays) */
function _buildTableFromRows_(rows){
  try{
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      // Return a minimal, valid HTML block so the client never sees a null/empty table
      var emptyHtml = '<div class="ai-table-wrap"><div class="ai-table" style="padding:8px;font:14px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#334155;background:#fff;border:1px solid #e5e7eb;border-radius:6px">No matching rows for the requested query.</div></div>';
      return { headers: [], rows: [], csv: ' ', html: emptyHtml };
    }
    var headers = [];
    var isObj = typeof rows[0] === 'object' && !Array.isArray(rows[0]);
    if (isObj){
      var hmap = {};
      rows.forEach(function(r){ Object.keys(r||{}).forEach(function(k){ if (!(k in hmap)) hmap[k]=true; }); });
      headers = Object.keys(hmap);
      var tableRows = rows.map(function(r){ return headers.map(function(h){ return (r && r[h] != null) ? String(r[h]) : ''; }); });
    } else if (Array.isArray(rows[0])){
      // assume first row is header if strings
      headers = rows[0].slice(0);
      var tableRows = rows.slice(1).map(function(r){ return r.map(function(c){ return (c==null)?'':String(c); }); });
    } else {
      // primitive rows
      headers = ['value'];
      var tableRows = rows.map(function(r){ return [String(r)]; });
    }

    // Build CSV
    function escCsv(v){ if (v==null) return ''; v = String(v); if (/[",\n]/.test(v)) return '"' + v.replace(/"/g,'""') + '"'; return v; }
    var csv = headers.map(escCsv).join(',') + '\n' + tableRows.map(function(r){ return r.map(escCsv).join(','); }).join('\n');

    // Build HTML table (simple)
  var ths = headers.map(function(h){ return '<th>' + _escHtml(String(h)) + '</th>'; }).join('');
  var trs = tableRows.map(function(r){ return '<tr>' + r.map(function(c){ return '<td>' + _escHtml(String(c)) + '</td>'; }).join('') + '</tr>'; }).join('');
    var html = '<div class="ai-table-wrap"><table class="ai-table" style="border-collapse:collapse;width:100%"><thead><tr>' + ths + '</tr></thead><tbody>' + trs + '</tbody></table></div>';
    return { headers: headers, rows: tableRows, csv: csv, html: html };
  }catch(e){ console.error('Table build error:', e); return { headers:[], rows:[], csv:'', html:'' }; }
}

/** Diagnostic: create a sample Output: table for current month and return table.csv length and preview */
function test_outputTable(){
  var now = new Date();
  var plan = { type: 'aggregate', metric: 'total', groupBy: 'project', filters: { timeframe: { month: now.getMonth(), year: now.getFullYear() } } };
  var data = executePlan(plan);
  var rows = data && data.rows ? data.rows : [];
  var tbl = _buildTableFromRows_(rows);
  return { ok:true, rows: rows.length || 0, csvSize: (tbl.csv||'').length, htmlPreview: (tbl.html||'').slice(0, 200) };
}

/** Unique projects (served from index when possible) */
function getProjects() {
  try {
    const idx = getProjectTeamIndex(); // warm path
    return idx.projects;
  } catch (e) {
    const rows = _readDD_compact_().filter(r => _up(r.inuse) === 'IN USE' && r.project);
    const map = new Map();
    rows.forEach(r => { if (!map.has(r.project)) map.set(r.project, { id: r.project, name: r.project }); });
    const out = Array.from(map.values()).sort((a,b)=>a.name.localeCompare(b.name, undefined, {sensitivity:'base'}));
    Logger.log('getProjects(fallback) -> %s', out.length);
    return out;
  }
}

/** Teams for a project (served from index when possible) */
function getTeams(projectId) {
  try {
    const proj = _norm(projectId);
    if (!proj) return [];
    const idx = getProjectTeamIndex();
    return idx.teamsByProject[proj] || [];
  } catch (e) {
    const proj = _norm(projectId);
    if (!proj) return [];
    const rows = _readDD_compact_().filter(r => _up(r.inuse) === 'IN USE' && r.project === proj);
    const teamToSet = new Map();
    rows.forEach(r => {
      if (!r.beneficiary) return; // Only skip if no beneficiary, allow empty teams
      const team = r.team || 'Unassigned'; // Use 'Unassigned' for empty teams
      if (!teamToSet.has(team)) teamToSet.set(team, new Set());
      teamToSet.get(team).add(r.beneficiary.toLowerCase());
    });
    const out = [];
    teamToSet.forEach((set, team) => {
      if (set.size > 0) {
        const displayName = team === 'Unassigned' ? 'Unassigned Team' : team;
        out.push({ id: team, name: displayName, count: set.size });
      }
    });
    out.sort((a,b)=>a.name.localeCompare(b.name, undefined, {sensitivity:'base'}));
    Logger.log('getTeams(fallback)(%s) -> %s', proj, out.length);
    return out;
  }
}

/** Beneficiaries */
function getBeneficiaries(projectId, teamIds, includeAllStatuses) {
  try {
    let proj = _norm(projectId);
    const projKey = _normProjectKey(projectId);
    const haveTeams = Array.isArray(teamIds) && teamIds.length > 0;
    if (!proj && !haveTeams) return [];
    const teamSet = new Set((teamIds || []).map(_norm));
    const teamAliasSet = new Set((teamIds || [])
      .map(function(t){ return _normTeamKey(t); })
      .filter(function(val){ return val !== ''; }));
    const includeAll = (function(flag){
      if (flag === true) return true;
      if (flag === false || flag === null || typeof flag === 'undefined') return false;
      if (typeof flag === 'string') {
        const norm = flag.trim().toLowerCase();
        return norm === 'true' || norm === 'all' || norm === 'any';
      }
      return false;
    })(includeAllStatuses);

    const rows = _readDD_compact_().filter(r => {
      if (!includeAll && _up(r.inuse) !== 'IN USE') return false;
      if (projKey) {
        if (_normProjectKey(r.project) !== projKey) return false;
      } else if (!haveTeams) {
        return false;
      }
      if (haveTeams) {
        const rowTeamNorm = _norm(r.team);
        const rowTeamAlias = _normTeamKey(r.team);
        const directMatch = teamSet.has(rowTeamNorm);
        const aliasMatch = rowTeamAlias ? teamAliasSet.has(rowTeamAlias) : false;
        if (!directMatch && !aliasMatch) return false;
      }
      return !!r.beneficiary;
    });

    if (!proj && rows.length) {
      const firstWithProject = rows.find(function(r){
        return r && r.project;
      });
      if (firstWithProject && firstWithProject.project) {
        proj = _norm(firstWithProject.project);
      }
    }

    const seen = new Set();
    const out  = [];
    for (const r of rows) {
      const key = r.beneficiary.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);

      const ben   = r.beneficiary;
      const acct  = r.account;
      const team  = r.team || '';
      const projN = r.project || '';

      out.push({
        beneficiary:   ben,
        accountHolder: acct,
        designation:   r.designation || '',
        erdaAmt:       0, // Always start with 0, user must fill manually
        project:       projN,
        projectName:   projN,
        team:          team,
        teamId:        team,
        teamName:      team,
        diffNames:     (ben && acct) ? (ben.toLowerCase() !== acct.toLowerCase()) : false
      });
    }

    out.sort((a,b)=> (a.beneficiary||'').localeCompare(b.beneficiary||'', undefined, {sensitivity:'base'}));
    Logger.log('getBeneficiaries(%s, %s teams) -> %s', proj || '(team only)', (teamIds||[]).length, out.length);
    return out;
  } catch (e) {
    console.error('getBeneficiaries error', e);
    throw new Error('getBeneficiaries: ' + e);
  }
}

/** Get project and team information for a specific beneficiary */
function getProjectTeamFromDD(beneficiaryName) {
  try {
    const name = _norm(beneficiaryName);
    if (!name) return { error: 'No beneficiary name provided' };

    const rows = _readDD_compact_().filter(r => {
      if (_up(r.inuse) !== 'IN USE') return false;
      return _norm(r.beneficiary) === name;
    });

    if (rows.length === 0) {
      return { error: 'Beneficiary not found in DD sheet' };
    }

    const row = rows[0]; // Take first match
    return {
      beneficiary: row.beneficiary,
      project: row.project,
      team: row.team,
      designation: row.designation,
      accountHolder: row.account,
      defaultDa: row.defaultDa
    };
  } catch (e) {
    console.error('getProjectTeamFromDD error', e);
    return { error: 'Error fetching beneficiary data: ' + e.message };
  }
}



/** diagnostic */
function getBeneficiariesDebug(projectId, teamIds){
  const list = getBeneficiaries(projectId, teamIds);
  const byTeam = {};
  list.forEach(x => { byTeam[x.teamName] = (byTeam[x.teamName]||0)+1; });
  return { ok:true, total:list.length, teams:Object.keys(byTeam).length, byTeam, sample:list.slice(0,5) };
}

/**
 * Get unique beneficiary counts from DD.
 * Returns a concise text summary covering IN USE and All rows, plus top projects.
 */
function getUniqueBeneficiaryStatsFromDD() {
  try {
    const rows = _readDD_compact_();
    if (!rows || !rows.length) return 'No DD data available to compute beneficiary counts.';

    const allSet = new Set();
    const inUseSet = new Set();
    const projCountsInUse = new Map();
    const teamSetInUse = new Set();
    const projSetInUse = new Set();

    rows.forEach(r => {
      const ben = _norm(r.beneficiary);
      if (!ben) return;
      allSet.add(ben);
      if (_up(r.inuse) === 'IN USE') {
        inUseSet.add(ben);
        const proj = _norm(r.project);
        const team = _norm(r.team);
        if (proj) projSetInUse.add(proj);
        if (team) teamSetInUse.add(team);
        if (proj) projCountsInUse.set(proj, (projCountsInUse.get(proj) || 0) + 1);
      }
    });

    const topProjects = Array.from(projCountsInUse.entries())
      .sort((a,b)=> b[1]-a[1])
      .slice(0,5)
      .map(([p,c],i)=> `${i+1}. ${p}: ${c}`);

    const stats = {
      inUse: inUseSet.size,
      all: allSet.size,
      projInUse: projSetInUse.size,
      teamInUse: teamSetInUse.size,
      topProjects: topProjects
    };
    return formatUniqueBeneficiaryStats(stats);
  } catch (e) {
    console.error('getUniqueBeneficiaryStatsFromDD error:', e);
    return 'Error computing unique beneficiary counts.';
  }
}

/* -------------------------- Chat System -------------------------- */

/**
 * Handle POST requests for chat functionality
 */
function doPost(e) {
  try {
    try { ensureSpreadsheetTZ(); } catch(_tz) {}
    const postData = JSON.parse(e.postData.contents);
    
    if (postData.action === 'chat') {
      const response = (typeof Chat !== 'undefined' && Chat && typeof Chat.processChatQuery === 'function')
        ? Chat.processChatQuery(postData.message)
        : processChatQuery(postData.message);
      return ContentService
        .createTextOutput(JSON.stringify({ response: response }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('doPost error:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Server error: ' + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Global variable to store conversation context
var chatContext = {
  lastQuery: '',
  lastIntent: '',
  lastResponse: '',
  awaitingFollowup: false,
  followupType: ''
};

// Global variables for self-learning system
var conversationLog = [];
var userFeedback = {};
var intentAccuracy = {};

/**
 * Enhanced AI-like chat query processor with intelligent intent analysis and self-learning
 * Handles natural language, contextual responses, follow-up questions, and learns from interactions
 */
function processChatQuery(message, convId) {
  const startTime = new Date();
  // Ensure default RAG folder is configured once
  try { RAG_ensureDefaultFolderConfigured(); } catch(_e) {}
  
  if (!message || typeof message !== 'string') {
    return 'Hello! I\'m your Fund Request Assistant. Ask me about expenses, beneficiaries, teams, vehicles, or any fund-related questions.';
  }
  
  const query = message.toLowerCase().trim();
  console.log('Processing enhanced chat query:', query);
  console.log('Current context:', chatContext);
  
  try {
    // Detect explicit output preference anywhere (e.g., 'output: table') and strip it
    let localMessage = String(message||'');
    let localExplicitOutput = null;
    try{
      const om = localMessage.match(/\boutput\s*[:\-]\s*(table|summary|total)\b/ig);
      if (om && om.length){ const last = om[om.length-1]; const mm = last.match(/\boutput\s*[:\-]\s*(table|summary|total)\b/i); if (mm && mm[1]) localExplicitOutput = mm[1].toLowerCase(); localMessage = localMessage.replace(/\boutput\s*[:\-]\s*(table|summary|total)\b/ig, '').trim(); }
    }catch(_eOut){}
    const wantsTable = (localExplicitOutput === 'table');
    
    // Check for rating feedback first
    if (query.startsWith('rate ')) {
      return handleRatingFeedback(message);
    }

    // Early routing: handle docs: and ai: prefixes BEFORE any quick intents
    var isDocs = /^\s*docs\s*[:\-]/i.test(message);
    var isForcedOpen = /^\s*ai\s*[:\-]/i.test(message);
    if (isDocs) {
      var cleanDocsMsg = message.replace(/^\s*docs\s*[:\-]/i, '').trim();
      try {
        console.log('Docs-mode forced');
        var ctxDocs = RAG_retrieveContextForQuery(cleanDocsMsg, 5);
        if (!ctxDocs) ctxDocs = 'No indexed snippets matched the question.';
        var responseDocs = (typeof Chat !== 'undefined' && Chat && typeof Chat.llmAnswer === 'function')
          ? Chat.llmAnswer(cleanDocsMsg, ctxDocs, null)
          : 'LLM (grounded docs mode) is not available. Please configure LLM settings.';
        const convDocs = logConversation(cleanDocsMsg, responseDocs, 'docs_mode', {}, startTime);
        return addFeedbackPrompt(responseDocs, convDocs);
      } catch (errDocs) {
        console.error('Docs-mode error:', errDocs);
        // fall through to regular flow if docs fails
      }
    }
    if (isForcedOpen) {
      var cleanMsg = message.replace(/^\s*ai\s*[:\-]/i, '').trim();
      try {
        console.log('Open-mode forced (Play)');
        var responseOpen = (typeof Chat !== 'undefined' && Chat && typeof Chat.llmAnswerOpen === 'function')
          ? Chat.llmAnswerOpen(cleanMsg, '', null)
          : 'LLM (open mode) is not available. Please configure LLM settings.';
        const convIdOpen = logConversation(cleanMsg, responseOpen, 'open_mode', {}, startTime);
        return addFeedbackPrompt(responseOpen, convIdOpen);
      } catch (errOpen) {
        console.error('Open-mode error:', errOpen);
        // fall through to regular flow if open fails
      }
    }

    // Quick intent: beneficiary count from DD (only if not in docs/open forced modes)
    if (/\b(how\s*many|count|number\s*of)\b.*\b(beneficiar)/i.test(message)) {
      const summary = getUniqueBeneficiaryStatsFromDD();
      const convIdCnt = logConversation(message, summary, 'beneficiary_count', {}, startTime);
      return addFeedbackPrompt(summary, convIdCnt);
    }
    
    // Initialize CarT_P sheet if it doesn't exist
    addSampleCarData();

    // Targeted: "Which vehicle team X is using now" → latest IN USE vehicles for that team
    try {
      var teamVehicleNow = null;
      var qRaw = String(message||'');
      var qlc = qRaw.toLowerCase();
      if (/(car|vehicle)/i.test(qRaw) && /(using|in\s*use)/i.test(qRaw) && /\bteam\b/i.test(qRaw)) {
        // Extract team name appearing after the word "team"
        var mTeam = /\bteam\s+([^\n\r,;:!?]+?)(?:\s+(?:is|are)\b|\s+(?:currently\s*)?(?:using|in\s*use)\b|$)/i.exec(qRaw);
        if (mTeam && mTeam[1]) {
          var teamName = mTeam[1].trim().replace(/[\s]+$/,'');
          if (teamName) {
            var respTeam = getTeamVehiclesInUse(teamName);
            var convTeam = logConversation(message, respTeam, 'vehicle_team_in_use', { team: teamName }, startTime);
            return addFeedbackPrompt(respTeam, convTeam);
          }
        }
      }
    } catch(_teamParseErr) { /* fall through */ }

    // Targeted: "Beneficiary <name> fuel amount in <Month Year>"
    try {
      var hasBenWord = /\bbeneficiar/i.test(message);
      var hasFuelWord = /\bfuel\b/i.test(message);
      if (hasBenWord && hasFuelWord) {
        // Use entity extractor to find known beneficiary and month/year
        var entTmp = extractEntities(query);
        var myTmp = entTmp.monthYear || parseMonthYearLoose(message);
        if (entTmp && Array.isArray(entTmp.names) && entTmp.names.length === 1 && myTmp) {
          var bname = entTmp.names[0];
          var respFuel = getBeneficiaryMetricForPeriod(bname, 'fuel', myTmp.month, myTmp.year);
          var convFuel = logConversation(message, respFuel, 'beneficiary_fuel_period', { beneficiary: bname, month: myTmp.month, year: myTmp.year }, startTime);
          return addFeedbackPrompt(respFuel, convFuel);
        }
      }
    } catch(_bfErr) { /* fall through */ }
    
    // Check for follow-up responses first
    if (chatContext.awaitingFollowup && isFollowupResponse(query)) {
      const followupResponse = handleFollowupResponse(query, chatContext.followupType);
      // Log followup conversation
      const conversationId = logConversation(message, followupResponse, 'followup', {}, startTime);
      // Reset context after handling followup
      chatContext = {
        lastQuery: message,
        lastIntent: 'followup',
        lastResponse: followupResponse,
        awaitingFollowup: false,
        followupType: ''
      };
      return addFeedbackPrompt(followupResponse, conversationId);
    }
    
    // Intent analysis with learning enhancement
    const intent = analyzeIntentWithLearning(query);
    const entities = extractEntities(query);
    
    console.log('Detected intent:', intent);
    console.log('Extracted entities:', entities);
    
    let response = '';
    let handledSpecific = false;

    // Early: beneficiary + explicit date range or month/year targeted question
    try {
      const expenseHint = /(expense|expenses|spend|cost|fuel|transport|airtime|er\s?da|vehicle|car|misc)/i.test(message);
      // If a date range is present + one known name, respect it first
      if (!handledSpecific && entities && entities.names && entities.names.length === 1 && entities.dateRange){
        const target = entities.names[0];
        const hasFuel  = /\bfuel\b/i.test(message);
        // Accept both "ER DA" and plain "DA"
        const hasERDA  = /(\ber\s?da\b|\bda\b)/i.test(message);
        const hasCar   = /\b(vehicle\s*rent|car\b)/i.test(message);
        const hasAir   = /\bair(time)?\b/i.test(message);
        const hasTrans = /\btransport|travel\b/i.test(message);
        const hasMisc  = /\bmisc|miscellaneous\b/i.test(message);
        if (hasFuel || hasERDA || hasCar || hasAir || hasTrans || hasMisc){
          const metric = hasFuel?'fuel': (hasERDA?'erda': (hasCar?'vehicleRent': (hasAir?'airtime': (hasTrans?'transport': 'misc'))));
          // Use updated helper; request table output if wanted
          console.log('beneficiary+dateRange branch - target:', target, 'metric:', metric, 'wantsTable:', wantsTable);
          try{ console.log('entities.dateRange:', entities.dateRange); }catch(_e){}
          const resp = getBeneficiaryMetricForDateRange(target, metric, entities.dateRange.start, entities.dateRange.end, { output: wantsTable ? 'table' : 'summary' });
          // If user asked for a table, prefer returning the structured object so the chat UI renders HTML
          if (wantsTable) {
            if (resp && typeof resp === 'object') {
              // Ensure we always have a minimal non-empty HTML so the client never sees null
              if (!resp.table || !resp.table.html) {
                const safeHtml = '<div class="ai-table-wrap"><div class="ai-table" style="padding:8px;font:14px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#334155;background:#fff;border:1px solid #e5e7eb;border-radius:6px">No matching rows for the requested query.</div></div>';
                resp.table = resp.table || { headers:[], rows:[], csv:'', html: safeHtml };
              }
              logConversation(message, '[table]', 'beneficiary_metric_range_table', { beneficiary: target, metric: metric }, startTime);
              return resp; // object with { message, table: { html, csv } }
            }
            // Fallback: wrap string response as a table-like object for consistent rendering
            try {
              const msg = (typeof resp === 'string') ? resp : '';
              const safeHtml = '<div class="ai-table-wrap"><div class="ai-table" style="padding:8px;font:14px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#334155;background:#fff;border:1px solid #e5e7eb;border-radius:6px">' + _escHtml(msg) + '</div></div>';
              const wrapped = { ok:true, message: msg, table: { headers:[], rows:[], csv:'', html: safeHtml } };
              logConversation(message, '[table]', 'beneficiary_metric_range_table_wrapped', { beneficiary: target, metric: metric }, startTime);
              return wrapped;
            } catch(eTblObj){ console.error('Table object fallback error:', eTblObj); }
          }
          // Non-table path or no table available: use message/string fallback
          response = (typeof resp === 'string' ? resp : (resp && resp.message) || '');
          handledSpecific = true;
        } else {
          const aggR = getBeneficiaryExpensesForDateRange(target, entities.dateRange.start, entities.dateRange.end, { output: wantsTable ? 'table' : 'summary' });
          if (wantsTable && aggR && aggR.ok && aggR.table && aggR.table.html){
            logConversation(message, '[table]', 'beneficiary_expense_range_table', { beneficiary: target }, startTime);
            return aggR; // return object for chat UI to render table
          }
          if (aggR && aggR.ok){ response = aggR.message; handledSpecific = true; }
        }
      }
      const wantsBenWise = /(beneficiary\s*[- ]?wise|by\s+beneficiary)/i.test(message);

      // Beneficiary-wise summary for a month
      if (!handledSpecific && wantsBenWise) {
        const now = new Date();
        const myParsed = parseMonthYearLoose(message);
        const mm = entities.monthYear ? entities.monthYear.month : (myParsed ? myParsed.month : now.getMonth());
        const yy = entities.monthYear ? entities.monthYear.year  : (myParsed ? myParsed.year  : now.getFullYear());
        response = getBeneficiaryMonthlySummary(mm, yy);
        handledSpecific = true;
      }
      if (!handledSpecific && entities && entities.names && entities.names.length === 1 && (entities.monthYear || expenseHint)) {
        const target = entities.names[0];
        // monthYear may be null; prefer explicit parse from message before defaulting to current
        const now = new Date();
        const myParsed2 = parseMonthYearLoose(message);
        const mm = entities.monthYear ? entities.monthYear.month : (myParsed2 ? myParsed2.month : now.getMonth());
        const yy = entities.monthYear ? entities.monthYear.year  : (myParsed2 ? myParsed2.year  : now.getFullYear());
        const agg = getBeneficiaryPeriodExpenses(target, mm, yy);
        if (agg && agg.ok) {
          response = agg.message;
          handledSpecific = true;
        }
      }
    } catch(_es){ /* ignore and let normal flow continue */ }
    
    // Handle specific intents with contextual responses
    if (!handledSpecific) switch (intent.type || intent) {
      case 'monthly_expenses':
        response = handleMonthlyExpensesIntent(entities, query);
        break;
      
      case 'fuel_expenses':
        response = handleFuelExpensesIntent(entities, query);
        break;
      
      case 'transport_expenses':
        response = handleTransportExpensesIntent(entities, query);
        break;
      
      case 'beneficiary_info':
        response = handleBeneficiaryIntent(entities, query);
        // Set context for potential followup
        if (response.includes('break this down by specific teams')) {
          chatContext.awaitingFollowup = true;
          chatContext.followupType = 'team_breakdown';
        }
        break;
      
      case 'team_info':
        response = handleTeamIntent(entities, query);
        break;
      
      case 'vehicle_info':
        response = handleVehicleIntent(entities, query);
        break;
      
      case 'summary_report':
        response = handleSummaryIntent(entities, query);
        break;
      
      case 'greeting':
        response = handleGreeting(query);
        break;
      
      case 'help':
        response = handleHelpRequest(query);
        break;
      
      case 'comparison':
        response = handleComparisonIntent(entities, query);
        break;
      
      case 'time_based':
        response = handleTimeBasedIntent(entities, query);
        break;
      
      default:
        try {
          // Use the global RAG index only (no per-conversation pinned context)
          var ctxToUse = RAG_retrieveContextForQuery(message, 5) || '';
          // Work mode should remain businesslike: always use llmAnswer, even if context is empty
          var responseWork = (typeof Chat !== 'undefined' && Chat && typeof Chat.llmAnswer === 'function')
            ? Chat.llmAnswer(message, ctxToUse, null)
            : null;
          response = responseWork || handleUnknownIntentWithLearning(query, entities);
        } catch (llmErr2) {
          console.error('LLM work-mode fallback error:', llmErr2);
          response = handleUnknownIntentWithLearning(query, entities);
        }
        break;
    }
    
    // Log conversation for learning
    const responseTime = new Date() - startTime;
    const conversationId = logConversation(message, response, intent.type || intent, entities, startTime, responseTime);
    
    // Update context
    chatContext = {
      lastQuery: message,
      lastIntent: intent.type || intent,
      lastResponse: response,
      awaitingFollowup: chatContext.awaitingFollowup,
      followupType: chatContext.followupType
    };
    
    // Add feedback prompt to response
    return addFeedbackPrompt(response, conversationId);
    
  } catch (error) {
    console.error('Enhanced chat query processing error:', error);
    const errorResponse = 'I apologize, but I encountered an issue processing your request. I\'m learning from this to improve future responses.';
    const conversationId = logConversation(message, errorResponse, 'error', {}, startTime);
    return addFeedbackPrompt(errorResponse, conversationId);
  }
}

/**
 * ---------- RAG (Retrieval-Augmented Generation) helpers ----------
 * Index and retrieve Drive + Sheets content for grounded LLM answers.
 * Configure with Script Properties:
 *  - RAG_FOLDER_ID (optional): crawl this folder recursively
 *  - RAG_SHEET_IDS (optional): comma-separated list of Google Sheet file IDs
 */

function RAG_indexFolder(folderId) {
  // Allow calling without an argument: use configured or default folder ID
  if (!folderId || String(folderId).trim() === '') {
    try { RAG_ensureDefaultFolderConfigured(); } catch(_){ }
    const sp = PropertiesService.getScriptProperties();
    folderId = (sp.getProperty('RAG_FOLDER_ID') || RAG_DEFAULT_FOLDER_ID || '').trim();
  }
  if (!folderId) throw new Error('RAG_indexFolder: folderId required (set Script Property RAG_FOLDER_ID or RAG_DEFAULT_FOLDER_ID)');

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('RAG_Index') || ss.insertSheet('RAG_Index');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['source','fileId','name','mimeType','path','chunkIndex','text']);
  }

  const folder = DriveApp.getFolderById(folderId);
  let count = 0;

  function appendChunks(file, path, text) {
    if (!text) return;
    const chunks = RAG_chunk(text, 1500);
    const rows = chunks.map((c, i) => ['drive', file.getId(), file.getName(), file.getMimeType(), path, i, c]);
    if (rows.length) sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
    count += rows.length;
  }

  function walk(f, path) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      try {
        const text = RAG_extractText(file);
        appendChunks(file, path, text);
      } catch (e) {
        console.log('RAG extract error for ' + file.getName() + ': ' + e);
      }
    }
    const sub = f.getFolders();
    while (sub.hasNext()) {
      const sf = sub.next();
      walk(sf, path + '/' + sf.getName());
    }
  }

  walk(folder, '/' + folder.getName());
  // Record last index time
  try { PropertiesService.getScriptProperties().setProperty('RAG_INDEXED_AT', String(Date.now())); } catch(_){ }
  return 'Indexed chunks: ' + count;
}

function RAG_indexSheets(sheetIdsCsv) {
  if (!sheetIdsCsv) throw new Error('RAG_indexSheets: sheetIdsCsv required');
  const ids = String(sheetIdsCsv).split(',').map(s => s.trim()).filter(Boolean);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('RAG_Index') || ss.insertSheet('RAG_Index');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['source','fileId','name','mimeType','path','chunkIndex','text']);
  }
  let count = 0;
  ids.forEach(id => {
    try{
      const ext = RAG_extractTextFromSheetId(id, 250);
      const chunks = RAG_chunk(ext.text, 1500);
      const rows = chunks.map((c,i)=>['sheet', id, ext.title, 'GOOGLE_SHEETS', ext.path, i, c]);
      if (rows.length) sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
      count += rows.length;
    }catch(e){ console.log('RAG sheet extract error for '+id+': '+e); }
  });
  try { PropertiesService.getScriptProperties().setProperty('RAG_INDEXED_AT', String(Date.now())); } catch(_){ }
  return 'Indexed sheet chunks: ' + count;
}

function RAG_extractText(file) {
  const mt = file.getMimeType();
  try {
    if (mt === MimeType.GOOGLE_DOCS) {
      return DocumentApp.openById(file.getId()).getBody().getText();
    }
    if (mt === MimeType.GOOGLE_SHEETS) {
      const ext = RAG_extractTextFromSheetId(file.getId(), 200);
      return ext.text;
    }
    if (mt === MimeType.PLAIN_TEXT || mt === MimeType.CSV || mt === MimeType.JSON) {
      return file.getBlob().getDataAsString();
    }
    // Optional: Slides text extraction if SlidesApp is enabled
  } catch (e) {
    console.log('RAG_extractText failed: ' + e);
  }
  return '';
}

function RAG_extractTextFromSheetId(fileId, maxRowsPerTab) {
  const s = SpreadsheetApp.openById(fileId);
  const parts = s.getSheets().map(sh => {
    const vals = sh.getDataRange().getDisplayValues();
    const rows = vals.slice(0, Math.min(vals.length, maxRowsPerTab || 200));
    const body = rows.map(r => r.join(' | ')).join('\n');
    return 'TAB: ' + sh.getName() + '\n' + body;
  });
  return { title: s.getName(), path: '/' + s.getName(), text: parts.join('\n\n') };
}

function RAG_chunk(text, maxLen) {
  const out = [];
  if (!text) return out;
  let i = 0;
  const n = Math.max(500, maxLen || 1500);
  while (i < text.length) {
    out.push(text.substring(i, i + n));
    i += n;
  }
  return out;
}

function RAG_retrieveContext(query, topK) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('RAG_Index');
  if (!sh) return '';
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return '';
  const header = vals[0];
  const rows = vals.slice(1);
  const COL_TEXT = header.indexOf('text');
  const COL_PATH = header.indexOf('path');
  const terms = String(query||'').toLowerCase().split(/\W+/).filter(Boolean);
  function score(t) {
    const lower = String(t||'').toLowerCase();
    var s = 0;
    for (var i=0;i<terms.length;i++) if (lower.indexOf(terms[i]) >= 0) s++;
    return s;
  }
  const scored = rows
    .map(r => ({ text: r[COL_TEXT], path: r[COL_PATH], s: score(r[COL_TEXT]) }))
    .filter(o => o.s > 0)
    .sort((a,b) => b.s - a.s)
    .slice(0, topK||5);
  if (!scored.length) return '';
  return scored.map((o,i)=> 'Snippet ' + (i+1) + ' (' + (o.path||'') + '):\n' + o.text).join('\n\n');
}

function RAG_retrieveContextForQuery(query, topK){
  // Refresh index if stale (auto-pick new/changed files in the folder)
  try {
    var sp0 = PropertiesService.getScriptProperties();
    var ttl = Number(sp0.getProperty('RAG_INDEX_TTL_MIN') || '240'); // default 4h
    if (!ttl || isNaN(ttl)) ttl = 240;
    RAG_refreshIndexIfStale(ttl);
  } catch(_e) {}

  // If there is an index sheet, use it directly
  var ctx = RAG_retrieveContext(query, topK);
  if (ctx) return ctx;
  // On-the-fly extraction from configured sources if no index exists
  var sp = PropertiesService.getScriptProperties();
  var folderId = (sp.getProperty('RAG_FOLDER_ID') || '').trim();
  var sheetIdsCsv = (sp.getProperty('RAG_SHEET_IDS') || '').trim();
  // Try Sheets first (quick, bounded)
  if (sheetIdsCsv) {
    try{
      var tmp = RAG_indexSheets(sheetIdsCsv);
      ctx = RAG_retrieveContext(query, topK);
      if (ctx) return ctx;
    }catch(e){ console.log('RAG_retrieve sheets on-the-fly failed: '+e); }
  }
  if (folderId) {
    try{
      var tmp2 = RAG_indexFolder(folderId); // may be slow for big trees
      ctx = RAG_retrieveContext(query, topK);
      if (ctx) return ctx;
    }catch(e){ console.log('RAG_retrieve folder on-the-fly failed: '+e); }
  }
  return '';
}

// Clear RAG index sheet safely
function RAG_clearIndex(){
  try{
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('RAG_Index');
    if (!sh) return 'No index sheet to clear';
    sh.clear();
    sh.appendRow(['source','fileId','name','mimeType','path','chunkIndex','text']);
    return 'Cleared RAG_Index';
  }catch(e){ return 'Clear failed: ' + String(e); }
}

// Reindex folder/sheets if the index is older than maxAgeMinutes
function RAG_refreshIndexIfStale(maxAgeMinutes){
  var sp = PropertiesService.getScriptProperties();
  var last = Number(sp.getProperty('RAG_INDEXED_AT') || '0');
  var now = Date.now();
  var maxAgeMs = Math.max(5, Number(maxAgeMinutes)||240) * 60 * 1000;
  var need = false;
  try{
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('RAG_Index');
    if (!sh || sh.getLastRow() <= 1) need = true; // no index yet
  }catch(_){ need = true; }
  if (!need && last && (now - last) < maxAgeMs) return { refreshed:false, reason:'fresh' };

  var folderId = (sp.getProperty('RAG_FOLDER_ID') || '').trim();
  var sheetIdsCsv = (sp.getProperty('RAG_SHEET_IDS') || '').trim();
  if (!folderId && !sheetIdsCsv) return { refreshed:false, reason:'no sources' };

  // Rebuild index (clear then index)
  RAG_clearIndex();
  var res = [];
  if (sheetIdsCsv){ try{ res.push(RAG_indexSheets(sheetIdsCsv)); }catch(e){ res.push('Sheets error: '+e); } }
  if (folderId){    try{ res.push(RAG_indexFolder(folderId)); }catch(e){ res.push('Folder error: '+e); } }
  return { refreshed:true, result: res.join(' | ') };
}

// If no RAG_FOLDER_ID is set, set it to default once (does not overwrite a custom value)
function RAG_ensureDefaultFolderConfigured(){
  try{
    var sp = PropertiesService.getScriptProperties();
    var cur = (sp.getProperty('RAG_FOLDER_ID') || '').trim();
    if (!cur && RAG_DEFAULT_FOLDER_ID && RAG_DEFAULT_FOLDER_ID.trim()){
      sp.setProperty('RAG_FOLDER_ID', RAG_DEFAULT_FOLDER_ID.trim());
    }
  }catch(_){ }
}

/** Force a RAG index refresh: clear index, reset timestamp, and re-run indexing (folder/sheets).
 * Run this from Apps Script to force RAG to pick up latest document/sheet changes.
 */
function RAG_forceRefresh(){
  try{
    var sp = PropertiesService.getScriptProperties();
    // Clear index sheet
    try{ RAG_clearIndex(); }catch(_){ }
    // Reset timestamp so refresh will run
    try{ sp.setProperty('RAG_INDEXED_AT','0'); }catch(_){ }
    // Attempt to reindex based on configured sources (will use RAG_FOLDER_ID or RAG_SHEET_IDS)
    var folderId = (sp.getProperty('RAG_FOLDER_ID') || '').trim();
    var sheetIdsCsv = (sp.getProperty('RAG_SHEET_IDS') || '').trim();
    var res = [];
    if (sheetIdsCsv){ try{ res.push(RAG_indexSheets(sheetIdsCsv)); }catch(e){ res.push('Sheets error: '+String(e)); } }
    if (folderId){ try{ res.push(RAG_indexFolder(folderId)); }catch(e){ res.push('Folder error: '+String(e)); } }
    // Ensure timestamp is set
    try{ sp.setProperty('RAG_INDEXED_AT', String(Date.now())); }catch(_){ }
    return { ok:true, result: res };
  }catch(e){ console.error('RAG_forceRefresh error', e); return { ok:false, error:String(e) }; }
}



// Convenience config helpers (no secrets)
function RAG_setFolder(folderId){
  var id = String(folderId||'').trim();
  if (!id) {
    try {
      if (typeof RAG_DEFAULT_FOLDER_ID !== 'undefined' && RAG_DEFAULT_FOLDER_ID && String(RAG_DEFAULT_FOLDER_ID).trim()) {
        id = String(RAG_DEFAULT_FOLDER_ID).trim();
      }
    } catch(_){ /* ignore */ }
  }
  if (!id) throw new Error('folderId required — pass an ID or set RAG_DEFAULT_FOLDER_ID');
  PropertiesService.getScriptProperties().setProperty('RAG_FOLDER_ID', id);
  return 'RAG_FOLDER_ID set to: ' + id;
}
function RAG_setSheetsCsv(sheetIdsCsv){
  if (!sheetIdsCsv) throw new Error('sheetIdsCsv required');
  PropertiesService.getScriptProperties().setProperty('RAG_SHEET_IDS', String(sheetIdsCsv));
  return 'RAG_SHEET_IDS set.';
}

/**
 * Decide whether to use LLM fallback (DeepSeek/OpenRouter)
 * Controlled by Script Property CHAT_USE_LLM=true|false or query prefix 'ai:'
 */
function _shouldUseLLM_(query){
  try {
    var sp = PropertiesService.getScriptProperties();
    var flag = (sp.getProperty('CHAT_USE_LLM') || '').toLowerCase();
    if (flag === 'true' || flag === '1' || flag === 'yes') return true;
  } catch(_){ }
  return /^\s*ai\s*[:\-]/i.test(String(query||''));
}

/**
 * Analyze user intent with confidence scoring
 */
function analyzeIntent(query) {
  const intents = [
    {
      type: 'greeting',
      keywords: ['hello', 'hi', 'hey', 'good morning', 'good afternoon', 'greetings'],
      weight: 1.0
    },
    {
      type: 'help',
      keywords: ['help', 'what can you do', 'how to', 'guide', 'assist', 'support'],
      weight: 1.0
    },
    {
      type: 'monthly_expenses',
      keywords: ['monthly', 'month', 'expense', 'spending', 'cost this month'],
      weight: 0.8
    },
    {
      type: 'fuel_expenses',
      keywords: ['fuel', 'petrol', 'gas', 'gasoline', 'diesel'],
      weight: 0.9
    },
    {
      type: 'transport_expenses',
      keywords: ['transport', 'travel', 'trip', 'journey', 'commute'],
      weight: 0.9
    },
    {
      type: 'beneficiary_info',
      keywords: ['beneficiary', 'beneficiaries', 'recipient', 'who received', 'which beneficiary', 'largest', 'highest', 'most', 'top'],
      weight: 0.9
    },
    {
      type: 'team_info',
      keywords: ['team', 'teams', 'group', 'department'],
      weight: 0.8
    },
    {
      type: 'vehicle_info',
      keywords: ['car', 'vehicle', 'available', 'release', 'automobile'],
      weight: 0.9
    },
    {
      type: 'summary_report',
      keywords: ['summary', 'total', 'overview', 'report', 'breakdown'],
      weight: 0.8
    },
    {
      type: 'comparison',
      keywords: ['compare', 'vs', 'versus', 'difference', 'higher', 'lower', 'more', 'less'],
      weight: 0.7
    },
    {
      type: 'time_based',
      keywords: ['last week', 'this week', 'yesterday', 'today', 'recent', 'latest'],
      weight: 0.7
    }
  ];
  
  let bestMatch = { type: 'unknown', confidence: 0 };
  
  for (const intent of intents) {
    let score = 0;
    let matches = 0;
    
    for (const keyword of intent.keywords) {
      if (query.includes(keyword)) {
        score += intent.weight;
        matches++;
      }
    }
    
    const confidence = matches > 0 ? (score / intent.keywords.length) * matches : 0;
    
    if (confidence > bestMatch.confidence) {
      bestMatch = { type: intent.type, confidence: confidence };
    }
  }
  
  return bestMatch;
}

/**
 * Enhanced intent analysis with learning capabilities
 */
function analyzeIntentWithLearning(query) {
  const baseIntent = analyzeIntent(query);
  
  // Check if we have learned patterns for this query
  const learnedIntent = checkLearnedPatterns(query);
  if (learnedIntent && learnedIntent.confidence > baseIntent.confidence) {
    return learnedIntent;
  }
  
  return baseIntent;
}

/**
 * Check learned patterns from previous interactions
 */
function checkLearnedPatterns(query) {
  try {
    // Simple pattern matching based on conversation history
    for (const conversation of conversationLog) {
      if (conversation.query && conversation.intent) {
        const similarity = calculateSimilarity(query, conversation.query);
        if (similarity > 0.7) { // 70% similarity threshold
          return {
            type: conversation.intent,
            confidence: similarity * 2, // Boost learned patterns
            learned: true
          };
        }
      }
    }
  } catch (error) {
    console.log('Error checking learned patterns:', error);
  }
  
  return null;
}

/**
 * Calculate similarity between two strings (simple implementation)
 */
function calculateSimilarity(str1, str2) {
  const words1 = str1.toLowerCase().split(' ');
  const words2 = str2.toLowerCase().split(' ');
  
  let matches = 0;
  for (const word1 of words1) {
    if (words2.includes(word1)) {
      matches++;
    }
  }
  
  return matches / Math.max(words1.length, words2.length);
}

/**
 * Log conversation for learning purposes
 */
function logConversation(query, response, intent, entities, startTime, responseTime) {
  const conversationId = 'conv_' + new Date().getTime();
  
  const logEntry = {
    id: conversationId,
    timestamp: new Date().toISOString(),
    query: query,
    response: response,
    intent: intent,
    entities: entities,
    responseTime: responseTime || 0,
    feedback: null
  };
  
  // Keep only last 100 conversations to manage memory
  if (conversationLog.length >= 100) {
    conversationLog.shift();
  }
  
  conversationLog.push(logEntry);
  console.log('Logged conversation:', conversationId);
  
  return conversationId;
}

/**
 * Handle user rating feedback
 */
function handleRatingFeedback(message) {
  try {
    const parts = message.split(' ');
    if (parts.length >= 3) {
      const rating = parseInt(parts[1]);
      const conversationId = parts[2];
      
      if (rating >= 1 && rating <= 5) {
        // Store feedback
        userFeedback[conversationId] = {
          rating: rating,
          timestamp: new Date().toISOString()
        };
        
        // Update conversation log with feedback
        const conversation = conversationLog.find(c => c.id === conversationId);
        if (conversation) {
          conversation.feedback = rating;
          
          // Update intent accuracy tracking
          if (!intentAccuracy[conversation.intent]) {
            intentAccuracy[conversation.intent] = { total: 0, positive: 0 };
          }
          intentAccuracy[conversation.intent].total++;
          if (rating >= 4) {
            intentAccuracy[conversation.intent].positive++;
          }
        }
        
  // Delegated to responses.gs
  return formatRatingThankYou(rating);
      }
    }
    
    return 'Please provide feedback in format: "rate [1-5] [conversation_id]"';
  } catch (error) {
    console.error('Error handling rating feedback:', error);
    return 'Sorry, I had trouble processing your feedback. Please try again.';
  }
}

/**
 * Add feedback prompt to response
 */
function addFeedbackPrompt(response, conversationId) {
  // Previously appended an inline rating instruction. We no longer add that server-side;
  // the client renders a compact star-based rating UI instead.
  return response;
}

/**
 * Enhanced unknown intent handler with learning
 */
function handleUnknownIntentWithLearning(query, entities) {
  // Check if this is similar to previous queries
  const similarConversations = findSimilarConversations(query);
  
  if (similarConversations.length > 0) {
    const bestMatch = similarConversations[0];
    return formatUnknownIntentWithLearning(bestMatch);
  }
  // Default response delegated to responses.gs
  return formatUnknownIntentWithLearning(null);
}

/**
 * Find similar conversations from history
 */
function findSimilarConversations(query) {
  const similarities = [];
  
  for (const conversation of conversationLog) {
    if (conversation.feedback && conversation.feedback >= 4) { // Only consider well-rated responses
      const similarity = calculateSimilarity(query, conversation.query);
      if (similarity > 0.5) {
        similarities.push({
          ...conversation,
          similarity: similarity
        });
      }
    }
  }
  
  return similarities.sort((a, b) => b.similarity - a.similarity);
}

/**
 * Extract entities (numbers, dates, names) from query
 */
function extractEntities(query) {
  const entities = {
    amounts: [],
    dates: [],
    names: [],
    timeframes: [],
    monthYear: null,
    dateRange: null
  };
  
  // Extract monetary amounts
  const amountRegex = /\$?([0-9,]+(?:\.[0-9]{2})?)/g;
  let match;
  while ((match = amountRegex.exec(query)) !== null) {
    entities.amounts.push(parseFloat(match[1].replace(',', '')));
  }
  
  // Extract time references
  const timeframes = ['january', 'february', 'march', 'april', 'may', 'june',
                     'july', 'august', 'september', 'october', 'november', 'december',
                     'last month', 'this month', 'next month', 'week', 'year'];
  
  for (const timeframe of timeframes) {
    if (query.includes(timeframe)) {
      entities.timeframes.push(timeframe);
    }
  }

  // Parse month/year including month-only phrases; accept loose parsing
  try {
    const my = parseMonthYearLoose(query);
    if (my) entities.monthYear = my; // { month:0-11, year }
  } catch(_e) {}

  // Parse explicit date ranges like "02-Sep-25 to 25-Sep-25" or "02/09/2025 - 25/09/2025"
  try {
    const dr = parseDateRangeFromText(query);
    if (dr && dr.start && dr.end) entities.dateRange = dr; // { start:Date, end:Date }
  } catch(_e2) {}

  // Detect beneficiary names from DD + Submissions
  try {
    const known = getKnownBeneficiariesSet(); // Set of canonical names
    if (known && known.size) {
      const q = query.toLowerCase();
      // Prefer longer names first to avoid partial overlaps
      const allNames = Array.from(known.values()).sort((a,b)=> b.length - a.length);
      const matched = [];
      const used = new Set();
      allNames.forEach(name => {
        const nlc = name.toLowerCase();
        if (used.has(nlc)) return;
        // word-boundary loose match; fallback to indexOf
        const re = new RegExp('(^|[^a-z0-9])' + nlc.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '([^a-z0-9]|$)');
        if (re.test(q)) {
          matched.push(name);
          used.add(nlc);
        }
      });
      entities.names = matched;
    }
  } catch(_e2) {}

  return entities;
}

/**
 * ---------------- NL → Plan → Execute (lightweight) ----------------
 * A minimal planner and executor that work over your existing Sheets.
 * Plan types: aggregate | rank | compare | lookup
 * Metrics: total | fuel | erda | vehicleRent | airtime | transport | misc
 */

/** Map metric synonyms to canonical metric keys used by executor */
function _metricFromText_(q){
  q = String(q||'').toLowerCase();
  if (/\bfuel|petrol|diesel\b/.test(q)) return 'fuel';
  if (/\b(er\s?da|da)\b/.test(q)) return 'erda';
  if (/(vehicle\s*rent|car\s*amt|car\b)/.test(q)) return 'vehicleRent';
  if (/\bair(time)?\b/.test(q)) return 'airtime';
  if (/\btransport|travel\b/.test(q)) return 'transport';
  if (/\bmisc|miscellaneous\b/.test(q)) return 'misc';
  if (/\btotal|overall|expense|spend|cost\b/.test(q)) return 'total';
  return 'total';
}

/** Detect groupBy (“by team/project/beneficiary”) */
function _groupByFromText_(q){
  q = String(q||'').toLowerCase();
  if (/\bby\s+team|team[-\s]?wise\b/.test(q)) return 'team';
  if (/\bby\s+project|project[-\s]?wise\b/.test(q)) return 'project';
  if (/\bby\s+beneficiary|beneficiary[-\s]?wise\b/.test(q)) return 'beneficiary';
  return 'none';
}

/** Extract a topN like "top 5" */
function _topNFromText_(q){
  const m = String(q||'').toLowerCase().match(/top\s+(\d{1,3})/);
  return m ? Math.max(1, parseInt(m[1],10)) : null;
}

/** Convert timeframe token to {month,year} using existing parseMonthYear */
function _resolveTimeframe_(q){
  const now = new Date();
  const ql = String(q||'').toLowerCase();
  const my = parseMonthYear(ql);
  if (my) return my; // {month,year}
  if (/\bthis\s+month\b/.test(ql)) return { month: now.getMonth(), year: now.getFullYear() };
  if (/\blast\s+month\b/.test(ql)) { const d=new Date(now.getFullYear(),now.getMonth()-1,1); return {month:d.getMonth(),year:d.getFullYear()}; }
  if (/\bnext\s+month\b/.test(ql)) { const d=new Date(now.getFullYear(),now.getMonth()+1,1); return {month:d.getMonth(),year:d.getFullYear()}; }
  // default current month
  return { month: now.getMonth(), year: now.getFullYear() };
}

/** Build a small plan JSON from a natural-language message */
function planAnalysis(message){
  try{
    let q = String(message||'');
    // Detect explicit output preference anywhere in the message: "Output: table|summary|total"
    // Accept variants like 'output:table', 'Output - table', case-insensitive. Prefer last occurrence.
    let explicitOutput = null;
    try{
      const outMatches = q.match(/\boutput\s*[:\-]\s*(table|summary|total)\b/ig);
      if (outMatches && outMatches.length){
        const last = outMatches[outMatches.length - 1];
        const m = last.match(/\boutput\s*[:\-]\s*(table|summary|total)\b/i);
        if (m && m[1]) explicitOutput = m[1].toLowerCase();
        // remove all output directives from the message before further analysis
        q = q.replace(/\boutput\s*[:\-]\s*(table|summary|total)\b/ig, '').trim();
      }
    }catch(_eOut){}
    const ql = q.toLowerCase();
    const entities = extractEntities(ql);

    // Team-based vehicle queries
    // e.g., "Which vehicle team H-Mas-TI is using now?"
    const teamMatch = /\bteam\s+([a-z0-9\-_/\. ]{2,})/i.exec(q);
    const teamToken = teamMatch ? teamMatch[1].trim() : null;
    const asksUsingNow = /(which\s+vehicle.*using|using\s+now|currently\s+using|in\s+use\s+now)/i.test(q);
    const asksLast3 = /\blast\s*3\b|\b3\s*(entries|records|vehicles)\b/i.test(q);
    const want2Release1InUse = /(2\s*releases?\s*and\s*1\s*in\s*use|two\s*releases?\s*and\s*one\s*in\s*use)/i.test(q);

    if (teamToken && asksUsingNow){
      return {
        type: 'lookup',
        entity: 'vehicle',
        filters: { team: teamToken, status: 'IN USE' },
        order: 'desc',
        latestPer: 'Vehicle Number',
        limit: 10,
        join: { vehicleMaster: true },
        output: 'list'
      };
    }
    // Latest vehicle number for a team (e.g., "latest vehicle number of H-Mas-TI")
    const asksLatest = /(latest|current|recent|newest)/i.test(q);
    const mentionsVehNum = /(vehicle\s*number|car\s*number|reg(istration)?)/i.test(q);
    if (teamToken && asksLatest && mentionsVehNum){
      return {
        type: 'lookup',
        entity: 'vehicle_latest_for_team',
        filters: { team: teamToken, preferStatus: 'IN USE' },
        order: 'desc',
        output: 'one'
      };
    }
    if (teamToken && asksLast3){
      const pattern = want2Release1InUse ? ['RELEASE','RELEASE','IN USE'] : null;
      return {
        type: 'lookup',
        entity: 'vehicle_history',
        filters: { team: teamToken },
        order: 'desc',
        historyPattern: pattern || ['RELEASE','RELEASE','IN USE'],
        limit: 3,
        join: { vehicleMaster: true },
        output: 'list'
      };
    }

    // Vehicle lookups
    if (/\b(available|release|released)\b/.test(ql) && /(car|vehicle)s?\b/.test(ql)){
      return { type:'lookup', entity:'vehicle', filters:{ latestStatus:'RELEASE' }, output:'list' };
    }
    if (/\b(history)\b/.test(ql) && /(car|vehicle)\b/.test(ql)){
      // Requires car number in follow-up; return a partial plan
      return { type:'lookup', entity:'vehicle_history', filters:{ carNumber: '' }, output:'sentence' };
    }

    // Compare
    if (/(compare|vs|versus)/.test(ql)){
      const tf = _resolveTimeframe_(ql);
      const wantFuel = /fuel|petrol|diesel/.test(ql);
      const wantTrans = /transport|travel/.test(ql);
      const metrics = (wantFuel && wantTrans) ? ['fuel','transport'] : ['fuel','transport'];
      return { type:'compare', metrics, filters:{ timeframe: tf }, output:'summary' };
    }

    // Rank / aggregate over submissions
    const metric = _metricFromText_(ql);
    const groupBy = _groupByFromText_(ql);
    const topN = _topNFromText_(ql);
  const tf = _resolveTimeframe_(ql);
  // Detect explicit date ranges like 'from X to Y' anywhere in the query
  const explicitRange = parseDateRangeFromText(q);

    // Beneficiary filter if a single known name was detected
    let filterBen = null;
    if (entities && Array.isArray(entities.names) && entities.names.length === 1){
      filterBen = entities.names[0];
    }

  const base = { metric, groupBy, filters:{ timeframe: tf }, output: explicitOutput || 'summary' };
    if (filterBen) base.filters.beneficiary = filterBen;
    // If user mentioned a specific metric and provided an explicit date range, prefer metric-range filtering
    if (explicitRange && metric && metric !== 'total'){
      base.filters.metricRange = { metric: metric, start: explicitRange.start, end: explicitRange.end };
      // also set a flag so executor uses metric ranges instead of Timestamp if available
      base.filters.useMetricRange = true;
    } else if (explicitRange){
      // If no metric explicitly mentioned, attach timeframe from explicit range as filters.timeframeExact
      base.filters.timeframeExact = explicitRange;
    }

    if (groupBy !== 'none' || topN){
      if (topN) base.topN = topN;
      return Object.assign({ type:'rank' }, base);
    }
    return Object.assign({ type:'aggregate' }, base);
  }catch(e){
    console.error('planAnalysis error:', e);
    return null;
  }
}

/** Header indices for submissions sheet (tolerant to both schemas) */
function _submissionsHeaderIndices_(head){
  const IX = _headerIndex_(head);
  const get = (labels, req)=>{ try{ return IX.get(labels); }catch(e){ if(req) throw e; return -1; } };
  const iTs   = get(['Timestamp','Date','Date and time of entry','Date and time'], true);
  const iBen  = get(['Beneficiary'], false);
  const iTeam = get(['Team','Team Name'], false);
  const iProj = get(['Project','Project Name'], false);
    const m = {
      total:       get(['Row Total','Total Amount','Total'], false),
      fuel:        get(['Fuel Amt','Fuel'], false),
      erda:        get(['ER DA Amt','ER DA'], false),
      // Avoid plain 'Vehicle Rent' which may match range headers; prefer Amt/Amount variants
      vehicleRent: get(['Car Amt','Car Amount','Vehicle Rent Amount','Vehicle Rent Amt'], false),
      airtime:     get(['Airtime Amount','Airtime Amt','Airtime'], false),
      transport:   get(['Transport Amount','Transport Amt','Transport'], false),
      misc:        get(['Misc Amt','Misc','Miscellaneous'], false)
    };
  // Post-fix: if vehicleRent points to a likely date/range column, try to refine.
  try{
    if (m.vehicleRent >= 0){
      const hvr = String(head[m.vehicleRent]||'').toLowerCase();
      if (/(from|to|date)/.test(hvr)){
        // Find a better match explicitly containing amt/amount and not from/to/date
        for (let i=0;i<head.length;i++){
          const h = String(head[i]||'').toLowerCase();
          if ((/vehicle/.test(h) || /car/.test(h)) && (/(amt|amount)/.test(h)) && !/(from|to|date)/.test(h)){
            m.vehicleRent = i; break;
          }
        }
      }
    }
  }catch(_eVR){}
  // If users provide a fixed column layout (by letter), allow a direct column index override
  // Common mapping provided by user: Total=G(7), Fuel=K(11), DA=N(14), Vehicle Rent=R(18), Airtime=U(21), Transport=X(24), Misc=AA(27)
  try{
    const headLen = Array.isArray(head) ? head.length : 0;
    const fixed = { total:7, fuel:11, erda:14, vehicleRent:18, airtime:21, transport:24, misc:27 };
    const needCols = Math.max(fixed.total, fixed.fuel, fixed.erda, fixed.vehicleRent, fixed.airtime, fixed.transport, fixed.misc);
    if (headLen >= needCols){
      // override with 0-based indices from provided column letters
      Object.keys(fixed).forEach(k => { m[k] = fixed[k] - 1; });
    }
  }catch(_e){}
  return { iTs, iBen, iTeam, iProj, metrics:m };
}

/**
 * Discover metric-specific From/To columns (best-effort) from header row.
 * Returns maps: rangeStarts and rangeEnds with metric keys (fuel, erda, ...)
 */
function _submissionsMetricRangeIndices_(head){
  const res = { rangeStarts: {}, rangeEnds: {} };
  try{
    const H = head.map(h => String(h||'').toLowerCase());
    const metricTokens = {
      fuel: ['fuel','petrol','diesel'],
      erda: ['er da','erda','da'],
      vehicleRent: ['car amt','car','vehicle rent','vehicle'],
      airtime: ['airtime','air time'],
      transport: ['transport','travel'],
      misc: ['misc','miscellaneous']
    };

    Object.keys(metricTokens).forEach(metric => {
      const tokens = metricTokens[metric];
      const tokenRegexes = tokens.map(t => new RegExp('\\b' + t.replace(/[.*+?^${}()|[\\]\\]/g,'\\$&') + '\\b'));
      let startIdx = -1; let endIdx = -1;
      const candidateDateCols = [];
      const tokenCols = [];

      // First pass: find columns that mention the metric (word-boundary) and tag obvious from/to/date markers
      for (let i=0;i<H.length;i++){
        const h = H[i];
        const hasToken = tokenRegexes.some(rx => rx.test(h));
        if (!hasToken) continue;
        tokenCols.push(i);
        if (startIdx < 0 && (/\bfrom\b/.test(h) || /\bstart\b/.test(h) || h.includes('from date') || h.includes('start date') || h.includes('_from') || h.includes('from_'))) startIdx = i;
        if (endIdx < 0 && (/\bto\b/.test(h) || /\bend\b/.test(h) || h.includes('to date') || h.includes('end date') || h.includes('_to') || h.includes('to_') || h.includes('end_'))) endIdx = i;
        if (h.includes('date') || h.includes('day')) candidateDateCols.push(i);
      }

      // Second pass: if explicit from/to not found, look at neighbouring columns around token columns
      if ((startIdx < 0 || endIdx < 0) && tokenCols.length){
        for (const i of tokenCols){
          // scan up to 2 columns left/right for 'from'/'to' or 'date' markers
          for (let d=-2; d<=2; d++){
            if (d===0) continue;
            const j = i + d; if (j < 0 || j >= H.length) continue;
            const hj = H[j];
            if (startIdx < 0 && (/\bfrom\b/.test(hj) || /\bstart\b/.test(hj) || hj.includes('from date') || hj.includes('start date'))) startIdx = j;
            if (endIdx < 0 && (/\bto\b/.test(hj) || /\bend\b/.test(hj) || hj.includes('to date') || hj.includes('end date'))) endIdx = j;
            if (startIdx >= 0 && endIdx >= 0) break;
          }
          if (startIdx >= 0 && endIdx >= 0) break;
        }
      }

      // Third fallback: use any 'date' columns near token columns
      if ((startIdx < 0 || endIdx < 0) && candidateDateCols.length && tokenCols.length){
        // pick the date column closest to the first token column
        const base = tokenCols[0];
        candidateDateCols.sort((a,b)=> Math.abs(a-base) - Math.abs(b-base));
        const pick = candidateDateCols[0];
        if (startIdx < 0) startIdx = pick;
        if (endIdx < 0) endIdx = pick;
      }

      // Final normalization: if still missing, leave as -1
      res.rangeStarts[metric] = (typeof startIdx === 'number') ? startIdx : -1;
      res.rangeEnds[metric] = (typeof endIdx === 'number') ? endIdx : -1;
    });
  }catch(_){ }
  return res;
}

/**
 * Given a row and submission header indices (IX), return an object {start:Date,end:Date}
 * for the provided metric (e.g., 'fuel'). Falls back to Timestamp as a single-day range.
 */
function _rowMetricRange_(row, dispRow, IX, metric){
  try{
    metric = String(metric||'').trim();
    if (!metric) metric = 'total';
    // IX may not have range maps; build from head if available
    const headRange = (IX && IX._rangeCached_) ? IX._rangeCached_ : null;
    let rangeStarts = null, rangeEnds = null;
    if (headRange){ rangeStarts = headRange.rangeStarts; rangeEnds = headRange.rangeEnds; }
    // If not cached, try to build from IX._head if available
    if (!rangeStarts){
      if (IX && Array.isArray(IX._head)){
        const rr = _submissionsMetricRangeIndices_(IX._head);
        rangeStarts = rr.rangeStarts; rangeEnds = rr.rangeEnds;
        IX._rangeCached_ = rr;
      }
    }
    // Attempt to read metric-specific from/to
    const sIdx = (rangeStarts && typeof rangeStarts[metric] === 'number') ? rangeStarts[metric] : -1;
    const eIdx = (rangeEnds && typeof rangeEnds[metric] === 'number') ? rangeEnds[metric] : -1;
    function parseCellToDate(v){
      if (!v && v !== 0) return null;
      if (v instanceof Date) return v;
      const s = String(v||'').trim();
      const dt = parseDateTokenFlexible(s);
      if (dt) return dt;
      const d2 = new Date(s);
      if (!isNaN(d2.getTime())) return d2;
      return null;
    }

  // Helper: try parse combined ranges like "01-Sep-25 to 10-Sep-25" or "01/09/2025 - 10/09/2025"
    function parseRangeFromString(s){
      if (!s) return null;
      const str = String(s).trim();
      // common separators: to, -, –, —
      const m = str.match(/(\d{1,2}[\-/][A-Za-z0-9]{1,}?[\-/]\d{2,4}|[A-Za-z]{3,}\s*\d{1,2}[,\s]*\d{2,4}|\d{1,2}\s*[A-Za-z]{3}\s*\d{2,4})(?:\s*(?:to|\-|–|—)\s*)(\d{1,2}[\-/][A-Za-z0-9]{1,}?[\-/]\d{2,4}|[A-Za-z]{3,}\s*\d{1,2}[,\s]*\d{2,4}|\d{1,2}\s*[A-Za-z]{3}\s*\d{2,4})/i);
      if (m && m[1] && m[2]){
        const a = parseDateTokenFlexible(m[1]);
        const b = parseDateTokenFlexible(m[2]);
        if (a && b) return { start: a, end: b };
      }
      // single date: treat as both start and end
      const single = parseDateTokenFlexible(str);
      if (single) return { start: single, end: single };
      return null;
    }

    let start = null, end = null;
    if (sIdx >=0){ start = parseCellToDate(row[sIdx] || dispRow[sIdx]); }
    if (eIdx >=0){ end = parseCellToDate(row[eIdx] || dispRow[eIdx]); }

    // If there is a fixed column mapping for certain metrics, prefer those and be strict
    // Strict behavior: if the mapped From/To cells are blank or unparsable, treat the row as not qualifying (return null).
    function colLetterToIndex(letter){
      if (!letter) return -1;
      const s = String(letter||'').toUpperCase().trim();
      let idx = 0;
      for (let i=0;i<s.length;i++){ const c = s.charCodeAt(i) - 64; if (c>0) idx = idx*26 + c; }
      return idx > 0 ? idx - 1 : -1;
    }
    // Fixed column fallbacks for common schemas (0-based letters shown as A1 letters for readability)
    // These are used only if header-driven detection fails, mirroring the reliable DA behavior.
    // Fuel:   From=I, To=J
    // ER DA:  From=L, To=M
    // Vehicle Rent: From=O, To=P (common schema variant where Car Amt is at R)
    const FIXED_RANGE_COLS = {
      fuel:        { from: 'I', to: 'J' },
      erda:        { from: 'L', to: 'M' },
      vehicleRent: { from: 'O', to: 'P' },
      airtime:     { from: 'S', to: 'T' },
      transport:   { from: 'V', to: 'W' },
      misc:        { from: 'Y', to: 'Z' }
    };
    if (FIXED_RANGE_COLS[metric]){
      const fromIdx = colLetterToIndex(FIXED_RANGE_COLS[metric].from);
      const toIdx = colLetterToIndex(FIXED_RANGE_COLS[metric].to);
      let sVal = null, eVal = null;
      try{ sVal = (row && typeof row[fromIdx] !== 'undefined') ? row[fromIdx] : (dispRow && dispRow[fromIdx]); }catch(_e){}
      try{ eVal = (row && typeof row[toIdx] !== 'undefined') ? row[toIdx] : (dispRow && dispRow[toIdx]); }catch(_e){}
      const sDate = parseCellToDate(sVal);
      const eDate = parseCellToDate(eVal);
      if (!sDate || !eDate) return null; // strict: leave row when missing/unparsable
      return { start: sDate, end: eDate };
    }

    if (sIdx >=0){ start = parseCellToDate(row[sIdx] || dispRow[sIdx]); }
    if (eIdx >=0){ end = parseCellToDate(row[eIdx] || dispRow[eIdx]); }

    // If missing start/end, try parsing any cell in the row for a combined range string
    if (!start || !end){
      try{
        const len = Math.max(Array.isArray(row)?row.length:0, Array.isArray(dispRow)?dispRow.length:0);
        for (let i=0;i<len && (!start || !end); i++){
          const raw = (dispRow && dispRow[i]) || row[i] || '';
          if (!raw) continue;
          const cand = parseRangeFromString(raw);
          if (cand){ start = start || cand.start; end = end || cand.end; break; }
        }
      }catch(_e){}
    }
    // If we have only one side, normalize
    if (start && !end) end = start;
    if (!start && end) start = end;
    // If still missing, fallback to Timestamp
    if (!start || !end){
      const tsIdx = (IX && typeof IX.iTs !== 'undefined') ? IX.iTs : -1;
      if (tsIdx >= 0){
        const rawTs = (row[tsIdx] instanceof Date) ? row[tsIdx] : (row[tsIdx] || dispRow[tsIdx] || '');
        const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
        if (!isNaN(d.getTime())){ start = start || d; end = end || d; }
      }
    }
    if (start && end) return { start: start, end: end };
  }catch(_){ }
  return null;
}

/** Return true if ranges [aStart,aEnd] and [bStart,bEnd] overlap (inclusive) */
function _rangesOverlap_(aStart,aEnd,bStart,bEnd){
  if (!aStart || !aEnd || !bStart || !bEnd) return false;
  try {
    // Normalize dates to Tanzania local day boundaries for stable comparisons
    const fmtDate = (d)=> Utilities.formatDate(new Date(d), TZ(), 'yyyy-MM-dd');
    const aS = fmtDate(aStart), aE = fmtDate(aEnd);
    const bS = fmtDate(bStart), bE = fmtDate(bEnd);
    return (aS <= bE) && (bS <= aE);
  } catch(e) {
    console.error('_rangesOverlap_ error:', e);
    return false;
  }
}

/** Test helper: log first 10 submissions with metric-range overlap checks for Fuel */
function test_metricRangeFiltering(){
  const sh = getSheet('submissions'); if (!sh) return 'no sheet';
  const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn(); if (lastRow<=1) return 'no data';
  const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const rng = sh.getRange(2,1,Math.min(20,lastRow-1), lastCol);
  const vals = rng.getValues(); const disp = rng.getDisplayValues();
  const IX = _submissionsHeaderIndices_(head);
  IX._head = head;
  const metric = 'fuel';
  const sampleStart = parseDateTokenFlexible('01-Sep-2025');
  const sampleEnd = parseDateTokenFlexible('15-Sep-2025');
  const out = [];
  for (let r=0;r<vals.length;r++){
    const row = vals[r];
    const rr = _rowMetricRange_(row, disp[r], IX, metric);
    out.push({row: r+2, metricRange: rr, overlaps: rr ? _rangesOverlap_(rr.start, rr.end, sampleStart, sampleEnd) : false});
  }
  return out;
}



/** Core aggregator over submissions with filters/grouping */
function _aggregateSubmissions_({metric, groupBy, filters}){
  const sh = getSheet('submissions');
  if (!sh) return { ok:false, reason:'no_submissions' };
  const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn();
  if (lastRow <= 1 || lastCol <= 0) return { ok:true, rows:[], groups: new Map() };
  const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const IX = _submissionsHeaderIndices_(head);
  const rng = sh.getRange(2,1,lastRow-1,lastCol);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();

  const tf = (filters && filters.timeframe) ? filters.timeframe : null;
  const month = (tf && (typeof tf.month === 'number')) ? tf.month : null;
  const year  = (tf && (typeof tf.year === 'number'))  ? tf.year  : null;
  const wantBen  = (filters && filters.beneficiary) ? String(filters.beneficiary).toLowerCase() : null;
  const wantTeam = (filters && filters.team) ? String(filters.team).toLowerCase() : null;
  const wantProj = (filters && filters.project) ? String(filters.project).toLowerCase() : null;

  const midx = IX.metrics[metric] >= 0 ? IX.metrics[metric] : -1;
  const gkey = (row)=>{
    if (groupBy === 'beneficiary' && IX.iBen>=0) return String(row[IX.iBen]||'').trim()||'Unknown';
    if (groupBy === 'team' && IX.iTeam>=0) return String(row[IX.iTeam]||'').trim()||'Unknown Team';
    if (groupBy === 'project' && IX.iProj>=0) return String(row[IX.iProj]||'').trim()||'Unknown Project';
    return 'Total';
  };

  const map = new Map();
  // Attempt to detect metric-range columns for the sheet head (cached)
  IX._head = IX._head || head; // store head for range detection helper
  const metricRanges = _submissionsMetricRangeIndices_(head);
  IX._rangeCached_ = metricRanges;

  for (let r=0;r<vals.length;r++){
    const row = vals[r];
    // Time filter: either by Timestamp (legacy) or by metric-specific From/To if requested
    if (filters && filters.useMetricRange && filters.metricRange && filters.metricRange.metric){
      const wantMetric = String(filters.metricRange.metric || '').trim();
      const rowRange = _rowMetricRange_(row, disp[r], IX, wantMetric);
      if (!rowRange) continue; // no usable range on this row
      const qStart = filters.metricRange.start; const qEnd = filters.metricRange.end;
      if (!qStart || !qEnd) continue;
      if (!_rangesOverlap_(rowRange.start, rowRange.end, qStart, qEnd)) continue;
    } else {
      const rawTs = (row[IX.iTs] instanceof Date) ? row[IX.iTs] : (row[IX.iTs] || disp[r][IX.iTs] || '');
      const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
      if (d && !isNaN(d.getTime())){
        if (month!=null && d.getMonth() !== month) continue;
        if (year!=null && d.getFullYear() !== year) continue;
      }
    }
    // Dim filters
    if (wantBen && IX.iBen>=0 && String(row[IX.iBen]||'').trim().toLowerCase() !== wantBen) continue;
    if (wantTeam && IX.iTeam>=0 && String(row[IX.iTeam]||'').trim().toLowerCase() !== wantTeam) continue;
    if (wantProj && IX.iProj>=0 && String(row[IX.iProj]||'').trim().toLowerCase() !== wantProj) continue;

  // Value
  function numAt(i){ if (i<0) return 0; return parseAmount(row[i]); }
    // If metric is 'total' but IX.metrics.total is missing, build total by summing known category cols
    let val = 0;
    if (midx >= 0){
      val = numAt(midx);
    } else if (metric === 'total' || (IX.metrics && IX.metrics.total < 0)){
      // sum categories if explicit total column isn't present
      const catIdxs = [IX.metrics.fuel, IX.metrics.erda, IX.metrics.vehicleRent, IX.metrics.airtime, IX.metrics.transport, IX.metrics.misc];
      for (let ci=0; ci<catIdxs.length; ci++){ val += numAt(catIdxs[ci]); }
    } else {
      val = numAt(IX.metrics.total);
    }
    const key = gkey(row);
    map.set(key, (map.get(key)||0) + val);
  }
  return { ok:true, groups: map };
}

/** Execute a plan and return a human-readable string plus raw data */
function executePlan(plan){
  try{
    if (!plan || !plan.type) return { ok:false, message:'No plan to execute.' };
    switch (plan.type){
      case 'aggregate':
      case 'rank': {
        const metric = plan.metric || 'total';
        const groupBy = plan.groupBy || 'none';
        const res = _aggregateSubmissions_({ metric, groupBy, filters: plan.filters||{} });
        if (!res.ok) return { ok:false, message:'No submissions data available.' };
        const arr = Array.from(res.groups.entries()).map(([k,v])=>({ key:k, value:v }));
        // sort desc by value (rank) but keep single aggregate as one item
        arr.sort((a,b)=> b.value - a.value);
        const top = (plan.type==='rank' && plan.topN) ? arr.slice(0, plan.topN) : arr;
  const tf = plan.filters && plan.filters.timeframe ? plan.filters.timeframe : null;
        const label = tf ? Utilities.formatDate(new Date(tf.year, tf.month, 1),TZ(),'MMMM yyyy') : 'current period';
        const fmt = (n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
        let lines = [];
        if (groupBy==='none'){
          const total = top.length? top[0].value : 0;
          lines.push(`💰 ${metric} — ${label}: ${fmt(total)}`);
        } else {
          lines.push(`📊 ${metric} by ${groupBy} — ${label}`);
          top.forEach((it,i)=>{ lines.push(`${i+1}. ${it.key}: ${fmt(it.value)}`); });
        }
        // If the plan explicitly asked to use metric-range filtering, add a short explanatory note
        if (plan.filters && plan.filters.useMetricRange) {
          lines.push('');
          lines.push('Note: results include rows whose category-specific From/To ranges overlap your requested window.');
        }
        var result = { ok:true, message: lines.join('\n'), rows: top };
        if (plan.output === 'table'){
          // If this was a single-metric + beneficiary request and user asked for a table,
          // prefer returning the per-row metric table (detailed rows) instead of the aggregate total.
          if (plan.metric && plan.filters && plan.filters.metricRange && plan.filters.beneficiary){
            try{
              const ben = plan.filters.beneficiary;
              const metric = plan.metric;
              const mr = plan.filters.metricRange || {};
              const s = (mr.start instanceof Date) ? mr.start : new Date(mr.start);
              const e = (mr.end instanceof Date) ? mr.end : new Date(mr.end);
              const detail = getBeneficiaryMetricForDateRange(ben, metric, s, e, { output: 'table' });
              // If detail returned a table, return that response instead
              if (detail && (detail.table || detail.rows)) return detail;
            }catch(_){ /* fallback to aggregate table below */ }
          }
          // rows is an array of {key,value} for groupBy or metric pairs
          var rowsForTable = top.map(function(it){ return { key: it.key, value: it.value }; });
          var tbl = _buildTableFromRows_(rowsForTable);
          result.table = tbl;
        }
        return result;
      }
      case 'compare': {
        const tf = plan.filters && plan.filters.timeframe ? plan.filters.timeframe : _resolveTimeframe_('');
        const a = plan.metrics && plan.metrics[0] || 'fuel';
        const b = plan.metrics && plan.metrics[1] || 'transport';
        const A = _aggregateSubmissions_({ metric:a, groupBy:'none', filters:{ timeframe: tf } });
        const B = _aggregateSubmissions_({ metric:b, groupBy:'none', filters:{ timeframe: tf } });
        const fmt = (n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
        const label = Utilities.formatDate(new Date(tf.year, tf.month, 1),TZ(),'MMMM yyyy');
        const va = (A.ok && A.groups.size) ? Array.from(A.groups.values())[0] : 0;
        const vb = (B.ok && B.groups.size) ? Array.from(B.groups.values())[0] : 0;
        const lines = [
          `⚖️ Comparison — ${label}`,
          `${a}: ${fmt(va)}`,
          `${b}: ${fmt(vb)}`,
          (va>vb? `➡️ Higher: ${a}` : (vb>va? `➡️ Higher: ${b}` : '➡️ Equal'))
        ];
  var cmpRes = { ok:true, message: lines.join('\n'), rows:[{metric:a,value:va},{metric:b,value:vb}] };
  if (plan.output === 'table') cmpRes.table = _buildTableFromRows_(cmpRes.rows);
  return cmpRes;
      }
      case 'lookup': {
        if (plan.entity==='vehicle'){
          const needAdvanced = (plan.filters && (plan.filters.team || plan.filters.status)) || plan.latestPer || plan.limit || (plan.join && plan.join.vehicleMaster);
          if (needAdvanced){
            let rows = _readCarTP_objects_();
            if (plan.filters && plan.filters.team){
              const team = String(plan.filters.team).toLowerCase();
              rows = rows.filter(r => String(r.Team||'').toLowerCase() === team);
            }
            if (plan.filters && plan.filters.status){
              const status = _normStatus_(plan.filters.status);
              rows = rows.filter(r => _normStatus_(r.Status) === status);
            }
            if ((plan.order||'').toLowerCase() === 'desc') rows.sort((a,b)=> b._ts - a._ts); else rows.sort((a,b)=> a._ts - b._ts);
            if (plan.latestPer){
              const key = plan.latestPer;
              const seen = new Set(); const dedup = [];
              for (const r of rows){ const k = String(r[key]||'').trim(); if (!k || seen.has(k)) continue; seen.add(k); dedup.push(r); }
              rows = dedup;
            }
            if (plan.limit && rows.length > plan.limit) rows = rows.slice(0, plan.limit);
            if (plan.join && plan.join.vehicleMaster){ const idx = _readVehicleMasterIndex_(); rows = rows.map(r => _enrichVehicle_(r, idx)); }
            const out = rows.map(r => ({ team:r.Team, vehicleNumber:r['Vehicle Number'], status:_normStatus_(r.Status), datetime:r['Date and time of entry'], make:r.Make||r._vm_make||'', model:r.Model||r._vm_model||'', owner:r.Owner||r._vm_owner||'', category:r.Category||r._vm_category||'', usageType:r['Usage Type']||r._vm_usageType||'', ratings:r.Ratings, remarks:r['Last Users remarks'] }));
            const lines = ['🚗 Vehicles:']; out.forEach(v => { lines.push(`• ${v.vehicleNumber} — ${v.make} ${v.model} (${v.status})`); });
            var outRes = { ok:true, message: lines.join('\n'), rows: out };
            if (plan.output === 'table') outRes.table = _buildTableFromRows_(out);
            return outRes;
          } else {
            const list = getLatestReleaseCars();
            if (!list || !list.length) return { ok:true, message:'No available vehicles (latest status RELEASE).', rows:[] };
            const top = list.slice(0, Math.min(10, list.length));
            const lines = ['🚗 Available Vehicles (latest = RELEASE):'];
            top.forEach(c=>{ lines.push(`• ${c.carNumber} — ${c.make||''} ${c.model||''}`.trim()); });
            var simpleRes = { ok:true, message: lines.join('\n'), rows: top };
            if (plan.output === 'table') simpleRes.table = _buildTableFromRows_(top);
            return simpleRes;
          }
        }
        if (plan.entity==='vehicle_latest_for_team'){
          var team = plan.filters && plan.filters.team ? String(plan.filters.team) : '';
          var prefer = plan.filters && plan.filters.preferStatus ? String(plan.filters.preferStatus) : '';
          var resVL = getLatestVehicleForTeam(team, prefer);
          if (!resVL || !resVL.ok) return { ok:false, message: resVL && resVL.message ? resVL.message : 'No match' };
          var line = `Newest for ${team}: ${resVL.vehicleNumber}`;
          if (resVL.make || resVL.model) line += ` — ${[resVL.make,resVL.model].filter(Boolean).join(' ')}`;
          if (resVL.status) line += ` (${resVL.status})`;
          if (resVL.dateTime) line += ` — ${Utilities.formatDate(new Date(resVL.dateTime), TZ(), 'dd-MMM-yy HH:mm')}`;
          return { ok:true, message: line, rows:[resVL] };
        }
        if (plan.entity==='vehicle_history'){
          // car-specific history
          const car = plan.filters && plan.filters.carNumber;
          if (car){
            const h = getCarHistory(car);
            if (!h || (!h.users.length && !h.feedback.length && !h.ratings.length)){
              return { ok:true, message:`No recent history found for ${car}.`, rows:[] };
            }
            const lines = [`📜 History for ${car}:`];
            if (h.users.length) lines.push(`Users: ${h.users.join(', ')}`);
            if (h.feedback.length) lines.push(`Feedback: ${h.feedback.join(' | ')}`);
            if (h.ratings.length) lines.push(`Ratings: ${h.ratings.join(', ')}`);
            return { ok:true, message: lines.join('\n'), rows:h };
          }
          // team-based history with pattern/limit
          if (plan.filters && plan.filters.team){
            let rows = _readCarTP_objects_();
            const team = String(plan.filters.team).toLowerCase();
            rows = rows.filter(r => String(r.Team||'').toLowerCase() === team);
            rows.sort((a,b)=> b._ts - a._ts);
            if (Array.isArray(plan.historyPattern) && plan.historyPattern.length){
              const need = plan.historyPattern.map(_normStatus_);
              const out = []; let i=0;
              for (const r of rows){ if (i>=need.length) break; if (_normStatus_(r.Status) === need[i]){ out.push(r); i++; } }
              rows = out;
            }
            if (plan.limit && rows.length > plan.limit) rows = rows.slice(0, plan.limit);
            if (plan.join && plan.join.vehicleMaster){ const idx = _readVehicleMasterIndex_(); rows = rows.map(r => _enrichVehicle_(r, idx)); }
            const out = rows.map(r => ({ team:r.Team, vehicleNumber:r['Vehicle Number'], status:_normStatus_(r.Status), datetime:r['Date and time of entry'], make:r.Make||r._vm_make||'', model:r.Model||r._vm_model||'', owner:r.Owner||r._vm_owner||'', category:r.Category||r._vm_category||'', usageType:r['Usage Type']||r._vm_usageType||'', ratings:r.Ratings, remarks:r['Last Users remarks'] }));
            const lines = [`📜 Last ${out.length} for team ${plan.filters.team}:`]; out.forEach(v => lines.push(`• ${v.vehicleNumber} — ${v.status}`));
            return { ok:true, message: lines.join('\n'), rows: out };
          }
          return { ok:false, message:'Please provide a car number or team for history.' };
        }
        return { ok:false, message:'Unknown lookup entity.' };
      }
      default:
        return { ok:false, message:'Unknown plan type.' };
    }
  }catch(e){
    console.error('executePlan error:', e);
    return { ok:false, message:'Plan execution failed.' };
  }
}

/** Public NL entrypoint: build plan and execute */
function askQuestion(message){
  try{
    const plan = planAnalysis(message);
    if (!plan) return { ok:false, plan:null, data:[], note:'Could not build a plan for this question.' };
    const data = executePlan(plan);
    // If metric-range filtering was requested, append a short note for clarity across plan types
    if (plan && plan.filters && plan.filters.useMetricRange && data && data.message){
      data.message = data.message + '\n\nNote: results include rows whose category-specific From/To ranges overlap your requested window.';
    }
    return { ok:true, plan, data, note:'' };
  }catch(e){
    return { ok:false, plan:null, data:[], note:String(e && e.message || e) };
  }
}

/** Read CarT_P as array of objects (tolerant headers) + parsed timestamp */
function _readCarTP_objects_(){
  const sh = _openCarTP_(); if (!sh) return [];
  const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn(); if (lastRow<=1) return [];
  const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  const IX = _headerIndex_(head);
  function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
  const iRef  = idx(['Ref','Reference Number','Ref Number'], false);
  const iDate = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
  const iProj = idx(['Project'], false);
  const iTeam = idx(['Team'], false);
  let iCarNo = -1; try{ iCarNo = idx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car'], false);}catch(_){ iCarNo=-1; }
  if (iCarNo<0) iCarNo = _findCarNumberColumn_(head);
  const iMake  = idx(['Make','Car Make','Brand'], false);
  const iModel = idx(['Model','Car Model'], false);
  const iCat   = idx(['Category','Vehicle Category','Cat'], false);
  const iUse   = idx(['Usage Type','Usage','Use Type'], false);
  const iOwner = idx(['Owner','Owner Name','Owner Info'], false);

  const normalizedHead = head.map(function(h){ return String(h || '').trim().toLowerCase(); });
  const sanitizedHead = normalizedHead.map(function(h){ return h.replace(/[^a-z0-9]+/g, ''); });
  function findHeaderIndex(predicate) {
    for (let i = 0; i < normalizedHead.length; i++) {
      if (predicate(normalizedHead[i], sanitizedHead[i], i)) {
        return i;
      }
    }
    return -1;
  }

  let iBeneficiary = -1;
  const beneficiaryAliases = ['R.Beneficiary','R Beneficiary','R_Beneficiary','Team Members','Members','Member Names','Beneficiaries'];
  for (let b = 0; b < beneficiaryAliases.length && iBeneficiary < 0; b++) {
    try { iBeneficiary = IX.get(beneficiaryAliases[b]); } catch (_err) { iBeneficiary = -1; }
  }
  if (iBeneficiary < 0) {
    iBeneficiary = findHeaderIndex(function(norm, san){
      if (!norm) return false;
      if (norm.indexOf('responsible') !== -1) return false;
      if (san === 'rben') return false;
      if (norm.indexOf('beneficiary') !== -1) return true;
      return /\bmember\b/.test(norm);
    });
  }

  let iRespShort = -1;
  const respShortAliases = ['R. Ben','R Ben','RBen'];
  for (let s = 0; s < respShortAliases.length && iRespShort < 0; s++) {
    try { iRespShort = IX.get(respShortAliases[s]); } catch (_err) { iRespShort = -1; }
  }
  if (iRespShort < 0) {
    iRespShort = findHeaderIndex(function(_norm, san){
      return san === 'rben';
    });
  }

  let iRespFull = -1;
  const respFullAliases = ['Responsible Beneficiary','ResponsibleBeneficiary','Name of Responsible beneficiary','Responsible beneficiary','Name of responsible beneficiary'];
  for (let f = 0; f < respFullAliases.length && iRespFull < 0; f++) {
    try { iRespFull = IX.get(respFullAliases[f]); } catch (_err) { iRespFull = -1; }
  }
  if (iRespFull < 0) {
    iRespFull = findHeaderIndex(function(norm, san){
      if (!norm) return false;
      return norm.indexOf('responsible') !== -1 && norm.indexOf('benefici') !== -1;
    });
  }

  const iStat  = idx(['In Use/Release','In Use / release','In Use','Status'], false);
  const iRem   = idx(['Last Users remarks','Remarks','Feedback'], false);
  const iRate  = idx(['Ratings','Stars','Rating'], false);
  const iSubmit= idx(['Submitter username','Submitter','User'], false);
  const iRespTime = idx(['R.Ben Time','R.Ben timestamp','Responsible Beneficiary Time'], false);

  const rng = sh.getRange(2,1,lastRow-1,lastCol);
  const data = rng.getValues();
  const disp = rng.getDisplayValues();
  const out = [];
  for (let r=0;r<data.length;r++){
    const row = data[r];
    const beneficiaryValue = iBeneficiary>=0 ? (row[iBeneficiary] || disp[r][iBeneficiary] || '') : '';
    const respShortValue = iRespShort>=0 ? (row[iRespShort] || disp[r][iRespShort] || '') : '';
    const respFullValue = iRespFull>=0 ? (row[iRespFull] || disp[r][iRespFull] || '') : '';

    let responsibleCandidate = respShortValue || respFullValue || '';
    let responsibleValue = _sanitizeResponsibleName(responsibleCandidate);
    if (!responsibleValue && responsibleCandidate) {
      const pieces = _splitBeneficiaryNames_(responsibleCandidate)
        .map(_sanitizeResponsibleName)
        .filter(Boolean);
      if (pieces.length) responsibleValue = pieces[0];
    }
    if (!responsibleValue && respShortValue) {
      const shortSanitized = _sanitizeResponsibleName(respShortValue);
      if (shortSanitized) responsibleValue = shortSanitized;
      else if (!responsibleCandidate) responsibleValue = respShortValue;
    }
    if (!responsibleValue && respFullValue) {
      const fullSanitized = _sanitizeResponsibleName(respFullValue);
      if (fullSanitized) responsibleValue = fullSanitized;
      else if (!responsibleCandidate) responsibleValue = respFullValue;
    }
    if (!responsibleValue && beneficiaryValue) {
      const memberPieces = _splitBeneficiaryNames_(beneficiaryValue)
        .map(_sanitizeResponsibleName)
        .filter(Boolean);
      if (memberPieces.length) {
        responsibleValue = memberPieces[0];
      }
    }

    const obj = {
      Ref: iRef>=0 ? (row[iRef] || disp[r][iRef] || '') : '',
      'Date and time of entry': iDate>=0 ? (row[iDate] || disp[r][iDate] || '') : '',
      Project: iProj>=0 ? (row[iProj] || disp[r][iProj] || '') : '',
      Team: iTeam>=0 ? (row[iTeam] || disp[r][iTeam] || '') : '',
      'Vehicle Number': iCarNo>=0 ? (row[iCarNo] || disp[r][iCarNo] || '') : '',
      Make: iMake>=0 ? (row[iMake] || disp[r][iMake] || '') : '',
      Model: iModel>=0 ? (row[iModel] || disp[r][iModel] || '') : '',
      Category: iCat>=0 ? (row[iCat] || disp[r][iCat] || '') : '',
      'Usage Type': iUse>=0 ? (row[iUse] || disp[r][iUse] || '') : '',
      Owner: iOwner>=0 ? (row[iOwner] || disp[r][iOwner] || '') : '',
      'R.Beneficiary': beneficiaryValue,
      'R. Ben': responsibleValue || '',
      Status: iStat>=0 ? (row[iStat] || disp[r][iStat] || '') : '',
      'Last Users remarks': iRem>=0 ? (row[iRem] || disp[r][iRem] || '') : '',
      Ratings: iRate>=0 ? (row[iRate] || disp[r][iRate] || '') : '',
      'Submitter username': iSubmit>=0 ? (row[iSubmit] || disp[r][iSubmit] || '') : '',
      'R.Ben Time': iRespTime>=0 ? (row[iRespTime] || disp[r][iRespTime] || '') : ''
    };
    obj._rowIndex = r + 2;
    if (respFullValue) {
      obj['Responsible Beneficiary'] = respFullValue;
      obj['Name of Responsible beneficiary'] = respFullValue;
    } else if (responsibleValue) {
      obj['Responsible Beneficiary'] = responsibleValue;
      obj['Name of Responsible beneficiary'] = responsibleValue;
    }
    if (responsibleValue) {
      obj.responsibleBeneficiary = responsibleValue;
      obj['R. Ben'] = responsibleValue;
      obj.rBenShort = responsibleValue;
    } else if (!obj.responsibleBeneficiary && obj['Responsible Beneficiary']) {
      const sanitized = _sanitizeResponsibleName(obj['Responsible Beneficiary']);
      if (sanitized) {
        obj.responsibleBeneficiary = sanitized;
        obj['R. Ben'] = sanitized;
        obj.rBenShort = sanitized;
      }
    }
    const ts = _parseTs_(obj['Date and time of entry']);
    obj._ts = ts;
    out.push(obj);
  }
  return out;
}

/** Best-effort vehicle master index from Vehicle sheet */
function _readVehicleMasterIndex_(){
  const sh = _openVehicleSheet_(); if (!sh) return {};
  const vals = sh.getDataRange().getDisplayValues(); if (!vals || vals.length<2) return {};
  const head = vals[0];
  const IX = _headerIndex_(head);
  function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
  let iKey = -1; try{ iKey = idx(['Vehicle Number','Car Number','Vehicle','Number','Reg','Registration'], false);}catch(_){ iKey=-1; }
  if (iKey<0){
    // fallback simple contains scan
    for (let i=0;i<head.length;i++){ const h=String(head[i]||'').toLowerCase(); if (h.includes('vehicle') && (h.includes('number')||h.includes('no')||h.includes('reg'))) { iKey=i; break; } }
  }
  if (iKey<0) return {};
  const iMake  = idx(['Make','Brand','Maker'], false);
  const iModel = idx(['Model','Variant'], false);
  const iOwner = idx(['Owner','Ownership'], false);
  const iCat   = idx(['Category','Type','Class'], false);
  const iUse   = idx(['Usage Type','Usage','Contract Type'], false);
  const idxMap = {};
  for (let r=1;r<vals.length;r++){
    const row = vals[r]; const key = String(row[iKey]||'').trim(); if (!key) continue;
    idxMap[key] = {
      make:  iMake>=0 ? row[iMake] : '',
      model: iModel>=0 ? row[iModel] : '',
      owner: iOwner>=0 ? row[iOwner] : '',
      category: iCat>=0 ? row[iCat] : '',
      usageType: iUse>=0 ? row[iUse] : ''
    };
  }
  return idxMap;
}

function _enrichVehicle_(r, master){
  const key = String(r['Vehicle Number']||'').trim(); const m = master[key]; if (!m) return r;
  r._vm_make = m.make; r._vm_model = m.model; r._vm_owner = m.owner; r._vm_category = m.category; r._vm_usageType = m.usageType; return r;
}

function _parseTs_(v){
  if (!v) return 0;
  try{
    if (v instanceof Date) return v.getTime();
    const d = new Date(v); if (!isNaN(d.getTime())) return d.getTime();
  }catch(_){ }
  return 0;
}
function _parseDateTimeFlexible_(value){
  if (!value && value !== 0) return 0;
  if (value instanceof Date) {
    const ts = value.getTime();
    return isNaN(ts) ? 0 : ts;
  }
  if (typeof value === 'number' && !isNaN(value)) {
    if (value > 1e12) {
      return value;
    }
    if (value > 0) {
      const serialMs = Math.round((value - 25569) * 86400000);
      if (serialMs > 0) {
        const serialDate = new Date(serialMs);
        if (!isNaN(serialDate.getTime())) {
          return serialDate.getTime();
        }
      }
    }
  }

  const text = String(value || '').trim();
  if (!text) return 0;

  const normalized = text.replace(/T/, ' ').replace(/-(?=\d)/g, '/');
  const direct = Date.parse(normalized);
  if (!isNaN(direct)) return direct;

  const dmy = normalized.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})(?:\s+(\d{1,2})(?::(\d{2})(?::(\d{2}))?)?)?$/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    const year = Number(dmy[3].length === 2 ? ('20' + dmy[3]) : dmy[3]);
    const hour = Number(dmy[4] || '0');
    const minute = Number(dmy[5] || '0');
    const second = Number(dmy[6] || '0');
    const date = new Date(year, month - 1, day, hour, minute, second);
    if (!isNaN(date.getTime())) {
      return date.getTime();
    }
  }

  const monthMap = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };

  const alpha = text.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{2,4})(?:\s+(\d{1,2})(?::(\d{2})(?::(\d{2}))?)?)?$/);
  if (alpha) {
    const day = Number(alpha[1]);
    const monthName = alpha[2].toLowerCase();
    const month = monthMap.hasOwnProperty(monthName) ? monthMap[monthName] : null;
    if (month !== null) {
      const year = Number(alpha[3].length === 2 ? ('20' + alpha[3]) : alpha[3]);
      const hour = Number(alpha[4] || '0');
      const minute = Number(alpha[5] || '0');
      const second = Number(alpha[6] || '0');
      const date = new Date(year, month, day, hour, minute, second);
      if (!isNaN(date.getTime())) {
        return date.getTime();
      }
    }
  }

  return 0;
}

function _normStatus_(s){
  const raw = String(s||'').trim().toUpperCase();
  if (!raw) return '';
  const compact = raw.replace(/\s+/g, ' ');
  if (compact === 'IN USE' || raw === 'INUSE' || raw === 'IN-USE' || raw === 'IN_USE') return 'IN USE';
  if (compact === 'RELEASE' || compact === 'RELEASED') return 'RELEASE';
  if (/IN\s*USE/.test(raw)) return 'IN USE';
  if (/RELEAS/.test(raw)) return 'RELEASE';
  return compact;
}

function _vehicleKey_(value){
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) return '';
  const stripped = raw.replace(/[^A-Z0-9]/g, '');
  return stripped || raw;
}

function _beneficiaryKey_(value){
  if (!value && value !== 0) return '';
  return String(value)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function _teamBeneficiaryKey_(team, beneficiary){
  const teamKey = _normTeamKey(team);
  const benKey = _beneficiaryKey_(beneficiary);
  if (!teamKey && !benKey) return '';
  return teamKey + '|' + benKey;
}

function _splitBeneficiaryNames_(value) {
  if (!value && value !== 0) return [];
  if (Array.isArray(value)) {
    return Array.from(new Set(value.map(function(v){ return String(v || '').trim(); }).filter(Boolean)));
  }
  let str = String(value || '')
    .replace(/[&]/g, ',')
    .replace(/\band\b/gi, ',')
    .replace(/[()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!str) return [];
  return Array.from(new Set(
    str.split(/[,;/\n]+/).map(function(part){ return part.trim(); }).filter(Boolean)
  ));
}

function _beneficiaryNamesFromRow_(row) {
  const sources = [
    row['R.Beneficiary'],
    row.responsibleBeneficiary,
    row['Responsible Beneficiary'],
    row['Name of Responsible beneficiary']
  ];
  const names = [];
  sources.forEach(function(source){
    _splitBeneficiaryNames_(source).forEach(function(name){
      if (!names.includes(name)) names.push(name);
    });
  });
  if (!names.length && row['Submitter username']) {
    const fallback = String(row['Submitter username']).trim();
    if (fallback) names.push(fallback);
  }
  return names;
}


/**
 * Parse month/year phrases like 'sept 2025', 'sep 2025', 'september 2025',
 * as well as 'this month', 'last month', 'next month'. Returns {month:0-11, year} or null.
 */
function parseMonthYear(text){
  const q = String(text||'').toLowerCase();
  const now = new Date();
  const monthMap = {
    jan:0,january:0,
    feb:1,february:1,
    mar:2,march:2,
    apr:3,april:3,
    may:4,
    jun:5,june:5,
    jul:6,july:6,
    aug:7,august:7,
    sep:8,sept:8,september:8,
    oct:9,october:9,
    nov:10,november:10,
    dec:11,december:11
  };

  // relative words
  if (/\bthis\s+month\b/.test(q)) return { month: now.getMonth(), year: now.getFullYear() };
  if (/\blast\s+month\b/.test(q)) {
    const d = new Date(now.getFullYear(), now.getMonth()-1, 1);
    return { month: d.getMonth(), year: d.getFullYear() };
  }
  if (/\bnext\s+month\b/.test(q)) {
    const d = new Date(now.getFullYear(), now.getMonth()+1, 1);
    return { month: d.getMonth(), year: d.getFullYear() };
  }

  // explicit month + year (month name first)
  const m1 = q.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b\s+(\d{4})/);
  if (m1) {
    const m = monthMap[m1[1].slice(0,3)];
    const y = parseInt(m1[2],10);
    if (!isNaN(m) && !isNaN(y)) return { month: m, year: y };
  }

  // year + month name
  const m2 = q.match(/\b(\d{4})\b\s+(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)/);
  if (m2) {
    const y = parseInt(m2[1],10);
    const m = monthMap[m2[2].slice(0,3)];
    if (!isNaN(m) && !isNaN(y)) return { month: m, year: y };
  }

  return null;
}

/**
 * Loose month/year parser to handle formats like:
 *  - Sep, 2025 or September, 2025
 *  - Sep-2025, September-2025
 *  - 09/2025, 9/2025, 2025/09, 2025-09
 */
function parseMonthYearLoose(text){
  var r = parseMonthYear(text);
  if (r) return r;
  try{
    var s = String(text||'');
    var now = new Date();
    // Allow comma or hyphen between month name and year
    var mA = s.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b\s*[,-]?\s*(\d{4})/i);
    if (mA){
      var mm = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}[mA[1].slice(0,3).toLowerCase()];
      var yy = parseInt(mA[2],10);
      if (!isNaN(mm) && !isNaN(yy)) return { month:mm, year:yy };
    }
    // Numeric MM/YYYY or M/YYYY
    var mB = s.match(/\b(\d{1,2})\s*[\/\-]\s*(\d{4})\b/);
    if (mB){
      var mnum = parseInt(mB[1],10);
      var y = parseInt(mB[2],10);
      if (mnum>=1 && mnum<=12) return { month:mnum-1, year:y };
    }
    // Numeric YYYY/MM or YYYY-M
    var mC = s.match(/\b(\d{4})\s*[\/\-]\s*(\d{1,2})\b/);
    if (mC){
      var y2 = parseInt(mC[1],10);
      var mnum2 = parseInt(mC[2],10);
      if (mnum2>=1 && mnum2<=12) return { month:mnum2-1, year:y2 };
    }
    // Month name only (no explicit year) → assume current year
    var mOnly = s.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t|tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b/i);
    if (mOnly){
      var mo = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}[mOnly[1].slice(0,3).toLowerCase()];
      if (typeof mo === 'number') return { month: mo, year: now.getFullYear() };
    }
  }catch(_){ }
  return null;
}

/**
* Parse a date token in multiple formats into a Date (TZ Africa/Dar_es_Salaam):
 *  - DD-MMM-YY or DD-MMM-YYYY (e.g., 02-Sep-25, 02-Sep-2025)
 *  - DD/MM/YYYY or DD/MM/YY
 *  - YYYY-MM-DD
 */
function parseDateTokenFlexible(token){
  if (!token) return null;
  const s = String(token).trim();
  if (!s) return null;

  function createTZDate(y,m,d) {
    // Create date in TZ timezone to avoid UTC conversion issues
    try {
      const dateStr = `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}T12:00:00`;
      const tz = TZ();
      // Create date at noon to avoid any timezone boundary issues
      const dt = new Date(dateStr);
      // Adjust for timezone
      dt.setMinutes(dt.getMinutes() + dt.getTimezoneOffset());
      return dt;
    } catch(e) {
      console.error('createTZDate error:', e);
      return new Date(y,m,d);
    }
  }

  // YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m){
    const y = parseInt(m[1],10), mo = parseInt(m[2],10)-1, d = parseInt(m[3],10);
    const dt = createTZDate(y,mo,d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  // DD/MM/YYYY or DD/MM/YY
  m = s.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})$/);
  if (m){
    let y = parseInt(m[3],10); if (y < 100) y = 2000 + y;
    const mo = parseInt(m[2],10)-1, d = parseInt(m[1],10);
    const dt = createTZDate(y,mo,d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  // DD-MMM-YY or DD-MMM-YYYY
  m = s.match(/^(\d{1,2})[-\s](jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-\s](\d{2,4})$/i);
  if (m){
    const d = parseInt(m[1],10);
    const monKey = m[2].slice(0,3).toLowerCase();
    const map = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const mo = map[monKey];
    let y = parseInt(m[3],10); if (y < 100) y = 2000 + y;
    if (typeof mo === 'number'){
      const dt = createTZDate(y,mo,d);
      return isNaN(dt.getTime()) ? null : dt;
    }
  }
  return null;
}

/**
 * Parse a date range from free text, recognizing separators 'to' or '-'
 * Returns { start: Date, end: Date } or null
 */
function parseDateRangeFromText(text){
  const s = String(text||'');
  // Common separators
  const parts = s.split(/\bto\b|\s-\s|\s–\s|\s—\s|\s+to\s+/i);
  if (parts.length >= 2){
    // Find two date-like tokens near the split
    // Try the last two tokens after trimming
    const left = parts[0].split(/\s+/).slice(-1)[0] || parts[0];
    const right = parts[1].split(/\s+/)[0] || parts[1];
    let d1 = parseDateTokenFlexible(left.replace(/[,]+$/,''));
    let d2 = parseDateTokenFlexible(right.replace(/^[,]+/,''));
    // Fallback: scan for any date tokens in left/right chunks
    if (!d1){
      const mL = parts[0].match(/(\d{1,2}[-\/][A-Za-z]{3,}|\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{4}-\d{2}-\d{2})/);
      if (mL) d1 = parseDateTokenFlexible(mL[1]);
    }
    if (!d2){
      const mR = parts[1].match(/(\d{1,2}[-\/][A-Za-z]{3,}|\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}|\d{4}-\d{2}-\d{2})/);
      if (mR) d2 = parseDateTokenFlexible(mR[1]);
    }
    if (d1 && d2){
      // Normalize ordering
      if (d2 < d1){ const tmp = d1; d1 = d2; d2 = tmp; }
      return { start: d1, end: d2 };
    }
  }
  return null;
}

/** Return a Set of known beneficiary names (from DD + submissions). */
function getKnownBeneficiariesSet(){
  const set = new Set();
  try{
    const dd = _readDD_compact_();
    dd.forEach(r => { if (r && r.beneficiary) set.add(String(r.beneficiary).trim()); });
  }catch(_){ }
  try{
    const sh = getSheet('submissions');
    if (sh) {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow >= 1 && lastCol >= 1){
        const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
        const iBen = head.indexOf('Beneficiary');
        if (iBen >= 0 && lastRow > 1){
          const vals = sh.getRange(2, iBen+1, lastRow-1, 1).getValues();
          vals.forEach(v => { const n = String(v[0]||'').trim(); if (n) set.add(n); });
        }
      }
    }
  }catch(_){ }
  return set;
}

/**
 * Aggregate expenses for a beneficiary within a given month (0-11) and year.
 * Header-tolerant for either the detailed or compact submissions schema.
 */
function getBeneficiaryPeriodExpenses(beneficiaryName, month, year){
  try{
    const name = String(beneficiaryName||'').trim();
    if (!name) return { ok:false, message:'Please provide a beneficiary name.' };
    const sh = getSheet('submissions');
    if (!sh) return { ok:false, message:'Submissions data not available.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) return { ok:true, message:`No submissions found for ${name} in ${Utilities.formatDate(new Date(year,month,1),TZ(),'MMMM yyyy')}.` };

    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h=>String(h||''));
    const idx = (labels)=>{
      for (let i=0;i<labels.length;i++){
        const lab = labels[i];
        const j = head.indexOf(lab);
        if (j !== -1) return j;
      }
      return -1;
    };

    const iBen = idx(['Beneficiary']);
    const iTs  = idx(['Timestamp','Date','Date and time of entry','Date and time']);
    if (iBen < 0 || iTs < 0) return { ok:false, message:'Required columns not found in submissions sheet.' };

    // Category columns (detailed vs compact)
    const iFuel = idx(['Fuel Amt','Fuel']);
    const iERDA = idx(['ER DA Amt','ER DA']);
    const iCar  = idx(['Car Amt','Vehicle Rent']);
    const iAir  = idx(['Airtime Amt','Airtime']);
    const iTrans= idx(['Transport Amt','Transport']);
    const iMisc = idx(['Misc Amt','Misc','Miscellaneous']);
    const iTotal= idx(['Row Total','Total Amount','Total']);
    // Fallback to explicit column-letter mapping if header detection fails
    const fixedCols = { total:7, fuel:11, erda:14, vehicleRent:18, airtime:21, transport:24, misc:27 };
    try{
      const headLen = Array.isArray(head) ? head.length : 0;
      if (headLen >= 27){
        if (iTotal < 0) iTotal = fixedCols.total - 1;
        if (iFuel  < 0) iFuel  = fixedCols.fuel - 1;
        if (iERDA  < 0) iERDA  = fixedCols.erda - 1;
        if (iCar   < 0) iCar   = fixedCols.vehicleRent - 1;
        if (iAir   < 0) iAir   = fixedCols.airtime - 1;
        if (iTrans < 0) iTrans = fixedCols.transport - 1;
        if (iMisc  < 0) iMisc  = fixedCols.misc - 1;
      }
    }catch(_e){}

    const rng = sh.getRange(2,1,lastRow-1,lastCol);
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();

    const nmLC = name.toLowerCase();
    let sum = { fuel:0, erda:0, car:0, air:0, transport:0, misc:0, total:0, count:0 };

    for (let r=0;r<vals.length;r++){
      const row = vals[r];
      const ben = String(row[iBen]||'').trim().toLowerCase();
      if (ben !== nmLC) continue;
      const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
      const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
      if (isNaN(d.getTime())) continue;
      if (d.getMonth() !== month || d.getFullYear() !== year) continue;

      function numAt(i){ if (i < 0) return 0; return parseAmount(row[i]); }

      sum.fuel      += numAt(iFuel);
      sum.erda      += numAt(iERDA);
      sum.car       += numAt(iCar);
      sum.air       += numAt(iAir);
      sum.transport += numAt(iTrans);
      sum.misc      += numAt(iMisc);
      // If explicit total column missing, derive total by summing category columns
      let rowTot = 0;
      if (iTotal >= 0) {
        rowTot = numAt(iTotal);
      } else {
        rowTot = numAt(iFuel) + numAt(iERDA) + numAt(iCar) + numAt(iAir) + numAt(iTrans) + numAt(iMisc);
      }
      sum.total     += rowTot;
      sum.count++;
    }

    const monthLabel = Utilities.formatDate(new Date(year,month,1),TZ(),'MMMM yyyy');
    const formatCurrency = (amount)=>{
      const num = parseFloat(amount)||0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    const lines = [
      `👤 ${name} — ${monthLabel}`,
      `💰 Total: ${formatCurrency(sum.total)} (${sum.count} submissions)`,
      '',
      `Breakdown:`,
  `⛽ Fuel: ${formatCurrency(sum.fuel)}`,
  `🏠 DA: ${formatCurrency(sum.erda)}`,
      `🚗 Vehicle Rent: ${formatCurrency(sum.car)}`,
      `📱 Airtime: ${formatCurrency(sum.air)}`,
      `🚌 Transport: ${formatCurrency(sum.transport)}`,
      `📋 Misc: ${formatCurrency(sum.misc)}`
    ];
    return { ok:true, message: lines.join('\n') };
  }catch(e){
    console.error('getBeneficiaryPeriodExpenses error:', e);
    return { ok:false, message: 'Error retrieving beneficiary expenses.' };
  }
}

/**
 * Aggregate a single metric for a beneficiary within an explicit date range [start..end].
 * metric keys: 'fuel','erda','vehicleRent','airtime','transport','misc','total'
 */
function getBeneficiaryMetricForDateRange(beneficiaryName, metric, startDate, endDate, opts){
  try{
    const name = String(beneficiaryName||'').trim();
    if (!name) return 'Please provide a beneficiary name.';
    const sh = getSheet('submissions');
    if (!sh) return 'Submissions data not available.';
    const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) return 'No submissions found.';

    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    // Use tolerant header index resolution used elsewhere in the codebase
    const IX = _submissionsHeaderIndices_(head);
    IX._head = head;
    const iBen = (typeof IX.iBen !== 'undefined' && IX.iBen >= 0) ? IX.iBen : -1;
    const iTs  = (typeof IX.iTs !== 'undefined' && IX.iTs >= 0) ? IX.iTs : -1;
    if (iBen < 0 || iTs < 0) return 'Required columns not found in submissions sheet.';
    // Prefer the header-tolerant metric indices discovered by _submissionsHeaderIndices_
    const mapMetricCol = (IX && IX.metrics) ? IX.metrics : { total:-1, fuel:-1, erda:-1, vehicleRent:-1, airtime:-1, transport:-1, misc:-1 };
    // Fallback to fixed column mapping (letters provided by user) if necessary
    try{
      const fixed = { total:6, fuel:10, erda:13, vehicleRent:17, airtime:20, transport:23, misc:26 };
      if (lastCol >= 27){ Object.keys(fixed).forEach(k => { if (!mapMetricCol[k] || mapMetricCol[k] < 0) mapMetricCol[k] = fixed[k]; }); }
    }catch(_e){}
    const col = (mapMetricCol && typeof mapMetricCol[metric] === 'number' && mapMetricCol[metric] >= 0) ? mapMetricCol[metric] : (mapMetricCol && typeof mapMetricCol.total === 'number' ? mapMetricCol.total : -1);

    const rng = sh.getRange(2,1,lastRow-1,lastCol);
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();
    const nmLC = name.toLowerCase();
    // Compare dates in Tanzania timezone to avoid off-by-one
    function tzDateKey(date) {
      const tz = TZ();
      const formatted = Utilities.formatDate(new Date(date), tz, 'yyyyMMdd');
      return Number(formatted);
    }
    const startKey = tzDateKey(startDate);
    const endKey = tzDateKey(endDate);
    let sum = 0;
    // Strict mapping for specific metrics as confirmed by user (1-based letters):
    // Vehicle Rent: O,P | Airtime: S,T | Transport: V,W | Misc: Y,Z
    const STRICT_METRICS = ['vehicleRent','airtime','transport','misc'];
    function colLetterToIndex(letter){
      if (!letter) return -1; const s = String(letter||'').toUpperCase().trim();
      let idx = 0; for (let i=0;i<s.length;i++){ const c = s.charCodeAt(i)-64; if (c>0) idx = idx*26 + c; }
      return idx>0 ? idx-1 : -1;
    }
    const STRICT_RANGE_COLS = {
      vehicleRent: { from: colLetterToIndex('O'), to: colLetterToIndex('P') },
      airtime:     { from: colLetterToIndex('S'), to: colLetterToIndex('T') },
      transport:   { from: colLetterToIndex('V'), to: colLetterToIndex('W') },
      misc:        { from: colLetterToIndex('Y'), to: colLetterToIndex('Z') }
    };
    const isStrict = STRICT_METRICS.indexOf(String(metric||'').trim()) !== -1;
    // Heuristic header-based indices (used only for non-strict metrics like fuel/erda)
    const metricRanges = isStrict ? null : _submissionsMetricRangeIndices_(head);
    const sIdx = isStrict ? -1 : ((metricRanges && metricRanges.rangeStarts && typeof metricRanges.rangeStarts[metric] === 'number') ? metricRanges.rangeStarts[metric] : -1);
    const eIdx = isStrict ? -1 : ((metricRanges && metricRanges.rangeEnds && typeof metricRanges.rangeEnds[metric] === 'number') ? metricRanges.rangeEnds[metric] : -1);
    // helper: robust cell→Date for strict parsing
    function parseCellToDate(v){
      if (v instanceof Date) return v;
      if (v == null || v === '') return null;
      const d = parseDateTokenFlexible(String(v));
      if (d) return d;
      const d2 = new Date(v);
      return isNaN(d2.getTime()) ? null : d2;
    }

    for (let r=0;r<vals.length;r++){
      const row = vals[r];
      const ben = String(row[iBen]||'').trim().toLowerCase();
      if (ben !== nmLC) continue;
      if (isStrict){
        const map = STRICT_RANGE_COLS[metric];
        const fIdx = map ? map.from : -1;
        const tIdx = map ? map.to : -1;
        if (fIdx < 0 || tIdx < 0) continue; // cannot evaluate
        const fRaw = (typeof row[fIdx] !== 'undefined' && row[fIdx] !== '') ? row[fIdx] : (disp[r] && disp[r][fIdx]);
        const tRaw = (typeof row[tIdx] !== 'undefined' && row[tIdx] !== '') ? row[tIdx] : (disp[r] && disp[r][tIdx]);
        const rs = parseCellToDate(fRaw);
        const re = parseCellToDate(tRaw);
        if (!rs || !re) continue; // strict: exclude if missing/unparsable
        if (!_rangesOverlap_(rs, re, startDate, endDate)) continue;
      } else {
        // Non-strict: header-derived ranges if available, else Timestamp
        if (sIdx >=0 || eIdx >=0){
          const rowStart = (sIdx>=0 && row[sIdx] instanceof Date) ? row[sIdx] : (sIdx>=0 ? parseDateTokenFlexible(String(row[sIdx]||disp[r][sIdx]||'')) : null);
          const rowEnd = (eIdx>=0 && row[eIdx] instanceof Date) ? row[eIdx] : (eIdx>=0 ? parseDateTokenFlexible(String(row[eIdx]||disp[r][eIdx]||'')) : null);
          if (!rowStart && !rowEnd){
            const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
            const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
            if (isNaN(d.getTime())) continue;
            const dtKey = tzDateKey(d);
            if (dtKey < startKey || dtKey > endKey) continue;
          } else {
            const rs = rowStart || rowEnd; const re = rowEnd || rowStart;
            if (!_rangesOverlap_(rs, re, startDate, endDate)) continue;
          }
        } else {
          const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
          const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
          if (isNaN(d.getTime())) continue;
          const dtKey = tzDateKey(d);
          if (dtKey < startKey || dtKey > endKey) continue;
        }
      }
      const v = row[col];
      sum += parseAmount(v);
    }
  const fmtShortYMD = (date) => Utilities.formatDate(date, TZ(), 'dd-MMM-yy');
    const fmtMoney = (n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
    const label = `${fmtShortYMD(startDate)} to ${fmtShortYMD(endDate)}`;
  const metricLabel = (metric==='vehicleRent'?'Vehicle Rent': (metric==='erda'?'DA': metric.charAt(0).toUpperCase()+metric.slice(1)));
    const text = `👤 ${name} — ${label}\n${metricLabel}: ${fmtMoney(sum)}`;
    try{
      const wantsTable = opts && (opts.output === 'table' || opts.wantsTable);
      if (wantsTable){
        const rowsOut = [];
        // Try to find a sensible amount header for this metric using tolerant header index
        const HIX = _headerIndex_(head.map(h=>String(h||'')));
        function findHeaderIndex(labels){ try{ return HIX.get(labels); }catch(e){ return -1; } }
        const amtCandidates = (function(m){
          if (m==='fuel') return ['Fuel Amt','Fuel Amount','Fuel'];
          if (m==='erda') return ['DA Amount','DA Amt','DA','ER DA Amt','ER DA Amount','ER DA','ERDA','er da amt','erda amt'];
          // Avoid plain 'Vehicle Rent' to prevent matching 'Vehicle Rent From/To'
          if (m==='vehicleRent') return ['Car Amt','Car Amount','Vehicle Rent Amount','Vehicle Rent Amt'];
          if (m==='airtime') return ['Airtime Amount','Airtime Amt','Airtime'];
          if (m==='transport') return ['Transport Amount','Transport Amt','Transport'];
          return ['Misc Amt','Misc Amount','Misc','Miscellaneous'];
        })(metric);
        let amtIdx = findHeaderIndex(amtCandidates);
        // fallback: if header-search fails, try the metric column discovered earlier (mapMetricCol)
        if (amtIdx < 0 && mapMetricCol && typeof mapMetricCol[metric] === 'number' && mapMetricCol[metric] >= 0) amtIdx = mapMetricCol[metric];
        for (let r=0;r<vals.length;r++){
          const row = vals[r];
          const ben = String(row[iBen]||'').trim().toLowerCase();
          if (ben !== name.toLowerCase()) continue;
          let rs = null, re = null;
          if (isStrict){
            const map = STRICT_RANGE_COLS[metric];
            const fIdx = map ? map.from : -1;
            const tIdx = map ? map.to : -1;
            if (fIdx < 0 || tIdx < 0) continue;
            const fRaw = (typeof row[fIdx] !== 'undefined' && row[fIdx] !== '') ? row[fIdx] : (disp[r] && disp[r][fIdx]);
            const tRaw = (typeof row[tIdx] !== 'undefined' && row[tIdx] !== '') ? row[tIdx] : (disp[r] && disp[r][tIdx]);
            rs = parseCellToDate(fRaw);
            re = parseCellToDate(tRaw);
            if (!rs || !re) continue;
            if (!_rangesOverlap_(rs, re, startDate, endDate)) continue;
          } else {
            const rowStart = (sIdx>=0 && row[sIdx] instanceof Date)? row[sIdx] : (sIdx>=0 ? parseDateTokenFlexible(String(row[sIdx]||disp[r][sIdx]||'')) : null);
            const rowEnd = (eIdx>=0 && row[eIdx] instanceof Date)? row[eIdx] : (eIdx>=0 ? parseDateTokenFlexible(String(row[eIdx]||disp[r][eIdx]||'')) : null);
            rs = rowStart || rowEnd; re = rowEnd || rowStart;
            if (!rs && !re){
              const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
              const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
              if (isNaN(d.getTime())) continue;
              if (d < startDate || d > endDate) continue;
            } else {
              if (!_rangesOverlap_(rs, re, startDate, endDate)) continue;
            }
          }
          // Prefer raw value but fall back to display value (formatted currency) if needed
          let amt = '';
          if (amtIdx>=0){
            if (typeof vals[r][amtIdx] !== 'undefined' && vals[r][amtIdx] !== '') amt = vals[r][amtIdx];
            else if (disp[r] && typeof disp[r][amtIdx] !== 'undefined') amt = disp[r][amtIdx];
          }
          rowsOut.push({ row: r+2, from: rs? Utilities.formatDate(new Date(rs), TZ(), 'dd-MMM-yy') : '', to: re? Utilities.formatDate(new Date(re), TZ(), 'dd-MMM-yy') : '', amount: amt });
        }
  const tbl = _buildTableFromRows_(rowsOut);
  const matchedCount = (rowsOut && rowsOut.length) ? rowsOut.length : 0;
  const msgWithCount = `Matched ${matchedCount} rows for the requested range.\n\n` + text;
  return { ok:true, message: msgWithCount, rows: rowsOut, table: tbl };
      }
    }catch(_eTbl){ /* ignore table-build errors */ }
    return text;
  }catch(e){
    console.error('getBeneficiaryMetricForDateRange error:', e);
    return 'Error retrieving beneficiary metric for date range.';
  }
}

/**
 * Beneficiary expense breakdown across categories within explicit date range.
 */
function getBeneficiaryExpensesForDateRange(beneficiaryName, startDate, endDate, opts){
  try{
    const name = String(beneficiaryName||'').trim();
    if (!name) return { ok:false, message:'Please provide a beneficiary name.' };
    const sh = getSheet('submissions');
    if (!sh) return { ok:false, message:'Submissions data not available.' };
    const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) return { ok:true, message:`No submissions found for ${name}.` };

  const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  // tolerant header indices
  const IX = _submissionsHeaderIndices_(head);
  IX._head = head;
  const iBen = (typeof IX.iBen !== 'undefined' && IX.iBen >= 0) ? IX.iBen : -1;
  const iTs  = (typeof IX.iTs !== 'undefined' && IX.iTs >= 0) ? IX.iTs : -1;
  // category indices from IX.metrics (fallback to naive search if unavailable)
  const iFuel = (IX && IX.metrics && typeof IX.metrics.fuel === 'number') ? IX.metrics.fuel : -1;
  const iERDA = (IX && IX.metrics && typeof IX.metrics.erda === 'number') ? IX.metrics.erda : -1;
  const iCar  = (IX && IX.metrics && typeof IX.metrics.vehicleRent === 'number') ? IX.metrics.vehicleRent : -1;
  const iAir  = (IX && IX.metrics && typeof IX.metrics.airtime === 'number') ? IX.metrics.airtime : -1;
  const iTrans= (IX && IX.metrics && typeof IX.metrics.transport === 'number') ? IX.metrics.transport : -1;
  const iMisc = (IX && IX.metrics && typeof IX.metrics.misc === 'number') ? IX.metrics.misc : -1;
  const iTotal= (IX && IX.metrics && typeof IX.metrics.total === 'number') ? IX.metrics.total : -1;
  if (iBen<0 || iTs<0) return { ok:false, message:'Required columns not found in submissions sheet.' };

    const rng = sh.getRange(2,1,lastRow-1,lastCol);
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();
    const nmLC = name.toLowerCase();
    // Compare by calendar day keys normalized to Tanzania timezone
    const key = (y,m,d)=> (y*10000) + ((m+1)*100) + d;
    const fmtDateKey = (d)=> {
      try {
        const tz = TZ();
        const fmtd = Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd');
        const [y,m,d] = fmtd.split('-').map(n=>parseInt(n,10));
        return key(y,m-1,d); // Adjust month back to 0-based for consistency
      } catch(e) { return null; }
    };
    const startKey = fmtDateKey(startDate);
    const endKey = fmtDateKey(endDate);
    let sum = { fuel:0, erda:0, car:0, air:0, transport:0, misc:0, total:0, count:0 };
  // Detect metric-specific ranges for the sheet
  const metricRanges = _submissionsMetricRangeIndices_(head);
    const anyRangeExists = Object.values(metricRanges.rangeStarts||{}).some(i=> i>=0) || Object.values(metricRanges.rangeEnds||{}).some(i=> i>=0);
    for (let r=0;r<vals.length;r++){
      const row = vals[r];
      const ben = String(row[iBen]||'').trim().toLowerCase();
      if (ben !== nmLC) continue;
      let include = false;
      if (anyRangeExists){
        // If any metric range exists, try to find a matching metric range cell for this row
        // We'll accept the row if any of the category-specific ranges overlap the requested window
        const categories = ['fuel','erda','vehicleRent','airtime','transport','misc'];
        for (let c=0;c<categories.length;c++){
          const cat = categories[c];
          const sIdx = (metricRanges.rangeStarts && typeof metricRanges.rangeStarts[cat] === 'number') ? metricRanges.rangeStarts[cat] : -1;
          const eIdx = (metricRanges.rangeEnds && typeof metricRanges.rangeEnds[cat] === 'number') ? metricRanges.rangeEnds[cat] : -1;
          if (sIdx<0 && eIdx<0) continue;
          const rowStart = (sIdx>=0 && row[sIdx] instanceof Date) ? row[sIdx] : (sIdx>=0 ? parseDateTokenFlexible(String(row[sIdx]||disp[r][sIdx]||'')) : null);
          const rowEnd = (eIdx>=0 && row[eIdx] instanceof Date) ? row[eIdx] : (eIdx>=0 ? parseDateTokenFlexible(String(row[eIdx]||disp[r][eIdx]||'')) : null);
          if (!rowStart && !rowEnd) continue;
          const rs = rowStart || rowEnd; const re = rowEnd || rowStart;
          if (_rangesOverlap_(rs,re,startDate,endDate)) { include = true; break; }
        }
        if (!include) {
          // fallback: test Timestamp
          const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
          const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
          if (isNaN(d.getTime())) continue;
          const dtKey = Number(Utilities.formatDate(new Date(d), TZ(), 'yyyyMMdd'));
          if (dtKey < startKey || dtKey > endKey) continue;
        }
      } else {
  const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
  const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
  if (isNaN(d.getTime())) continue;
  const dtKey = Number(Utilities.formatDate(new Date(d), TZ(), 'yyyyMMdd'));
  if (dtKey < startKey || dtKey > endKey) continue;
      }
  function numAt(i){ if (i<0) return 0; return parseAmount(row[i]); }
      sum.fuel      += numAt(iFuel);
      sum.erda      += numAt(iERDA);
      sum.car       += numAt(iCar);
      sum.air       += numAt(iAir);
      sum.transport += numAt(iTrans);
      sum.misc      += numAt(iMisc);
      sum.total     += numAt(iTotal);
      sum.count++;
    }
    const fmtMoney = (n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
  const label = `${Utilities.formatDate(new Date(startDate), TZ(), 'dd-MMM-yy')} to ${Utilities.formatDate(new Date(endDate), TZ(), 'dd-MMM-yy')}`;
    const lines = [
      `👤 ${name} — ${label}`,
      `💰 Total: ${fmtMoney(sum.total)} (${sum.count} submissions)`,
      '',
      `Breakdown:`,
  `⛽ Fuel: ${fmtMoney(sum.fuel)}`,
  `🏠 DA: ${fmtMoney(sum.erda)}`,
      `🚗 Vehicle Rent: ${fmtMoney(sum.car)}`,
      `📱 Airtime: ${fmtMoney(sum.air)}`,
      `🚌 Transport: ${fmtMoney(sum.transport)}`,
      `📋 Misc: ${fmtMoney(sum.misc)}`
    ];
    const text = lines.join('\n');
    try{
      const wantsTable = opts && (opts.output === 'table' || opts.wantsTable);
      if (wantsTable){
  // Build per-row table of overlapping category ranges
  const rowsOut = [];
  const categories = ['fuel','erda','vehicleRent','airtime','transport','misc'];
  // helper to find header index using tolerant header index helper
  const HIX = _headerIndex_(head.map(h=>String(h||'')));
  function findHeaderIndex(labels){ try{ return HIX.get(labels); }catch(e){ return -1; } }
        for (let r=0;r<vals.length;r++){
          const row = vals[r];
          const ben = String(row[iBen]||'').trim().toLowerCase();
          if (ben !== nmLC) continue;
          for (let c=0;c<categories.length;c++){
            const cat = categories[c];
            const sIdx = (metricRanges.rangeStarts && typeof metricRanges.rangeStarts[cat] === 'number') ? metricRanges.rangeStarts[cat] : -1;
            const eIdx = (metricRanges.rangeEnds && typeof metricRanges.rangeEnds[cat] === 'number') ? metricRanges.rangeEnds[cat] : -1;
            if (sIdx<0 && eIdx<0) continue;
            const rowStart = (sIdx>=0 && row[sIdx] instanceof Date) ? row[sIdx] : (sIdx>=0 ? parseDateTokenFlexible(String(row[sIdx]||disp[r][sIdx]||'')) : null);
            const rowEnd = (eIdx>=0 && row[eIdx] instanceof Date) ? row[eIdx] : (eIdx>=0 ? parseDateTokenFlexible(String(row[eIdx]||disp[r][eIdx]||'')) : null);
            if (!rowStart && !rowEnd) continue;
            const rs = rowStart || rowEnd; const re = rowEnd || rowStart;
            if (_rangesOverlap_(rs,re,startDate,endDate)){
              let amtIdx = (function(cat){
                if (cat==='fuel') return findHeaderIndex(['Fuel Amt','Fuel Amount','Fuel']);
                if (cat==='erda') return findHeaderIndex(['DA Amount','DA Amt','DA','ER DA Amt','ER DA Amount','ER DA','ERDA','er da amt','erda amt']);
                // Avoid plain 'Vehicle Rent' to prevent matching 'Vehicle Rent From/To'
                if (cat==='vehicleRent') return findHeaderIndex(['Car Amt','Car Amount','Vehicle Rent Amount','Vehicle Rent Amt']);
                if (cat==='airtime') return findHeaderIndex(['Airtime Amount','Airtime Amt','Airtime']);
                if (cat==='transport') return findHeaderIndex(['Transport Amount','Transport Amt','Transport']);
                return findHeaderIndex(['Misc Amt','Misc Amount','Misc','Miscellaneous']);
              })(cat);
              // fallback to IX.metrics if header search fails
              if (amtIdx < 0 && IX && IX.metrics && typeof IX.metrics[cat] === 'number' && IX.metrics[cat] >= 0) amtIdx = IX.metrics[cat];
              let amt = '';
              if (amtIdx>=0){ if (typeof vals[r][amtIdx] !== 'undefined' && vals[r][amtIdx] !== '') amt = vals[r][amtIdx]; else if (disp[r] && typeof disp[r][amtIdx] !== 'undefined') amt = disp[r][amtIdx]; }
              rowsOut.push({ row: r+2, category: cat, from: rs? Utilities.formatDate(new Date(rs), TZ(), 'dd-MMM-yy') : '', to: re? Utilities.formatDate(new Date(re), TZ(), 'dd-MMM-yy') : '', amount: amt });
            }
          }
        }
  const tbl = _buildTableFromRows_(rowsOut);
  const matchedCount = (rowsOut && rowsOut.length) ? rowsOut.length : 0;
  const msgWithCount = `Matched ${matchedCount} rows for the requested range.\n\n` + text;
  return { ok:true, message: msgWithCount, rows: rowsOut, table: tbl };
      }
    }catch(_eT){ }
    return { ok:true, message: text };
  }catch(e){
    console.error('getBeneficiaryExpensesForDateRange error:', e);
    return { ok:false, message:'Error retrieving beneficiary expenses for date range.' };
  }
}
/**
 * Return a single metric total (e.g., fuel) for a beneficiary in a given month/year.
 */
function getBeneficiaryMetricForPeriod(beneficiaryName, metric, month, year){
  try{
    const name = String(beneficiaryName||'').trim();
    if (!name) return 'Please provide a beneficiary name.';
    const res = _aggregateSubmissions_({ metric: String(metric||'total'), groupBy: 'none', filters: { timeframe: { month: month, year: year }, beneficiary: name } });
    if (!res.ok) return 'Submissions data not available.';
    const val = res.groups.size ? Array.from(res.groups.values())[0] : 0;
  const label = Utilities.formatDate(new Date(year, month, 1), TZ(), 'MMMM yyyy');
  const metricLabel = (metric==='fuel'?'Fuel':metric);
  return formatBeneficiaryMetric(name, metricLabel, label, val);
  }catch(e){
    console.error('getBeneficiaryMetricForPeriod error:', e);
    return 'Error retrieving beneficiary metric.';
  }
}

/** Diagnostic helper: returns detailed info about metric table resolution for a beneficiary and date range
 * Usage (run from Apps Script): debugBeneficiaryMetricTable('Mompati','erda','01-Sep-2025','30-Sep-2025')
 * Returns an object with head, IX, metricRanges, indices, matchedRows and table (if any)
 */
  function debugBeneficiaryMetricTable(beneficiaryName, metric, startText, endText){
  try{
    const startDate = (startText instanceof Date) ? startText : parseDateTokenFlexible(String(startText||''));
    const endDate = (endText instanceof Date) ? endText : parseDateTokenFlexible(String(endText||''));
    const sh = getSheet('submissions'); if (!sh) return { ok:false, reason:'no_sheet' };
    const lastRow = sh.getLastRow(); const lastCol = sh.getLastColumn(); if (lastRow<=1) return { ok:false, reason:'no_data' };
    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const rng = sh.getRange(2,1,lastRow-1,lastCol);
    const vals = rng.getValues(); const disp = rng.getDisplayValues();
    const IX = _submissionsHeaderIndices_(head); IX._head = head;
    const metricRanges = _submissionsMetricRangeIndices_(head);
    const iBen = (typeof IX.iBen==='number')?IX.iBen:-1; const iTs = (typeof IX.iTs==='number')?IX.iTs:-1;
    const mapMetricCol = (IX && IX.metrics) ? IX.metrics : {};
    const sIdx = (metricRanges.rangeStarts && typeof metricRanges.rangeStarts[metric] === 'number') ? metricRanges.rangeStarts[metric] : -1;
    const eIdx = (metricRanges.rangeEnds && typeof metricRanges.rangeEnds[metric] === 'number') ? metricRanges.rangeEnds[metric] : -1;
    const HIX = _headerIndex_(head.map(h=>String(h||'')));
    function findHeaderIndex(labels){ try{ return HIX.get(labels); }catch(e){ return -1; } }
    const amtCandidates = (function(m){
      if (m==='fuel') return ['Fuel Amt','Fuel Amount','Fuel'];
      if (m==='erda') return ['DA Amount','DA Amt','DA','ER DA Amt','ER DA Amount','ER DA','ERDA','er da amt','erda amt'];
      // Avoid plain 'Vehicle Rent' to prevent matching 'Vehicle Rent From/To'
      if (m==='vehicleRent') return ['Car Amt','Car Amount','Vehicle Rent Amount','Vehicle Rent Amt'];
      if (m==='airtime') return ['Airtime Amount','Airtime Amt','Airtime'];
      if (m==='transport') return ['Transport Amount','Transport Amt','Transport'];
      return ['Misc Amt','Misc Amount','Misc','Miscellaneous'];
    })(metric);
  let amtIdx = findHeaderIndex(amtCandidates);
  // fallback to mapMetricCol/IX metrics if header search fails
  if (amtIdx < 0 && mapMetricCol && typeof mapMetricCol[metric] === 'number' && mapMetricCol[metric] >= 0) amtIdx = mapMetricCol[metric];
    const outRows = [];
    const nmLC = String(beneficiaryName||'').toLowerCase();
    for (let r=0;r<vals.length;r++){
      const row = vals[r];
      const ben = String(row[iBen]||'').trim().toLowerCase(); if (ben !== nmLC) continue;
      // determine row metric range
      const rowStart = (sIdx>=0 && row[sIdx] instanceof Date) ? row[sIdx] : (sIdx>=0 ? parseDateTokenFlexible(String(row[sIdx]||disp[r][sIdx]||'')) : null);
      const rowEnd = (eIdx>=0 && row[eIdx] instanceof Date) ? row[eIdx] : (eIdx>=0 ? parseDateTokenFlexible(String(row[eIdx]||disp[r][eIdx]||'')) : null);
      const rs = rowStart || rowEnd; const re = rowEnd || rowStart;
      let include = false;
      if (sIdx>=0 || eIdx>=0){
        if (!rs && !re){ // fallback to timestamp
          const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
          const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
          if (!isNaN(d.getTime())){
            include = !(tzDateKey(d) < tzDateKey(startDate) || tzDateKey(d) > tzDateKey(endDate));
          }
        } else {
          include = _rangesOverlap_(rs,re,startDate,endDate);
        }
      } else {
        const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
        const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
        if (!isNaN(d.getTime())){
          include = !(tzDateKey(d) < tzDateKey(startDate) || tzDateKey(d) > tzDateKey(endDate));
        }
      }
      if (!include) continue;
  let amt = '';
  if (amtIdx>=0){ if (typeof vals[r][amtIdx] !== 'undefined' && vals[r][amtIdx] !== '') amt = vals[r][amtIdx]; else if (disp[r] && typeof disp[r][amtIdx] !== 'undefined') amt = disp[r][amtIdx]; }
  outRows.push({ row: r+2, from: rs? Utilities.formatDate(new Date(rs), TZ(), 'dd-MMM-yy') : '', to: re? Utilities.formatDate(new Date(re), TZ(), 'dd-MMM-yy') : '', amount: amt });
    }
    const tbl = _buildTableFromRows_(outRows);
    return { ok:true, head: head, IX: IX, metricRanges: metricRanges, indices: { iBen:iBen, iTs:iTs, sIdx:sIdx, eIdx:eIdx, amtIdx:amtIdx }, amtCandidates: amtCandidates, matchedRows: outRows.length, rows: outRows, table: tbl };
  }catch(e){ console.error('debugBeneficiaryMetricTable error', e); return { ok:false, error:String(e) }; }
}

/**
 * Group by Beneficiary for a given month/year and return a ranked summary.
 * Works with both detailed and compact submissions schemas.
 */
function getBeneficiaryMonthlySummary(month, year){
  try{
    const sh = getSheet('submissions');
    if (!sh) return 'Submissions data not available.';
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0){
  const label = Utilities.formatDate(new Date(year,month,1), TZ(), 'MMMM yyyy');
      return `No submissions found for ${label}.`;
    }

    const head = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h=>String(h||''));
    const idx = (labels)=>{ for (let i=0;i<labels.length;i++){ const j=head.indexOf(labels[i]); if (j!==-1) return j; } return -1; };
    const iBen  = idx(['Beneficiary']);
    const iTs   = idx(['Timestamp','Date','Date and time of entry','Date and time']);
    const iFuel = idx(['Fuel Amt','Fuel']);
    const iERDA = idx(['ER DA Amt','ER DA']);
    const iCar  = idx(['Car Amt','Vehicle Rent']);
    const iAir  = idx(['Airtime Amt','Airtime']);
    const iTrans= idx(['Transport Amt','Transport']);
    const iMisc = idx(['Misc Amt','Misc','Miscellaneous']);
    const iTotal= idx(['Row Total','Total Amount','Total']);
    if (iBen<0 || iTs<0) return 'Required columns not found in submissions sheet.';

    const rng = sh.getRange(2,1,lastRow-1,lastCol);
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();

    const map = new Map(); // name -> { total, fuel, erda, car, air, transport, misc, count }
    function numAt(row, i){ if (i<0) return 0; return parseAmount(row[i]); }
    for (let r=0;r<vals.length;r++){
      const row = vals[r];
      const ben = String(row[iBen]||'').trim();
      if (!ben) continue;
      const rawTs = (row[iTs] instanceof Date) ? row[iTs] : (row[iTs]||disp[r][iTs]||'');
      const d = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
      if (isNaN(d.getTime())) continue;
      if (d.getMonth() !== month || d.getFullYear() !== year) continue;
      const acc = map.get(ben) || { fuel:0, erda:0, car:0, air:0, transport:0, misc:0, total:0, count:0 };
      acc.fuel      += numAt(row, iFuel);
      acc.erda      += numAt(row, iERDA);
      acc.car       += numAt(row, iCar);
      acc.air       += numAt(row, iAir);
      acc.transport += numAt(row, iTrans);
      acc.misc      += numAt(row, iMisc);
      acc.total     += numAt(row, iTotal);
      acc.count++;
      map.set(ben, acc);
    }

  const label = Utilities.formatDate(new Date(year,month,1), TZ(), 'MMMM yyyy');
    if (map.size === 0) return `No submissions found for ${label}.`;
    const list = Array.from(map.entries()).sort((a,b)=> (b[1].total||0) - (a[1].total||0));
    return formatBeneficiaryMonthlySummary(label, list);
  }catch(e){
    console.error('getBeneficiaryMonthlySummary error:', e);
    return 'Error retrieving beneficiary-wise summary.';
  }
}
// Intent handlers and AI-facing response helpers have been moved to `handlers.gs` to keep `code.gs` smaller.
// See `handlers.gs` for implementations of handleBeneficiaryIntent, handleTeamIntent, etc.

function handleUnknownIntent(query, entities) {
  // Intelligent fallback with context awareness
  let suggestions = [];
  
  if (query.includes('money') || query.includes('cost') || query.includes('spend')) {
    suggestions.push('💰 Try: "What are the monthly expenses?"');
  }
  
  if (query.includes('who') || query.includes('person') || query.includes('people')) {
    suggestions.push('👥 Try: "Show me beneficiary expenses"');
  }
  
  if (query.includes('how much') || query.includes('total')) {
    suggestions.push('📊 Try: "Give me an expense summary"');
  }
  
  // Delegate fallback message building to centralized helper in responses.gs
  return getFallbackResponse(suggestions);
}

/**
 * Get monthly vehicle releases summary
 */
function getMonthlyExpenses() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const findCol = (names) => {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const timestampCol = findCol(['Date and time of entry','Date','Timestamp']);
    
    if (timestampCol === -1) {
      return 'Unable to find timestamp column in the data.';
    }
    
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    let requestCount = 0;
    const vehicles = [];
    const carCol = findCol(['Vehicle Number','Car Number']);
    const projectCol = findCol(['Project']);
    
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][timestampCol]);
      
      if (timestamp.getMonth() === currentMonth && timestamp.getFullYear() === currentYear) {
        requestCount++;
        const carNumber = carCol >= 0 ? (data[i][carCol] || 'Unknown') : 'Unknown';
        const project = projectCol >= 0 ? (data[i][projectCol] || 'Unknown') : 'Unknown';
        vehicles.push({ car: carNumber, project: project });
      }
    }
    
    let response = `This month's vehicle releases: ${requestCount} vehicles released`;
    
    if (vehicles.length > 0 && vehicles.length <= 5) {
      response += '\n\nVehicles released:';
      vehicles.forEach(v => {
        response += `\n• ${v.car} (${v.project})`;
      });
    }
    
    return response;
    
  } catch (error) {
    console.error('getMonthlyExpenses error:', error);
    return 'Sorry, I could not retrieve monthly vehicle data.';
  }
}

/**
 * Get vehicles currently in use
 */
function getPendingRequests() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const findCol = (names) => {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const statusCol = findCol(['Status','In Use/Release','In Use / release']);
    const carCol = findCol(['Vehicle Number','Car Number']);
    const projectCol = findCol(['Project']);
    const teamCol = findCol(['Team']);
    
    if (statusCol === -1 || carCol === -1) {
      return 'Unable to find required columns for vehicle usage.';
    }

    let inUseCount = 0;
    const inUseVehicles = [];
    
    for (let i = 1; i < data.length; i++) {
      const status = (data[i][statusCol] || '').toString().toLowerCase();
      if (status.includes('use') || status.includes('in use')) {
        inUseCount++;
        
        if (inUseVehicles.length < 5) {
          inUseVehicles.push({
            car: data[i][carCol] || 'Unknown',
            project: data[i][projectCol] || 'Unknown',
            team: data[i][teamCol] || 'Unknown'
          });
        }
      }
    }
    
    let response = `Found ${inUseCount} vehicles currently in use`;
    
    if (inUseVehicles.length > 0) {
      response += '\n\nVehicles in use:';
      inUseVehicles.forEach(vehicle => {
        response += `\n• ${vehicle.car} - ${vehicle.project} (${vehicle.team})`;
      });
    }
    
    return response;
    
  } catch (error) {
    console.error('getPendingRequests error:', error);
    return 'Sorry, I could not retrieve vehicle usage data.';
  }
}

/**
 * Latest IN USE vehicles for a specific team (by latest entry per vehicle)
 */
function getTeamVehiclesInUse(teamName) {
  try {
    var tName = String(teamName||'').trim();
    if (!tName) return 'Please provide a team name.';
    var rows = _readCarTP_objects_();
    if (!rows || !rows.length) return 'No vehicle records found.';
    var tLc = tName.toLowerCase();
    // Filter to team, newest first
    rows = rows.filter(function(r){ return String(r.Team||'').trim().toLowerCase() === tLc; });
    if (!rows.length) return 'No records found for team ' + tName + '.';
    rows.sort(function(a,b){ return (b._ts||0) - (a._ts||0); });
    // Deduplicate to latest row per vehicle number
    var seen = new Set();
    var latest = [];
    for (var i=0;i<rows.length;i++){
      var vnum = String(rows[i]['Vehicle Number']||'').trim();
      if (!vnum || seen.has(vnum)) continue;
      seen.add(vnum);
      latest.push(rows[i]);
    }
    // Keep only IN USE
    var inUse = latest.filter(function(r){ return _normStatus_(r.Status) === 'IN USE'; });
    if (!inUse.length) return 'No vehicles currently in use for team ' + tName + '.';
    var lines = ['Vehicles currently in use for ' + tName + ':'];
    for (var j=0;j<inUse.length;j++){
      var v = inUse[j];
      var mm = [String(v.Make||'').trim(), String(v.Model||'').trim()].filter(function(x){return x;}).join(' ');
      var line = '• ' + String(v['Vehicle Number']||'').trim();
      if (mm) line += ' — ' + mm;
      lines.push(line);
    }
    return lines.join('\n');
  } catch(e) {
    console.error('getTeamVehiclesInUse error:', e);
    return 'Error retrieving vehicles in use for team ' + String(teamName||'') + '.';
  }
}

/**
 * Latest vehicle record for a team.
 * - Prefers rows with Status='IN USE' if preferStatus is provided and exists.
 * - Otherwise returns the newest row by Date/Time for that team.
 */
function getLatestVehicleForTeam(teamName, preferStatus){
  try{
    var tName = String(teamName||'').trim();
    if (!tName) return { ok:false, message:'Please provide a team name.' };
    var rows = _readCarTP_objects_();
    if (!rows || !rows.length) return { ok:false, message:'No vehicle records found.' };
    var tLc = tName.toLowerCase();
    rows = rows.filter(function(r){ return String(r.Team||'').trim().toLowerCase() === tLc; });
    if (!rows.length) return { ok:false, message:'No records found for team ' + tName };
    // Sort newest first by parsed timestamp
    rows.sort(function(a,b){ return (b._ts||0) - (a._ts||0); });
    // Optionally prefer a status (e.g., IN USE) if available among newest entries
    var target = null;
    var want = _normStatus_((preferStatus||''));
    if (want){
      for (var i=0;i<rows.length;i++){
        if (_normStatus_(rows[i].Status) === want){ target = rows[i]; break; }
      }
    }
    if (!target) target = rows[0];
    return {
      ok:true,
      team: tName,
      vehicleNumber: String(target['Vehicle Number']||'').trim(),
      status: _normStatus_(target.Status||''),
      make: String(target.Make||'').trim(),
      model: String(target.Model||'').trim(),
      category: String(target.Category||'').trim(),
      responsibleBeneficiary: String(target['R.Beneficiary'] || target['R. Ben'] || target.responsibleBeneficiary || '').trim(),
      dateTime: target['Date and time of entry'] || null
    };
  }catch(e){
    console.error('getLatestVehicleForTeam error:', e);
    return { ok:false, message:'Error retrieving latest vehicle for team.' };
  }
}

/**
 * Return the most recent IN USE vehicle per requested team.
 * Falls back to the latest row for that team if no IN USE entry exists.
 */
function getLatestInUseVehiclesForTeams(teamNames){
  try{
    if (!Array.isArray(teamNames)) return { ok:false, error:'Invalid team list' };

    const requested = [];
    const seenKeys = new Set();
    teamNames.forEach(function(name){
      const original = String(name || '').trim();
      if (!original) return;
      const key = original.toLowerCase();
      if (seenKeys.has(key)) return;
      seenKeys.add(key);
      requested.push({ original: original, key: key });
    });
    if (!requested.length) return { ok:true, teams:{} };

    const rows = _readCarTP_objects_();
    if (!rows.length) return { ok:true, teams:{} };

    const requestedKeySet = new Set(requested.map(function(t){ return t.key; }));
    const teamMap = new Map(); // key -> { inUse:{ts,record}, latest:{ts,record} }

    rows.forEach(function(row){
      const team = String(row.Team || '').trim();
      if (!team) return;
      const key = team.toLowerCase();
      if (!requestedKeySet.has(key)) return;

      var ts = 0;
      if (typeof row._ts === 'number' && !isNaN(row._ts)) {
        ts = row._ts;
      } else {
        ts = _parseTs_(row['Date and time of entry']) || 0;
      }

      var dateIso = '';
      if (row['Date and time of entry'] instanceof Date) {
        dateIso = row['Date and time of entry'].toISOString();
      } else if (ts) {
        try { dateIso = new Date(ts).toISOString(); } catch(_){ /* ignore */ }
      }

      const record = {
        team: team,
        project: String(row.Project || row.project || '').trim(),
        carNumber: String(row['Vehicle Number'] || '').trim(),
        make: String(row.Make || '').trim(),
        model: String(row.Model || '').trim(),
        usageType: String(row['Usage Type'] || '').trim(),
        category: String(row.Category || '').trim(),
        owner: String(row.Owner || '').trim(),
        status: _normStatus_(row.Status || ''),
        remarks: String(row['Last Users remarks'] || '').trim(),
       stars: Number(row.Ratings || 0) || 0,
       dateTime: dateIso,
        responsibleBeneficiary: String(row['R.Beneficiary'] || row['R. Ben'] || row.responsibleBeneficiary || '').trim()
      };

      const entry = teamMap.get(key) || { inUse:null, latest:null };
      if (!entry.latest || ts >= entry.latest.ts) {
        entry.latest = { ts: ts, record: record };
      }
      if (record.status === 'IN USE' && (!entry.inUse || ts >= entry.inUse.ts)) {
        entry.inUse = { ts: ts, record: record };
      }
      teamMap.set(key, entry);
    });

    const out = {};
    requested.forEach(function(req){
      const entry = teamMap.get(req.key);
      if (!entry) return;
      const preferred = entry.inUse && entry.inUse.record ? entry.inUse.record : (entry.latest && entry.latest.record ? entry.latest.record : null);
      if (!preferred) return;
      const status = _normStatus_(preferred.status || '');
      if (status !== 'IN USE') return;
      if (!preferred.carNumber) return;
      out[req.original] = {
        carNumber: preferred.carNumber,
        status: status,
        project: preferred.project,
        make: preferred.make,
        model: preferred.model,
        usageType: preferred.usageType,
        category: preferred.category,
        owner: preferred.owner,
        remarks: preferred.remarks,
        stars: preferred.stars,
        team: preferred.team,
        dateTime: preferred.dateTime,
        responsibleBeneficiary: preferred.responsibleBeneficiary,
        source: entry.inUse && entry.inUse.record ? 'car_tp_inuse' : 'car_tp_latest'
      };
    });

    if (!Object.keys(out).length) {
      try {
        const summary = getVehicleInUseData();
        if (summary && summary.ok && Array.isArray(summary.assignments)) {
          const summaryMap = summary.assignments.reduce(function(map, row){
            if (!row) return map;
            const key = String(row.team || '').trim().toLowerCase();
            if (!key || map.has(key)) return map;
            map.set(key, row);
            return map;
          }, new Map());

          requested.forEach(function(req){
            if (out[req.original]) return;
            const row = summaryMap.get(req.key);
            if (!row) return;
            const status = _normStatus_(row.status || 'IN USE');
            if (status !== 'IN USE') return;
            const vehicleNumber = String(row.vehicleNumber || '').trim();
            if (!vehicleNumber) return;
            out[req.original] = {
              carNumber: vehicleNumber,
              status: status,
              project: row.project || '',
              make: row.make || '',
              model: row.model || '',
              usageType: row.usageType || '',
              category: row.category || '',
              owner: row.owner || '',
              remarks: row.remarks || '',
              stars: row.ratings ? Number(row.ratings) || 0 : 0,
              team: row.team || '',
              dateTime: row.generatedAt || row.updatedAt || row.entryDate || '',
              responsibleBeneficiary: row.responsibleBeneficiary || row.beneficiary || '',
              source: 'vehicle_in_use_summary'
            };
          });
        }
      } catch (fallbackErr) {
        console.warn('getLatestInUseVehiclesForTeams summary fallback failed:', fallbackErr);
      }
    }

    return { ok:true, teams: out };
  }catch(e){
    console.error('getLatestInUseVehiclesForTeams error:', e);
    return { ok:false, error:String(e) };
  }
}

/**
 * Return the list of beneficiaries currently marked as IN USE for a team.
 * Reads the Vehicle_InUse summary sheet and extracts the R.Beneficiary column.
 */
function getTeamMembersCurrentlyInUse(teamName, projectName) {
  try {
    const team = String(teamName || '').trim();
    if (!team) return { ok: true, members: [] };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName('Vehicle_InUse');
    if (!sh) return { ok: true, members: [] };

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return { ok: true, members: [] };

    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    function idx(labels, required) {
      try {
        return IX.get(labels);
      } catch (e) {
        if (required) throw e;
        return -1;
      }
    }

    const iTeam = idx(['Team', 'Team Name'], true);
    const iResp = idx(['R.Beneficiary', 'Responsible Beneficiary', 'R Beneficiary', 'Responsible', 'R. Ben', 'R Ben'], false);
    const iStatus = idx(['Status', 'In Use/Release', 'In Use'], false);

    const iDate = idx(['Date and time of entry', 'Date and time', 'Timestamp', 'Date'], false);
    const iCar = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car'], false);

    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const targetTeam = team.toLowerCase();
    const nowMs = Date.now();
    const records = new Map();

    data.forEach(row => {
      const rowTeam = String(row[iTeam] || '').trim().toLowerCase();
      if (!rowTeam) return;
      if (rowTeam !== targetTeam) return;

      if (iStatus >= 0) {
        const status = String(row[iStatus] || '').trim().toUpperCase();
        if (status && status !== 'IN USE' && status.indexOf('IN USE') === -1) {
          return;
        }
      }

      const rawNames = iResp >= 0 ? String(row[iResp] || '').trim() : '';
      if (!rawNames) return;

      let entryMs = 0;
      if (iDate >= 0) {
        const entryVal = row[iDate];
        if (entryVal instanceof Date) {
          entryMs = entryVal.getTime();
        } else if (entryVal) {
          const parsed = new Date(entryVal);
          if (!isNaN(parsed.getTime())) entryMs = parsed.getTime();
        }
      }

      const vehicleNumber = iCar >= 0 ? String(row[iCar] || '').trim() : '';
      const daysInUse = entryMs ? Math.max(0, Math.floor((nowMs - entryMs) / 86400000)) : null;

      rawNames.split(/[,;/\n]+/).forEach(part => {
        const name = String(part || '').trim();
        if (!name) return;
        const key = name.toLowerCase();
        const existing = records.get(key);
        if (!existing || entryMs > existing.entryMs) {
          records.set(key, {
            entryMs,
            name,
            vehicleNumber,
            daysInUse
          });
        }
      });
    });

    const members = Array.from(records.values()).map(rec => ({
      name: rec.name,
      vehicleNumber: rec.vehicleNumber,
      daysInUse: rec.daysInUse
    }));

    return { ok: true, members };
  } catch (e) {
    console.error('getTeamMembersCurrentlyInUse error:', e);
    return { ok: false, error: String(e), members: [] };
  }
}

function _deserializeVehicleReleasedWrapper_(raw, sourceLabel) {
  if (!raw) return null;
  try {
    const wrapper = JSON.parse(raw);
    if (!wrapper || typeof wrapper !== 'object' || !wrapper.payload) {
      return null;
    }
    const payload = Object.assign({}, wrapper.payload);
    payload.cached = true;
    payload.cacheSource = sourceLabel || null;
    payload.cacheVersion = wrapper.version || null;
    return payload;
  } catch (err) {
    console.warn('Vehicle_Released cache parse failed:', err);
    return null;
  }
}

function _loadVehicleReleasedDropdownPayload_() {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  let props = null;
  try {
    props = PropertiesService.getScriptProperties();
  } catch (_err) {
    props = null;
  }

  let version = null;
  if (props) {
    try {
      version = props.getProperty(VEHICLE_RELEASED_VERSION_PROP_KEY) || null;
    } catch (_err) {
      version = null;
    }
  }

  function fromCache(raw, label) {
    const parsed = _deserializeVehicleReleasedWrapper_(raw, label);
    if (!parsed) return null;
    if (version && parsed.cacheVersion && parsed.cacheVersion !== version) {
      return null;
    }
    return parsed;
  }

  let payload = null;

  if (!payload && cache) {
    try {
      const cachedRaw = cache.get(VEHICLE_RELEASED_CACHE_KEY);
      if (cachedRaw) {
        payload = fromCache(cachedRaw, 'CacheService');
      }
    } catch (cacheErr) {
      console.warn('Vehicle_Released cache lookup failed:', cacheErr);
    }
  }

  if (!payload && props) {
    try {
      const storedRaw = props.getProperty(VEHICLE_RELEASED_PROP_KEY);
      if (storedRaw) {
        payload = _deserializeVehicleReleasedWrapper_(storedRaw, 'PropertiesService');
        if (payload) {
          version = payload.cacheVersion || version;
          if (cache) {
            try {
              cache.put(VEHICLE_RELEASED_CACHE_KEY, storedRaw, VEHICLE_RELEASED_CACHE_TTL_SECONDS);
            } catch (cachePutErr) {
              console.warn('Vehicle_Released cache backfill failed:', cachePutErr);
            }
          }
        }
      }
    } catch (propErr) {
      console.warn('Vehicle_Released properties lookup failed:', propErr);
    }
  }

  let lock = null;
  let locked = false;
  if (!payload && typeof LockService !== 'undefined') {
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(5000);
      locked = true;
    } catch (lockErr) {
      console.warn('Vehicle_Released cache lock acquisition failed:', lockErr);
    }
  }

  try {
    if (!payload && cache) {
      try {
        const cachedRaw = cache.get(VEHICLE_RELEASED_CACHE_KEY);
        if (cachedRaw) {
          payload = fromCache(cachedRaw, 'CacheService');
        }
      } catch (cacheRetryErr) {
        console.warn('Vehicle_Released cache retry failed:', cacheRetryErr);
      }
    }

    if (!payload && props) {
      try {
        const storedRaw = props.getProperty(VEHICLE_RELEASED_PROP_KEY);
        if (storedRaw) {
          payload = _deserializeVehicleReleasedWrapper_(storedRaw, 'PropertiesService');
          if (payload && cache) {
            try {
              cache.put(VEHICLE_RELEASED_CACHE_KEY, storedRaw, VEHICLE_RELEASED_CACHE_TTL_SECONDS);
            } catch (cachePutErr) {
              console.warn('Vehicle_Released cache backfill failed (post-lock):', cachePutErr);
            }
          }
        }
      } catch (propRetryErr) {
        console.warn('Vehicle_Released properties retry failed:', propRetryErr);
      }
    }
  } finally {
    if (locked && lock) {
      try { lock.releaseLock(); } catch (_releaseErr) { /* ignore */ }
    }
  }

  if (!payload) {
    const fresh = _buildVehicleReleasedDropdownPayload_();
    payload = Object.assign({}, fresh, {
      cached: false,
      cacheSource: 'fresh',
      cacheVersion: null
    });

    if (fresh && fresh.ok) {
      const versionStamp = String(Date.now());
      payload.cacheVersion = versionStamp;
      const wrapper = { version: versionStamp, payload: fresh };
      const serialized = JSON.stringify(wrapper);
      if (!props) {
        try {
          props = PropertiesService.getScriptProperties();
        } catch (_propInitErr) {
          props = null;
        }
      }
      if (props) {
        try {
          props.setProperty(VEHICLE_RELEASED_PROP_KEY, serialized);
          props.setProperty(VEHICLE_RELEASED_VERSION_PROP_KEY, versionStamp);
        } catch (propStoreErr) {
          console.error('Failed to persist Vehicle_Released payload to PropertiesService:', propStoreErr);
        }
      }
      if (cache) {
        try {
          cache.put(VEHICLE_RELEASED_CACHE_KEY, serialized, VEHICLE_RELEASED_CACHE_TTL_SECONDS);
        } catch (cacheStoreErr) {
          console.warn('Failed to persist Vehicle_Released payload to CacheService:', cacheStoreErr);
        }
      }
    }
  }

  if (payload && payload.cached && payload.ok !== false) {
    const vehicles = Array.isArray(payload.vehicles) ? payload.vehicles : [];
    if (!vehicles.length) {
      invalidateVehicleReleasedCache('Cached Vehicle_Released payload empty, forcing rebuild');
      const fresh = _buildVehicleReleasedDropdownPayload_();
      payload = Object.assign({}, fresh, {
        cached: false,
        cacheSource: 'fresh',
        cacheVersion: null
      });
      if (fresh && fresh.ok) {
        const versionStamp = String(Date.now());
        payload.cacheVersion = versionStamp;
        const wrapper = { version: versionStamp, payload: fresh };
        const serialized = JSON.stringify(wrapper);
        if (!props) {
          try {
            props = PropertiesService.getScriptProperties();
          } catch (_propInitErr) {
            props = null;
          }
        }
        if (props) {
          try {
            props.setProperty(VEHICLE_RELEASED_PROP_KEY, serialized);
            props.setProperty(VEHICLE_RELEASED_VERSION_PROP_KEY, versionStamp);
          } catch (propStoreErr) {
            console.error('Failed to persist Vehicle_Released payload to PropertiesService (rebuild):', propStoreErr);
          }
        }
        if (cache) {
          try {
            cache.put(VEHICLE_RELEASED_CACHE_KEY, serialized, VEHICLE_RELEASED_CACHE_TTL_SECONDS);
          } catch (cacheStoreErr) {
            console.warn('Failed to persist Vehicle_Released payload to CacheService (rebuild):', cacheStoreErr);
          }
        }
      }
    }
  }

  return payload;
}

function _loadVehicleDropdownPayload_(sheetName) {
  const cacheKey = 'VEHICLE_' + sheetName.toUpperCase() + '_CACHE_KEY';
  const propKey = 'VEHICLE_' + sheetName.toUpperCase() + '_PROP_KEY';
  const versionPropKey = 'VEHICLE_' + sheetName.toUpperCase() + '_VERSION_PROP_KEY';
  const cacheTtl = VEHICLE_RELEASED_CACHE_TTL_SECONDS || 3600;

  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  let props = null;
  try {
    props = PropertiesService.getScriptProperties();
  } catch (_err) {
    props = null;
  }

  let version = null;
  if (props) {
    try {
      version = props.getProperty(versionPropKey) || null;
    } catch (_err) {
      version = null;
    }
  }

  function fromCache(raw, label) {
    const parsed = _deserializeVehicleReleasedWrapper_(raw, label);
    if (!parsed) return null;
    if (version && parsed.cacheVersion && parsed.cacheVersion !== version) {
      return null;
    }
    return parsed;
  }

  let payload = null;

  if (!payload && cache) {
    try {
      const cachedRaw = cache.get(cacheKey);
      if (cachedRaw) {
        payload = fromCache(cachedRaw, 'CacheService');
      }
    } catch (cacheErr) {
      console.warn(sheetName + ' cache lookup failed:', cacheErr);
    }
  }

  if (!payload && props) {
    try {
      const storedRaw = props.getProperty(propKey);
      if (storedRaw) {
        payload = _deserializeVehicleReleasedWrapper_(storedRaw, 'PropertiesService');
        if (payload) {
          version = payload.cacheVersion || version;
          if (cache) {
            try {
              cache.put(cacheKey, storedRaw, cacheTtl);
            } catch (cachePutErr) {
              console.warn(sheetName + ' cache backfill failed:', cachePutErr);
            }
          }
        }
      }
    } catch (propErr) {
      console.warn(sheetName + ' properties lookup failed:', propErr);
    }
  }

  let lock = null;
  let locked = false;
  if (!payload && typeof LockService !== 'undefined') {
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(5000);
      locked = true;
    } catch (lockErr) {
      console.warn(sheetName + ' cache lock acquisition failed:', lockErr);
    }
  }

  try {
    if (!payload && cache) {
      try {
        const cachedRaw = cache.get(cacheKey);
        if (cachedRaw) {
          payload = fromCache(cachedRaw, 'CacheService');
        }
      } catch (cacheRetryErr) {
        console.warn(sheetName + ' cache retry failed:', cacheRetryErr);
      }
    }

    if (!payload && props) {
      try {
        const storedRaw = props.getProperty(propKey);
        if (storedRaw) {
          payload = _deserializeVehicleReleasedWrapper_(storedRaw, 'PropertiesService');
          if (payload && cache) {
            try {
              cache.put(cacheKey, storedRaw, cacheTtl);
            } catch (cachePutErr) {
              console.warn(sheetName + ' cache backfill failed (post-lock):', cachePutErr);
            }
          }
        }
      } catch (propRetryErr) {
        console.warn(sheetName + ' properties retry failed:', propRetryErr);
      }
    }
  } finally {
    if (locked && lock) {
      try { lock.releaseLock(); } catch (_releaseErr) { /* ignore */ }
    }
  }

  if (!payload) {
    const fresh = _buildVehicleDropdownPayload_(sheetName);
    payload = Object.assign({}, fresh, {
      cached: false,
      cacheSource: 'fresh',
      cacheVersion: null
    });

    if (fresh && fresh.ok) {
      const versionStamp = String(Date.now());
      payload.cacheVersion = versionStamp;
      const wrapper = { version: versionStamp, payload: fresh };
      const serialized = JSON.stringify(wrapper);
      if (!props) {
        try {
          props = PropertiesService.getScriptProperties();
        } catch (_propInitErr) {
          props = null;
        }
      }
      if (props) {
        try {
          props.setProperty(propKey, serialized);
          props.setProperty(versionPropKey, versionStamp);
        } catch (propStoreErr) {
          console.error('Failed to persist ' + sheetName + ' payload to PropertiesService:', propStoreErr);
        }
      }
      if (cache) {
        try {
          cache.put(cacheKey, serialized, cacheTtl);
        } catch (cacheStoreErr) {
          console.warn('Failed to persist ' + sheetName + ' payload to CacheService:', cacheStoreErr);
        }
      }
    }
  }

  if (payload && payload.cached && payload.ok !== false) {
    const vehicles = Array.isArray(payload.vehicles) ? payload.vehicles : [];
    if (!vehicles.length) {
      invalidateVehicleCache(sheetName, 'Cached ' + sheetName + ' payload empty, forcing rebuild');
      const fresh = _buildVehicleDropdownPayload_(sheetName);
      payload = Object.assign({}, fresh, {
        cached: false,
        cacheSource: 'fresh',
        cacheVersion: null
      });
      if (fresh && fresh.ok) {
        const versionStamp = String(Date.now());
        payload.cacheVersion = versionStamp;
        const wrapper = { version: versionStamp, payload: fresh };
        const serialized = JSON.stringify(wrapper);
        if (!props) {
          try {
            props = PropertiesService.getScriptProperties();
          } catch (_propInitErr) {
            props = null;
          }
        }
        if (props) {
          try {
            props.setProperty(propKey, serialized);
            props.setProperty(versionPropKey, versionStamp);
          } catch (propStoreErr) {
            console.error('Failed to persist ' + sheetName + ' payload to PropertiesService (rebuild):', propStoreErr);
          }
        }
        if (cache) {
          try {
            cache.put(cacheKey, serialized, cacheTtl);
          } catch (cacheStoreErr) {
            console.warn('Failed to persist ' + sheetName + ' payload to CacheService (rebuild):', cacheStoreErr);
          }
        }
      }
    }
  }

  return payload;
}

function _buildVehicleReleasedDropdownPayload_() {
  _maybeAutoRefreshCarTPSummaries_(10);

  const summary = getVehicleReleasedSummary();
  const generatedAt = new Date().toISOString();

  if (!summary.ok) {
    return {
      ok: false,
      source: 'Vehicle_Released',
      count: 0,
      vehicles: [],
      cached: false,
      updatedAt: summary.updatedAt || '',
      error: summary.error || 'Vehicle_Released summary unavailable',
      notes: summary.notes || null,
      sheetId: summary.sheetId || null,
      tried: summary.tried || null,
      summaryMessage: summary.message || null,
      generatedAt: generatedAt
    };
  }

  const vehicles = (summary.vehicles || []).map(function(entry) {
    const carNumber = String(entry.vehicleNumber || entry.carNumber || '').trim();
    const status = _normStatus_(entry.status || 'RELEASE') || 'RELEASE';
    return Object.assign({}, entry, {
      carNumber: carNumber,
      vehicleNumber: carNumber,
      latestRelease: entry.latestRelease || summary.updatedAt || '',
      status: status
    });
  }).filter(function(entry) {
    return String(entry.carNumber || '').trim() !== '';
  });

  if (!vehicles.length) {
    console.warn('[BACKEND] getVehiclePickerData via cache builder returned 0 vehicles (Vehicle_Released empty)', {
      summaryMessage: summary.message || null,
      notes: summary.notes || null,
      tried: summary.tried || null,
      sheetId: summary.sheetId || null,
      headers: summary.headerRow || null
    });
  } else {
    console.log(`[BACKEND] getVehiclePickerData cache builder loaded ${vehicles.length} vehicles from Vehicle_Released (sheet ${summary.sheetLabel || summary.sheetId || 'unknown'})`);
  }

  return {
    ok: true,
    source: 'Vehicle_Released',
    count: vehicles.length,
    vehicles: vehicles,
    cached: false,
    updatedAt: summary.updatedAt || new Date().toISOString(),
    generatedAt: generatedAt,
  notes: summary.notes || null,
  headers: summary.headerRow || null,
    sheetId: summary.sheetId || null,
    tried: summary.tried || null,
    summaryMessage: summary.message || null
  };
}

function _buildVehicleDropdownPayload_(sheetName) {
  const summary = getVehicleSummaryRows(sheetName);
  const generatedAt = new Date().toISOString();

  if (!summary.ok) {
    return {
      ok: false,
      source: sheetName,
      count: 0,
      vehicles: [],
      cached: false,
      updatedAt: summary.updatedAt || '',
      error: summary.error || sheetName + ' summary unavailable',
      notes: summary.notes || null,
      sheetId: summary.sheetId || null,
      tried: summary.tried || null,
      summaryMessage: summary.message || null,
      generatedAt: generatedAt
    };
  }

  if (!summary.rows.length) {
    if (summary.notes || summary.tried) {
      console.warn('[BACKEND] ' + sheetName + ' summary empty', { notes: summary.notes, tried: summary.tried });
    }
    return { ok: true, source: sheetName, vehicles: [], updatedAt: summary.updatedAt || '', message: sheetName + ' summary empty' };
  }

  const IX = summary.headerIndex;
  const idx = function(labels) {
    if (!IX) return -1;
    try { return IX.get(labels); } catch (_err) { return -1; }
  };

  let vehicles = [];

  if (sheetName === 'vehicle') {
    const fallback = function(index, fallbackIndex) {
      return index >= 0 ? index : fallbackIndex;
    };

    const makeIdx = fallback(idx(['Make', 'Vehicle Make', 'Car Make']), 25);
    const modelIdx = fallback(idx(['Model', 'Vehicle Model', 'Car Model']), 26);
    const categoryIdx = fallback(idx(['Category', 'Vehicle Category']), 27);
    const usageIdx = fallback(idx(['Usage Type', 'Usage', 'Vehicle Usage']), 28);
    const ownerIdx = fallback(idx(['Owner', 'Vehicle Owner']), 29);
    const vehicleIdx = fallback(idx(['Vehicle Number', 'Vehicle No', 'Car Number', 'Vehicle']), 30);

    vehicles = summary.rows.map(function(row) {
      const safeGet = function(i) {
        return i >= 0 && i < row.length ? row[i] : '';
      };
      return {
        vehicleNumber: safeGet(vehicleIdx),
        make: safeGet(makeIdx),
        model: safeGet(modelIdx),
        category: safeGet(categoryIdx),
        usageType: safeGet(usageIdx),
        owner: safeGet(ownerIdx),
        status: 'AVAILABLE'
      };
    }).filter(function(entry) {
      return String(entry.vehicleNumber || '').trim() !== '';
    });
  } else {
    // For other sheets, use the existing logic or adapt as needed
    const vehicleIdx = idx(['Vehicle Number', 'Car Number', 'Vehicle No', 'Car No', 'Car #', 'Car', 'Vehicle']);
    const statusIdx = idx(['Status', 'Release Status']);
    const makeIdx = idx(['Make']);
    const modelIdx = idx(['Model']);
    const categoryIdx = idx(['Category']);
    const usageIdx = idx(['Usage Type']);
    const ownerIdx = idx(['Owner']);

    vehicles = summary.rows.map(function(row) {
      return {
        vehicleNumber: vehicleIdx >= 0 ? row[vehicleIdx] : '',
        status: statusIdx >= 0 ? row[statusIdx] : 'RELEASE',
        make: makeIdx >= 0 ? row[makeIdx] : '',
        model: modelIdx >= 0 ? row[modelIdx] : '',
        category: categoryIdx >= 0 ? row[categoryIdx] : '',
        usageType: usageIdx >= 0 ? row[usageIdx] : '',
        owner: ownerIdx >= 0 ? row[ownerIdx] : ''
      };
    }).filter(function(entry) {
      return String(entry.vehicleNumber || '').trim() !== '';
    });
  }

  if (!vehicles.length) {
    console.warn('[BACKEND] getVehiclePickerData via cache builder returned 0 vehicles (' + sheetName + ' empty)', {
      summaryMessage: summary.message || null,
      notes: summary.notes || null,
      tried: summary.tried || null,
      sheetId: summary.sheetId || null,
      headers: summary.headerRow || null
    });
  } else {
    console.log(`[BACKEND] getVehiclePickerData cache builder loaded ${vehicles.length} vehicles from ${sheetName} (sheet ${summary.sheetLabel || summary.sheetId || 'unknown'})`);
  }

  return {
    ok: true,
    source: sheetName,
    count: vehicles.length,
    vehicles: vehicles,
    cached: false,
    updatedAt: summary.updatedAt || new Date().toISOString(),
    generatedAt: generatedAt,
    notes: summary.notes || null,
    headers: summary.headerRow || null,
    sheetId: summary.sheetId || null,
    tried: summary.tried || null,
    summaryMessage: summary.message || null
  };
}

function getVehiclePickerData(isNewCar){
  try {
    if (isNewCar) {
      return _loadVehicleDropdownPayload_('vehicle');
    } else {
      return _loadVehicleReleasedDropdownPayload_();
    }
  } catch (e) {
    console.error('getVehiclePickerData failed:', e);
    return { ok:false, source: isNewCar ? 'vehicle' : 'Vehicle_Released', error:String(e) };
  }
}

function getVehicleReleasedCacheVersion(){
  try {
    let version = null;
    let updatedAt = '';

    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    if (cache) {
      try {
        const cached = cache.get(VEHICLE_RELEASED_CACHE_KEY);
        if (cached) {
          const parsed = JSON.parse(cached);
          if (parsed) {
            if (!version && parsed.version) version = String(parsed.version);
            const payload = parsed.payload || parsed;
            if (!version && payload && payload.cacheVersion) version = String(payload.cacheVersion);
            if (!updatedAt && payload && payload.generatedAt) updatedAt = String(payload.generatedAt);
            if (!updatedAt && payload && payload.updatedAt) updatedAt = String(payload.updatedAt);
          }
        }
      } catch (cacheErr) {
        console.warn('getVehicleReleasedCacheVersion cache parse failed:', cacheErr);
      }
    }

    let props = null;
    try {
      props = PropertiesService.getScriptProperties();
    } catch (_propErr) {
      props = null;
    }

    if (props) {
      if (!version || !updatedAt) {
        const stored = props.getProperty(VEHICLE_RELEASED_PROP_KEY);
        if (stored) {
          try {
            const parsed = JSON.parse(stored);
            if (parsed) {
              if (!version && parsed.version) version = String(parsed.version);
              const payload = parsed.payload || parsed;
              if (!version && payload && payload.cacheVersion) version = String(payload.cacheVersion);
              if (!updatedAt && payload && payload.generatedAt) updatedAt = String(payload.generatedAt);
              if (!updatedAt && payload && payload.updatedAt) updatedAt = String(payload.updatedAt);
            }
          } catch (propParseErr) {
            console.warn('getVehicleReleasedCacheVersion properties parse failed:', propParseErr);
          }
        }
      }
      if (!version) {
        try {
          const propVersion = props.getProperty(VEHICLE_RELEASED_VERSION_PROP_KEY);
          if (propVersion) version = String(propVersion);
        } catch (_versionReadErr) {
          // ignore
        }
      }
    }

    return { ok: true, version: version, updatedAt: updatedAt };
  } catch (err) {
    console.error('getVehicleReleasedCacheVersion failed:', err);
    return { ok: false, error: String(err) };
  }
}

/**
 * Get team vehicle usage summary
 */
function getTeamSummary() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const teamCol = headers.indexOf('Team');
    const starsCol = headers.indexOf('Stars');
    
    if (teamCol === -1) {
      return 'Unable to find team column in the data.';
    }
    
    const teamStats = {};
    
    for (let i = 1; i < data.length; i++) {
      const team = data[i][teamCol] || 'Unknown';
      const stars = parseFloat(data[i][starsCol]) || 0;
      
      if (!teamStats[team]) {
        teamStats[team] = { count: 0, totalStars: 0, avgRating: 0 };
      }
      
      teamStats[team].count++;
      teamStats[team].totalStars += stars;
      teamStats[team].avgRating = teamStats[team].totalStars / teamStats[team].count;
    }
    
    let response = 'Team Vehicle Usage Summary:\n';
    const sortedTeams = Object.entries(teamStats)
      .sort(([,a], [,b]) => b.count - a.count)
      .slice(0, 5);
    
    sortedTeams.forEach(([team, stats]) => {
      const avgRating = stats.avgRating > 0 ? ` (Avg rating: ${stats.avgRating.toFixed(1)}★)` : '';
      response += `\n• ${team}: ${stats.count} vehicle releases${avgRating}`;
    });
    
    return response;
    
  } catch (error) {
    console.error('getTeamSummary error:', error);
    return 'Sorry, I could not retrieve team summary data.';
  }
}

/**
 * Get vehicle information
 */
function getBeneficiaryInfo(query) {
  try {
    // Extract potential vehicle number from query
    const words = query.split(' ');
    let vehicleNumber = '';
    
    // Look for vehicle numbers after common question words
    const questionWords = ['what', 'about', 'vehicle', 'car', 'who', 'using'];
    for (let i = 0; i < words.length; i++) {
      if (questionWords.includes(words[i].toLowerCase()) && i + 1 < words.length) {
        vehicleNumber = words.slice(i + 1).join(' ');
        break;
      }
    }
    
    if (!vehicleNumber) {
      return formatVehicleMissingId();
    }
    
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const findCol = (names) => {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const carCol = findCol(['Vehicle Number','Car Number']);
    const projectCol = findCol(['Project']);
    const teamCol = findCol(['Team']);
    const statusCol = findCol(['Status','In Use/Release','In Use / release']);
    const starsCol = findCol(['Stars','Ratings']);
    const remarksCol = findCol(['Last Users remarks','Remarks']);
    const responsibleCol = findCol(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible','R. Ben','R Ben']);

    if (carCol === -1) {
      return 'Unable to find vehicle column in the data.';
    }
    
    // Find the vehicle
    for (let i = 1; i < data.length; i++) {
      const carNumber = (data[i][carCol] || '').toString().toLowerCase();
      if (carNumber.includes(vehicleNumber.toLowerCase())) {
        const project = data[i][projectCol] || 'Unknown';
        const team = data[i][teamCol] || 'Unknown';
        const status = statusCol >= 0 ? (data[i][statusCol] || 'Unknown') : 'Unknown';
        const stars = starsCol >= 0 ? (data[i][starsCol] || 'No rating') : 'No rating';
        const remarks = remarksCol >= 0 ? (data[i][remarksCol] || 'No remarks') : 'No remarks';
        const responsible = responsibleCol >= 0 ? (data[i][responsibleCol] || 'Not set') : 'Not set';

  return formatVehicleInfo({ car: data[i][carCol], project: project, team: team, status: status, stars: stars, remarks: remarks, responsibleBeneficiary: responsible });
      }
    }
    
  return formatVehicleNotFound(vehicleNumber);
    
  } catch (error) {
    console.error('getBeneficiaryInfo error:', error);
    return 'Sorry, I could not retrieve vehicle information.';
  }
}

/**
 * Get vehicle makes and models summary
 */
function getExpenseCategories() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const makeCol = headers.indexOf('Make');
    const modelCol = headers.indexOf('Model');
    
    if (makeCol === -1) {
      return 'Unable to find vehicle make column in the data.';
    }
    
    const vehicleStats = {};
    
    for (let i = 1; i < data.length; i++) {
      const make = data[i][makeCol] || 'Unknown Make';
      const model = data[i][modelCol] || 'Unknown Model';
      const vehicle = `${make} ${model}`.trim();
      
      if (!vehicleStats[vehicle]) {
        vehicleStats[vehicle] = { count: 0 };
      }
      
      vehicleStats[vehicle].count++;
    }
    
    let response = 'Vehicle Makes & Models:\n';
    const sortedVehicles = Object.entries(vehicleStats)
      .sort(([,a], [,b]) => b.count - a.count)
      .slice(0, 5);
    
    sortedVehicles.forEach(([vehicle, stats]) => {
      response += `\n• ${vehicle}: ${stats.count} vehicles`;
    });
    
    return response;
    
  } catch (error) {
    console.error('getExpenseCategories error:', error);
    return 'Sorry, I could not retrieve vehicle data.';
  }
}

/**
 * Get highest rated vehicles
 */
function getHighAmountRequests() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const findCol = (names) => {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const starsCol = findCol(['Stars','Ratings']);
    const carCol = findCol(['Vehicle Number','Car Number']);
    const projectCol = findCol(['Project']);
    const remarksCol = findCol(['Last Users remarks','Remarks']);
    
    if (starsCol === -1 || carCol === -1) {
      return 'Unable to find required columns for rating analysis.';
    }
    
    const vehicles = [];
    
    for (let i = 1; i < data.length; i++) {
      const stars = parseFloat(data[i][starsCol]) || 0;
      if (stars > 0) {
        vehicles.push({
          car: carCol >= 0 ? (data[i][carCol] || 'Unknown') : 'Unknown',
          project: projectCol >= 0 ? (data[i][projectCol] || 'Unknown') : 'Unknown',
          stars: stars,
          remarks: remarksCol >= 0 ? (data[i][remarksCol] || 'No remarks') : 'No remarks'
        });
      }
    }
    
    // Sort by rating descending and take top 5
    const topVehicles = vehicles
      .sort((a, b) => b.stars - a.stars)
      .slice(0, 5);
    
    let response = 'Highest Rated Vehicles:\n';
    topVehicles.forEach((vehicle, index) => {
      response += `\n${index + 1}. ${vehicle.car} (${vehicle.project}): ${vehicle.stars}★`;
    });
    
    return response;
    
  } catch (error) {
    console.error('getHighAmountRequests error:', error);
    return 'Sorry, I could not retrieve vehicle rating data.';
  }
}

/**
 * Get recent vehicle releases
 */
function getRecentRequests() {
  try {
    const sheet = SpreadsheetApp.openById(CAR_SHEET_ID).getSheetByName('CarT_P');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const findCol = (names) => {
      for (let i = 0; i < names.length; i++) {
        const idx = headers.indexOf(names[i]);
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const timestampCol = findCol(['Date and time of entry','Timestamp','Date']);
    const carCol = findCol(['Vehicle Number','Car Number']);
    const projectCol = findCol(['Project']);
    const starsCol = findCol(['Stars','Ratings']);
    
    if (timestampCol === -1 || carCol === -1) {
      return 'Unable to find required columns for recent vehicle releases.';
    }
    
    const releases = [];
    
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][timestampCol]);
      if (!isNaN(timestamp.getTime())) {
        releases.push({
          car: carCol >= 0 ? (data[i][carCol] || 'Unknown') : 'Unknown',
          project: projectCol >= 0 ? (data[i][projectCol] || 'Unknown') : 'Unknown',
          stars: starsCol >= 0 ? (parseFloat(data[i][starsCol]) || 0) : 0,
          timestamp: timestamp
        });
      }
    }
    
    // Sort by timestamp descending and take top 5
    const recentReleases = releases
      .sort((a, b) => b.timestamp - a.timestamp)
      .slice(0, 5);
    
    let response = 'Recent Vehicle Releases:\n';
    recentReleases.forEach((release, index) => {
      const dateStr = release.timestamp.toLocaleDateString();
      const rating = release.stars > 0 ? ` (${release.stars}★)` : '';
      response += `\n${index + 1}. ${release.car} - ${release.project}${rating} (${dateStr})`;
    });
    
    return response;
    
  } catch (error) {
    console.error('getRecentRequests error:', error);
    return 'Sorry, I could not retrieve recent vehicle release data.';
  }
}

/**
 * Default chat response with suggestions
 */
function getDefaultChatResponse() {
  return formatDefaultChatResponse();
}

// Helper function to get sheet by name
function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    // If submissions sheet doesn't exist, create it
    if (!sheet && sheetName === 'submissions') {
      sheet = ss.insertSheet(SUBMISSIONS_SHEET_NAME);
      // Add headers for submissions sheet
      const headers = [
        'Timestamp', 'Email', 'Beneficiary', 'Project', 'Account Holder', 'Team', 'Designation',
        'Total Amount', 'Fuel', 'DA', 'Vehicle Rent', 'Airtime', 'Transport', 'Miscellaneous'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    return sheet;
  } catch (error) {
    console.error(`Error accessing sheet '${sheetName}':`, error);
    return null;
  }
}

// Add sample submission data if sheet is empty
function addSampleSubmissionData() {
  try {
    // Gate seeding behind script property SEED_SAMPLES=true|1|yes
    try {
      const sp = PropertiesService.getScriptProperties();
      const flag = (sp.getProperty('SEED_SAMPLES') || '').toLowerCase();
      const enabled = (flag === 'true' || flag === '1' || flag === 'yes');
      if (!enabled) return; // do not seed unless explicitly enabled
    } catch(_e) { /* ignore, default: no seeding */ }

    const sheet = getSheet('submissions');
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { // Only headers or empty
      const sampleData = [
        [new Date(), 'john@example.com', 'John Doe', 'Project Alpha', 'John Doe', 'Team A', 'Manager', 1500, 300, 400, 200, 100, 300, 200],
        [new Date(), 'jane@example.com', 'Jane Smith', 'Project Beta', 'Jane Smith', 'Team B', 'Coordinator', 1200, 250, 350, 150, 80, 250, 117],
        [new Date(), 'mike@example.com', 'Mike Johnson', 'Project Gamma', 'Mike Johnson', 'Team A', 'Assistant', 800, 200, 200, 100, 50, 150, 100],
        [new Date(Date.now() - 86400000), 'sarah@example.com', 'Sarah Wilson', 'Project Alpha', 'Sarah Wilson', 'Team C', 'Analyst', 950, 180, 300, 120, 70, 200, 80],
        [new Date(Date.now() - 172800000), 'david@example.com', 'David Brown', 'Project Beta', 'David Brown', 'Team B', 'Developer', 1100, 220, 280, 180, 90, 220, 110]
      ];
      
      sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    }
  } catch (error) {
    console.error('Error adding sample submission data:', error);
  }
}

// Submissions Sheet Query Functions
function getMonthlySubmissionExpenses() {
  try {
    // Ensure sample data exists
    addSampleSubmissionData();
    
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    
    let totalExpenses = 0;
    let monthlyData = {
      fuel: 0,
      erda: 0,
      vehicleRent: 0,
      airtime: 0,
      transport: 0,
      misc: 0,
      count: 0
    };
    
    // Build index map from header row (use exact header names; fallback to known positions)
    const hdr = headers.map(h=> String(h||'').trim());
    const find = (name, fallback)=>{
      const i = hdr.indexOf(name);
      return i >= 0 ? i : (typeof fallback === 'number' ? fallback : -1);
    };
    // Fallback positions (0-based) from AI Language.md: Total G=6, Fuel K=10, ER DA N=13, Vehicle R=17, Airtime U=20, Transport X=23, Misc AA=26
    const IDX = {
      ts: find('Timestamp', 0),
      total: find('Total Expense', 6),
      fuel: find('Fuel Amount', 10),
      erda: find('ER DA Amount', 13),
      vehicleRent: find('Vehicle Rent Amount', 17),
      airtime: find('Airtime Amount', 20),
      transport: find('Transport Amount', 23),
      misc: find('Misc Amount', 26)
    };
  function toNum(v){ return parseAmount(v); }
    rows.forEach(row => {
      const timestamp = new Date(row[IDX.ts]);
      if (timestamp.getMonth() === currentMonth && timestamp.getFullYear() === currentYear) {
        monthlyData.fuel += toNum(row[IDX.fuel]);
        monthlyData.erda += toNum(row[IDX.erda]);
        monthlyData.vehicleRent += toNum(row[IDX.vehicleRent]);
        monthlyData.airtime += toNum(row[IDX.airtime]);
        monthlyData.transport += toNum(row[IDX.transport]);
        monthlyData.misc += toNum(row[IDX.misc]);
        totalExpenses += toNum(row[IDX.total]);
        monthlyData.count++;
      }
    });
    
    // Delegated to responses.gs
    return formatMonthlyExpenseSummary(currentYear, totalExpenses, monthlyData);
  } catch (error) {
    console.error('Error getting monthly submission expenses:', error);
    return 'Error retrieving monthly expense data.';
  }
}

function getBeneficiaryExpenses() {
  try {
    addSampleSubmissionData();
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const beneficiaryTotals = {};
    
    // Build header indices from the sheet headers
    const headers = data[0] || [];
    const hdr = headers.map(h=> String(h||'').trim());
    const find = (name, fallback)=>{ const i = hdr.indexOf(name); return i>=0? i : (typeof fallback === 'number' ? fallback : -1); };
    const IDX = {
      beneficiary: find('Beneficiary', 2),
      total: find('Total Expense', 6),
      fuel: find('Fuel Amount', 10),
      erda: find('ER DA Amount', 13),
      vehicleRent: find('Vehicle Rent Amount', 17),
      airtime: find('Airtime Amount', 20),
      transport: find('Transport Amount', 23),
      misc: find('Misc Amount', 26)
    };
  function toNum(v){ return parseAmount(v); }
    rows.forEach(row => {
      const beneficiary = row[IDX.beneficiary] || 'Unknown';
      let total = toNum(row[IDX.total]);
      if (total === 0){ // fallback: sum category columns
        total = toNum(row[IDX.fuel]) + toNum(row[IDX.erda]) + toNum(row[IDX.vehicleRent]) + toNum(row[IDX.airtime]) + toNum(row[IDX.transport]) + toNum(row[IDX.misc]);
      }
      if (!beneficiaryTotals[beneficiary]) {
        beneficiaryTotals[beneficiary] = { total: 0, count: 0 };
      }
      beneficiaryTotals[beneficiary].total += total;
      beneficiaryTotals[beneficiary].count++;
    });
    
    const sorted = Object.entries(beneficiaryTotals)
      .sort(([,a], [,b]) => b.total - a.total)
      .slice(0, 10);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = '👥 **Top Beneficiaries by Expense Amount:**\n\n';
    sorted.forEach(([name, data], index) => {
      result += `${index + 1}. **${name}**\n   💰 ${formatCurrency(data.total)} (${data.count} submissions)\n\n`;
    });
    
    return result;
  } catch (error) {
    console.error('Error getting beneficiary expenses:', error);
    return 'Error retrieving beneficiary expense data.';
  }
}

function getTeamExpenseBreakdown() {
  try {
    addSampleSubmissionData();
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const teamTotals = {};
    
    const hdr2 = (data[0]||[]).map(h=>String(h||'').trim());
    const find2 = (name, fb)=>{ const i = hdr2.indexOf(name); return i>=0? i : (typeof fb === 'number'? fb : -1); };
    const IDX2 = {
      team: find2('Team', 5),
      total: find2('Total Expense', 6),
      fuel: find2('Fuel Amount', 10),
      erda: find2('ER DA Amount', 13),
      vehicleRent: find2('Vehicle Rent Amount', 17),
      airtime: find2('Airtime Amount', 20),
      transport: find2('Transport Amount', 23),
      misc: find2('Misc Amount', 26)
    };
    rows.forEach(row => {
      const team = row[IDX2.team] || 'Unknown Team';
      let total = toNum(row[IDX2.total]);
      if (total === 0){ total = toNum(row[IDX2.fuel]) + toNum(row[IDX2.erda]) + toNum(row[IDX2.vehicleRent]) + toNum(row[IDX2.airtime]) + toNum(row[IDX2.transport]) + toNum(row[IDX2.misc]); }
      if (!teamTotals[team]) { teamTotals[team] = { total: 0, count: 0 }; }
      teamTotals[team].total += total;
      teamTotals[team].count++;
    });
    
    const sorted = Object.entries(teamTotals)
      .sort(([,a], [,b]) => b.total - a.total);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = '🏢 **Team Expense Breakdown:**\n\n';
    sorted.forEach(([team, data], index) => {
      result += `${index + 1}. **${team}**\n   💰 ${formatCurrency(data.total)} (${data.count} submissions)\n\n`;
    });
    
    return result;
  } catch (error) {
    console.error('Error getting team expense breakdown:', error);
    return 'Error retrieving team expense data.';
  }
}

function getFuelExpenses() {
  try {
    addSampleSubmissionData();
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    let totalFuel = 0;
    let fuelEntries = 0;
    const fuelByBeneficiary = {};
    
    rows.forEach(row => {
  const hdr = (data[0]||[]).map(h=>String(h||'').trim());
  const find = (name, fb)=>{ const i = hdr.indexOf(name); return i>=0? i : (typeof fb==='number'? fb : -1); };
  const IDX = { fuel: find('Fuel Amount', 10) };
  function toNum(v){ return parseAmount(v); }
  const fuel = toNum(row[IDX.fuel]);
      if (fuel > 0) {
        totalFuel += fuel;
        fuelEntries++;
        
        const beneficiary = row[2] || 'Unknown';
        if (!fuelByBeneficiary[beneficiary]) {
          fuelByBeneficiary[beneficiary] = 0;
        }
        fuelByBeneficiary[beneficiary] += fuel;
      }
    });
    
    const sorted = Object.entries(fuelByBeneficiary)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = `⛽ **Fuel Expense Summary:**\n\n`;
    result += `💰 **Total Fuel Costs:** ${formatCurrency(totalFuel)}\n`;
    result += `📝 **Number of Fuel Entries:** ${fuelEntries}\n\n`;
    result += `**Top Fuel Users:**\n`;
    
    sorted.forEach(([name, amount], index) => {
      result += `${index + 1}. ${name}: ${formatCurrency(amount)}\n`;
    });
    
    return result;
  } catch (error) {
    console.error('Error getting fuel expenses:', error);
    return 'Error retrieving fuel expense data.';
  }
}

function getTransportExpenses() {
  // Implementation moved to `responses.gs` to centralize response formatting.
  // See responses.gs:getTransportExpenses for the full implementation.
}

function getExpenseSummary() {
  // Implementation moved to `responses.gs` to centralize response formatting.
  // See responses.gs:getExpenseSummary for the full implementation.
}

/* -------------------------- LLM admin menu -------------------------- */

function onOpen(){
  // Add a lightweight LLM menu in the Spreadsheet UI (if available)
  try{
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('LLM')
      .addItem('Show Config', 'LLM_showConfig')
      .addSeparator()
      .addItem('Ping (Auto Provider)', 'LLM_pingAuto')
      .addItem('Ping OpenRouter', 'LLM_pingOpenRouter')
      .addItem('Ping DeepSeek', 'LLM_pingDeepSeek')
      .addSeparator()
      .addItem('Toggle Fallback (CHAT_USE_LLM)', 'LLM_toggleFallback')
      .addToUi();

    // RAG control menu
    ui.createMenu('RAG')
      .addItem('Index Folder', 'RAG_menuIndexFolder')
      .addItem('Refresh Index (auto TTL)', 'RAG_menuRefresh')
      .addSeparator()
      .addItem('Clear Index', 'RAG_menuClearIndex')
      .addToUi();
  }catch(e){
    // Not a spreadsheet-bound project or UI not available; ignore.
  }
}

function LLM_showConfig(){
  var summary = (typeof getLLMConfigSummary === 'function') ? getLLMConfigSummary() : { error:'getLLMConfigSummary not available' };
  var json = JSON.stringify(summary, null, 2);
  _LLM_showDialog_('LLM Config', json);
}

function LLM_pingAuto(){
  var res;
  try{
    var prov = (typeof _llmProvider_ === 'function') ? _llmProvider_() : 'none';
    if (prov === 'deepseek' && typeof debugDeepSeekPing === 'function') res = debugDeepSeekPing();
    else if (prov === 'openrouter' && typeof debugOpenRouterPing === 'function') res = debugOpenRouterPing();
    else res = { ok:false, error:'No LLM configured' };
  }catch(e){ res = { ok:false, error:String(e) }; }
  _LLM_showDialog_('LLM Ping (Auto)', JSON.stringify(res, null, 2));
}

function LLM_pingOpenRouter(){
  var res = (typeof debugOpenRouterPing === 'function') ? debugOpenRouterPing() : { ok:false, error:'debugOpenRouterPing not available' };
  _LLM_showDialog_('LLM Ping (OpenRouter)', JSON.stringify(res, null, 2));
}

function LLM_pingDeepSeek(){
  var res = (typeof debugDeepSeekPing === 'function') ? debugDeepSeekPing() : { ok:false, error:'debugDeepSeekPing not available' };
  _LLM_showDialog_('LLM Ping (DeepSeek)', JSON.stringify(res, null, 2));
}

function LLM_toggleFallback(){
  try{
    var sp = PropertiesService.getScriptProperties();
    var cur = (sp.getProperty('CHAT_USE_LLM') || '').toLowerCase();
    var next = (cur === 'true' || cur === '1' || cur === 'yes') ? 'false' : 'true';
    sp.setProperty('CHAT_USE_LLM', next);
    _LLM_showDialog_('LLM Fallback', 'CHAT_USE_LLM set to: ' + next);
  }catch(e){
    _LLM_showDialog_('LLM Fallback', 'Error: ' + String(e));
  }
}

function _LLM_showDialog_(title, body){
  try{
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutput('<div style="white-space:pre-wrap;font:12px/1.4 -apple-system,Segoe UI,Roboto,Arial,sans-serif">'+ _escHtml(body) +'</div>')
      .setWidth(600).setHeight(420);
    ui.showModalDialog(html, title);
  }catch(_){
    // UI not available; log for developers
    Logger.log(title + ':\n' + body);
  }
}

// Helper to run the debug helper for Mompati and log the result for easy copy-paste
function runDebugForMompati(){
  try{
    const res = debugBeneficiaryMetricTable('Mompati','erda','01-Sep-2025','30-Sep-2025');
    try{ Logger.log(JSON.stringify(res, null, 2)); }catch(_){ Logger.log(String(res)); }
    return res;
  }catch(e){ Logger.log('runDebugForMompati error: ' + String(e)); return { ok:false, error:String(e) }; }
}

// Helper to run the natural-language query and request a table result for DA
function runAskForDA_table(){
  try{
    const q = 'DA FROM 01-Sep-2025 TO 30-Sep-2025 Mompati output:table';
    const res = askQuestion(q);
    try{ Logger.log(JSON.stringify(res, null, 2)); }catch(_){ Logger.log(String(res)); }
    return res;
  }catch(e){ Logger.log('runAskForDA_table error: ' + String(e)); return { ok:false, error:String(e) }; }
}

// Helper to run the debug helper for Vehicle Rent and log the result
function runDebugForVehicleRent(){
  try{
    const res = debugBeneficiaryMetricTable('Mompati','vehicleRent','01-Sep-2025','30-Sep-2025');
    try{ Logger.log(JSON.stringify(res, null, 2)); }catch(_){ Logger.log(String(res)); }
    return res;
  }catch(e){ Logger.log('runDebugForVehicleRent error: ' + String(e)); return { ok:false, error:String(e) }; }
}

// Helper to run the natural-language query and request a table result for Vehicle Rent
function runAskForVehicleRent_table(){
  try{
    const q = 'Vehicle Rent FROM 01-Sep-2025 TO 30-Sep-2025 Mompati output:table';
    const res = askQuestion(q);
    try{ Logger.log(JSON.stringify(res, null, 2)); }catch(_){ Logger.log(String(res)); }
    return res;
  }catch(e){ Logger.log('runAskForVehicleRent_table error: ' + String(e)); return { ok:false, error:String(e) }; }
}

function _escHtml(s){
  try{
    return String(s||'').replace(/[&<>"{]/g, function(c){
      return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','{':'&#123;'}[c]) || c;
    });
  }catch(_){
    return String(s||'');
  }
}

// ---------------- RAG menu handlers ----------------
function RAG_menuIndexFolder(){
  try{
    RAG_ensureDefaultFolderConfigured();
    var sp = PropertiesService.getScriptProperties();
    var folderId = (sp.getProperty('RAG_FOLDER_ID') || RAG_DEFAULT_FOLDER_ID || '').trim();
    if (!folderId) return _LLM_showDialog_('RAG Index', 'No folder configured. Set RAG_FOLDER_ID first.');
    var res = RAG_indexFolder(folderId);
    _LLM_showDialog_('RAG Index', 'Folder indexed.\n' + res);
  }catch(e){ _LLM_showDialog_('RAG Index', 'Error: ' + String(e)); }
}

function RAG_menuClearIndex(){
  try{
    var msg = RAG_clearIndex();
    _LLM_showDialog_('RAG Clear Index', msg);
  }catch(e){ _LLM_showDialog_('RAG Clear Index', 'Error: ' + String(e)); }
}

function RAG_menuRefresh(){
  try{
    RAG_ensureDefaultFolderConfigured();
    var sp = PropertiesService.getScriptProperties();
    var ttl = Number(sp.getProperty('RAG_INDEX_TTL_MIN') || '240');
    var res = RAG_refreshIndexIfStale(ttl);
    _LLM_showDialog_('RAG Refresh', JSON.stringify(res, null, 2));
  }catch(e){ _LLM_showDialog_('RAG Refresh', 'Error: ' + String(e)); }
}
