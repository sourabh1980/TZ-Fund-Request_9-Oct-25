/**
 * Gets vehicle assignment data from Vehicle_InUse tab
 * Returns latest IN USE vehicle mapped per beneficiary so UI can prefill rows.
 */
function getVehicleInUseData() {
  const CACHE_KEY = 'vehicle_in_use_payload_v3';
  const PROP_KEY = 'vehicle_in_use_payload_v3_json';
  try {
    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    let props = null;
    try {
      props = PropertiesService.getScriptProperties();
    } catch (_propErr) {
      props = null;
    }
    if (cache) {
      const cached = cache.get(CACHE_KEY);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          if (parsed && parsed.ok) {
            parsed.cached = true;
            return parsed;
          }
        } catch (_err) {
          cache.remove(CACHE_KEY);
        }
      }
    }

    _maybeAutoRefreshCarTPSummaries_(5);

    if (props) {
      const stored = props.getProperty(PROP_KEY);
      if (stored) {
        try {
          const parsed = JSON.parse(stored);
          if (parsed && parsed.ok) {
            parsed.cached = true;
            parsed.fromProperties = true;
            return parsed;
          }
        } catch (_propParseErr) {
          try { props.deleteProperty(PROP_KEY); } catch (_){ /* ignore */ }
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
      const emptyPayload = { ok: true, source: 'Vehicle_InUse', assignments: [], updatedAt: '', message: 'No IN USE assignments found' };
      if (cache) cache.put(CACHE_KEY, JSON.stringify(emptyPayload), 15);
      if (props) {
        try { props.setProperty(PROP_KEY, JSON.stringify(emptyPayload)); } catch (_err) { /* ignore */ }
      }
      return emptyPayload;
    }

    const lastCol = sheet.getLastColumn();
    const header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(header);
    const idx = (labels, required) => {
      try {
        return IX.get(labels);
      } catch (err) {
        if (required) throw err;
        return -1;
      }
    };

    const iBeneficiary = idx(['R.Beneficiary', 'Responsible Beneficiary', 'Beneficiary', 'Name'], true);
    const iVehicle = idx(['Vehicle Number', 'Car Number', 'Vehicle'], true);
    const iProject = idx(['Project', 'Project Name'], false);
    const iTeam = idx(['Team', 'Team Name'], false);
    const iStatus = idx(['Status', 'In Use/Release', 'In Use'], false);
    const iDate = idx(['Date and time of entry', 'Date and time', 'Timestamp', 'Date'], false);
    const iMake = idx(['Make', 'Car Make', 'Brand'], false);
    const iModel = idx(['Model', 'Car Model'], false);
    const iCategory = idx(['Category', 'Vehicle Category'], false);
    const iUsage = idx(['Usage Type', 'Usage'], false);
    const iOwner = idx(['Owner', 'Owner Name'], false);
    const iRemarks = idx(['Last Users remarks', 'Remarks', 'Feedback'], false);
    const iRatings = idx(['Ratings', 'Stars', 'Rating'], false);
    const iRef = idx(['Ref', 'Reference Number'], false);
    const iSubmitter = idx(['Submitter username', 'Submitter', 'User'], false);

    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();

    const assignments = [];
    let newestTs = 0;

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const disp = displayValues[r];

      const beneficiary = String((row[iBeneficiary] != null ? row[iBeneficiary] : disp[iBeneficiary]) || '').trim();
      const vehicleNumber = String((row[iVehicle] != null ? row[iVehicle] : disp[iVehicle]) || '').trim();
      if (!beneficiary || !vehicleNumber) {
        continue;
      }

      const rawStatus = iStatus >= 0 ? (row[iStatus] || disp[iStatus] || '') : 'IN USE';
      const status = _normStatus_(rawStatus) || 'IN USE';
      if (status !== 'IN USE') {
        continue;
      }

      const rawDate = iDate >= 0 ? (row[iDate] || disp[iDate] || '') : '';
      const ts = _parseTs_(rawDate);
      if (ts > newestTs) {
        newestTs = ts;
      }

      assignments.push({
        beneficiary: beneficiary,
        responsibleBeneficiary: beneficiary,
        vehicleNumber: vehicleNumber,
        project: iProject >= 0 ? (row[iProject] || disp[iProject] || '') : '',
        team: iTeam >= 0 ? (row[iTeam] || disp[iTeam] || '') : '',
        status: status,
        entryDate: rawDate,
        entryTimestamp: ts,
        rowNumber: r + 2,
        sheet: 'Vehicle_InUse',
        ref: iRef >= 0 ? (row[iRef] || disp[iRef] || '') : '',
        make: iMake >= 0 ? (row[iMake] || disp[iMake] || '') : '',
        model: iModel >= 0 ? (row[iModel] || disp[iModel] || '') : '',
        category: iCategory >= 0 ? (row[iCategory] || disp[iCategory] || '') : '',
        usageType: iUsage >= 0 ? (row[iUsage] || disp[iUsage] || '') : '',
        owner: iOwner >= 0 ? (row[iOwner] || disp[iOwner] || '') : '',
        remarks: iRemarks >= 0 ? (row[iRemarks] || disp[iRemarks] || '') : '',
        ratings: iRatings >= 0 ? (row[iRatings] || disp[iRatings] || '') : '',
        submitter: iSubmitter >= 0 ? (row[iSubmitter] || disp[iSubmitter] || '') : ''
      });
    }

    assignments.sort(function(a, b) {
      const aTs = typeof a.entryTimestamp === 'number' ? a.entryTimestamp : 0;
      const bTs = typeof b.entryTimestamp === 'number' ? b.entryTimestamp : 0;
      if (aTs !== bTs) return bTs - aTs;
      return (b.rowNumber || 0) - (a.rowNumber || 0);
    });

    const updatedAt = newestTs > 0 ? new Date(newestTs).toISOString() : new Date().toISOString();
    const payload = {
      ok: true,
      source: 'Vehicle_InUse',
      assignments: assignments,
      updatedAt: updatedAt,
      generatedAt: updatedAt
    };

    if (cache) {
      try { cache.put(CACHE_KEY, JSON.stringify(payload), 15); } catch (_err) { /* ignore */ }
    }
    if (props) {
      try { props.setProperty(PROP_KEY, JSON.stringify(payload)); } catch (_err) { /* ignore */ }
    }

    return payload;
  } catch (error) {
    console.error('[BACKEND] getVehicleInUseData failed:', error);
    return { ok: false, source: 'Vehicle_InUse', error: String(error) };
  }
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

    const values = sh.getRange(1, 2, lastRow, columnCount).getDisplayValues();
    const header = values[0];
    const dataRows = values.slice(1);

    const CAR_INDEX = 4;   // Column F (Vehicle Number) relative to range starting at B
    const TEAM_INDEX = 2;  // Column D (Team)
    const STATUS_INDEX = 10; // Column L (Status)

    const targetCar = normalizedCar.toLowerCase();
    const filteredRows = dataRows.filter(row => {
      const carValue = String(row[CAR_INDEX] || '').trim().toLowerCase();
      const teamValue = String(row[TEAM_INDEX] || '').trim().toLowerCase();
      const statusValue = String(row[STATUS_INDEX] || '').trim().toUpperCase();
      const carMatch = carValue === targetCar;
      const teamMatch = normalizedTeam ? teamValue === normalizedTeam : true;
      const statusMatch = statusValue === 'IN USE';
      return carMatch && teamMatch && statusMatch;
    });

    console.log(`Filtered ${filteredRows.length} rows for car ${normalizedCar} and team ${normalizedTeam || '(any)'}`);

    if (filteredRows.length === 0) {
      const teamMessage = normalizedTeam ? ` and team: ${teamName}` : '';
      return [['Info', `No IN USE history found for vehicle: ${normalizedCar}${teamMessage}`]];
    }

    const result = [header].concat(filteredRows);
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
    
    if (!carNumber) {
      console.log('No car number provided');
      return [['Error', 'No car number provided']];
    }
    
    // Try to access CarT_P sheet
    const sh = _openCarTP_();
    if (!sh) {
      console.log('CarT_P sheet not found - returning info message');
      return [['Info', 'CarT_P sheet not accessible - please check sheet permissions']];
    }
    
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) {
      console.log('CarT_P sheet has no data');
      return [['Info', 'CarT_P sheet contains no data']];
    }
    
    const lastCol = sh.getLastColumn();
    const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const header = data[0];
    
    // Find column indices
    const carIdx = header.findIndex(h => 
      String(h || '').toLowerCase().includes('vehicle') || 
      String(h || '').toLowerCase().includes('car')
    );
    const statusIdx = header.findIndex(h => 
      String(h || '').toLowerCase().includes('status') ||
      String(h || '').toLowerCase().includes('release') ||
      String(h || '').toLowerCase().includes('use')
    );
    
    if (carIdx === -1) {
      console.log('Vehicle column not found in CarT_P sheet');
      return [['Error', 'Vehicle Number column not found in CarT_P sheet']];
    }
    
    // Filter for RELEASE status and matching car number
    const filteredRows = data.slice(1).filter(row => {
      const rowCarNumber = String(row[carIdx] || '').trim();
      const rowStatus = String(row[statusIdx] || '').trim().toUpperCase();
      return rowCarNumber === carNumber && rowStatus === 'RELEASE';
    });
    
    console.log(`Found ${filteredRows.length} RELEASE entries for car ${carNumber}`);
    
    if (filteredRows.length === 0) {
      return [['Info', `No RELEASE history found for vehicle: ${carNumber}`]];
    }
    
    // Return header + filtered rows
    const result = [header].concat(filteredRows);
    console.log('Returning real RELEASE history result with', result.length, 'rows');
    return result;
    
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
function debugCarTPData() {
  try {
    console.log('=== Debugging CarT_P Data ===');
    
    // Check if CarT_P sheet exists
    const sh = _openCarTP_();
    if (!sh) {
      console.log('CarT_P sheet not found');
      return { ok: false, error: 'CarT_P sheet not found' };
    }
    
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    console.log(`CarT_P sheet: ${lastRow} rows, ${lastCol} columns`);
    
    if (lastRow <= 1) {
      console.log('CarT_P sheet has no data rows');
      return { ok: false, error: 'CarT_P sheet has no data' };
    }
    
    // Get header and sample data
    const data = sh.getRange(1, 1, Math.min(lastRow, 5), lastCol).getValues();
    const header = data[0];
    console.log('Header:', header);
    
    // Find car and status columns
    const carIdx = header.findIndex(h => 
      String(h || '').toLowerCase().includes('vehicle') || 
      String(h || '').toLowerCase().includes('car')
    );
    const statusIdx = header.findIndex(h => 
      String(h || '').toLowerCase().includes('status') ||
      String(h || '').toLowerCase().includes('release') ||
      String(h || '').toLowerCase().includes('use')
    );
    
    console.log(`Car column index: ${carIdx}, Status column index: ${statusIdx}`);
    
    // Show sample rows
    const sampleRows = data.slice(1);
    console.log('Sample rows:', sampleRows);
    
    // Test with first car found
    let testCarNumber = null;
    for (let i = 1; i < data.length; i++) {
      const carNumber = String(data[i][carIdx] || '').trim();
      if (carNumber) {
        testCarNumber = carNumber;
        break;
      }
    }
    
    console.log(`Testing with car number: ${testCarNumber}`);
    
    let releaseHistory = [];
    let inUseHistory = [];
    
    if (testCarNumber) {
      releaseHistory = getVehicleReleaseHistory(testCarNumber);
      inUseHistory = getInUseHistoryForcedWorking(testCarNumber, '');
    }
    
    return {
      ok: true,
      sheetInfo: {
        rows: lastRow,
        cols: lastCol,
        carColumnIndex: carIdx,
        statusColumnIndex: statusIdx
      },
      header: header,
      sampleData: sampleRows,
      testCarNumber: testCarNumber,
      releaseHistoryLength: releaseHistory.length,
      inUseHistoryLength: inUseHistory.length,
      releaseHistory: releaseHistory,
      inUseHistory: inUseHistory
    };
    
  } catch (error) {
    console.error('Error in debugCarTPData:', error);
    return { ok: false, error: String(error) };
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
function refreshVehicleStatusSheets() {
  const allCarRows = _readCarTP_objects_();
  if (!allCarRows.length) {
    console.log('No CarT_P data found.');
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
    if (names.length) {
      names.forEach(function(name){
        const key = name.toLowerCase();
        const prev = latestByBeneficiary.get(key);
        if (!prev || row._ts > prev._ts || (row._ts === prev._ts && row._rowIndex >= prev._rowIndex)) {
          const clone = Object.assign({}, row);
          clone['R.Beneficiary'] = name;
          clone.responsibleBeneficiary = name;
          clone.__beneficiaryKey = key;
          latestByBeneficiary.set(key, clone);
        }
      });
    } else {
      const beneficiary = String(
        row['R.Beneficiary'] ||
        row.responsibleBeneficiary ||
        row['Responsible Beneficiary'] ||
        row['Name of Responsible beneficiary'] ||
        ''
      ).trim();
      if (beneficiary) {
        const key = beneficiary.toLowerCase();
        const prev = latestByBeneficiary.get(key);
        if (!prev || row._ts > prev._ts || (row._ts === prev._ts && row._rowIndex >= prev._rowIndex)) {
          latestByBeneficiary.set(key, row);
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
  const header = [
    'Ref','Date and time of entry','Project','Team','R.Beneficiary','Vehicle Number',
    'Make','Model','Category','Usage Type','Owner','Status','Last Users remarks','Ratings','Submitter username'
  ];
  sh.clearContents();
  sh.getRange(1,1,1,header.length).setValues([header]);
  sh.setFrozenRows(1);
  if (rows.length) {
    const values = rows.map(r => [
      r.Ref || r['Reference Number'] || '',
      r['Date and time of entry'] || '',
      r.Project || '',
      r.Team || '',
      r['R.Beneficiary'] || r.responsibleBeneficiary || '',
      r['Vehicle Number'] || '',
      r.Make || '',
      r.Model || '',
      r.Category || '',
      r['Usage Type'] || '',
      r.Owner || '',
      _normStatus_(r.Status) || (r.Status || ''),
      r['Last Users remarks'] || '',
      r.Ratings || '',
      r['Submitter username'] || ''
    ]);
    sh.getRange(2,1,values.length,header.length).setValues(values);
  }
  try { sh.autoResizeColumns(1, header.length); } catch(_){}
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
 * Get Responsible Beneficiary and other team members for a car from CarT_P
 */
function getCarReleaseDetails(carNumber) {
  const ss = SpreadsheetApp.openById(CAR_SHEET_ID);
  const sheet = ss.getSheetByName('CarT_P');
  if (!sheet) return { ok:false, error:'CarT_P sheet not found' };
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const carIdx = header.indexOf('Vehicle Number');
  const teamIdx = header.indexOf('Team');
  const statusIdx = header.indexOf('Status');
  const rbenIdx = header.indexOf('R.Beneficiary');
  if (carIdx < 0 || teamIdx < 0 || statusIdx < 0 || rbenIdx < 0) return { ok:false, error:'Required columns missing' };

  // Find all IN USE rows for this car
  const dateIdx = header.indexOf('Date and time of entry');
  if (dateIdx < 0) return { ok:false, error:'Date and time of entry column missing' };

  // Gather all IN USE rows for this car
  let inUseRows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][carIdx]).trim() === carNumber && String(data[i][statusIdx]).trim().toUpperCase() === 'IN USE') {
      inUseRows.push(data[i]);
    }
  }
  if (inUseRows.length === 0) return { ok:false, error:'No IN USE row found for this car' };

  // Find the most recent date/time
  let maxDate = null;
  inUseRows.forEach(row => {
    let dt = new Date(row[dateIdx]);
    if (!maxDate || dt > maxDate) maxDate = dt;
  });
  // Filter to only rows with the most recent date/time
  let recentRows = inUseRows.filter(row => {
    let dt = new Date(row[dateIdx]);
    return dt.getTime() === maxDate.getTime();
  });

  // Gather all names from R.Beneficiary in recent rows
  let allNames = [];
  recentRows.forEach(row => {
    const rbenCell = String(row[rbenIdx]).trim();
    const names = rbenCell.split(/[,;\n]+/).map(n => n.trim()).filter(n => n);
    if (names.length > 0) {
      allNames.push(...names);
    }
  });
  if (allNames.length === 0) return { ok:false, error:'No beneficiary names found' };

  // The FIRST name from the first recent row is the responsible beneficiary
  const responsible = recentRows.length > 0 ? String(recentRows[0][rbenIdx]).trim().split(/[,;\n]+/).map(n => n.trim()).filter(n => n)[0] : allNames[0];
  // All other names except responsible, deduplicated
  let teamMembersSet = new Set(allNames.filter(n => n !== responsible));
  const teamMembers = Array.from(teamMembersSet);
  return { ok:true, responsible, teamMembers };
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
        'Submitter username'
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
    if (Array.isArray(releaseData.teamMembers) && releaseData.teamMembers.length) {
      releaseData.teamMembers.forEach(addTeamMember);
    } else if (Array.isArray(releaseData.lastUsers) && releaseData.lastUsers.length) {
      releaseData.lastUsers.forEach(addTeamMember);
    } else if (typeof releaseData.lastUsers === 'string' && releaseData.lastUsers) {
      releaseData.lastUsers.split(/[,;\n]+/).forEach(addTeamMember);
    }
    if (Array.isArray(releaseData.otherTeamMembers) && releaseData.otherTeamMembers.length) {
      releaseData.otherTeamMembers.forEach(addTeamMember);
    }

    const teamMembersString = teamMembersList.join(', ');
    const responsibleBeneficiary = teamMembersList.length
      ? teamMembersList[0]
      : String(releaseData.responsibleBeneficiary || '').trim();
    releaseData.responsibleBeneficiary = responsibleBeneficiary;
    releaseData.teamMembers = teamMembersList;

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
    const rBeneficiaryValue = teamMembersString || responsibleBeneficiary;

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
    assign(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], rBeneficiaryValue, false);
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

    // Update existing IN USE rows for this vehicle/team to mark all team members as released
    const idxVehicleCol = safeIdx(['Vehicle Number','Car Number','Vehicle No','Car No','Car #','Car']);
    const idxStatusCol = safeIdx(['In Use/Release','In Use / release','In Use','Status']);
    const idxTeamCol = safeIdx(['Team','Team Name']);
    const idxRemarksCol = safeIdx(['Last Users remarks','Remarks','Feedback']);
    const idxStarsCol = safeIdx(['Stars','Ratings','Rating']);
    const idxSubmitCol = safeIdx(['Submitter username','Submitter','User']);
    const idxRBCol = safeIdx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible']);

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

        const updated = rowValue.slice();
        updated[idxStatusCol] = 'RELEASE_USER';
        if (idxRBCol >= 0 && rBeneficiaryValue) {
          updated[idxRBCol] = rBeneficiaryValue;
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
        updates.push({ rowNumber: r + 2, values: updated });
      }

      for (let i = 0; i < updates.length; i++) {
        const update = updates[i];
        sh.getRange(update.rowNumber, 1, 1, header.length).setValues([update.values]);
      }
    }

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, 1, rowValues.length).setValues([rowValues]);

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

    return {
      ok: true,
      refNumber: refNumber,
      submitted: 1,
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
  const userName = String(payload.user || payload.releasedUser || '').trim();

  if (!carNumber) {
    return { ok: false, error: 'Missing car number' };
  }
  if (!userName) {
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
    const iDate = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    const iRemarks = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars = idx(['Ratings','Stars','Rating'], false);
    const iSubmit = idx(['Submitter username','Submitter','User'], false);

    const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    const rows = dataRange.getValues();
    const display = dataRange.getDisplayValues();

    const carKey = carNumber;
    const userKey = userName.toLowerCase();
    let targetIndex = -1;
    const activeUsers = new Map();

    for (let r = rows.length - 1; r >= 0; r--) {
      const row = rows[r];
      const carValue = String(row[iCar] || display[r][iCar] || '').trim().toUpperCase();
      if (carValue !== carKey) continue;

      const statusValue = String(row[iStatus] || display[r][iStatus] || '').trim().toUpperCase();
      const beneficiaryRaw = row[iBeneficiary] || display[r][iBeneficiary] || '';
      const beneficiary = String(beneficiaryRaw).trim();
      const beneficiaryKey = beneficiary.toLowerCase();

      if (statusValue === 'IN USE') {
        if (!activeUsers.has(beneficiaryKey)) {
          activeUsers.set(beneficiaryKey, { name: beneficiary, indices: [] });
        }
        activeUsers.get(beneficiaryKey).indices.push(r);
        if (beneficiaryKey === userKey && targetIndex === -1) {
          targetIndex = r;
        }
      }
    }

    if (targetIndex === -1) {
      return { ok: false, error: 'The selected user is not currently assigned to this vehicle' };
    }

    const remainingUsers = Array.from(activeUsers.keys()).filter(key => key !== userKey);
    const hasOtherUsers = remainingUsers.length > 0;
    const requireFeedback = !!payload.requireFeedback || !hasOtherUsers;

    if (requireFeedback) {
      const remarks = String(payload.remarks || '').trim();
      if (remarks.length < 10) {
        return { ok: false, error: 'Please provide at least 10 characters of remarks' };
      }
      const rating = Number(payload.stars);
      if (isNaN(rating) || rating < 1 || rating > 5) {
        return { ok: false, error: 'Stars rating must be between 1 and 5' };
      }
    }

    const rowValues = rows[targetIndex].slice();
    const releaseTimestamp = new Date();
    let remarksNote = String(payload.remarks || '').trim();
    if (!remarksNote) {
      remarksNote = `Released on ${Utilities.formatDate(releaseTimestamp, TZ(), 'dd-MMM-yyyy HH:mm')}`;
    }

    rowValues[iStatus] = 'RELEASE_USER';
    if (iRemarks >= 0) {
      rowValues[iRemarks] = remarksNote;
    }
    if (iStars >= 0) {
      rowValues[iStars] = payload.stars ? Number(payload.stars) : rowValues[iStars];
    }
    if (iSubmit >= 0 && payload.submitter) {
      rowValues[iSubmit] = payload.submitter;
    }

    // Preserve original assignment timestamp to maintain ordering
    if (iDate >= 0 && !requireFeedback) {
      // leave as-is for partial release; optional future enhancement could log release time separately
    }

    sh.getRange(targetIndex + 2, 1, 1, lastCol).setValues([rowValues]);

    if (!requireFeedback) {
      try { refreshVehicleStatusSheets(); } catch (err) { console.error('Partial user release refresh failed:', err); }
      try { syncVehicleSheetFromCarTP(); } catch (err) { console.error('Partial user release sync failed:', err); }
      try { CacheService.getScriptCache().remove('VEH_PICKER_V1'); } catch (_e) { /* ignore */ }
      return { ok: true, partial: true, releasedUser: userName };
    }

    // Last user released – perform full vehicle release using existing workflow
    const fullReleasePayload = {
      project: payload.project || '',
      projectName: payload.projectName || payload.project || '',
      team: payload.team || '',
      teamName: payload.teamName || payload.team || '',
      category: payload.category || '',
      carNumber: carNumber,
      make: payload.make || '',
      model: payload.model || '',
      usageType: payload.usageType || '',
      owner: payload.owner || '',
      status: 'RELEASE',
      lastUsers: Array.isArray(payload.lastUsers) && payload.lastUsers.length
        ? payload.lastUsers.join(', ')
        : userName,
      remarks: remarksNote,
      stars: Number(payload.stars || 0),
      submitter: payload.submitter || '',
      responsibleBeneficiary: payload.responsibleBeneficiary || userName
    };

    lock.releaseLock();
    lockHeld = false;
    const releaseResult = submitCarRelease(fullReleasePayload);
    return { ...releaseResult, partial: false, releasedUser: userName };
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
    const iResp     = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
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
    const sh = _openCarTP_();
    if (!sh) { console.log('CarT_P sheet not found'); return []; }
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iProject  = idx(['Project'], false);
    const iTeam     = idx(['Team'], false);
    let iCarNo = -1; try { iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); } catch(_){ iCarNo = -1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    if (iCarNo < 0) throw new Error('Car Number column not found');
    const iMake     = idx(['Make','Car Make','Brand'], false);
    const iModel    = idx(['Model','Car Model'], false);
    const iCategory = idx(['Category','Vehicle Category','Cat'], false);
    const iUsage    = idx(['Usage Type','Usage','Use Type'], false); // often column H
    const iUsage2   = idx(['Usage Type','Usage','Use Type','Contract Type','Contract','Agreement Type'], false); // possible column I
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner    = idx(['Owner','Owner Name','Owner Info'], false);
    const iResp     = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
    const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iRemarks  = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars    = idx(['Stars','Rating'], false);

    const rng = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = rng.getValues();
    const disp = rng.getDisplayValues();

    const map = new Map(); // key: CAR UPPER, value: record
    for (let r = data.length - 1; r >= 0; r--) {
      const row = data[r];
      const carRaw = row[iCarNo];
      const carDisp = disp[r][iCarNo];
      const car = String((carRaw == null || String(carRaw).trim()==='') ? carDisp : carRaw).trim();
      if (!car) continue;
      const key = car.toUpperCase();
      if (map.has(key)) continue; // already captured a later occurrence
      const responsible = iResp>=0 ? (row[iResp] || disp[r][iResp] || '') : '';
      map.set(key, {
        carNumber: car,
        make: iMake>=0 ? row[iMake] : '',
        model: iModel>=0 ? row[iModel] : '',
        // Map: category from explicit Category or fallback to Usage (H)
        category: iCategory>=0 ? row[iCategory] : (iUsage>=0 ? row[iUsage] : ''),
        // Usage Type prefers second usage/contract (I) then contract then usage(H)
        usageType: iUsage2>=0 ? row[iUsage2] : (iContract>=0 ? row[iContract] : (iUsage>=0 ? row[iUsage] : '')),
        contractType: iContract>=0 ? row[iContract] : '',
        owner: iOwner>=0 ? (row[iOwner] || disp[r][iOwner] || '') : '',
        remarks: iRemarks>=0 ? row[iRemarks] : '',
        project: iProject>=0 ? row[iProject] : '',
        team: iTeam>=0 ? row[iTeam] : '',
        status: iStatus>=0 ? (row[iStatus] || disp[r][iStatus]) : '',
        stars: iStars>=0 ? row[iStars] : 0,
        responsibleBeneficiary: responsible,
        'R.Beneficiary': responsible
      });
    }

    // Sort by car number for a stable dropdown
    const out = Array.from(map.values()).sort((a,b)=> String(a.carNumber||'').localeCompare(String(b.carNumber||'')));
    console.log(`getAllUniqueCars -> ${out.length} unique vehicles`);
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
    const sh = _openCarTP_();
    if (!sh) { console.log('CarT_P sheet not found'); return []; }
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iProject  = idx(['Project'], false);
    const iTeam     = idx(['Team'], false);
    let iCarNo = -1; try { iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); } catch(_){ iCarNo = -1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    if (iCarNo < 0) throw new Error('Car Number column not found');
    const iMake     = idx(['Make','Car Make','Brand'], false);
    const iModel    = idx(['Model','Car Model'], false);
    const iCategory = idx(['Category','Vehicle Category','Cat'], false);
    const iUsage    = idx(['Usage Type','Usage','Use Type'], false);
    const iUsage2   = idx(['Usage Type','Usage','Use Type','Contract Type','Contract','Agreement Type'], false);
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner    = idx(['Owner','Owner Name','Owner Info'], false);
    const iResp     = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
    const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iRemarks  = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars    = idx(['Stars','Rating'], false);

    const rng = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = rng.getValues();
    const disp = rng.getDisplayValues();

    const map = new Map();
    for (let r = data.length - 1; r >= 0; r--) {
      const row = data[r];
      const carRaw = row[iCarNo];
      const carDisp = disp[r][iCarNo];
      const car = String((carRaw == null || String(carRaw).trim()==='') ? carDisp : carRaw).trim();
      if (!car) continue;
      const s = iStatus>=0 ? String(row[iStatus] || disp[r][iStatus] || '').trim().toUpperCase() : '';
      const compactStatus = s.replace(/[\s_\-/]+/g,'');
      const isFullRelease = compactStatus === 'RELEASE' || compactStatus === 'RELEASED';
      const hasUse = compactStatus.indexOf('INUSE') !== -1 || compactStatus.indexOf('UNUSE') !== -1;
      if (!isFullRelease || hasUse) continue; // filter to latest entry that is RELEASE only
      const key = car.toUpperCase();
      if (map.has(key)) continue; // keep first (latest by position) RELEASE row per car
      const responsible = iResp>=0 ? (row[iResp] || disp[r][iResp] || '') : '';
      map.set(key, {
        carNumber: car,
        make: iMake>=0 ? row[iMake] : '',
        model: iModel>=0 ? row[iModel] : '',
        category: iCategory>=0 ? row[iCategory] : (iUsage>=0 ? row[iUsage] : ''),
        usageType: iUsage2>=0 ? row[iUsage2] : (iContract>=0 ? row[iContract] : (iUsage>=0 ? row[iUsage] : '')),
        contractType: iContract>=0 ? row[iContract] : '',
        owner: iOwner>=0 ? (row[iOwner] || disp[r][iOwner] || '') : '',
        remarks: iRemarks>=0 ? row[iRemarks] : '',
        project: iProject>=0 ? row[iProject] : '',
        team: iTeam>=0 ? row[iTeam] : '',
        status: s,
        stars: iStars>=0 ? row[iStars] : 0,
        responsibleBeneficiary: responsible,
        'R.Beneficiary': responsible
      });
    }

    const out = Array.from(map.values()).sort((a,b)=> String(a.carNumber||'').localeCompare(String(b.carNumber||'')));
    console.log(`getUniqueReleaseCars -> ${out.length} unique RELEASE vehicles`);
    return out;
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
    const sh = _openCarTP_();
    if (!sh) { console.log('CarT_P sheet not found'); return []; }
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iProject  = idx(['Project'], false);
    const iTeam     = idx(['Team'], false);
    const iDate     = idx(['Date and time of entry','Date and time','Timestamp','Date'], true);
    let iCarNo = -1; try { iCarNo = idx(['Car Number','Car No','Vehicle Number','Vehicle No','Car #','Car'], false); } catch(_){ iCarNo = -1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    if (iCarNo < 0) throw new Error('Car Number column not found');
    const iMake     = idx(['Make','Car Make','Brand'], false);
    const iModel    = idx(['Model','Car Model'], false);
    const iCategory = idx(['Category','Vehicle Category','Cat'], false);
    const iUsage    = idx(['Usage Type','Usage','Use Type'], false);
    const iUsage2   = idx(['Usage Type','Usage','Use Type','Contract Type','Contract','Agreement Type'], false);
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner    = idx(['Owner','Owner Name','Owner Info'], false);
    const iResp     = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
    const iStatus   = idx(['In Use/Release','In Use / release','In Use','Status'], false);
    const iRemarks  = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars    = idx(['Stars','Rating'], false);

    const rng = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = rng.getValues();
    const disp = rng.getDisplayValues();

    const map = new Map(); // key -> { ts, rec }
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const carRaw = row[iCarNo];
      const carDisp = disp[r][iCarNo];
      const car = String((carRaw == null || String(carRaw).trim()==='') ? carDisp : carRaw).trim();
      if (!car) continue;
      const dt = row[iDate] || disp[r][iDate] || '';
      let ts = NaN;
      if (dt instanceof Date) {
        ts = dt.getTime();
      } else if (dt !== '') {
        const t = new Date(dt);
        ts = isNaN(t.getTime()) ? NaN : t.getTime();
      }
      if (isNaN(ts)) {
        ts = r + 1; // fallback to row order when timestamp is missing/unparseable
      }
      const key = car.toUpperCase();
      const prev = map.get(key);
      if (!prev || ts >= prev.ts) {
        map.set(key, {
          ts,
          rec: {
            carNumber: car,
            make: iMake>=0 ? row[iMake] : '',
            model: iModel>=0 ? row[iModel] : '',
            category: iCategory>=0 ? row[iCategory] : (iUsage>=0 ? row[iUsage] : ''),
            usageType: iUsage2>=0 ? row[iUsage2] : (iContract>=0 ? row[iContract] : (iUsage>=0 ? row[iUsage] : '')),
            contractType: iContract>=0 ? row[iContract] : '',
            owner: iOwner>=0 ? (row[iOwner] || disp[r][iOwner] || '') : '',
            responsibleBeneficiary: iResp>=0 ? (row[iResp] || disp[r][iResp] || '') : '',
            'R.Beneficiary': iResp>=0 ? (row[iResp] || disp[r][iResp] || '') : '',
            remarks: iRemarks>=0 ? row[iRemarks] : '',
            project: iProject>=0 ? row[iProject] : '',
            team: iTeam>=0 ? row[iTeam] : '',
            status: iStatus>=0 ? (row[iStatus] || disp[r][iStatus]) : '',
            stars: iStars>=0 ? row[iStars] : 0,
            dateTime: dt
          }
        });
      }
    }

    // Keep only those whose latest row's status is RELEASE/RELEASED
    const out = [];
    map.forEach(({rec}) => {
      const s = String(rec.status || '').trim().toUpperCase();
      const compactStatus = s.replace(/[\s_\-/]+/g,'');
      const isFullRelease = compactStatus === 'RELEASE' || compactStatus === 'RELEASED';
      if (isFullRelease) {
        out.push(rec);
      }
    });
    out.sort((a,b)=> String(a.carNumber||'').localeCompare(String(b.carNumber||'')));
    console.log(`getLatestReleaseCars -> ${out.length} vehicles (latest entry by date is RELEASE)`);
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
      'Submitter username'
    ];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    console.log('Added header to CarT_P sheet');
    
    // Add sample car data - mix of RELEASE and IN USE entries (anonymized vehicle numbers)
    const sampleData = [
      ['REF-001', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'RELEASE', 'This car was good for city driving', 4, 'Jane Smith'],
      ['REF-002', new Date(), 'Project Falcon', 'Beta Team', 'Jane Smith', 'TEST-VEHICLE-B', 'Honda', 'Civic', 'Sedan', 'Rental', 'Jane Smith', 'IN USE', 'Currently assigned to team', 5, 'John Doe'],
      ['REF-003', new Date(), 'Project Eagle', 'Gamma Squad', 'Mike Johnson', 'TEST-VEHICLE-C', 'Nissan', 'Sentra', 'Sedan', 'Rental', 'Mike Johnson', 'RELEASE', 'Excellent vehicle for long trips', 5, 'Sarah Wilson'],
      ['REF-004', new Date(), 'Project Falcon', 'Delta Force', 'Sarah Wilson', 'TEST-VEHICLE-D', 'Toyota', 'Camry', 'Sedan', 'Rental', 'Sarah Wilson', 'IN USE', 'Comfortable and reliable vehicle', 4, 'Mike Johnson'],
      ['REF-005', new Date(), 'Project Eagle', 'Echo Team', 'Alex Brown', 'TEST-VEHICLE-E', 'Honda', 'Accord', 'Sedan', 'Rental', 'Alex Brown', 'RELEASE', 'Great car for team transportation', 5, 'Lisa Davis'],
      ['REF-006', new Date(), 'Project Falcon', 'Alpha Ops', 'John Doe', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Sedan', 'Rental', 'John Doe', 'IN USE', 'Reassigned to same team', 4, 'Jane Smith']
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
        'Car Number',
        'Make',
        'Model',
        'Usage Type',
        'Contract Type',
        'Owner',
        'In Use/Release',
        'Last Users remarks',
        'Stars',
        'Submitter username'
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
      ['REF-001', new Date(), 'Project Falcon', 'Alpha Ops', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Rental', 'Monthly', 'John Doe', 'RELEASE', 'This car was good for city driving', 4, 'Jane Smith'],
      ['REF-002', new Date(), 'Project Falcon', 'Beta Team', 'TEST-VEHICLE-B', 'Honda', 'Civic', 'Rental', 'Monthly', 'Jane Smith', 'IN USE', 'Currently assigned to team', 5, 'John Doe'],
      ['REF-003', new Date(), 'Project Eagle', 'Gamma Squad', 'TEST-VEHICLE-C', 'Nissan', 'Sentra', 'Rental', 'Weekly', 'Mike Johnson', 'RELEASE', 'Excellent vehicle for long trips', 5, 'Sarah Wilson'],
      ['REF-004', new Date(), 'Project Falcon', 'Delta Force', 'TEST-VEHICLE-D', 'Toyota', 'Camry', 'Rental', 'Monthly', 'Sarah Wilson', 'IN USE', 'Comfortable and reliable vehicle', 4, 'Mike Johnson'],
      ['REF-005', new Date(), 'Project Eagle', 'Echo Team', 'TEST-VEHICLE-E', 'Honda', 'Accord', 'Rental', 'Weekly', 'Alex Brown', 'RELEASE', 'Great car for team transportation', 5, 'Lisa Davis'],
      ['REF-006', new Date(), 'Project Falcon', 'Alpha Ops', 'TEST-VEHICLE-A', 'Toyota', 'Corolla', 'Rental', 'Monthly', 'John Doe', 'IN USE', 'Reassigned to same team', 4, 'Jane Smith']
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
        const responsible = String(r['R.Beneficiary'] || r.responsibleBeneficiary || '').trim();
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
      'test@example.com'
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

    const lastCol = sh.getLastColumn();
    const hasHeader = sh.getLastRow() >= 1 && lastCol >= 1;
    if(!hasHeader || sh.getLastRow() === 0){
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
        const currentHead = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
        const hasResponsible = currentHead.some(function(h){
          const norm = String(h||'').trim().toLowerCase().replace(/[^a-z]/g,'');
          return norm === 'rbeneficiary' || norm === 'responsiblebeneficiary';
        });
        if (hasResponsible) return;
        const teamIdx = currentHead.findIndex(function(h){
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
        console.warn('Unable to ensure R.Beneficiary column', err);
      }
    };

    ensureResponsibleColumn();
    const head = sh.getRange(1,1,1,Math.max(15, sh.getLastColumn())).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }
    const iRef   = idx(['Ref','Reference Number','Ref Number'], false);
    const iDate  = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    let iCarNo   = -1; try{ iCarNo = idx(['Vehicle Number','Car Number','Car No','Vehicle No','Car #','Car'], false);}catch(_){ iCarNo=-1; }
    if(iCarNo<0) iCarNo = _findCarNumberColumn_(head);
    const iProj  = idx(['Project'], false);
    const iTeam  = idx(['Team','Team Name'], false);
    const iResp  = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
    const iMake  = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iCat   = idx(['Category','Vehicle Category','Cat'], false);
    const iUse   = idx(['Usage Type','Usage','Use Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);
    const iStat  = idx(['Status','In Use/Release','In Use / release','In Use'], false);
    const iRem   = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iRate  = idx(['Ratings','Stars','Rating'], false);
    const iSub   = idx(['Submitter username','Submitter','User'], false);

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
function _findCarNumberColumn_(headRow){
  try{
    // headRow is array of header display strings
    for (var i=0;i<headRow.length;i++){
      var h = String(headRow[i]||'').toLowerCase().replace(/\s+/g,' ').trim();
      // must mention car/vehicle and number/no/# or registration/plate
      var hasVehicle = /(car|vehicle)/.test(h);
      var hasNumberish = /(number|no\b|#|registration|reg|plate)/.test(h);
      if (hasVehicle && hasNumberish) return i;
    }
  }catch(e){ /* ignore */ }
  return -1;
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

  const minIdx = Math.min(iName, iDesig, iDA, iProj, iTeam, iAcct, iUse);
  const maxIdx = Math.max(iName, iDesig, iDA, iProj, iTeam, iAcct, iUse);
  const width  = maxIdx - minIdx + 1;

  const shBody = sh.getRange(2, minIdx + 1, lastRow - 1, width).getValues();

  const out = new Array(shBody.length);
  for (let r = 0; r < shBody.length; r++) {
    const row = shBody[r];
    const get = (abs) => row[abs - minIdx];
    out[r] = {
      beneficiary: _norm(get(iName)),
      designation: _norm(get(iDesig)),
      defaultDa:   _toNum(get(iDA)),
      project:     _norm(get(iProj)),
      team:        _norm(get(iTeam)),
      account:     _norm(get(iAcct)),
      inuse:       _norm(get(iUse))
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
  const iResp  = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);
  const iStat  = idx(['In Use/Release','In Use / release','In Use','Status'], false);
  const iRem   = idx(['Last Users remarks','Remarks','Feedback'], false);
  const iRate  = idx(['Ratings','Stars','Rating'], false);
  const iSubmit= idx(['Submitter username','Submitter','User'], false);

  const rng = sh.getRange(2,1,lastRow-1,lastCol);
  const data = rng.getValues();
  const disp = rng.getDisplayValues();
  const out = [];
  for (let r=0;r<data.length;r++){
    const row = data[r];
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
      'R.Beneficiary': iResp>=0 ? (row[iResp] || disp[r][iResp] || '') : '',
      Status: iStat>=0 ? (row[iStat] || disp[r][iStat] || '') : '',
      'Last Users remarks': iRem>=0 ? (row[iRem] || disp[r][iRem] || '') : '',
      Ratings: iRate>=0 ? (row[iRate] || disp[r][iRate] || '') : '',
      'Submitter username': iSubmit>=0 ? (row[iSubmit] || disp[r][iSubmit] || '') : ''
    };
    obj._rowIndex = r + 2;
    if (!obj.responsibleBeneficiary) {
      obj.responsibleBeneficiary = obj['R.Beneficiary'] || '';
    }
    if (!obj['Responsible Beneficiary'] && obj['R.Beneficiary']) {
      obj['Responsible Beneficiary'] = obj['R.Beneficiary'];
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
      responsibleBeneficiary: String(target['R.Beneficiary'] || target.responsibleBeneficiary || '').trim(),
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
        responsibleBeneficiary: String(row['R.Beneficiary'] || row.responsibleBeneficiary || '').trim()
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
    const iResp = idx(['R.Beneficiary', 'Responsible Beneficiary', 'R Beneficiary', 'Responsible'], false);
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

/**
 * Optimized vehicle feed for the picker modal. Returns the best available
 * RELEASE list (latest-first, then unique RELEASE, then all uniques) in a
 * single Apps Script call. Results are briefly cached to keep popup loads
 * well under two seconds even with repeated usage.
 */
function _vehicleSheetReleaseVehicles(){
  try {
    const sh = _openVehicleSheet_();
    if (!sh) return [];
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const lastCol = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const IX = _headerIndex_(head);
    function idx(labels, required){ try{ return IX.get(labels); }catch(e){ if(required) throw e; return -1; } }

    const iDate  = idx(['Date and time of entry','Date and time','Timestamp','Date'], false);
    const iProj  = idx(['Project','Project Name'], false);
    const iTeam  = idx(['Team','Team Name'], false);
    let iCarNo   = -1; try{ iCarNo = idx(['Vehicle Number','Car Number','Vehicle No','Car No','Car #','Car'], false);}catch(_){ iCarNo = -1; }
    if (iCarNo < 0) iCarNo = _findCarNumberColumn_(head);
    if (iCarNo < 0) throw new Error('Vehicle Number column not found in Vehicle sheet');
    const iMake  = idx(['Make','Car Make','Brand'], false);
    const iModel = idx(['Model','Car Model'], false);
    const iUsage = idx(['Usage Type','Usage','Use Type'], false);
    const iContract = idx(['Contract Type','Contract','Agreement Type'], false);
    const iOwner = idx(['Owner','Owner Name','Owner Info'], false);
    const iCat   = idx(['Category','Vehicle Category','Cat'], false);
    const iStatus= idx(['Status','In Use/Release','In Use / release','In Use'], false);
    const iRemarks = idx(['Last Users remarks','Remarks','Feedback'], false);
    const iStars   = idx(['Ratings','Stars','Rating'], false);
    const iResp    = idx(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible'], false);

    const rng = sh.getRange(2, 1, lastRow - 1, lastCol);
    const data = rng.getValues();
    const disp = rng.getDisplayValues();
    const out = [];

    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const dv = disp[r];
      const rawCar = row[iCarNo];
      const carNumber = String((rawCar == null || String(rawCar).trim()==='') ? dv[iCarNo] : rawCar).trim();
      if (!carNumber) continue;

      const statusRaw = iStatus >= 0 ? (row[iStatus] || dv[iStatus] || '') : '';
      const status = _normStatus_(statusRaw);
      if (!status) continue;
      const compact = status.replace(/[\s_\-/]+/g,'');
      const isFullRelease = compact === 'RELEASE' || compact === 'RELEASED';
      const hasUse = compact.indexOf('INUSE') !== -1 || compact.indexOf('UNUSE') !== -1;
      if (!isFullRelease || hasUse) continue; // skip non-release or mixed status rows

      const dateVal = iDate >= 0 ? (row[iDate] || dv[iDate] || null) : null;
      let dateIso = '';
      try {
        if (dateVal instanceof Date) dateIso = dateVal.toISOString();
        else if (dateVal) {
          const parsed = new Date(dateVal);
          if (!isNaN(parsed.getTime())) dateIso = parsed.toISOString();
        }
      } catch(_){ /* ignore date parse errors */ }

      const project = iProj >= 0 ? (row[iProj] || dv[iProj] || '') : '';
      const team = iTeam >= 0 ? (row[iTeam] || dv[iTeam] || '') : '';
      const responsible = iResp >= 0 ? (row[iResp] || dv[iResp] || '') : '';

      out.push({
        carNumber: carNumber,
        project: project,
        team: team,
        make: iMake >= 0 ? (row[iMake] || dv[iMake] || '') : '',
        model: iModel >= 0 ? (row[iModel] || dv[iModel] || '') : '',
        usageType: iUsage >= 0 ? (row[iUsage] || dv[iUsage] || '') : '',
        contractType: iContract >= 0 ? (row[iContract] || dv[iContract] || '') : '',
        owner: iOwner >= 0 ? (row[iOwner] || dv[iOwner] || '') : '',
        category: iCat >= 0 ? (row[iCat] || dv[iCat] || '') : '',
        status: status,
        remarks: iRemarks >= 0 ? (row[iRemarks] || dv[iRemarks] || '') : '',
        stars: iStars >= 0 ? Number(row[iStars] || dv[iStars] || 0) || 0 : 0,
        dateTime: dateIso,
        responsibleBeneficiary: responsible,
        'R.Beneficiary': responsible
      });
    }

    out.sort((a,b)=> String(a.carNumber||'').localeCompare(String(b.carNumber||'')));
    return out;
  } catch (err) {
    console.error('vehicleSheetReleaseVehicles error:', err);
    return [];
  }
}

function getVehiclePickerData(){
  try {
    let source = 'vehicleRelease';
    let vehicles = _vehicleSheetReleaseVehicles();
    if (!vehicles || !vehicles.length) {
      source = 'latestRelease';
      vehicles = getLatestReleaseCars();
    }
    if (!vehicles || !vehicles.length) {
      source = 'uniqueRelease';
      vehicles = getUniqueReleaseCars();
    }
    if (!vehicles || !vehicles.length) {
      source = 'allUnique';
      vehicles = getAllUniqueCars();
    }

    return {
      ok: true,
      source: source,
      count: Array.isArray(vehicles) ? vehicles.length : 0,
      vehicles: Array.isArray(vehicles) ? vehicles : [],
      cached: false,
      generatedAt: new Date().toISOString()
    };
  } catch (e) {
    return { ok:false, error:String(e) };
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
    const responsibleCol = findCol(['R.Beneficiary','Responsible Beneficiary','R Beneficiary','Responsible']);

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
