/**
 * Response-building helpers (moved from code.gs)
 * These functions compute report summaries and return formatted
 * human-readable strings (used by intent handlers).
 */

function getFuelExpenses() {
  try {
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    let totalFuel = 0;
    let fuelEntries = 0;
    const fuelByBeneficiary = {};
    
    rows.forEach(row => {
      const fuel = parseFloat(row[10] || 0);
      if (fuel > 0) {
        totalFuel += fuel;
        fuelEntries++;
        const beneficiary = row[2] || 'Unknown';
        if (!fuelByBeneficiary[beneficiary]) fuelByBeneficiary[beneficiary] = 0;
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
  try {
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    let totalTransport = 0;
    let transportEntries = 0;
    const transportByBeneficiary = {};
    
    rows.forEach(row => {
      const transport = parseFloat(row[12] || 0);
      if (transport > 0) {
        totalTransport += transport;
        transportEntries++;
        
        const beneficiary = row[2] || 'Unknown';
        if (!transportByBeneficiary[beneficiary]) {
          transportByBeneficiary[beneficiary] = 0;
        }
        transportByBeneficiary[beneficiary] += transport;
      }
    });
    
    const sorted = Object.entries(transportByBeneficiary)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = `🚌 **Transport Expense Summary:**\n\n`;
     result += `💰 **Total Transport Costs:** ${formatCurrency(totalTransport)}\n`;
     result += `📝 **Number of Transport Entries:** ${transportEntries}\n\n`;
    
    result += `**Top Transport Users:**\n`;
    
    sorted.forEach(([name, amount], index) => {
      result += `${index + 1}. ${name}: ${formatCurrency(amount)}\n`;
    });
    
    return result;
  } catch (error) {
    console.error('Error getting transport expenses:', error);
    return 'Error retrieving transport expense data.';
  }
}

function getExpenseSummary() {
  try {
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    // Normalize headers: trim, remove NBSPs, collapse spaces, and lowercase for tolerant matching
    const headers = data[0].map(h => String(h || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase());
    function findCol(nameOrAliases) {
      if (!nameOrAliases) return -1;
      const aliases = Array.isArray(nameOrAliases) ? nameOrAliases : [nameOrAliases];
      for (const a of aliases) {
        const norm = String(a || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
        const idx = headers.indexOf(norm);
        if (idx !== -1) return idx;
      }
      return -1;
    }
    const IDX = {
      fuel: findCol(['Fuel Amount','Fuel Amt','Fuel']),
      da: findCol(['DA Amount','DA Amt','DA','ER DA Amount','ER DA Amt','ER DA','ERDA','er da amt','erda amt']),
      vehicleRent: findCol(['Vehicle Rent Amount','Vehicle Rent','Vehicle Rent Amt','Vehicle Rent Amt']),
      airtime: findCol(['Airtime Amount','Airtime']),
      transport: findCol(['Transport Amount','Transport']),
      misc: findCol(['Misc Amount','Miscellaneous','Misc']),
      total: findCol(['Total Expense','Total','Overall Total'])
    };
    // Ensure fallback indices using common exact headers (user-provided)
    const lowerHeaders = (data[0] || []).map(h => String(h || '').normalize ? String(h || '').normalize('NFKC') : String(h || ''))
      .map(s => s.replace(/\u00A0|\uFEFF|\u200B/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase());
    if (IDX.da === -1) {
      const alt = lowerHeaders.indexOf('da amount');
      if (alt !== -1) IDX.da = alt;
    }
    if (IDX.vehicleRent === -1) {
      const alt = lowerHeaders.indexOf('vehicle rent amount');
      if (alt !== -1) IDX.vehicleRent = alt;
    }

    // Safe-sum: if an index is still -1, treat values as 0 (avoid NaN)
    let totals = {
      fuel: 0,
      da: 0,
      vehicleRent: 0,
      airtime: 0,
      transport: 0,
      misc: 0,
      overall: 0,
      count: rows.length
    };
    function toNumSafe(i, row) {
      if (typeof i !== 'number' || i < 0) return 0;
      return parseAmount(row[i]);
    }
    rows.forEach(row => {
      totals.fuel += toNumSafe(IDX.fuel, row);
      totals.da += toNumSafe(IDX.da, row);
      totals.vehicleRent += toNumSafe(IDX.vehicleRent, row);
      totals.airtime += toNumSafe(IDX.airtime, row);
      totals.transport += toNumSafe(IDX.transport, row);
      totals.misc += toNumSafe(IDX.misc, row);
      let rowTotal = toNumSafe(IDX.total, row);
      if (rowTotal === 0) {
        rowTotal = toNumSafe(IDX.fuel, row) + toNumSafe(IDX.vehicleRent, row) + toNumSafe(IDX.misc, row) + toNumSafe(IDX.da, row) + toNumSafe(IDX.airtime, row) + toNumSafe(IDX.transport, row);
      }
      totals.overall += rowTotal;
    });
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    const formatPercentage = (part, total) => {
      if (!total || total === 0) return '0.0%';
      return `${((part/total)*100).toFixed(1)}%`;
    };
    return `📊 **Overall Expense Summary:**\n\n` +
           `💰 **Grand Total:** ${formatCurrency(totals.overall)}\n` +
           `📝 **Total Submissions:** ${totals.count}\n\n` +
           `**Category Breakdown:**\n` +
           `⛽ Fuel: ${formatCurrency(totals.fuel)} (${formatPercentage(totals.fuel, totals.overall)})\n` +
           `🏠 DA: ${formatCurrency(totals.da)} (${formatPercentage(totals.da, totals.overall)})\n` +
           `🚗 Vehicle Rent: ${formatCurrency(totals.vehicleRent)} (${formatPercentage(totals.vehicleRent, totals.overall)})\n` +
           `📱 Airtime: ${formatCurrency(totals.airtime)} (${formatPercentage(totals.airtime, totals.overall)})\n` +
           `🚌 Transport: ${formatCurrency(totals.transport)} (${formatPercentage(totals.transport, totals.overall)})\n` +
           `📋 Miscellaneous: ${formatCurrency(totals.misc)} (${formatPercentage(totals.misc, totals.overall)})`;
  } catch (error) {
    console.error('Error getting expense summary:', error);
    return 'Error retrieving expense summary data.';
  }
}

  /**
   * Debug helper: logs header info, resolved indices and sample DA values (first N rows).
   * Run this from the Apps Script editor (select function and Run) and check View → Logs.
   */
  function debugExpenseResolution(sampleRows) {
    try {
      const sheet = getSheet('submissions');
      if (!sheet) return 'No submissions sheet available.';
      const data = sheet.getDataRange().getValues();
      const headersRaw = (data[0] || []).map(h => String(h || ''));
      const normalize = s => (String(s || '').normalize ? String(s || '').normalize('NFKC') : String(s || ''))
        .replace(/\u00A0|\uFEFF|\u200B/g, ' ')
        .replace(/[\u0000-\u001F]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
      const headersNorm = headersRaw.map(normalize);
      const findColLocal = (aliases) => {
        const arr = Array.isArray(aliases) ? aliases : [aliases];
        for (const a of arr) {
          const n = normalize(a);
          const i = headersNorm.indexOf(n);
          if (i !== -1) return i;
        }
        return -1;
      };
      const IDX = {
        fuel: findColLocal(['Fuel Amount','Fuel Amt','Fuel']),
        da: findColLocal(['DA Amount','DA Amt','DA','ER DA Amount','ER DA Amt','ER DA','ERDA','er da amt','erda amt']),
        vehicleRent: findColLocal(['Vehicle Rent Amount','Vehicle Rent','Vehicle Rent Amt']),
        total: findColLocal(['Total Expense','Total','Overall Total'])
      };
      const lines = [];
      lines.push('Raw headers: ' + JSON.stringify(headersRaw));
      lines.push('Normalized headers: ' + JSON.stringify(headersNorm));
      lines.push('Resolved indices: ' + JSON.stringify(IDX));

      const rows = data.slice(1);
      const N = typeof sampleRows === 'number' && sampleRows > 0 ? sampleRows : 10;
      lines.push(`Showing first ${Math.min(N, rows.length)} rows (beneficiary, DA From, DA To, DA Amount raw, parsed):`);
      for (let i = 0; i < Math.min(N, rows.length); i++) {
        const r = rows[i];
        const beneficiary = String(r[2] || '');
        const daFrom = String(r[ headersNorm.indexOf('da from') >=0 ? headersNorm.indexOf('da from') : headersRaw.indexOf('DA From') ] || '');
        const daTo = String(r[ headersNorm.indexOf('da to') >=0 ? headersNorm.indexOf('da to') : headersRaw.indexOf('DA To') ] || '');
        const daRaw = String(IDX.da >= 0 ? r[IDX.da] : (r[ headersNorm.indexOf('da amount') >=0 ? headersNorm.indexOf('da amount') : headersRaw.indexOf('DA Amount') ] || ''));
        const parsed = parseAmount(daRaw);
        lines.push(`${i+1}. ${beneficiary} | DA From: ${daFrom} | DA To: ${daTo} | DA Raw: ${daRaw} | parsed: ${parsed}`);
      }
      const out = lines.join('\n');
      Logger.log(out);
      return out;
    } catch (e) {
      console.error('debugExpenseResolution error', e);
      Logger.log('debugExpenseResolution error ' + e);
      return 'Error running debugExpenseResolution: ' + String(e);
    }
  }

// Default suggestions used by fallback responses
const DEFAULT_SUGGESTIONS = [
  '💰 "What are this month\'s expenses?"',
  '🚗 "Show me available vehicles"',
  '👥 "Who are the top beneficiaries?"',
  '📊 "Give me an expense summary"'
];

/**
 * Return the standardized fallback response message.
 * If `suggestions` is empty or not provided, DEFAULT_SUGGESTIONS is used.
 * Returns the formatted message string (keeps backward compatibility).
 */
function getFallbackResponse(suggestions){
  const list = (Array.isArray(suggestions) && suggestions.length) ? suggestions : DEFAULT_SUGGESTIONS;
    return `🤔 I\'m not quite sure what you\'re looking for, but I\'d love to help!\n\n` +
           `Here are some things you can ask me:\n\n` +
           list.map(x=>`• ${x}`).join('\n') +
           `\n\nIf none of these match, try rephrasing your question.`;
}
  /**
   * Thank you message after rating feedback is recorded
   */
  function formatRatingThankYou(rating){
    return `Thank you for your feedback! I've recorded your ${rating}/5 rating and will use it to improve my responses.`;
  }

  /**
   * Enhanced unknown intent message with learning suggestion
   */
  function formatUnknownIntentWithLearning(bestMatch){
    if (bestMatch && bestMatch.intent) {
      return `I think you might be asking about ${bestMatch.intent}. ${bestMatch.response}\n\nIf this isn't what you meant, please rephrase your question.`;
    }
    return `I'm still learning to understand that type of question. I can help you with:\n\n` +
           `📊 Monthly expenses and spending\n` +
           `⛽ Fuel costs and vehicle expenses\n` +
           `👥 Beneficiary and team information\n` +
           `📈 Reports and summaries\n\n` +
           `Could you rephrase your question using these topics?`;
  }

  /**
   * Format a beneficiary metric line for display
   */
  function formatBeneficiaryMetric(name, metricLabel, label, val){
    const fmt = (n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
    return `👤 ${name} — ${label}\n${metricLabel}: ${fmt(val)}`;
  }

  /**
   * Format beneficiary monthly ranked summary
   */
  function formatBeneficiaryMonthlySummary(label, list){
    if (!list || list.length === 0) return `No submissions found for ${label}.`;
    const formatCurrency=(n)=>{ const x=parseFloat(n)||0; return isNaN(x)?'$0.00':`$${x.toFixed(2)}`; };
    const lines = [`👥 Beneficiary-wise Expenses — ${label}`,''];
    list.forEach(([name, s],i)=>{
      lines.push(`${i+1}. ${name}: ${formatCurrency(s.total)} (${s.count} submissions)`);
    });
    return lines.join('\n');
  }

  /**
   * Monthly expense summary formatter
   */
  function formatMonthlyExpenseSummary(currentYear, totalExpenses, monthlyData){
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    return `📊 **Monthly Expense Summary (${new Date().toLocaleString('default', { month: 'long' })} ${currentYear})**\n\n` +
           `💰 **Total Expenses:** ${formatCurrency(totalExpenses)}\n` +
           `📝 **Total Submissions:** ${monthlyData.count}\n\n` +
           `**Breakdown:**\n` +
           `⛽ Fuel: ${formatCurrency(monthlyData.fuel)}\n` +
     `🏠 DA: ${formatCurrency(monthlyData.erda)}\n` +
           `🚗 Vehicle Rent: ${formatCurrency(monthlyData.vehicleRent)}\n` +
           `📱 Airtime: ${formatCurrency(monthlyData.airtime)}\n` +
           `🚌 Transport: ${formatCurrency(monthlyData.transport)}\n` +
           `📋 Miscellaneous: ${formatCurrency(monthlyData.misc)}`;
  }

  /**
   * Vehicle info formatter
   */
  function formatVehicleInfo(rowData){
    // rowData: { car, project, team, status, stars, remarks, responsibleBeneficiary }
    const responsibleLine = rowData.responsibleBeneficiary ? `\nResponsible Beneficiary: ${rowData.responsibleBeneficiary}` : '';
    return `Vehicle: ${rowData.car}\nProject: ${rowData.project}\nTeam: ${rowData.team}${responsibleLine}\nStatus: ${rowData.status}\nRating: ${rowData.stars}★\nLast Remarks: ${rowData.remarks}`;
  }

  /**
   * Default chat/help response
   */
  function formatDefaultChatResponse(){
    return `I can help you with both expense and vehicle data! Try asking me:\n\n📊 **Expense Queries:**\n• "Show me monthly expenses"\n• "What are the fuel costs?"\n• "Show beneficiary expenses"\n• "Team expense breakdown"\n• "Transport expenses"\n\n🚗 **Vehicle Queries:**\n• "Show me vehicle releases for this month"\n• "Which vehicles are currently in use?"\n• "Show team vehicle usage summary"\n• "What vehicle is [car number]?"\n• "Show recent vehicle releases"\n\nWhat would you like to know?`;
  }

  /**
   * Vehicle prompt helpers
   */
  function formatVehicleMissingId(){
    return 'Please specify a vehicle number. For example: "What about T123ABC?" or "Who is using vehicle T456DEF?"';
  }

  function formatVehicleNotFound(vehicleNumber){
    return `I couldn't find vehicle "${vehicleNumber}" in the records.`;
  }

  /**
   * Format unique beneficiary stats
   * stats: { inUse, all, projInUse, teamInUse, topProjects: [string...] }
   */
  function formatUniqueBeneficiaryStats(stats){
    const topProjects = Array.isArray(stats.topProjects) ? stats.topProjects : [];
    const lines = [
      `👥 Unique Beneficiaries — DD`,
      `• IN USE: ${stats.inUse}`,
      `• All rows: ${stats.all}`,
      `• Projects (IN USE): ${stats.projInUse}`,
      `• Teams (IN USE): ${stats.teamInUse}`
    ];
    if (topProjects.length) {
      lines.push('', 'Top projects by unique beneficiaries (IN USE):', topProjects.join('\n'));
    }
    return lines.join('\n');
  }
