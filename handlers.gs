/**
 * Handler functions moved from code.gs to keep code.gs smaller and easier to navigate.
 * This file contains intent handlers, follow-up handling, and some AI-facing helpers
 * that produce conversational responses from the data layer.
 */

function handleMonthlyExpensesIntent(entities, query) {
  if (query.includes('fuel') || query.includes('petrol')) {
    return getFuelExpenses();
  }
  if (query.includes('transport')) {
    return getTransportExpenses();
  }
  return getMonthlySubmissionExpenses();
}

function handleFuelExpensesIntent(entities, query) {
  const response = getFuelExpenses();
  if (entities.timeframes.length > 0) {
    return response + `\n\n💡 *This data covers the current reporting period. For specific ${entities.timeframes[0]} data, please let me know if you need historical comparisons.*`;
  }
  return response;
}

function handleTransportExpensesIntent(entities, query) {
  return getTransportExpenses();
}

function handleBeneficiaryIntent(entities, query) {
  // Check if this is specifically about fuel costs
  if (query.includes('fuel') || query.includes('petrol') || query.includes('gas')) {
    const response = getFuelExpenses();
    chatContext.awaitingFollowup = true;
    chatContext.followupType = 'fuel_breakdown';
    return response + `\n\n❓ *Would you like me to break this down by specific teams or projects?*`;
  }

  // If a single beneficiary name is present, and we have a timeframe (or default to current month),
  // return targeted per-person monthly expenses
  try {
    if (entities && Array.isArray(entities.names) && entities.names.length === 1) {
      const name = entities.names[0];
      const now = new Date();
      const mm = entities.monthYear ? entities.monthYear.month : now.getMonth();
      const yy = entities.monthYear ? entities.monthYear.year  : now.getFullYear();
      const agg = getBeneficiaryPeriodExpenses(name, mm, yy);
      if (agg && agg.ok) return agg.message;
    }
  } catch(_e){ /* fall back */ }
  
  const response = getBeneficiaryExpenses();
  if (query.includes('who') || query.includes('which')) {
    chatContext.awaitingFollowup = true;
    chatContext.followupType = 'team_breakdown';
    return response + `\n\n❓ *Would you like me to break this down by specific teams or projects?*`;
  }
  return response;
}

function handleTeamIntent(entities, query) {
  return getTeamExpenseBreakdown();
}

function handleVehicleIntent(entities, query) {
  // Handle different vehicle-related queries
  if (query.includes('in use') || query.includes('currently') || query.includes('using')) {
    return getPendingRequests();
  }
  if (query.includes('make') || query.includes('model') || query.includes('type')) {
    return getExpenseCategories();
  }
  if (query.includes('high') && (query.includes('rated') || query.includes('rating') || query.includes('star'))) {
    return getHighAmountRequests();
  }
  if (query.includes('recent') || query.includes('latest')) {
    return getRecentRequests();
  }
  
  // Default vehicle info
  return getBeneficiaryInfo(query) + `\n\n❓ *Need details about a specific vehicle or want to check vehicle-related expenses?*`;
}

function handleSummaryIntent(entities, query) {
  return getExpenseSummary();
}

function handleGreeting(query) {
  const greetings = [
    'Hello! I\'m your Fund Request Assistant. How can I help you today?',
    'Hi there! Ready to help you with fund requests and expense queries.',
    'Greetings! What would you like to know about your fund requests?'
  ];
  
  const randomGreeting = greetings[Math.floor(Math.random() * greetings.length)];
  return randomGreeting + `\n\n💡 *Try asking: "What are this month\'s expenses?" or "Show me available vehicles"`;
}

function handleHelpRequest(query) {
  return `🤖 **I\'m your AI Fund Request Assistant!**\n\n` +
         `I can help you with:\n\n` +
         `💰 **Expenses**: "monthly expenses", "fuel costs", "transport spending"\n` +
         `👥 **People**: "beneficiary expenses", "team breakdown"\n` +
         `🚗 **Vehicles**: "available cars", "vehicle status"\n` +
         `📊 **Reports**: "expense summary", "cost breakdown"\n` +
         `🔍 **Analysis**: "compare expenses", "highest costs"\n\n` +
         `💬 **Natural Language**: Just ask naturally! I understand context and can have conversations.\n\n` +
         `❓ *What specific information are you looking for?*`;
}

function handleComparisonIntent(entities, query) {
  if (query.includes('fuel') && query.includes('transport')) {
    const fuelData = getFuelExpenses();
    const transportData = getTransportExpenses();
    return `📊 **Fuel vs Transport Comparison**\n\n${fuelData}\n\n---\n\n${transportData}\n\n💡 *Would you like me to analyze which category has higher spending?*`;
  }
  
  return getExpenseSummary() + `\n\n💡 *I can compare specific categories if you tell me which ones you\'re interested in.*`;
}

function handleTimeBasedIntent(entities, query) {
  const timeframe = entities.timeframes[0] || 'current period';
  return getMonthlySubmissionExpenses() + `\n\n📅 *This shows data for the ${timeframe}. Need historical comparisons or trends?*`;
}

/**
 * Check if the query is a follow-up response (yes/no/affirmative)
 */
function isFollowupResponse(query) {
  const affirmativeWords = ['yes', 'yeah', 'yep', 'sure', 'ok', 'okay', 'please', 'go ahead', 'continue'];
  const negativeWords = ['no', 'nope', 'not now', 'later', 'skip'];
  
  return affirmativeWords.some(word => query.includes(word)) || 
         negativeWords.some(word => query.includes(word));
}

/**
 * Handle follow-up responses based on context
 */
function handleFollowupResponse(query, followupType) {
  const isAffirmative = ['yes', 'yeah', 'yep', 'sure', 'ok', 'okay', 'please', 'go ahead', 'continue']
    .some(word => query.includes(word));
  
  if (!isAffirmative) {
    return '👍 No problem! Feel free to ask me anything else about your fund requests.';
  }
  
  switch (followupType) {
    case 'team_breakdown':
      return getTeamExpenseBreakdown();
    
    case 'fuel_breakdown':
      return getFuelExpensesByTeam();
    
    case 'project_breakdown':
      return getProjectExpenseBreakdown();
    
    default:
      return 'I\'d be happy to help! What specific breakdown would you like to see?';
  }
}

/**
 * Get fuel expenses broken down by team
 */
function getFuelExpensesByTeam() {
  try {
    addSampleSubmissionData();
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const teamFuelTotals = {};
    
    rows.forEach(row => {
      const team = row[5] || 'Unknown Team';
  // header-aware index resolution
  const hdr = (data[0]||[]).map(h=>String(h||'').trim());
  const find = (name, fb)=>{ const i = hdr.indexOf(name); return i>=0? i : (typeof fb==='number'? fb : -1); };
  const IDX = { fuel: find('Fuel Amount', 10) };
  function toNum(v){ return parseAmount(v); }
  const fuel = toNum(row[IDX.fuel]);
      const beneficiary = row[2] || 'Unknown';
      
      if (fuel > 0) {
        if (!teamFuelTotals[team]) {
          teamFuelTotals[team] = { total: 0, beneficiaries: new Set() };
        }
        teamFuelTotals[team].total += fuel;
        teamFuelTotals[team].beneficiaries.add(beneficiary);
      }
    });
    
    const sorted = Object.entries(teamFuelTotals)
      .sort(([,a], [,b]) => b.total - a.total);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = '⛽ **Fuel Expenses by Team:**\n\n';
    
    if (sorted.length === 0) {
      result += 'No fuel expenses found in the current data.';
    } else {
      sorted.forEach(([team, data], index) => {
        const beneficiaryList = Array.from(data.beneficiaries).join(', ');
        result += `${index + 1}. **${team}**\n`;
        result += `   💰 ${formatCurrency(data.total)}\n`;
        result += `   👥 Beneficiaries: ${beneficiaryList}\n\n`;
      });
      
      // Find the team with highest fuel cost
      const topTeam = sorted[0];
      result += `🏆 **Highest Fuel Costs:** ${topTeam[0]} with ${formatCurrency(topTeam[1].total)}`;
    }
    
    return result;
  } catch (error) {
    console.error('Error getting fuel expenses by team:', error);
    return 'Error retrieving fuel expense breakdown by team.';
  }
}

/**
 * Get expenses broken down by project
 */
function getProjectExpenseBreakdown() {
  try {
    addSampleSubmissionData();
    const sheet = getSheet('submissions');
    if (!sheet) return 'Submissions data not available.';
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const projectTotals = {};
    
    rows.forEach(row => {
      const project = row[4] || 'Unknown Project';
  const hdr = (data[0]||[]).map(h=>String(h||'').trim());
  const find = (name, fb)=>{ const i = hdr.indexOf(name); return i>=0? i : (typeof fb==='number'? fb : -1); };
  const IDX = { total: find('Total Expense', 6) };
  function toNum(v){ return parseAmount(v); }
  const total = toNum(row[IDX.total]);
      
      if (!projectTotals[project]) {
        projectTotals[project] = { total: 0, count: 0 };
      }
      projectTotals[project].total += total;
      projectTotals[project].count++;
    });
    
    const sorted = Object.entries(projectTotals)
      .sort(([,a], [,b]) => b.total - a.total);
    
    const formatCurrency = (amount) => {
      const num = parseFloat(amount) || 0;
      return isNaN(num) ? '$0.00' : `$${num.toFixed(2)}`;
    };
    
    let result = '🏗️ **Project Expense Breakdown:**\n\n';
    sorted.forEach(([project, data], index) => {
      result += `${index + 1}. **${project}**\n   💰 ${formatCurrency(data.total)} (${data.count} submissions)\n\n`;
    });
    
    return result;
  } catch (error) {
    console.error('Error getting project expense breakdown:', error);
    return 'Error retrieving project expense data.';
  }
}
