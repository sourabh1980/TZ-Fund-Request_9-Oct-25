/**
 * Code.gs — adds a safe route for ?view=dateslab while preserving your current app behavior.
 * - If the user is logged in and requests ?view=dateslab, we render DatesLab.html.
 * - If not logged in, we render EntryForm.html (your current flow).
 * - Existing Index flow and all server functions are unchanged.
 */

/** HTML include helper: <?!= include('FileName'); ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();

    // --- normalize route param (accept view/page/v, any case) ---
    var route = '';
    if (e && e.parameter) {
      var p = e.parameter;
      route = String(p.view || p.page || p.v || '').toLowerCase();
    }

    // session check (your existing logic)
    var isLoggedIn = checkLoginStatus();
    var userToken = scriptProperties.getProperty('userToken');
    var loggedInUser = scriptProperties.getProperty('loggedInUser');

    Logger.log('doGet: isLoggedIn=%s, token=%s, user=%s, route=%s',
               isLoggedIn, userToken, loggedInUser, route);

    if (isLoggedIn) {
      // If DatesLab is requested, serve it
      if (route === 'dateslab') {
        Logger.log('Serving DatesLab.html');
        var lab = HtmlService.createTemplateFromFile('DatesLab');
        lab.isLoggedIn = true;
        return lab.evaluate()
          .setTitle('Dates Lab')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      // Default: Index.html
      Logger.log('Serving Index.html');
      var template = HtmlService.createTemplateFromFile('Index');
      template.isLoggedIn = true;
      return template.evaluate()
        .setTitle('Fund Req TZ')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Not logged in → clear transient state and show EntryForm (your flow)
    scriptProperties.deleteProperty('pendingEmail');
    scriptProperties.deleteProperty('isOTPVerified');
    Logger.log('Not logged in → EntryForm.html');
    return HtmlService.createHtmlOutputFromFile('EntryForm')
      .setTitle('Login - Fund Req TZ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    Logger.log('Error in doGet: ' + err.message + '\n' + err.stack);
    return HtmlService.createHtmlOutput(
      '<h2>Error Loading Application</h2>' +
      '<p>Please try again or contact the administrator.</p>' +
      '<p>Error: ' + err.message + '</p>'
    )
      .setTitle('Application Error')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getIndexHtml() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var isLoggedIn = checkLoginStatus();
    var userToken = scriptProperties.getProperty('userToken');
    var loggedInUser = scriptProperties.getProperty('loggedInUser');
    Logger.log('getIndexHtml: isLoggedIn = ' + isLoggedIn + ', userToken = ' + userToken + ', loggedInUser = ' + loggedInUser);

    if (!isLoggedIn) {
      Logger.log('getIndexHtml: User not logged in, returning error');
      return { success: false, message: 'User not logged in' };
    }

    Logger.log('getIndexHtml: User is logged in, rendering Index.html');
    var template = HtmlService.createTemplateFromFile('Index');
    template.isLoggedIn = true;
    var htmlOutput = template.evaluate()
      .setTitle('Fund Req TZ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    var fullHtml = htmlOutput.getContent();

    // Extract all script contents and concatenate, remove scripts from html
    // (robust: supports <script> with attributes)
    var scriptContent = '';
    var htmlContent = fullHtml.replace(/<script\b[^>]*>([\s\S]*?)<\/script>/gi, function(match, p1) {
      scriptContent += p1 + '\n';
      return '';
    });

    // Trim
    scriptContent = scriptContent.trim();
    htmlContent = htmlContent.trim();

    // (optional) Escape </script> if your client uses innerHTML for injection.
    // If your client uses textContent/Blob injection, this escaping is harmless.
    scriptContent = scriptContent.replace(/<\/script>/gi, '<\\/script>');

    // >>> Boot hook: guarantees dropdown + split-flap mount after client injects
    scriptContent += '\n;(function(){ try{ if (window.__afterIndexRender) window.__afterIndexRender(); }catch(e){} try{ if (typeof scanAllDateInputs==="function") scanAllDateInputs(); }catch(e){} })();\n';

    Logger.log('getIndexHtml: Index.html rendered successfully, HTML length: ' + htmlContent.length + ', Script length: ' + scriptContent.length);
    return { success: true, html: htmlContent, script: scriptContent };

  } catch (e) {
    Logger.log('Error in getIndexHtml: ' + e.message + ', Stack: ' + e.stack);
    return { success: false, message: 'Error rendering Index.html: ' + e.message };
  }
}

function validateEmail(email) {
  try {
    Logger.log('Starting email validation for: ' + email);
    var ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    var sheet = ss.getSheetByName('Valid');
    if (!sheet) {
      Logger.log('Valid sheet not found.');
      return { success: false, message: 'Internal error: Valid sheet not found.' };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('Valid sheet is empty or has only headers.');
      return { success: false, message: 'Internal error: Valid sheet is empty.' };
    }

    var emailClean = email.trim().toLowerCase();
    Logger.log('Validating cleaned email: ' + emailClean);

    for (var i = 1; i < data.length; i++) {
      var sheetEmail = data[i][2] ? String(data[i][2]).trim().toLowerCase() : '';
      if (sheetEmail && sheetEmail === emailClean) {
        var offEmail = data[i][3] ? String(data[i][3]).trim() : '';
        if (!offEmail) {
          Logger.log('No Off Email found for email: ' + emailClean);
          return { success: false, message: 'No official email address found for this account.' };
        }
        // Generate and send OTP
        var otp = generateOTP();
        var cache = CacheService.getUserCache();
        cache.put('otp_' + emailClean, otp, 300); // 5-minute expiration
        cache.put('otp_attempts_' + emailClean, '0', 300); // Initialize attempts
        cache.put('off_email_' + emailClean, offEmail, 300); // Store Off Email
        sendOTP(emailClean, offEmail, otp);

        // Store pending email in script properties
        PropertiesService.getScriptProperties().setProperty('pendingEmail', emailClean);
        Logger.log('Email validated, OTP sent to: ' + offEmail);
        return { success: true, message: 'OTP sent to your official email address: ' + offEmail };
      }
    }

    Logger.log('Email not found in Valid sheet: ' + emailClean);
    return { success: false, message: 'Email not found in authorized list.' };
  } catch (e) {
    Logger.log('Error in validateEmail: ' + e.message + ', Stack: ' + e.stack);
    return { success: false, message: 'Error validating email: ' + e.message };
  }
}

function regenerateOTP(email) {
  try {
    var emailClean = email.trim().toLowerCase();
    Logger.log('Regenerating OTP for email: ' + emailClean);
    var cache = CacheService.getUserCache();
    var offEmail = cache.get('off_email_' + emailClean);

    if (!offEmail) {
      Logger.log('No Off Email found in cache for email: ' + emailClean);
      return { success: false, message: 'Session expired. Please start over by entering your email.' };
    }

    // Generate and send new OTP
    var otp = generateOTP();
    cache.put('otp_' + emailClean, otp, 300); // 5-minute expiration
    cache.put('otp_attempts_' + emailClean, '0', 300); // Reset attempts
    cache.put('off_email_' + emailClean, offEmail, 300); // Refresh Off Email cache
    sendOTP(emailClean, offEmail, otp);

    Logger.log('New OTP sent to: ' + offEmail);
    return { success: true, message: 'A new OTP has been sent to your official email address: ' + offEmail };
  } catch (e) {
    Logger.log('Error in regenerateOTP: ' + e.message + ', Stack: ' + e.stack);
    return { success: false, message: 'Error regenerating OTP: ' + e.message };
  }
}

function generateOTP() {
  // Generate a 6-digit OTP
  return Math.floor(100000 + Math.random() * 900000).toString();
}

function sendOTP(email, offEmail, otp) {
  try {
    MailApp.sendEmail({
      to: offEmail,
      subject: 'Your OTP Code for Fund Req TZ',
      body:
        'Dear User,\n\nYour OTP for accessing the Fund Req TZ application is: ' + otp + '\n\n' +
        'This OTP is valid for 5 minutes. Please enter it in the application to proceed.\n\n' +
        'Best regards,\nFund Req TZ Team'
    });
    Logger.log('OTP sent to: ' + offEmail + ' for email: ' + email);
  } catch (e) {
    Logger.log('Error sending OTP to ' + offEmail + ': ' + e.message);
    throw new Error('Failed to send OTP: ' + e.message);
  }
}

function verifyOTP(email, otp) {
  try {
    var emailClean = email.trim().toLowerCase();
    var cache = CacheService.getUserCache();
    var storedOTP = cache.get('otp_' + emailClean);
    var attemptsKey = 'otp_attempts_' + emailClean;
    var attempts = parseInt(cache.get(attemptsKey) || '0', 10);

    if (!storedOTP) {
      Logger.log('No OTP found for email: ' + emailClean);
      return { success: false, message: 'OTP expired or invalid. Please request a new OTP.' };
    }

    attempts++;
    cache.put(attemptsKey, attempts.toString(), 300);

    if (attempts > 3) {
      cache.remove('otp_' + emailClean);
      cache.remove(attemptsKey);
      cache.remove('off_email_' + emailClean);
      Logger.log('Too many OTP attempts for email: ' + emailClean);
      return { success: false, message: 'Too many attempts. Please request a new OTP.' };
    }

    if (otp === storedOTP) {
      // OTP verified, clear cache and mark as authenticated
      cache.remove('otp_' + emailClean);
      cache.remove(attemptsKey);
      cache.remove('off_email_' + emailClean);
      PropertiesService.getScriptProperties().setProperty('isOTPVerified', 'true');
      Logger.log('OTP verified for email: ' + emailClean);
      return { success: true, message: 'OTP verified. Please enter your username and password.' };
    }

    Logger.log('Invalid OTP for email: ' + emailClean + ', Attempt: ' + attempts);
    return { success: false, message: 'Invalid OTP. ' + (3 - attempts) + ' attempts remaining.' };
  } catch (e) {
    Logger.log('Error in verifyOTP: ' + e.message + ', Stack: ' + e.stack);
    return { success: false, message: 'Error verifying OTP: ' + e.message };
  }
}

function loginUser(credentials) {
  try {
    // Check if OTP is verified
    var scriptProperties = PropertiesService.getScriptProperties();
    var isOTPVerified = scriptProperties.getProperty('isOTPVerified');
    if (isOTPVerified !== 'true') {
      Logger.log('Login attempt without OTP verification for username: ' + credentials.username);
      return { success: false, message: 'Please verify your email with OTP first.' };
    }

    var ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    var sheet = ss.getSheetByName('Valid');
    if (!sheet) {
      sheet = ss.insertSheet('Valid');
      sheet.appendRow(['Username', 'Password', 'Email', 'Off email']);
      sheet.appendRow(['admin', 'password123', 'user1@example.com', 'admin.official@company.com']);
      sheet.appendRow(['user1', 'pass456', 'user2@example.com', 'user1.official@company.com']);
      Logger.log('Created new "Valid" sheet with sample data.');
    }

    var data = sheet.getDataRange().getValues();
    var pendingEmail = scriptProperties.getProperty('pendingEmail');
    var pendingEmailClean = pendingEmail ? pendingEmail.trim().toLowerCase() : '';

    for (var i = 1; i < data.length; i++) {
      var sheetEmail = data[i][2] ? String(data[i][2]).trim().toLowerCase() : '';
      if (
        data[i][0] === credentials.username &&
        data[i][1] === credentials.password &&
        sheetEmail && sheetEmail === pendingEmailClean
      ) {
        var token = Utilities.getUuid();
        scriptProperties.setProperty('userToken', token);
        scriptProperties.setProperty('loggedInUser', credentials.username);
        // Clear OTP verification but keep userToken and loggedInUser
        scriptProperties.deleteProperty('isOTPVerified');
        scriptProperties.deleteProperty('pendingEmail');
        // Add a 2-second delay to ensure PropertiesService persists the state
        Utilities.sleep(2000);
        Logger.log('Login successful for username: ' + credentials.username + ', email: ' + pendingEmail + ', userToken: ' + token);
        return { success: true, token: token, username: credentials.username };
      }
    }

    Logger.log('Login failed for username: ' + credentials.username + ', email: ' + pendingEmail);
    return { success: false, message: 'Invalid username, password, or email mismatch.' };
  } catch (e) {
    Logger.log('Error in loginUser: ' + e.message + ', Stack: ' + e.stack);
    return { success: false, message: 'Error during login: ' + e.message };
  }
}

function updateLastActivityTimestamp() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastActivityTimestamp', Date.now().toString());
  Logger.log('Updated last activity timestamp');
}

function checkLoginStatus() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var token = scriptProperties.getProperty('userToken');
  var storedToken = scriptProperties.getProperty('userToken');
  var lastActivity = scriptProperties.getProperty('lastActivityTimestamp');

  // Check if token exists and is valid
  if (!token || token !== storedToken) {
    Logger.log('checkLoginStatus: No valid token found');
    return false;
  }

  // Check for session timeout (3 minutes)
  if (lastActivity) {
    var lastActivityTime = parseInt(lastActivity, 10);
    var currentTime = Date.now();
    var idleTimeLimit = 3 * 60 * 1000; // 3 minutes in milliseconds
    if (currentTime - lastActivityTime > idleTimeLimit) {
      Logger.log('checkLoginStatus: Session timed out due to inactivity');
      // Clear session properties on timeout
      scriptProperties.deleteProperty('userToken');
      scriptProperties.deleteProperty('loggedInUser');
      scriptProperties.deleteProperty('pendingEmail');
      scriptProperties.deleteProperty('isOTPVerified');
      scriptProperties.deleteProperty('lastActivityTimestamp');
      return false;
    }
  }

  Logger.log('checkLoginStatus: Session is active');
  return true;
}

function getUserToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('userToken');
}

function validateToken(token) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var storedToken = scriptProperties.getProperty('userToken');
  Logger.log("Stored Token: " + storedToken);
  return token === storedToken;
}

function logoutUser() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('userToken');
  scriptProperties.deleteProperty('loggedInUser');
  scriptProperties.deleteProperty('pendingEmail');
  scriptProperties.deleteProperty('isOTPVerified');
  scriptProperties.deleteProperty('lastActivityTimestamp'); // Clear last activity timestamp
  Logger.log("User logged out and all session properties cleared");
  return "Logged out successfully";
}

function getBeneficiaryData() {
  try {
    const sheetId = '14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs';
    const ss = SpreadsheetApp.openById(sheetId);
    let ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) {
      Logger.log('Sheet "DD" not found. Creating a new one with sample data.');
      ddSheet = ss.insertSheet('DD');
      ddSheet.appendRow(['Name of Beneficiary', 'Resource Designation', 'Default DA Amt', 'Project', 'Team', 'Account Holder', 'InUse/Release', 'Submitter', 'Timestamp', 'Mob No', 'Display Name']);
      ddSheet.appendRow(['John Doe', 'Manager', 50, 'Project A', 'Team1', 'John Doe', 'InUse', 'admin', new Date(), '123456789', 'John D']);
      ddSheet.appendRow(['Jane Smith', 'Engineer', 40, 'Project B', 'Team2', 'Jane Smith', 'InUse', 'admin', new Date(), '987654321', 'Jane S']);
      ddSheet.appendRow(['Bob Johnson', 'TL', 60, 'Project A', 'Team1', 'Bob Johnson', 'InUse', 'admin', new Date(), '555555555', 'Bob J']);
    }

    // Batch read: Get all DD data at once
    const ddData = ddSheet.getDataRange().getValues();

    // Sort data by timestamp descending (Column I, index 8)
    const sortedData = ddData.slice(1).sort((a, b) => {
      const tsA = a[8] instanceof Date ? a[8].getTime() : 0;
      const tsB = b[8] instanceof Date ? b[8].getTime() : 0;
      return tsB - tsA;
    });
    const names = [];
    const details = {};
    const projectMap = {};
    const projectsToTeams = {};
    const teamsToBeneficiaries = {};
    const accHoldersSet = new Set();
    const ddAccHolderDetails = {};
    const latestEntries = new Map();

    sortedData.forEach(row => {
      const name = row[0]?.trim() || ''; // A
      if (name && !latestEntries.has(name)) {
        latestEntries.set(name, row);
      }
    });

    latestEntries.forEach((row, name) => {
      const designation = row[1]?.trim() || ''; // B
      const defaultDA = parseFloat(row[2]) || 0; // C
      const project = row[3]?.trim() || ''; // D
      const team = row[4]?.trim() || ''; // E
      const accountHolder = row[5]?.trim() || ''; // F
      const status = row[6]?.trim() || ''; // G: InUse/Release
      const mobNo = row[9]?.toString().trim() || ''; // J (Mob No)
      const displayName = row[10]?.trim() || ''; // K (Display Name)
      const nationality = row[9]?.trim() || ''; // J for Nationality as per user

      if (status === 'InUse') {
        names.push(name);
        details[name] = { designation, defaultDA, project, team, accountHolder, nationality };

        if (project) {
          if (!projectMap[project]) {
            projectMap[project] = { count: 0, beneficiaries: [] };
          }
          projectMap[project].count++;
          projectMap[project].beneficiaries.push(name);

          if (!projectsToTeams[project]) {
            projectsToTeams[project] = new Set();
          }
          if (team) projectsToTeams[project].add(team);

          if (!teamsToBeneficiaries[project]) {
            teamsToBeneficiaries[project] = {};
          }
          if (!teamsToBeneficiaries[project][team]) {
            teamsToBeneficiaries[project][team] = [];
          }
          teamsToBeneficiaries[project][team].push(name);
        }
      }

      if (accountHolder) {
        accHoldersSet.add(accountHolder);
        ddAccHolderDetails[accountHolder] = { mobNo, displayName };
      }
    });

    // Batch read: Fetch from Ops_P for latest mobNo and displayName
    let opsPSheet = ss.getSheetByName('Ops_P');
    const accHolderDetails = {};
    let operators = [];
    if (opsPSheet) {
      // Batch read all Ops_P data
      const opsPDataFull = opsPSheet.getDataRange().getValues();
      if (opsPDataFull.length > 1) {
        const opsPData = opsPDataFull.slice(1); // Skip header
        const latestMap = new Map();
        opsPData.forEach(row => {
          const timestamp = row[0] instanceof Date ? row[0] : new Date(row[0]);
          const acc = row[1]?.trim() || '';
          const mobNo = row[2]?.toString().trim() || '';
          const displayName = row[3]?.trim() || '';
          const operator = row[4]?.trim() || '';
          const changeReq = row[5]?.trim() || '';
          if (acc) {
            if (!latestMap.has(acc) || timestamp > latestMap.get(acc).timestamp) {
              latestMap.set(acc, { timestamp, mobNo, displayName, operator });
            }
          }
        });
        latestMap.forEach((details, acc) => {
          accHolderDetails[acc] = { mobNo: details.mobNo, displayName: details.displayName, operator: details.operator };
        });
      }
      // Operators from K1:K10 (batch read)
      const opData = opsPDataFull.map(row => row[10]).filter(v => v && v.trim()); // Column K index 10
      operators = [...new Set(opData)];
    }

    // Convert sets to arrays for JSON
    for (let proj in projectsToTeams) {
      projectsToTeams[proj] = Array.from(projectsToTeams[proj]);
    }

    const projects = Object.keys(projectMap).map(proj => ({
      name: proj,
      count: projectMap[proj].count
    }));

    Logger.log('Fetched names: ' + JSON.stringify(names));
    Logger.log('Fetched projects: ' + JSON.stringify(projects));
    return {
      names: names,
      details: details,
      projects: projects,
      beneficiariesPerProject: Object.fromEntries(
        Object.entries(projectMap).map(([proj, info]) => [proj, info.beneficiaries])
      ),
      projectsToTeams,
      teamsToBeneficiaries,
      accHolders: Array.from(accHoldersSet).sort(),
      accHolderDetails,
      operators
    };
  } catch (e) {
    Logger.log('Error in getBeneficiaryData: ' + e.message + ', Stack: ' + e.stack);
    return { names: [], details: {}, projects: [], beneficiariesPerProject: {}, projectsToTeams: {}, teamsToBeneficiaries: {}, accHolderDetails: {}, operators: [] };
  }
}

function getUniqueDDValues(column) {
  try {
    const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    const ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) return [];
    const values = ddSheet.getRange(2, column, ddSheet.getLastRow() - 1, 1).getValues().flat().filter(v => v && v.trim());
    return [...new Set(values)].sort();
  } catch (e) {
    Logger.log('Error in getUniqueDDValues: ' + e.message);
    return [];
  }
}

function checkBeneficiaryExists(name) {
  try {
    const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    const ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) return false;
    const names = ddSheet.getRange(2, 1, ddSheet.getLastRow() - 1, 1).getValues().flat().map(v => v.trim());
    return names.includes(name);
  } catch (e) {
    Logger.log('Error in checkBeneficiaryExists: ' + e.message);
    return false;
  }
}

function getReleasedBeneficiaries() {
  try {
    const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    const ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) return [];
    const data = ddSheet.getDataRange().getValues();
    const sortedData = data.slice(1).sort((a, b) => {
      const tsA = a[8] instanceof Date ? a[8].getTime() : 0;
      const tsB = b[8] instanceof Date ? b[8].getTime() : 0;
      return tsB - tsA;
    });
    const latestEntries = new Map();
    sortedData.forEach(row => {
      const name = row[0]?.trim() || '';
      if (name && !latestEntries.has(name)) {
        latestEntries.set(name, row);
      }
    });
    const released = [];
    latestEntries.forEach((row, name) => {
      const status = row[6]?.trim() || '';
      if (status === 'Release') {
        const releaseDate = row[8] instanceof Date ? row[8] : new Date(row[8]);
        const days = Math.max(0, Math.floor((new Date() - releaseDate) / (1000 * 60 * 60 * 24)));
        released.push({
          name: name,
          nameWithDays: `${name} (${days})`,
          designation: row[1]?.trim() || '',
          defaultDA: parseFloat(row[2]) || 0,
          accountHolder: row[5]?.trim() || '',
          nationality: row[9]?.trim() || ''
        });
      }
    });
    return released.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    Logger.log('Error in getReleasedBeneficiaries: ' + e.message);
    return [];
  }
}

function getNationalityOptions() {
  try {
    const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    const ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) return {options: [], defaultNat: 'Tz'};
    const values = ddSheet.getRange(2, 27, ddSheet.getLastRow() - 1, 1).getValues().flat().filter(v => v && v.trim()); // AA is column 27
    const options = [...new Set(values)].sort();
    const userTz = Session.getTimeZone();
    let defaultNat = 'Tz'; // Default
    if (userTz.includes('America')) defaultNat = 'US';
    else if (userTz.includes('Africa/Nairobi')) defaultNat = 'KE';
    // Add more mappings as needed
    if (!options.includes(defaultNat)) defaultNat = options.includes('Tz') ? 'Tz' : options[0] || 'Tz';
    return {options, defaultNat};
  } catch (e) {
    Logger.log('Error in getNationalityOptions: ' + e.message);
    return {options: [], defaultNat: 'Tz'};
  }
}

function saveNewBeneficiary(data) {
  try {
    const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
    const ddSheet = ss.getSheetByName('DD');
    if (!ddSheet) return 'Error: DD sheet not found';
    const username = getLoggedInUsername();
    ddSheet.appendRow([
      data.name,
      data.desig,
      parseFloat(data.da),
      data.project,
      data.team,
      data.acc,
      'InUse', // G: InUse/Release
      username, // H: Submitter
      new Date(), // I: Timestamp
      data.nationality, // J: Nationality
      '', // K: Mob No
      '' // L: Display Name
    ]);
    return 'Saved successfully';
  } catch (e) {
    Logger.log('Error in saveNewBeneficiary: ' + e.message);
    return 'Error saving: ' + e.message;
  }
}

// CHANGED: use script timezone / Africa/Dar_es_Salaam to avoid UTC day-shift
function formatDateForClient(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  var tz = Session.getScriptTimeZone() || 'Africa/Dar_es_Salaam';
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function getAllHistoricalDates(beneficiary) {
  var ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  var sheet = ss.getSheetByName('Submissions');
  var allDates = {
    Fuel: [],
    DA: [],
    CarRent: [],
    Airtime: [],
    Misc: []
  };
  if (!sheet) {
    Logger.log('Submissions sheet not found.');
    return allDates;
  }
  var data = sheet.getDataRange().getValues();
  const colMap = {
    Fuel:   { from: 7,  to: 8  },  // G: Fuel_F, H: Fuel_To
  DA:   { from: 10, to: 11 },  // J: DA_F, K: DA_To
    CarRent:{ from: 13, to: 14 },  // M: Car Rent_F, N: Car Rent_To
    Airtime:{ from: 17, to: 18 },  // Q: Airtime_F, R: Airtime_To
    Misc:   { from: 20, to: 21 }   // T: Misc_F, U: Misc_To
  };
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === beneficiary) {  // B: Name of Beneficiary (index 1)
      Object.keys(colMap).forEach(type => {
        const { from, to } = colMap[type];
        const dateFromRaw = data[i][from - 1];
        const dateToRaw = data[i][to - 1];
        if (dateFromRaw && dateToRaw) {
          var dateFrom = dateFromRaw instanceof Date ? dateFromRaw : new Date(dateFromRaw);
          var dateTo   = dateToRaw   instanceof Date ? dateToRaw   : new Date(dateToRaw);
          var formattedFrom = formatDateForClient(dateFrom); // now local TZ
          var formattedTo   = formatDateForClient(dateTo);   // now local TZ
          if (formattedFrom && formattedTo) {
            allDates[type].push({ from: formattedFrom, to: formattedTo });
          }
        }
      });
    }
  }
  Logger.log(`All historical dates for ${beneficiary}: ${JSON.stringify(allDates)}`);
  return allDates;
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getLoggedInUsername() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('loggedInUser') || 'Unknown';
}

function submitForm(formData) {
  var ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    sheet = ss.insertSheet('Submissions');
    sheet.appendRow([
      'DOE', 'Name of Beneficiary', 'Account Holder Name', 'Designation', 'Project', 'Team',
  'Fuel_F', 'Fuel_To', 'Fuel Amt', 'DA_F', 'DA_To', 'DA Amt',
      'Car Rent_F', 'Car Rent_To', 'Car Rent Amt', 'Car Num', 'Airtime_F', 'Airtime_To', 'Airtime Amt',
      'Misc_F', 'Misc_To', 'Misc Amt',
      'Mob No', 'Display Name', 'W/H Charges',
      'Remarks', 'Requester email', 'Total Amt', 'Submitter username'
    ]);
    // Set "DOE" column (A) as plain text
    sheet.getRange("A:A").setNumberFormat("@");
  }

  var beneficiaries = formData.map(row => row.beneficiary).filter(b => b);
  var uniqueBeneficiaries = new Set(beneficiaries);
  if (beneficiaries.length > uniqueBeneficiaries.size) {
    return "Error: Duplicate Beneficiary Detected";
  }

  var username = getLoggedInUsername();
  var userEmail = getUserEmail();
  for (var i = 0; i < formData.length; i++) {
    var row = formData[i];
    // Validate DOE format (YYYY-MM-DD)
    if (!/^\d{4}-\d{2}-\d{2}$/.test(row.doe)) {
      return "Error: Invalid DOE format. Expected YYYY-MM-DD";
    }
    sheet.appendRow([
      row.doe, row.beneficiary, row.accountHolder, row.designation, row.project, row.team,
      row.fuelF, row.fuelTo, row.fuelAmt, row.erDaF, row.erDaTo, row.erDaAmt,
      row.carRentF, row.carRentTo, row.carRentAmt, row.carNum, row.airtimeF, row.airtimeTo, row.airtimeAmt,
      row.miscF, row.miscTo, row.miscAmt,
      row.mobNo, row.displayName, row.whCharges,
      row.remarks, userEmail,
      row.rowTotal, username
    ]);
  }

  var sheetData = prepareSheetData(formData, row.doe, userEmail);
  sendSummaryAsExcel({
    sheetData: sheetData,
    doe: row.doe,
    userEmail: userEmail,
    project: formData[0].project
  });
  return "Submission successful";
}

function saveOpsData(accHolder, mobNo, displayName, operator, changeType) {
  const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  const sheetName = changeType === 'Temp' ? 'Ops_T' : 'Ops_P';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Timestamp', 'Account Holder', 'Mobile Number', 'Display Name', 'Operator', 'Submitter']);
  }
  const timestamp = new Date().toISOString();
  const username = getLoggedInUsername();
  sheet.appendRow([timestamp, accHolder, mobNo, displayName, operator, username]);
  return 'Saved';
}

function prepareSheetData(formData, doe, userEmail) {
  var headers = [
    'Name of Beneficiary', 'Account Holder Name', 'Total Amt', 'Fuel Dates', 'Fuel Amt',
    'Er DA Dates', 'Er DA Amt', 'Car Rent Dates', 'Car Rent Amt', 'Car Num',
    'Airtime Dates', 'Airtime Amt',
    'Misc Dates', 'Misc Amt', 'Mob No', 'Display Name', 'W/H Charges', 'Remarks'
  ];

  var dataRows = formData.map(row => [
    row.beneficiary, row.accountHolder,
    (parseFloat(row.fuelAmt || 0) + parseFloat(row.erDaAmt || 0) + parseFloat(row.carRentAmt || 0) + parseFloat(row.airtimeAmt || 0) + parseFloat(row.miscAmt || 0)).toFixed(2),
    (row.fuelF ? (row.fuelF + (row.fuelF === row.fuelTo ? '' : '\nTo\n' + row.fuelTo)) : ''),
    row.fuelAmt || '0.00',
    (row.erDaF ? (row.erDaF === row.erDaTo ? row.erDaF : row.erDaF + '\nTo\n' + row.erDaTo) : ''),
    row.erDaAmt || '0.00',
    (row.carRentF ? (row.carRentF === row.carRentTo ? row.carRentF : row.carRentF + '\nTo\n' + row.carRentTo) : ''),
    row.carRentAmt || '0.00',
    row.carNum || '',
    (row.airtimeF ? (row.airtimeF === row.airtimeTo ? row.airtimeF : row.airtimeF + '\nTo\n' + row.airtimeTo) : ''),
    row.airtimeAmt || '0.00',
    (row.miscF ? (row.miscF === row.miscTo ? row.miscF : row.miscF + '\nTo\n' + row.miscTo) : ''),
    row.miscAmt || '0.00',
    row.mobNo || '',
    row.displayName || '',
    row.whCharges || '0.00',
    row.remarks
  ]);

  var totalsRow = ['Field Total', '',
    dataRows.reduce((sum, row) => sum + parseFloat(row[2] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    'Fuel Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[4] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    '',
    'Er DA Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[6] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    '',
    'Car Rent Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[8] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    '',
    '',
    'Airtime Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[11] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    '',
    'Misc Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[13] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    '', '', '', 'W/H Charges Total: ' + dataRows.reduce((sum, row) => sum + parseFloat(row[16] || 0), 0).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    ''
  ];

  return [headers, ...dataRows, totalsRow];
}

function getOffEmail(username) {
  var ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  var sheet = ss.getSheetByName('Valid');
  if (!sheet) {
    Logger.log('Valid sheet not found.');
    return '';
  }
  var data = sheet.getDataRange().getValues();
  // Check the Username column (column A, index 0) and return Off email (column D, index 3)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return data[i][3] || '';
    }
  }
  Logger.log('Off email not found for username: ' + username);
  return '';
}

function sendSummaryAsExcel(payload) {
  var sheetData = payload.sheetData;
  var doe = payload.doe;
  var userEmail = payload.userEmail;
  var project = payload.project || 'Unknown Project';

  // Date formatting
  var dateObj = new Date(doe);
  if (isNaN(dateObj.getTime())) {
    Logger.log('Invalid doe value: ' + doe + '. Using current date as fallback.');
    dateObj = new Date();
  }
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var day = String(dateObj.getDate()).padStart(2, '0');
  var month = months[dateObj.getMonth()];
  var year = dateObj.getFullYear();
  var weekday = days[dateObj.getDay()];
  var formattedDate = day + " " + month + " " + year + ", " + weekday;

  // Get user metadata
  var username = getLoggedInUsername();
  var offEmail = getOffEmail(username);

  // Create temporary spreadsheet
  var ss = SpreadsheetApp.create('Fund Req TZ - ' + doe);
  var sheet = ss.getSheets()[0];
  sheet.setName('Fund Request Summary');
  var spreadsheetId = ss.getId();

  // Set page margins for PDF (narrower margins)
  var requests = [
    {
      updateSpreadsheetProperties: {
        properties: { title: 'Fund Req TZ - ' + doe },
        fields: 'title'
      }
    },
    {
      updateSheetProperties: {
        properties: {
          sheetId: sheet.getSheetId(),
          gridProperties: { frozenRowCount: 3 }, // Freeze first 3 rows
        },
        fields: 'gridProperties.frozenRowCount'
      }
    }
  ];
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
    Logger.log('Successfully set spreadsheet properties');
  } catch (e) {
    Logger.log('Error setting spreadsheet properties: ' + e.message);
  }

  // Title
  var titleRange = sheet.getRange("A1:R1"); // Adjusted for 18 columns
  titleRange.merge();
  titleRange.setValue("Fund Req TZ ; " + project + ", Submitted by " + username);
  titleRange.setFontSize(24);
  titleRange.setFontWeight("bold");
  titleRange.setFontFamily("Roboto");
  titleRange.setHorizontalAlignment("center");
  titleRange.setBackground("#ff8a80");
  titleRange.setFontColor("#ffffff");
  titleRange.setBorder(true, true, true, true, false, false, "#ffccbc", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  titleRange.setWrap(true);

  // Date
  var dateRange = sheet.getRange("A2:R2");
  dateRange.merge();
  dateRange.setValue("DOE: " + formattedDate);
  dateRange.setFontSize(18);
  dateRange.setFontFamily("Roboto");
  dateRange.setHorizontalAlignment("center");
  dateRange.setBackground("#b3e5fc");
  dateRange.setFontColor("#37474f");
  dateRange.setBorder(true, true, true, true, false, false, "#b0bec5", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  dateRange.setWrap(true);

  // Headers (Row 3)
  var headerRange = sheet.getRange("A3:R3");
  if (sheetData.length > 0) {
    sheet.getRange(3, 1, 1, sheetData[0].length).setValues([sheetData[0]]);
  } else {
    sheet.getRange(3, 1, 1, sheetData[0].length).setValue(['No data available']);
  }
  headerRange.setFontSize(12);
  headerRange.setFontWeight("bold");
  headerRange.setFontFamily("Roboto");
  headerRange.setFontColor("#263238");
  headerRange.setBorder(true, true, true, true, true, true, "#ffb300", SpreadsheetApp.BorderStyle.SOLID);
  headerRange.setHorizontalAlignment("center");
  headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Lighter backgrounds for header groups
  const headerColors = [
    '#ffe082', '#ffe082', '#ffe082', // A-C
    '#a5d6a7', '#a5d6a7', // D-E Fuel
    '#90caf9', '#90caf9', // F-G Er DA
    '#ffcc80', '#ffcc80', // H-I Car Rent
    '#ffcc80', // J Car Num
    '#ce93d8', '#ce93d8', // K-L Airtime
    '#ef9a9a', '#ef9a9a', // M-N Misc
    '#ffe082', '#ffe082', '#ffe082', // O-Q New fields
    '#ffe082' // R Remarks
  ];
  for (let col = 1; col <= headerColors.length; col++) {
    sheet.getRange(3, col).setBackground(headerColors[col - 1]);
  }

  // Data Rows
  var dataRows = sheetData.slice(1, -1) || []; // Start from 1 since 0 is headers
  if (dataRows.length > 0) {
    var dataRange = sheet.getRange(4, 1, dataRows.length, sheetData[0].length);
    dataRange.setValues(dataRows);
    var dataRangeStyled = sheet.getRange(4, 1, dataRows.length, sheetData[0].length);
    dataRangeStyled.setFontSize(12);
    dataRangeStyled.setFontFamily("Roboto");
    dataRangeStyled.setHorizontalAlignment("center");
    sheet.getRange(4, 1, dataRows.length, 1).setHorizontalAlignment("right");
    dataRangeStyled.setWrap(true);

    // Alternating row colors
    var rowBackgrounds = dataRows.map((_, i) => Array(sheetData[0].length).fill(i % 2 === 0 ? "#e0f7fa" : "#ffffff"));
    dataRangeStyled.setBackgrounds(rowBackgrounds);

    // Conditional formatting for high amounts (>100000) and mismatches
    var amountColumns = [3, 5, 7, 9, 12, 14, 17]; // Updated
    dataRows.forEach((row, rowIndex) => {
      // Mismatch highlight: Beneficiary (col 1) != Account Holder (col 2)
      const benefCell = sheet.getRange(4 + rowIndex, 1);
      const accHolderCell = sheet.getRange(4 + rowIndex, 2);
      if (row[0] !== row[1]) {
        benefCell.setBackground("#add8e6"); // Light blue for mismatch
        accHolderCell.setBackground("#add8e6");
      }

      amountColumns.forEach(colIndex => {
        var cellValue = parseFloat(row[colIndex - 1]) || 0;
        var cell = sheet.getRange(4 + rowIndex, colIndex);
        if (cellValue > 100000) {
          cell.setBackground("#fff59d"); // Light yellow for high amounts
          cell.setFontColor("#f57f17"); // Dark yellow text
        } else if (cellValue < 0 || isNaN(cellValue) && row[colIndex - 1]) {
          cell.setBackground("#f06292"); // Pastel pink for errors
          cell.setFontColor("#b71c1c"); // Dark red
        }
        // Format amounts with millions separator
        if (row[colIndex - 1]) {
          cell.setValue(parseFloat(row[colIndex - 1]).toLocaleString('en-US'));
          cell.setNumberFormat('#,##0.00');
        }
        // For date columns, ensure line breaks
        var dateColumns = [4, 6, 8, 11, 13]; // Updated
        dateColumns.forEach(dateCol => {
          sheet.getRange(4 + rowIndex, dateCol).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        });
      });
    });

    dataRangeStyled.setBorder(true, true, true, true, true, true, "#b0bec5", SpreadsheetApp.BorderStyle.SOLID);
  }

  // Totals Row
  var totalsRow = sheetData[sheetData.length - 1] || [];
  if (totalsRow.length > 0) {
    var totalRowNum = 4 + dataRows.length;
    var totalRange = sheet.getRange(totalRowNum, 1, 1, sheetData[0].length);
    sheet.getRange(totalRowNum, 1, 1, 3).merge(); // Merge A:C for "Field Total : Amt"
    var grandTotal = totalsRow[2]; // The grand total value
    sheet.getRange(totalRowNum, 1).setValue("Field Total : " + grandTotal);
    var adjustedTotals = totalsRow.slice(3); // keep positions from D onward
    sheet.getRange(totalRowNum, 4, 1, adjustedTotals.length).setValues([adjustedTotals]);
    var totalRangeStyled = sheet.getRange(totalRowNum, 1, 1, sheetData[0].length);
    totalRangeStyled.setFontSize(14);
    totalRangeStyled.setFontWeight("bold");
    totalRangeStyled.setFontFamily("Roboto");
    totalRangeStyled.setBackground("#ab47bc");
    totalRangeStyled.setFontColor("#ffffff");
    totalRangeStyled.setBorder(true, true, true, true, true, true, "#7b1fa2", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    totalRangeStyled.setHorizontalAlignment("center");

    // Highlight high amounts in totals row
    var amountColumnsTotals = [4, 6, 8, 11, 13, 17]; // Adjusted positions after merge (starting from D=4)
    amountColumnsTotals.forEach((colIndex, idx) => {
      var cellValue = parseFloat(adjustedTotals[idx].replace(/[^0-9.-]+/g, '')) || 0; // Strip labels
      var cell = sheet.getRange(totalRowNum, colIndex);
      if (cellValue > 100000) {
        cell.setBackground("#fff59d"); // Override purple with yellow for high
        cell.setFontColor("#f57f17");
      }
      cell.setNumberFormat('#,##0.00');
    });
  }

  // Set column widths and auto-resize with smart calculation
  setSmartColumnWidths(sheet, sheet.getLastRow(), sheetData[0].length);
  sheet.getRange(1, 1, sheet.getLastRow(), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Force wrap on column 1
  sheet.setColumnWidth(18, 250); // Keep Remarks wider
  sheet.autoResizeRows(1, sheet.getLastRow());

  // Flush changes
  SpreadsheetApp.flush();

  // Convert to PDF with landscape orientation and narrower margins
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf&format=pdf" +
    "&size=A4" +
    "&portrait=false" +
    "&scale=4" +
    "&top_margin=0.5" +
    "&bottom_margin=0.5" +
    "&left_margin=0.3" +
    "&right_margin=0.3" +
    "&sheetnames=false&printtitle=false&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=true" + // Repeat frozen rows on each page
    "&gid=" + sheet.getSheetId();

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });

  var blob = response.getBlob();
  var pdfFile = DriveApp.createFile(blob);
  pdfFile.setName('Fund Req TZ - ' + doe + '.pdf');

  // Send email
  var subject = 'Fund Req TZ Summary - ' + doe + ' - ' + project + ' - ' + username;
  var body = 'Dear Team,\n\n' +
             'Please find attached the Fund Request Summary submitted on ' + doe + ' by ' + userEmail + '.\n' +
             'Username: ' + username + '\n' +
             'Official Email: ' + (offEmail || 'Not found') + '\n\n' +
             'Best regards,\nFund Req TZ Team';
  var toRecipients = 'eteltanzaniafinance@gmail.com';
  var ccRecipients = offEmail || '';
  try {
    MailApp.sendEmail({
      to: toRecipients,
      cc: ccRecipients,
      subject: subject,
      body: body,
      attachments: [pdfFile]
    });
    Logger.log('Email sent successfully to ' + toRecipients + ' with CC: ' + ccRecipients + ' and PDF attachment: ' + pdfFile.getName());
  } catch (e) {
    Logger.log('Failed to send email: ' + e.message);
    throw new Error('Email sending failed: ' + e.message);
  }

  // Cleanup
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  pdfFile.setTrashed(true);
  return "Email sent successfully";
}

// New function for smart column widths
function setSmartColumnWidths(sheet, numRows, numCols) {
  for (let col = 1; col <= numCols; col++) {
    let maxLen = 0;
    for (let row = 1; row <= numRows; row++) {
      const val = sheet.getRange(row, col).getValue().toString();
      maxLen = Math.max(maxLen, val.length);
    }
    let width = Math.min(250, maxLen * 8 + 20); // 8px per char + padding, cap at 250px
    if (col === 1) width = Math.min(180, width); // Special cap for "Name of Beneficiary"
    if (col === 3) width = Math.min(200, width); // Extra space for "Total Amt" to avoid truncation
    sheet.setColumnWidth(col, width);
  }
}

function getLatestCarForTeam(team) {
  const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  const carTPSheet = ss.getSheetByName('CarT_P');
  if (!carTPSheet) return {};
  const data = carTPSheet.getDataRange().getValues();
  let teamRows = data.slice(1).filter(row => row[2].toString().trim() === team);
  teamRows.sort((a,b) => new Date(b[0]) - new Date(a[0]));
  if (teamRows.length > 0) {
    const latest = teamRows[0];
    if (latest[9].toString().trim() === 'Release') {
      return {}; // Return empty if latest is Release
    }
    return {
      carNumber: latest[3].toString().trim(),
      make: latest[4].toString().trim(),
      model: latest[5].toString().trim(),
      carType: latest[6].toString().trim(),
      contract: latest[7].toString().trim(),
      owner: latest[8].toString().trim(),
      inUseRelease: latest[9].toString().trim()
    };
  }
  return {};
}

function saveCarData(team, carNumber, make, model, carType, contract, owner, inUseRelease) {
  const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  const carTPSheet = ss.getSheetByName('CarT_P');
  if (!carTPSheet) {
    carTPSheet = ss.insertSheet('CarT_P');
    carTPSheet.appendRow(['Date and time of entry', 'Project', 'Team', 'Car Number', 'Make', 'Model', 'Type', 'Contract', 'Owner', 'In Use/Release', 'Submitter']);
  }
  const username = getLoggedInUsername();
  carTPSheet.appendRow([new Date(), '', team, carNumber, make, model, carType, contract, owner, inUseRelease, username]);
  return 'Saved';
}

function getCarLookup() {
  const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  const cardataSheet = ss.getSheetByName('Cardata');
  if (!cardataSheet) return {existing: [], modelByMake: {}, typeByModel: {}, makes: [], contracts: []};
  const lastRow = cardataSheet.getLastRow();
  const existing = cardataSheet.getRange(2, 3, lastRow - 1, 8).getValues().map(row => row.map(v => v.toString().trim())); // C-J: Car Number to In Use/Release
  const makeModelType = cardataSheet.getRange(2, 14, lastRow - 1, 3).getValues(); // N-P
  let modelByMake = {};
  let typeByModel = {};
  makeModelType.forEach(row => {
    const make = row[0].toString().trim();
    const model = row[1].toString().trim();
    const type = row[2].toString().trim();
    if (make && model) {
      if (!modelByMake[make]) modelByMake[make] = [];
      if (!modelByMake[make].includes(model)) modelByMake[make].push(model);
    }
    if (model && type) {
      if (!typeByModel[model]) typeByModel[model] = [];
      if (!typeByModel[model].includes(type)) typeByModel[model].push(type);
    }
  });
  const makes = Object.keys(modelByMake).sort();
  const contracts = [...new Set(cardataSheet.getRange(2, 17, lastRow - 1, 1).getValues().flat().map(v => v.toString().trim()).filter(v => v))].sort();
  return {existing, modelByMake, typeByModel, makes, contracts};
}

function logCarRelease(team) {
  const ss = SpreadsheetApp.openById('14GNxC-Oaqds4BvaKXtLbcSFLRzY3nJbF8JCqnzKpECs');
  const carTPSheet = ss.getSheetByName('CarT_P');
  if (!carTPSheet) {
    carTPSheet = ss.insertSheet('CarT_P');
    carTPSheet.appendRow(['Date and time of entry', 'Project', 'Team', 'Car Number', 'Make', 'Model', 'Type', 'Contract', 'Owner', 'In Use/Release', 'Submitter']);
  }
  const data = carTPSheet.getDataRange().getValues();
  let teamRows = data.slice(1).filter(row => row[2].toString().trim() === team);
  teamRows.sort((a,b) => new Date(b[0]) - new Date(a[0]));
  if (teamRows.length > 0) {
    const latest = teamRows[0];
    const username = getLoggedInUsername();
    carTPSheet.appendRow([
      new Date(),
      latest[1], // Project
      team,
      latest[3], // Car Number
      latest[4], // Make
      latest[5], // Model
      latest[6], // Type
      latest[7], // Contract
      latest[8], // Owner
      'Release',
      username
    ]);
    return 'Logged';
  }
  return 'No current car to release';
}

/**
 * Serve the DatesLab HTML as a string if needed on client
 */
function getDatesLabHtml() {
  return HtmlService.createHtmlOutputFromFile('DatesLab').getContent();
}