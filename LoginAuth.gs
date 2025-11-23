/**
 * Verifies user credentials against the 'users' sheet.
 * Assumed Columns:
 * A: S.No
 * B: Name
 * C: username
 * D: password
 * E: Secret code
 * F: email
 *
 * @param {Object} creds - { username, password, secretCode }
 * @return {Object} - { success: boolean, message: string, user: Object }
 */
function verifyLogin(creds) {
  try {
    if (!creds || !creds.username || !creds.password || !creds.secretCode) {
      return { success: false, message: 'All fields are required.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('users');
    if (!sheet) {
      return { success: false, message: 'System error: User database not found.' };
    }

    const data = sheet.getDataRange().getValues();
    // Assuming Row 1 is headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const uName = String(row[2] || '').trim(); // Column C
      const uPass = String(row[3] || '').trim(); // Column D
      const uCode = String(row[4] || '').trim(); // Column E

      if (uName === creds.username && uPass === creds.password && uCode === creds.secretCode) {
        return {
          success: true,
          message: 'Login successful',
          user: {
            name: row[1],
            email: row[5]
          }
        };
      }
    }

    return { success: false, message: 'Invalid credentials.' };

  } catch (e) {
    console.error('Login error:', e);
    return { success: false, message: 'An error occurred during login.' };
  }
}

/**
 * Serves the Login page.
 * You can use this function to test the page independently if needed.
 */
function doGetLogin() {
  return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('Portal Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns the URL of the published web app.
 * Used for client-side redirection.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
