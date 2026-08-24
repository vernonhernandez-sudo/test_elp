const SHEET_ID = '1bzJHxEfv0iHxIR-7pEFTFk4ujnpw6LWv_7dikkRMflQ';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('ELP').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ========== LOGIN & SESSION MANAGEMENT ==========

// Sales Partner / Lead Gen Specialist shift settings
const SHIFT_TIMEZONE = 'Asia/Manila';
const SHIFT_START_HOUR = 0; // 12:00 AM PHT
const SHIFT_END_HOUR = 10;   // 10:00 AM PHT

const SESSION_DURATION_MS = 10 * 60 * 60 * 1000 + 30 * 60 * 1000;

/**
 * Returns the current date/time in Philippine Standard Time.
 */
function getPHTNow() {
  return new Date(
    Utilities.formatDate(new Date(), SHIFT_TIMEZONE, "yyyy/MM/dd HH:mm:ss")
  );
}

/**
 * Returns the PHT shift key.
 *
 * Example:
 * 2026-08-21 03:00 PHT -> "2026-08-21"
 * 2026-08-21 10:00 PHT -> "2026-08-21"
 *
 * The shift itself is only available from 12 AM to 10 AM.
 */
function getCurrentShiftKey() {
  var now = new Date();

  return Utilities.formatDate(
    now,
    SHIFT_TIMEZONE,
    'yyyy-MM-dd'
  );
}

/**
 * Checks whether the current time is inside the
 * 12:00 AM - 10:00 AM PHT login window.
 */
function isWithinSalesShift() {
  var now = new Date();

  var hour = Number(
    Utilities.formatDate(now, SHIFT_TIMEZONE, 'H')
  );

  return hour >= SHIFT_START_HOUR && hour < SHIFT_END_HOUR;
}

/**
 * Gets the end time of the current PHT shift.
 *
 * The shift ends at 10:00 AM PHT on the current date.
 */
function getShiftEndTimestamp() {
  var dateString = Utilities.formatDate(
    new Date(),
    SHIFT_TIMEZONE,
    'yyyy-MM-dd'
  );

  var shiftEndString = dateString + ' 10:00:00';

  var parts = shiftEndString.match(
    /(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/
  );

  // Convert PHT local time to a real timestamp.
  var utcMillis = Date.UTC(
    Number(parts[1]),
    Number(parts[2]) - 1,
    Number(parts[3]),
    Number(parts[4]) - 8, // PHT = UTC+8
    Number(parts[5]),
    Number(parts[6])
  );

  return utcMillis;
}

/**
 * Stores the session in Script Properties.
 *
 * CacheService cannot safely hold an 8h10m session because
 * Apps Script CacheService has a maximum expiration of 6 hours.
 */
function saveSession(sessionId, sessionData) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      'SESSION_' + sessionId,
      JSON.stringify(sessionData)
    );
}

/**
 * Retrieves a session.
 */
function getStoredSession(sessionId) {
  if (!sessionId) return null;

  try {
    var raw = PropertiesService
      .getScriptProperties()
      .getProperty('SESSION_' + sessionId);

    if (!raw) return null;

    return JSON.parse(raw);

  } catch (e) {
    return null;
  }
}

/**
 * Removes a session.
 */
function removeStoredSession(sessionId) {
  if (!sessionId) return;

  try {
    PropertiesService
      .getScriptProperties()
      .deleteProperty('SESSION_' + sessionId);
  } catch (e) {}
}

/**
 * Determines whether a user is subject to shift restrictions.
 */
function isShiftRestrictedRole(role) {

  role = role ? role.toString().trim() : '';

  return role === 'Admin' ||
         role === 'Sales Partner' ||
         role === 'Lead Gen Specialist';
}

/**
 * Login
 */
function loginUser(email, password) {

  if (!email || !password) {
    return {
      success: false,
      message: 'Email and password are required.'
    };
  }

  try {

    var sheet = SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName('Agents');

    if (!sheet) {
      return {
        success: false,
        message: 'Agents sheet not found.'
      };
    }

    var data = sheet.getDataRange().getValues();

    // Remove header row
    data.shift();

    var loginEmail = email
      .toString()
      .trim()
      .toLowerCase();

    /*
     * =====================================================
     * SERVER-SIDE EMAIL DOMAIN VALIDATION
     *
     * This is important.
     *
     * JavaScript validation in Login.html can be bypassed,
     * so we also validate the domain here.
     * =====================================================
     */

    var emailPattern =
      /^[A-Z0-9._%+-]+@explorabooks\.com$/i;

    if (!emailPattern.test(loginEmail)) {

      return {
        success: false,
        message:
          'Please use your @explorabooks.com email address.'
      };
    }

    for (var i = 0; i < data.length; i++) {

      var row = data[i];

      /*
       * Column E = Email
       */
      var storedEmail = row[4]
        ? row[4].toString().trim().toLowerCase()
        : '';

      /*
       * Column C = Password
       */
      var storedPassword = row[2]
        ? row[2].toString().trim()
        : '';

      /*
       * Match Email + Password
       */
      if (
        storedEmail === loginEmail &&
        storedPassword === password
      ) {

        /*
         * =================================================
         * ACCOUNT STATUS
         * Column I
         * =================================================
         */

        if (
          row[8] &&
          row[8].toString().trim() !== 'Active'
        ) {

          return {
            success: false,
            message: 'Account is not active.'
          };
        }

        /*
         * =================================================
         * ROLE
         * Column H
         * =================================================
         */

        var role = row[7]
          ? row[7].toString().trim()
          : '';

        /*
         * =================================================
         * USER OBJECT
         * =================================================
         */

        var user = {

          id: row[0],

          name: row[3],

          email: row[4],

          role: role
        };

        /*
         * =================================================
         * SHIFT LOGIN WINDOW
         *
         * Sales Partner and Lead Gen Specialist can only
         * START a session between 12 AM and 10 AM PHT.
         *
         * IMPORTANT:
         *
         * There is NO one-login-per-shift restriction.
         *
         * They can logout and login again as many times as
         * they want during the 12 AM - 10 AM window.
         * =================================================
         */

        var shiftRestricted =
          isShiftRestrictedRole(role);

        if (shiftRestricted) {

          if (!isWithinSalesShift()) {

            return {
              success: false,
              message:
                'Your shift is currently closed. Admin, Sales Partner, and Lead Gen Specialist accounts can only log in from 12:00 AM to 10:00 AM Philippine Standard Time.'
            };
          }
        }

        /*
         * =================================================
         * CREATE SESSION
         * =================================================
         */

        var sessionId =
          Utilities.getUuid();

        var loginTime =
          Date.now();

        /*
         * Default session duration.
         *
         * We will still cap restricted users at 10 AM below.
         */
        var sessionExpiry =
          loginTime + SESSION_DURATION_MS;

        /*
         * =================================================
         * HARD 10:00 AM CUTOFF
         *
         * Even if SESSION_DURATION_MS goes beyond 10 AM,
         * restricted users can NEVER remain logged in
         * beyond the current shift.
         * =================================================
         */

        if (shiftRestricted) {

          var shiftEnd =
            getShiftEndTimestamp();

          if (shiftEnd < sessionExpiry) {

            sessionExpiry =
              shiftEnd;
          }
        }

        /*
         * =================================================
         * SESSION DATA
         * =================================================
         */

        var sessionData = {

          id: user.id,

          name: user.name,

          email: user.email,

          role: user.role,

          loginTime: loginTime,

          expiresAt: sessionExpiry,

          shiftRestricted: shiftRestricted
        };

        /*
         * Save session
         */
        saveSession(
          sessionId,
          sessionData
        );

        /*
         * =================================================
         * SUCCESS
         * =================================================
         */

        return {

          success: true,

          sessionId: sessionId,

          user: {

            id: user.id,

            name: user.name,

            email: user.email,

            role: user.role
          }
        };
      }
    }

    /*
     * No matching account
     */
    return {
      success: false,
      message: 'Invalid email or password.'
    };

  } catch (error) {

    return {
      success: false,
      message:
        'Error: ' + error.toString()
    };
  }
}

/**
 * Gets the currently logged-in user.
 *
 * This is now also responsible for checking:
 * - session expiration
 * - shift expiration
 * - whether the user belongs to the correct shift
 */
function getUserFromCache(sessionId) {

  if (!sessionId) {
    return null;
  }

  var session =
    getStoredSession(sessionId);

  if (!session) {
    return null;
  }

  var now =
    Date.now();

  /*
   * =====================================================
   * SESSION EXPIRATION
   * =====================================================
   */

  if (
    !session.expiresAt ||
    now >= Number(session.expiresAt)
  ) {

    removeStoredSession(sessionId);

    return null;
  }

  /*
   * =====================================================
   * ADDITIONAL SHIFT SAFETY CHECK
   *
   * If the user is a restricted role and the shift is
   * already over, immediately invalidate the session.
   * =====================================================
   */

  if (session.shiftRestricted) {

    if (!isWithinSalesShift()) {

      removeStoredSession(sessionId);

      return null;
    }
  }

  return session;
}

/**
 * Explicit frontend session validation.
 */
function validateSession(sessionId) {

  var user = getUserFromCache(sessionId);

  if (!user) {
    return {
      success: false,
      expired: true,
      message: 'Your session has expired. Please log in again.'
    };
  }

  return {
    success: true,
    expired: false,
    user: {
      id: user.id,
      name: user.name,
      email: user.email,
      role: user.role
    },
    expiresAt: user.expiresAt
  };
}

/**
 * Logout
 */
function logoutUser(sessionId) {

  removeStoredSession(sessionId);

  return {
    success: true
  };
}

// ========== ENTRY ACCESS CONTROL ==========

function canAccessEntry(user, entryAssignedAgentId) {
  if (!user) return false;

  var role = user.role ? user.role.toString().trim() : '';

  // Superadmin and Admin can access all entries
  if (role === 'Super Admin' || role === 'Admin') {
    return true;
  }

  // Sales Partner can ONLY access their own assigned leads
  if (role === 'Sales Partner') {
    return String(entryAssignedAgentId) === String(user.id);
  }

  // Lead Gen Specialist will be handled later
  if (role === 'Lead Gen Specialist') {
    return false;
  }

  return false;
}


// ========== FIND ENTRY BY ID ==========

function getEntryById(entryId) {

  var sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName('Entries');

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {

    if (String(data[i][0]) === String(entryId)) {

      return {
        row: i + 1,
        id: data[i][0],
        assignedAgentId: data[i][7],
        authorName: data[i][1],
        book: data[i][4]
      };

    }

  }

  return null;
}

    // ========== UNIQUE ENTRY ID ==========
    function generateEntryId(data) {
      var maxId = 0;

      for (var i = 1; i < data.length; i++) {
        var id = Number(data[i][0]);

        if (!isNaN(id) && id > maxId) {
          maxId = id;
        }
      }

      return maxId + 1;
    }

// ========== ENTRY MANAGEMENT (Preventing duplicate Entry IDs) ==========
function addEntry(sessionId, entryData) {
  try {

    var user = getUserFromCache(sessionId);

    if (!user) {
      return {
        success: false,
        message: 'Session expired.'
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Entries');
    var agentsSheet = ss.getSheetByName('Agents');

    if (!entryData.authorName || !entryData.book || !entryData.isbn) {
      return {
        success: false,
        message: 'Author Name, Book, and ISBN are required.'
      };
    }

    // Find the role of the person who mined the lead
    var userRole = '';

    var agents = agentsSheet.getDataRange().getValues();

    for (var i = 1; i < agents.length; i++) {
      if (agents[i][0] == user.id) {
        userRole = agents[i][7]
          ? agents[i][7].toString().trim()
          : '';
        break;
      }
    }

    var data = sheet.getDataRange().getValues();

    // Generate a unique ID
    var newEntryId = generateEntryId(data);

    /*
     * Initial status:
     *
     * Admin / Sales Partner mining their own lead
     *     → Exclusive
     *
     * Lead Gen Specialist
     *     → No Status
     *     → System distributes it
     */
    var initialStatus = '';

    var initialStatus = '';

    if (
      userRole === 'Admin' ||
      userRole === 'Sales Partner'
    ) {
      initialStatus = 'Exclusive';
    } else if (
      userRole === 'Lead Gen Specialist'
    ) {
      initialStatus = '';
    }

    // Format U.S. phone numbers before saving
    var formattedPhones =
      formatUSPhoneNumbers(entryData.phones);

    // Add the new entry
    sheet.appendRow([
      newEntryId,                  // A - ID
      entryData.authorName,        // B - Author
      formattedPhones,             // C - Phones
      entryData.email || '',       // D - Email
      entryData.book,              // E - Book
      entryData.isbn,              // F - ISBN
      entryData.address || '',     // G - Address
      '',                          // H - Assigned Agent
      initialStatus,               // I - Status
      new Date().toISOString(),    // J - Created/Assigned At
      'No',                        // K
      'No',                        // L
      user.id                      // M - Mined By
    ]);

    SpreadsheetApp.flush();

    logActivity(
      user.id,
      user.name,
      'Created entry #' + newEntryId + ': ' + entryData.authorName
    );

    /*
     * Only Lead Gen Specialist leads are automatically distributed.
     *
     * Admin and Sales Partner leads remain with the miner
     * and start as Exclusive.
     */
    if (userRole === 'Lead Gen Specialist') {
      distributeNewEntry(newEntryId);
    }

    return {
      success: true,
      message: 'Entry added!',
      entryId: newEntryId
    };

  } catch (e) {

    console.error('addEntry error:', e);

    return {
      success: false,
      message: 'Unable to add entry.'
    };
  }
}

// ========== PHONE NUMBER FORMATTING ==========

function formatUSPhoneNumber(phone) {

  if (!phone) return '';

  var original = phone.toString().trim();

  // Remove common formatting characters
  var digits = original.replace(/\D/g, '');

  /*
   * Only format U.S. numbers.
   *
   * 10 digits:
   * 9186917015
   *
   * 11 digits beginning with 1:
   * 19186917015
   */

  if (digits.length === 10) {

    return '(' +
      digits.substring(0, 3) +
      ') ' +
      digits.substring(3, 6) +
      '-' +
      digits.substring(6, 10);

  }

  if (
    digits.length === 11 &&
    digits.charAt(0) === '1'
  ) {

    digits = digits.substring(1);

    return '(' +
      digits.substring(0, 3) +
      ') ' +
      digits.substring(3, 6) +
      '-' +
      digits.substring(6, 10);

  }

  /*
   * Anything that doesn't clearly look like a
   * U.S. phone number is left unchanged.
   */

  return original;
}

function formatUSPhoneNumbers(phones) {

  if (!phones) return '';

  return phones
    .toString()
    .split(',')
    .map(function(phone) {
      return formatUSPhoneNumber(phone);
    })
    .join(', ');
}

function updateEntry(sessionId, entryId, entryData) {
  try {
    var u = getUserFromCache(sessionId);

    if (!u) {
      return {
        success: false,
        message: 'Session expired. Please sign in again.',
        code: 'SESSION_EXPIRED'
      };
    }

    var s = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries');

    if (!s) {
      console.error('updateEntry: Entries sheet not found.');
      return {
        success: false,
        message: 'Unable to update the entry right now.',
        code: 'SHEET_NOT_FOUND'
      };
    }

    var d = s.getDataRange().getValues();

    for (var i = 1; i < d.length; i++) {

      if (d[i][0] == entryId) {

        s.getRange(i + 1, 2).setValue(entryData.authorName || '');
        s.getRange(i + 1, 3).setValue(entryData.phones || '');
        s.getRange(i + 1, 4).setValue(entryData.email || '');
        s.getRange(i + 1, 5).setValue(entryData.book || '');
        s.getRange(i + 1, 6).setValue(entryData.isbn || '');
        s.getRange(i + 1, 7).setValue(entryData.address || '');

        SpreadsheetApp.flush();

        logActivity(
          u.id,
          u.name,
          'Updated entry #' + entryId
        );

        return {
          success: true,
          message: 'Updated!'
        };
      }
    }

    return {
      success: false,
      message: 'Entry not found.',
      code: 'ENTRY_NOT_FOUND'
    };

  } catch (e) {

    // Detailed error is visible in Apps Script execution logs,
    // but is NOT exposed to the user.
    console.error(
      'updateEntry failed. Entry ID: ' +
      entryId +
      ' | Error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );

    return {
      success: false,
      message: 'Unable to update the entry right now. Please try again.',
      code: 'UPDATE_FAILED'
    };
  }
}

// function deleteEntry(sessionId, entryId) {
//   try {
//     var u=getUserFromCache(sessionId); 
//     if(!u)return{success:false,message:'Session expired.'};
//     var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Entries'); 
//     var d=s.getDataRange().getValues();
//     for(var k=d.length-1;k>=1;k--){
//       if(d[k][0]==entryId){
//         s.deleteRow(k+1);
//         logActivity(u.id,u.name,'Deleted entry #'+entryId);
//         return{success:true,message:'Entry deleted.'};
//       }
//     }
//     return{success:false,message:'Not found.'};
//   } catch(e){return{success:false,message:'Error: '+e.toString()};}
// }

// ========== ENTRY ACCESS CONTROL ==========

function canAccessEntry(user, entryAssignedAgentId, entryMinedById) {

  if (!user) return false;

  var role =
    user.role
      ? user.role.toString().trim()
      : '';

  /*
   * =====================================================
   * SUPER ADMIN / ADMIN
   *
   * Can see every entry.
   * =====================================================
   */

  if (
    role === 'Super Admin' ||
    role === 'Admin'
  ) {
    return true;
  }

  /*
   * =====================================================
   * SALES PARTNER
   *
   * Can see:
   *
   * 1. Leads assigned to them
   * 2. Leads they personally mined
   *
   * This is important because Sales Partner-created
   * entries have no Assigned Agent yet.
   * =====================================================
   */

  if (role === 'Sales Partner') {

    var isAssignedToUser =
      String(entryAssignedAgentId) === String(user.id);

    var isMinedByUser =
      String(entryMinedById) === String(user.id);

    return isAssignedToUser || isMinedByUser;
  }

  /*
   * =====================================================
   * LEAD GEN SPECIALIST
   *
   * Will be handled separately when we implement
   * Lead Gen Specialist access/distribution.
   * =====================================================
   */

  if (role === 'Lead Gen Specialist') {
    return false;
  }

  return false;
}

// ========== GET ENTRIES ==========

function getEntries(sessionId) {

  try {

    var u = getUserFromCache(sessionId);

    if (!u) {
      return {
        success: false,
        entries: [],
        summary: {
          pipeCount: 0,
          exclusiveCount: 0,
          vmCount: 0,
          dncCount: 0,
          soldCount: 0,
          totalAssigned: 0
        }
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);

    var entrySheet = ss.getSheetByName('Entries');
    var agentSheet = ss.getSheetByName('Agents');

    var ed = entrySheet.getDataRange().getValues();

    if (ed.length <= 1) {
      return {
        success: true,
        entries: [],
        summary: {
          pipeCount: 0,
          exclusiveCount: 0,
          vmCount: 0,
          dncCount: 0,
          soldCount: 0,
          totalAssigned: 0
        }
      };
    }

    ed.shift();

    // Build Agent ID -> Name map
    var ad = agentSheet.getDataRange().getValues();
    ad.shift();

    var am = {};

    for (var k = 0; k < ad.length; k++) {
      am[ad[k][0]] = ad[k][3];
    }

    var entries = [];

    for (var i = 0; i < ed.length; i++) {

      //This line of codes checked both assigned leads and mined by
      var assignedAgentId = ed[i][7];
      var minedById = ed[i][12];

      if (!canAccessEntry(
        u,
        assignedAgentId,
        minedById
      )) {
        continue;
      }

      entries.push({
      id: ed[i][0],
      authorName: ed[i][1] || '',
      phones: ed[i][2] || '',
      email: ed[i][3] || '',
      book: ed[i][4] || '',
      isbn: ed[i][5] || '',
      address: ed[i][6] || '',
      assignedAgentId: ed[i][7],
      assignedAgentName: am[ed[i][7]] || 'Own Lead',
      status: ed[i][8] || '',
      createdAt: ed[i][9] || '',
      minedById: ed[i][12] || ''
    });

    }

    return {

      success: true,

      entries: entries,

      summary: {

        pipeCount: entries.filter(function(e) {
          return e.status === 'Pipe';
        }).length,

        exclusiveCount: entries.filter(function(e) {
          return e.status === 'Exclusive';
        }).length,

        vmCount: entries.filter(function(e) {
          return e.status === 'VM';
        }).length,

        dncCount: entries.filter(function(e) {
          return e.status === 'DNC';
        }).length,

        soldCount: entries.filter(function(e) {
          return e.status === 'Sold';
        }).length,

        totalAssigned: entries.length

      }

    };

  } catch (e) {

    return {
      success: false,
      entries: [],
      summary: {
        pipeCount: 0,
        exclusiveCount: 0,
        vmCount: 0,
        dncCount: 0,
        soldCount: 0,
        totalAssigned: 0
      }
    };

  }

}

// ========== STATUS MANAGEMENT ==========
function updateEntryStatus(sessionId, entryId, newStatus) {
  try {
    var u = getUserFromCache(sessionId);

    if (!u) {
      return {
        success: false,
        message: 'Session expired.'
      };
    }

    var validStatuses = ['', 'Pipe', 'Exclusive', 'Sold', 'DNC', 'VM'];

    if (validStatuses.indexOf(newStatus) === -1) {
      return {
        success: false,
        message: 'Invalid status.'
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var entrySheet = ss.getSheetByName('Entries');
    var agentsSheet = ss.getSheetByName('Agents');

    if (!entrySheet || !agentsSheet) {
      return {
        success: false,
        message: 'Required sheet not found.'
      };
    }

    var entryData = entrySheet.getDataRange().getValues();
    var agentData = agentsSheet.getDataRange().getValues();

    var entryRow = null;

    // Find the entry
    for (var i = 1; i < entryData.length; i++) {
      if (String(entryData[i][0]) === String(entryId)) {
        entryRow = entryData[i];
        break;
      }
    }

    if (!entryRow) {
      return {
        success: false,
        message: 'Entry not found.'
      };
    }

    /*
     * Entries columns:
     *
     * [0] ID
     * [1] Author
     * [2] Phones
     * [3] Email
     * [4] Book
     * [5] ISBN
     * [6] Address
     * [7] Assigned Agent ID
     * [8] Status
     * [9] Created/Assigned At
     * [10] ...
     * [11] ...
     * [12] Mined By
     */

    var currentStatus = entryRow[8]
      ? entryRow[8].toString().trim()
      : '';

          /*
      * =====================================================
      * SOLD IS FINAL
      *
      * Once an entry is Sold, its status cannot be changed.
      * =====================================================
      */

      if (currentStatus === 'Sold') {

        return {
          success: false,
          message: 'This entry is already marked as Sold and its status cannot be changed.'
        };
      }

    var minedById = entryRow[12]
      ? entryRow[12].toString().trim()
      : '';

    /*
     * Find the role of the logged-in user.
     */
    var userRole = '';

    for (var j = 1; j < agentData.length; j++) {
      if (String(agentData[j][0]) === String(u.id)) {
        userRole = agentData[j][7]
          ? agentData[j][7].toString().trim()
          : '';
        break;
      }
    }

    /*
     * SALES PARTNER RULE
     *
     * A Sales Partner cannot manually return a
     * Lead Gen Specialist lead to "- No Status -".
     *
     * A Sales Partner also cannot change a Lead Gen
     * Specialist lead to "Exclusive".
     *
     * Exclusive is only for leads mined by the
     * Sales Partner or Admin themselves.
     */
    if (userRole === 'Sales Partner') {

      var isOwnLead =
        String(minedById) === String(u.id);

      // Sales Partners cannot use "No Status"
      // on a lead that they did not mine.
      if (newStatus === '' && !isOwnLead) {
        return {
          success: false,
          message: 'Sales Partners cannot return a Lead Gen Specialist lead to No Status.'
        };
      }

      // Sales Partners cannot mark another person's
      // lead as Exclusive.
      if (newStatus === 'Exclusive' && !isOwnLead) {
        return {
          success: false,
          message: 'Only the person who mined the lead can mark it as Exclusive.'
        };
      }
    }

    /*
     * ADMIN RULE
     *
     * Admins are allowed to manually use:
     * - No Status
     * - Pipe
     * - Exclusive
     * - VM
     * - DNC
     * - Sold
     *
     * Therefore no restriction is applied here.
     */

    entrySheet.getRange(
      entryData.findIndex(function(row) {
        return String(row[0]) === String(entryId);
      }) + 1,
      9
    ).setValue(newStatus);

    logActivity(
      u.id,
      u.name,
      'Status #' + entryId + ' to ' + (newStatus || 'None')
    );

    addSystemRemark(
      entryId,
      u.name,
      'Status changed to ' + (newStatus || 'None')
    );

    return {
      success: true,
      message: 'Updated!'
    };

  } catch (e) {

    console.error('updateEntryStatus error:', e);

    return {
      success: false,
      message: 'Unable to update status.'
    };
  }
}

// ========== REMARKS ==========
function addRemark(sessionId, entryId, remarkText) {
  try { 
    var u=getUserFromCache(sessionId); 
    if(!u||!remarkText||!remarkText.trim())return{success:false,message:'Invalid.'}; 
    var s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Remarks'); 
    s.appendRow([s.getDataRange().getValues().length,entryId,new Date().toISOString(),u.id,u.name,remarkText.trim()]); 
    logActivity(u.id,u.name,'Added remark to #'+entryId); 
    return{success:true,message:'Remark added!'}; 
  } catch(e){return{success:false,message:'Error'};}
}

// ========== GET REMARKS ==========

function getRemarks(sessionId, entryId) {

  try {

    // Verify that the user is logged in
    var u = getUserFromCache(sessionId);

    if (!u) {
      return [];
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Remarks');

    if (!sheet) {
      console.error('getRemarks: Remarks sheet not found.');
      return [];
    }

    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    // Remove header row
    data.shift();

    var remarks = [];

    for (var i = 0; i < data.length; i++) {

      /*
       * IMPORTANT:
       *
       * Remarks belong to the Entry ID.
       *
       * We do NOT check the current assigned
       * Sales Partner here.
       *
       * Therefore, if a lead is transferred from
       * one Sales Partner to another, ALL previous
       * remarks remain visible.
       */

      if (String(data[i][1]) === String(entryId)) {

        remarks.push({

          id: data[i][0],

          entryId: data[i][1],

          timestamp: data[i][2],

          userId: data[i][3],

          userName: data[i][4],

          remark: data[i][5]

        });

      }

    }

    // Newest remark first
    remarks.sort(function(a, b) {

      return new Date(b.timestamp) -
             new Date(a.timestamp);

    });

    return remarks;

  } catch (e) {

    console.error(
      'getRemarks failed. Entry ID: ' +
      entryId +
      ' | Error: ' +
      e.message
    );

    return [];

  }

}

// ========== GET ENTRY ACTIVITY ==========

function getEntryActivity(sessionId, entryId) {

  try {

    /*
     * ==========================================
     * VERIFY LOGIN SESSION
     * ==========================================
     */

    var u = getUserFromCache(sessionId);

    if (!u) {
      return [];
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);

    var remarksSheet =
      ss.getSheetByName('Remarks');

    var activitySheet =
      ss.getSheetByName('ActivityLog');

    if (!remarksSheet || !activitySheet) {

      console.error(
        'getEntryActivity: Required sheet not found.'
      );

      return [];
    }

    var remarksData =
      remarksSheet.getDataRange().getValues();

    var activityData =
      activitySheet.getDataRange().getValues();

    /*
     * Remove header rows.
     */

    if (remarksData.length > 0) {
      remarksData.shift();
    }

    if (activityData.length > 0) {
      activityData.shift();
    }

    var items = [];

    /*
     * ==========================================
     * GET ALL REMARKS FOR THIS ENTRY
     * ==========================================
     *
     * IMPORTANT:
     *
     * We only check Entry ID.
     *
     * We do NOT check:
     *
     * - Current assigned Sales Partner
     * - Who currently owns the lead
     * - Who mined the lead
     *
     * Therefore all historical remarks remain
     * attached to this Entry.
     */

    for (var i = 0; i < remarksData.length; i++) {

      if (
        String(remarksData[i][1]) ===
        String(entryId)
      ) {

        items.push({

          type: 'remark',

          timestamp: remarksData[i][2],

          userName: remarksData[i][4],

          content: remarksData[i][5]

        });

      }

    }

    /*
     * ==========================================
     * GET ALL ACTIVITY FOR THIS ENTRY
     * ==========================================
     *
     * Your current ActivityLog stores the
     * Entry ID inside the Action text.
     *
     * Example:
     *
     * "Status #123 to Pipe"
     *
     * "Added remark to #123"
     *
     * Until we redesign the ActivityLog columns,
     * we will continue using this method.
     */

    for (var j = 0; j < activityData.length; j++) {

      var action =
        activityData[j][3]
          ? activityData[j][3].toString()
          : '';

      if (
        action.indexOf(
          '#' + entryId
        ) > -1
      ) {

        items.push({

          type: 'activity',

          timestamp: activityData[j][4],

          userName: activityData[j][2],

          content: action

        });

      }

    }

    /*
     * ==========================================
     * NEWEST FIRST
     * ==========================================
     */

    items.sort(function(a, b) {

      return new Date(b.timestamp) -
             new Date(a.timestamp);

    });

    /*
     * Limit the result to the newest 50
     * activities.
     */

    return items.slice(0, 50);

  } catch (e) {

    console.error(
      'getEntryActivity failed. Entry ID: ' +
      entryId +
      ' | Error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );

    return [];

  }

}

function logActivity(uid, un, action) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('ActivityLog');

    if (!sheet) {
      console.error('logActivity: ActivityLog sheet not found.');
      return;
    }

    var nextId = sheet.getLastRow();

    sheet.appendRow([
      nextId,
      uid,
      un,
      action,
      new Date().toISOString()
    ]);

  } catch (e) {

    console.error(
      'logActivity failed. Action: ' +
      action +
      ' | Error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );
  }
}

function addSystemRemark(eid, un, msg) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Remarks');

    if (!sheet) {
      console.error('addSystemRemark: Remarks sheet not found.');
      return;
    }

    var nextId = sheet.getLastRow();

    sheet.appendRow([
      nextId,
      eid,
      new Date().toISOString(),
      0,
      un,
      msg
    ]);

  } catch (e) {

    console.error(
      'addSystemRemark failed. Entry ID: ' +
      eid +
      ' | Error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );
  }
}

// ========== Round-Robin DISTRIBUTION ==========
function distributeNewEntry(entryId) {

  var lock = LockService.getScriptLock();

  try {

    // Prevent two simultaneous leads from getting the same agent
    lock.waitLock(10000);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var agentSheet = ss.getSheetByName('Agents');
    var entrySheet = ss.getSheetByName('Entries');

    if (!agentSheet || !entrySheet) {
      return;
    }

    var agentData = agentSheet.getDataRange().getValues();

    if (agentData.length <= 1) {
      return;
    }

    // Remove header row
    agentData.shift();

    /*
      Agents sheet columns:

      [0] = Agent ID
      [1] = Username
      [2] = Password
      [3] = Name
      [4] = Email
      [5] = ...
      [6] = ...
      [7] = Role
      [8] = Status
    */

    // Only Active Admins and Sales Partners can receive automatic leads
    var eligibleAgents = [];

    for (var i = 0; i < agentData.length; i++) {

      var agent = agentData[i];

      var agentId = agent[0];
      var agentName = agent[3];
      var role = agent[7] ? agent[7].toString().trim() : '';
      var status = agent[8] ? agent[8].toString().trim() : '';

      if (
        status === 'Active' &&
        (role === 'Admin' || role === 'Sales Partner')
      ) {

        eligibleAgents.push({
          id: agentId,
          name: agentName,
          role: role,
          rowIndex: i
        });

      }
    }

    // No eligible agents available
    if (eligibleAgents.length === 0) {
      return;
    }

    /*
      Store the ID of the last agent who received
      an automatically distributed lead.

      This property survives between function executions,
      so the round-robin sequence continues where it stopped.
    */
    var properties = PropertiesService.getScriptProperties();

    var lastAgentId = properties.getProperty('LAST_DISTRIBUTED_AGENT_ID');

    var nextAgentIndex = 0;

    if (lastAgentId !== null) {

      // Find the last agent in the current eligible list
      var lastIndex = -1;

      for (var j = 0; j < eligibleAgents.length; j++) {

        if (String(eligibleAgents[j].id) === String(lastAgentId)) {
          lastIndex = j;
          break;
        }

      }

      if (lastIndex !== -1) {

        // Continue with the next person
        nextAgentIndex = (lastIndex + 1) % eligibleAgents.length;

      } else {

        /*
          The previous agent is no longer eligible
          (possibly inactive, removed, or role changed).

          Start from the first eligible agent.
        */
        nextAgentIndex = 0;
      }
    }

    var selectedAgent = eligibleAgents[nextAgentIndex];

    /*
      Find the Entry and assign it to the selected agent.
    */
    var entryData = entrySheet.getDataRange().getValues();

    var entryFound = false;

    for (var k = 1; k < entryData.length; k++) {

      if (String(entryData[k][0]) === String(entryId)) {

        // Column 8 = Assigned Agent ID
        entrySheet.getRange(k + 1, 8).setValue(selectedAgent.id);

        // Column 10 = Assignment timestamp
        entrySheet
          .getRange(k + 1, 10)
          .setValue(new Date().toISOString());

        entryFound = true;
        break;
      }
    }

    if (!entryFound) {
      return;
    }

    /*
      Remember this agent.

      The NEXT lead will continue from the
      person immediately after this agent.
    */
    properties.setProperty(
      'LAST_DISTRIBUTED_AGENT_ID',
      String(selectedAgent.id)
    );

    SpreadsheetApp.flush();

    logActivity(
      0,
      'System',
      'Auto-assigned #' + entryId + ' to ' + selectedAgent.name
    );

    addSystemRemark(
      entryId,
      'System',
      'Lead auto-assigned to ' + selectedAgent.name
    );

  } catch (e) {

    console.error('Distribution error:', e);

  } finally {

    try {
      lock.releaseLock();
    } catch (lockError) {}

  }

}

function getAgentName(id) { 
  try { 
    var d=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues(); 
    for(var i=1;i<d.length;i++){if(d[i][0]==id)return d[i][3];} 
    return 'Unknown'; 
  } catch(e){return'Unknown';} 
}

// ========== TRANSFER TARGETS ==========

function getTransferTargets(sessionId) {
  try {

    var user = getUserFromCache(sessionId);

    if (!user) {
      return {
        success: false,
        targets: [],
        message: 'Session expired.'
      };
    }

    if (user.role !== 'Sales Partner') {
      return {
        success: false,
        targets: [],
        message: 'You are not authorized to request a transfer.'
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Agents');

    if (!sheet) {
      return {
        success: false,
        targets: [],
        message: 'Agents sheet not found.'
      };
    }

    var data = sheet.getDataRange().getValues();

    data.shift();

    var targets = [];

    for (var i = 0; i < data.length; i++) {

      var row = data[i];

      var agentId = row[0];

      var name = row[3]
        ? row[3].toString().trim()
        : '';

      var role = row[7]
        ? row[7].toString().trim()
        : '';

      var status = row[8]
        ? row[8].toString().trim()
        : '';

      if (
        status !== 'Active' ||
        (role !== 'Admin' && role !== 'Sales Partner')
      ) {
        continue;
      }

      if (String(agentId) === String(user.id)) {
        continue;
      }

      targets.push({
        id: agentId,
        name: name,
        role: role
      });
    }

    return {
      success: true,
      targets: targets
    };

  } catch (e) {

    console.error(
      'getTransferTargets error: ' +
      e.message
    );

    return {
      success: false,
      targets: [],
      message: 'Unable to load transfer targets: ' + e.message
    };
  }
}