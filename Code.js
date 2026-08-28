const SHEET_ID = '1bzJHxEfv0iHxIR-7pEFTFk4ujnpw6LWv_7dikkRMflQ';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('ELP').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ========== LOGIN & SESSION MANAGEMENT ==========

// Sales Partner / Lead Gen Specialist shift settings
const SHIFT_TIMEZONE = 'Asia/Manila';
const SHIFT_START_HOUR = 23;    // 11:45 PM PHT
const SHIFT_START_MINUTE = 45;  // 11:45 PM PHT

const SHIFT_END_HOUR = 10;      // 10:00 AM PHT
const SHIFT_END_MINUTE = 0;     // 10:00 AM PHT

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
    Utilities.formatDate(
      now,
      SHIFT_TIMEZONE,
      'H'
    )
  );

  var minute = Number(
    Utilities.formatDate(
      now,
      SHIFT_TIMEZONE,
      'm'
    )
  );

  var currentMinutes =
    (hour * 60) + minute;

  var startMinutes =
    (23 * 60) + 45; // 11:45 PM

  var endMinutes =
    10 * 60; // 10:00 AM

  /*
   * Shift crosses midnight:
   *
   * 11:45 PM → 11:59 PM
   * OR
   * 12:00 AM → 9:59 AM
   */

  return (
    currentMinutes >= startMinutes ||
    currentMinutes < endMinutes
  );
}

/**
 * Gets the end time of the current PHT shift.
 *
 * The shift ends at 10:00 AM PHT on the current date.
 */
function getShiftEndTimestamp() {

  var now = new Date();

  var currentHour = Number(
    Utilities.formatDate(
      now,
      SHIFT_TIMEZONE,
      'H'
    )
  );

  var currentMinute = Number(
    Utilities.formatDate(
      now,
      SHIFT_TIMEZONE,
      'm'
    )
  );

  var currentMinutes =
    (currentHour * 60) + currentMinute;

  var endMinutes =
    10 * 60; // 10:00 AM

  /*
   * If we are currently between
   * 12:00 AM and 9:59 AM,
   * the shift ends TODAY at 10:00 AM.
   *
   * If we are currently between
   * 11:45 PM and 11:59 PM,
   * the shift ends TOMORROW at 10:00 AM.
   */

  var endDate =
    new Date(now);

  if (currentMinutes >= 0 && currentMinutes < endMinutes) {

    // Same calendar day
    endDate.setDate(
      endDate.getDate()
    );

  } else {

    // Shift started at 11:45 PM,
    // so 10:00 AM is tomorrow.
    endDate.setDate(
      endDate.getDate() + 1
    );
  }

  var dateString =
    Utilities.formatDate(
      endDate,
      SHIFT_TIMEZONE,
      'yyyy-MM-dd'
    );

  var shiftEndString =
    dateString + ' 10:00:00';

  var parts =
    shiftEndString.match(
      /(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/
    );

  var utcMillis =
    Date.UTC(
      Number(parts[1]),
      Number(parts[2]) - 1,
      Number(parts[3]),
      Number(parts[4]) - 8,
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

// Login
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

          team: row[5],

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
                'Your shift is currently closed. Accounts can only log in from 11:45 PM to 10:00 AM PST.'
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

          team: user.team,

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

            team: user.team,

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

// // ========== ENTRY ACCESS CONTROL ==========

// function canAccessEntry(user, entryAssignedAgentId) {
//   if (!user) return false;

//   var role = user.role ? user.role.toString().trim() : '';

//   // Super Admin and Admin can access all entries
//   if (role === 'Super Admin' || role === 'Admin') {
//     return true;
//   }

//   // Sales Partner can ONLY access their own assigned leads
//   if (role === 'Sales Partner') {
//     return String(entryAssignedAgentId) === String(user.id);
//   }

//   // Lead Gen Specialist will be handled later
//   if (role === 'Lead Gen Specialist') {
//     return false;
//   }

//   return false;
// }


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

// ========== DUPLICATE LEAD CHECK ==========

function normalizeLeadName(value) {
  if (!value) return '';

  return value
    .toString()
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function normalizeLeadPhone(value) {
  if (!value) return '';

  return value
    .toString()
    .replace(/\D/g, '');
}

function normalizeLeadEmail(value) {
  if (!value) return '';

  return value
    .toString()
    .trim()
    .toLowerCase();
}

function isDuplicateLead(existingData, entryData, excludeEntryId) {

  var newName =
    normalizeLeadName(entryData.authorName);

  var newPhones =
    entryData.phones
      ? entryData.phones
          .toString()
          .split(',')
          .map(function(phone) {
            return normalizeLeadPhone(phone);
          })
          .filter(function(phone) {
            return phone !== '';
          })
      : [];

  var newEmail =
    normalizeLeadEmail(entryData.email);

  // Start after header row
  for (var i = 1; i < existingData.length; i++) {

    var row = existingData[i];

    // ==========================================
    // IGNORE THE ENTRY CURRENTLY BEING EDITED
    // ==========================================

    if (
      excludeEntryId !== undefined &&
      excludeEntryId !== null &&
      String(row[0]).trim() === String(excludeEntryId).trim()
    ) {
      continue;
    }

    var existingName =
      normalizeLeadName(row[1]);

    var existingPhones =
      row[2]
        ? row[2]
            .toString()
            .split(',')
            .map(function(phone) {
              return normalizeLeadPhone(phone);
            })
            .filter(function(phone) {
              return phone !== '';
            })
        : [];

    var existingEmail =
      normalizeLeadEmail(row[3]);

    // ==========================================
    // NAME MATCH
    // ==========================================

    if (
      newName &&
      existingName &&
      newName === existingName
    ) {
      return true;
    }

    // ==========================================
    // PHONE MATCH
    // ==========================================

    for (var p = 0; p < newPhones.length; p++) {

      for (var ep = 0; ep < existingPhones.length; ep++) {

        if (
          newPhones[p] &&
          existingPhones[ep] &&
          newPhones[p] === existingPhones[ep]
        ) {
          return true;
        }

      }

    }

    // ==========================================
    // EMAIL MATCH
    // Only check email when the new entry
    // actually contains an email.
    // ==========================================

    if (
      newEmail &&
      existingEmail &&
      newEmail === existingEmail
    ) {
      return true;
    }
  }

  return false;
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


if (isDuplicateLead(data, entryData)) {
  return {
    success: false,
    code: 'DUPLICATE_LEAD',
    message: 'The Leads You Entered Already Exist in the Database'
  };
}

var newEntryId = generateEntryId(data);

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

    var formattedPhones =
      formatUSPhoneNumbers(entryData.phones);

    var createdAt = new Date().toISOString();

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
      user.id,                      // M - Mined By
      ''                           // N - StatusStartedAt
    ]);

    SpreadsheetApp.flush();

    logActivity(
      user.id,
      user.name,
      'Created entry #' + newEntryId + ': ' + entryData.authorName
    );

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

// ========== TRANSFER TARGETS ==========

function getTransferTargets(sessionId) {

  try {

    var user = getUserFromCache(sessionId);

    if (!user) {
      return {
        success: false,
        message: 'Session expired.',
        targets: []
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Agents');

    if (!sheet) {
      return {
        success: false,
        message: 'Agents sheet not found.',
        targets: []
      };
    }

    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: true,
        targets: []
      };
    }

    data.shift();

    var targets = [];

    for (var i = 0; i < data.length; i++) {

      var row = data[i];

      /*
       * Agents columns:
       *
       * [0] AgentID
       * [1] Username
       * [2] Password
       * [3] Name
       * [4] Email
       * [5] Team
       * [6] Phone
       * [7] Role
       * [8] Status
       * [9] VerificationToken
       */

      var agentId = row[0];
      var name = row[3];
      var role = row[7]
        ? row[7].toString().trim()
        : '';

      var status = row[8]
        ? row[8].toString().trim()
        : '';

      /*
       * Only Active users can receive transfers.
       */
      if (status !== 'Active') {
        continue;
      }

      /*
       * Sales Partners can transfer to:
       * - Admin
       * - Other Sales Partners
       *
       * Super Admin is intentionally not included
       * because Super Admin can manage transfers directly.
       */
      if (
        role !== 'Sales Partner' &&
        role !== 'Admin'
      ) {
        continue;
      }

      /*
       * Do not show the currently logged-in Sales Partner
       * as a transfer target.
       */
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
      e.message +
      ' | Stack: ' +
      e.stack
    );

    return {
      success: false,
      message: 'Unable to load transfer targets.',
      targets: []
    };
  }
}

// ========== CREATE TRANSFER REQUEST ==========

function createTransferRequest(sessionId, entryId, targetAgentId, reason) {

  try {

    var user = getUserFromCache(sessionId);

    if (!user) {
      return {
        success: false,
        message: 'Session expired.'
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);

    var entriesSheet = ss.getSheetByName('Entries');
    var agentsSheet = ss.getSheetByName('Agents');
    var transferSheet = ss.getSheetByName('TransferRequests');

    if (!entriesSheet || !agentsSheet || !transferSheet) {
      return {
        success: false,
        message: 'Required sheet not found.'
      };
    }

    /*
     * =====================================================
     * FIND ENTRY
     * =====================================================
     */

    var entryData = entriesSheet.getDataRange().getValues();

    var entryRow = null;

    for (var i = 1; i < entryData.length; i++) {

      if (
        String(entryData[i][0]) ===
        String(entryId)
      ) {

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
     * =====================================================
     * CHECK CURRENT OWNER
     * =====================================================
     *
     * Entries:
     *
     * [7] Assigned Agent ID
     * [8] Status
     * [12] Mined By
     */

    var assignedAgentId = entryRow[7];
    
    var minedById = entryRow[12]
      ? entryRow[12].toString().trim()
      : '';

    /*
     * Only the current assigned Sales Partner
     * can request the transfer.
     */

    if (user.role === 'Sales Partner') {

      var isAssignedToUser =
        String(assignedAgentId) === String(user.id);

      var isOwnLead =
        !assignedAgentId &&
        String(minedById) === String(user.id);

      if (!isAssignedToUser && !isOwnLead) {
        return {
          success: false,
          message: 'You can only transfer leads assigned to you.'
        };
      }
    }

    /*
     * =====================================================
     * SOLD LEADS CANNOT BE TRANSFERRED
     * =====================================================
     */

    var currentStatus = entryRow[8]
      ? entryRow[8].toString().trim()
      : '';

    if (currentStatus === 'Sold') {

      return {
        success: false,
        message: 'Sold leads cannot be transferred.'
      };
    }

    /*
     * =====================================================
     * FIND TARGET AGENT
     * =====================================================
     */

    var agentData = agentsSheet.getDataRange().getValues();

    var targetAgent = null;

    for (var j = 1; j < agentData.length; j++) {

      if (
        String(agentData[j][0]) ===
        String(targetAgentId)
      ) {

        targetAgent = agentData[j];
        break;
      }
    }

    if (!targetAgent) {

      return {
        success: false,
        message: 'Transfer target not found.'
      };
    }

    var targetId = targetAgent[0];

    var targetName = targetAgent[3];

    var targetRole = targetAgent[7]
      ? targetAgent[7].toString().trim()
      : '';

    var targetStatus = targetAgent[8]
      ? targetAgent[8].toString().trim()
      : '';

    /*
     * =====================================================
     * TARGET VALIDATION
     * =====================================================
     */

    if (targetStatus !== 'Active') {

      return {
        success: false,
        message: 'The selected agent is not active.'
      };
    }

    if (
      targetRole !== 'Sales Partner' &&
      targetRole !== 'Admin'
    ) {

      return {
        success: false,
        message: 'Invalid transfer target.'
      };
    }

    /*
     * Cannot transfer to yourself.
     */

    if (
      String(targetId) ===
      String(user.id)
    ) {

      return {
        success: false,
        message: 'You cannot transfer a lead to yourself.'
      };
    }

    /*
     * =====================================================
     * CHECK FOR EXISTING PENDING REQUEST
     * =====================================================
     */

    var transferData =
      transferSheet.getDataRange().getValues();

    for (var k = 1; k < transferData.length; k++) {

      var existingEntryId = transferData[k][1];

      var existingStatus = transferData[k][9]
        ? transferData[k][9].toString().trim()
        : '';

      if (
        String(existingEntryId) ===
        String(entryId) &&
        existingStatus === 'Pending'
      ) {

        return {
          success: false,
          message: 'This lead already has a pending transfer request.'
        };
      }
    }

    /*
     * =====================================================
     * GENERATE REQUEST ID
     * =====================================================
     */

    var requestId =
      'TR-' +
      new Date().getTime();

    /*
     * =====================================================
     * SAVE REQUEST
     * =====================================================
     */

    transferSheet.appendRow([

      requestId,             // A RequestID
      entryId,               // B EntryID

      user.id,               // C FromAgentID
      user.name,             // D FromAgentName

      targetId,              // E ToAgentID
      targetName,            // F ToAgentName

      user.id,               // G RequestedByID
      user.name,             // H RequestedByName

      new Date().toISOString(), // I RequestedAt

      'Pending',             // J Status

      '',                    // K ReviewedByID
      '',                    // L ReviewedByName
      '',                    // M ReviewedAt

      reason
        ? reason.toString().trim()
        : '',                // N Reason

      ''                     // O ReviewNote
    ]);

    /*
     * =====================================================
     * ACTIVITY LOG
     * =====================================================
     */

    logActivity(
      user.id,
      user.name,
      'Requested transfer of #' +
      entryId +
      ' to ' +
      targetName
    );

    /*
     * SYSTEM REMARK
     */

    addSystemRemark(
      entryId,
      user.name,
      'Transfer requested to ' +
      targetName +
      '. Awaiting Admin approval.'
    );

    return {
      success: true,
      message: 'Transfer request submitted for approval.',
      requestId: requestId
    };

  } catch (e) {

    console.error(
      'createTransferRequest error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );

    return {
      success: false,
      message: 'Unable to create transfer request.'
    };
  }
}

// ========== CHECK PENDING TRANSFER ==========

function getPendingTransferForEntry(sessionId, entryId) {

  try {

    var user = getUserFromCache(sessionId);

    if (!user) {
      return {
        success: false,
        pending: false,
        message: 'Session expired.'
      };
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var transferSheet = ss.getSheetByName('TransferRequests');

    if (!transferSheet) {
      return {
        success: false,
        pending: false,
        message: 'TransferRequests sheet not found.'
      };
    }

    var data = transferSheet.getDataRange().getValues();

    /*
     * No transfer requests yet.
     */
    if (data.length <= 1) {

      return {
        success: true,
        pending: false
      };

    }

    /*
     * Remove header.
     */
    data.shift();

    for (var i = 0; i < data.length; i++) {

      var row = data[i];

      var requestEntryId =
        row[1] == null
          ? ''
          : String(row[1]).trim();

      var requestStatus =
        row[9] == null
          ? ''
          : String(row[9]).trim();

      if (
        requestEntryId === String(entryId).trim() &&
        requestStatus === 'Pending'
      ) {

        return {
          success: true,
          pending: true,

          request: {
            requestId: row[0],
            entryId: row[1],

            fromAgentId: row[2],
            fromAgentName: row[3],

            toAgentId: row[4],
            toAgentName: row[5],

            requestedById: row[6],
            requestedByName: row[7],

            requestedAt: row[8],

            status: row[9],

            reason: row[13] || '',
            reviewNote: row[14] || ''
          }
        };
      }
    }

    /*
     * No pending request found.
     */

    return {
      success: true,
      pending: false
    };

  } catch (e) {

    console.error(
      'getPendingTransferForEntry error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );

    return {
      success: false,
      pending: false,
      message: 'Unable to check transfer status.'
    };
  }
}

// ========== PHONE NUMBER FORMATTING ==========
function formatUSPhoneNumber(phone) {

  if (!phone) return '';

  var original = phone.toString().trim();

  var digits = original.replace(/\D/g, '');

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

// ==========================================
// DUPLICATE LEAD CHECK
// Ignore the entry currently being edited.
// ==========================================

if (isDuplicateLead(d, entryData, entryId)) {
  return {
    success: false,
    code: 'DUPLICATE_LEAD',
    message: 'The Leads You Entered Already Exist in the Database'
  };
}

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
    var transferSheet = ss.getSheetByName('TransferRequests');

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

    // ==========================================
    // BUILD AGENT ID -> NAME MAP
    // ==========================================

    var ad = agentSheet.getDataRange().getValues();
    ad.shift();

    var am = {};
    var arm = {};

    for (var k = 0; k < ad.length; k++) {

      am[ad[k][0]] = ad[k][3];

      arm[ad[k][0]] =
        ad[k][7]
          ? ad[k][7].toString().trim()
          : '';
    }

    for (var k = 0; k < ad.length; k++) {
      am[ad[k][0]] = ad[k][3];
    }

    // ==========================================
    // BUILD PENDING TRANSFER MAP
    // ==========================================

    var pendingTransfers = {};

    if (transferSheet) {

      var td = transferSheet.getDataRange().getValues();

      if (td.length > 1) {

        td.shift();

        for (var t = 0; t < td.length; t++) {

          var transferEntryId = td[t][1]; // Column B = EntryID
          var transferStatus = td[t][9];  // Column J = Status

          if (
            transferEntryId !== '' &&
            transferStatus &&
            transferStatus.toString().trim() === 'Pending'
          ) {

            pendingTransfers[String(transferEntryId)] = true;
          }
        }
      }
    }

    // ==========================================
    // BUILD ENTRIES
    // ==========================================

    var entries = [];

    for (var i = 0; i < ed.length; i++) {

      var assignedAgentId = ed[i][7];
      var minedById = ed[i][12];

      if (!canAccessEntry(
        u,
        assignedAgentId,
        minedById
      )) {
        continue;
      }

      // ==========================================
      // WRONG NUMBER LEADS
      // ==========================================

      var entryStatus =
        ed[i][8]
          ? ed[i][8].toString().trim()
          : '';

      if (entryStatus === 'Wrong Number') {
        continue;
      }

      // ========================================
      // ORIGINAL STATUS
      // ========================================

      var originalStatus =
        ed[i][8]
          ? ed[i][8].toString().trim()
          : '';

      // ========================================
      // CHECK PENDING TRANSFER
      // ========================================

      var isTransferPending =
        pendingTransfers[String(ed[i][0])] === true;

      /*
       * IMPORTANT:
       *
       * We do NOT change the Entries sheet.
       *
       * We only return "Pending" to the frontend
       * while a transfer request is awaiting approval.
       */

      var displayStatus =
        isTransferPending
          ? 'Pending'
          : originalStatus;

      entries.push({

        id: ed[i][0],

        authorName: ed[i][1] || '',

        phones: ed[i][2] || '',

        email: ed[i][3] || '',

        book: ed[i][4] || '',

        isbn: ed[i][5] || '',

        address: ed[i][6] || '',

        assignedAgentId: ed[i][7],

        assignedAgentName:
        ed[i][7]
          ? (am[ed[i][7]] || 'Own Lead')
          : 'Own Lead',

        /*
         * Status shown to frontend.
         *
         * Pending temporarily overrides the
         * original status while transfer is pending.
         */

        status: displayStatus,

        /*
         * Keep the real status available.
         *
         * This is useful later when Admin cancels
         * or approves the transfer.
         */

        originalStatus: originalStatus,

        /*
         * Frontend can use this to lock the entry.
         */

        transferPending: isTransferPending,

        createdAt: ed[i][9] || '',

        minedById: ed[i][12] || '',

        minedByRole:
        minedById
          ? (arm[minedById] || '')
          : ''

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

    console.error(
      'getEntries error: ' +
      e.message +
      ' | Stack: ' +
      e.stack
    );

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

    var entrySheetRow =
      entryData.findIndex(function(row) {
        return String(row[0]) === String(entryId);
      }) + 1;

    var statusStartedAt = '';

    if (
      newStatus === '' ||
      newStatus === 'Pipe' ||
      newStatus === 'VM'
    ) {
      statusStartedAt = new Date();
    }

    // Update Status
    entrySheet.getRange(entrySheetRow, 9).setValue(newStatus);

    // Update StatusStartedAt
    entrySheet.getRange(entrySheetRow, 14).setValue(statusStartedAt);

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

// ========== EXTEND LEAD GRACE PERIOD ==========
function extendLeadGracePeriod(sessionId, entryId) {
  try {

    var u = getUserFromCache(sessionId);

    if (!u) {
      return {
        success: false,
        message: 'Session expired.'
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

    var entryRowIndex = -1;
    var entryRow = null;

    // ==========================================
    // FIND ENTRY
    // ==========================================

    for (var i = 1; i < entryData.length; i++) {

      if (
        String(entryData[i][0]) ===
        String(entryId)
      ) {

        entryRowIndex = i;
        entryRow = entryData[i];

        break;
      }
    }

    if (entryRowIndex === -1 || !entryRow) {
      return {
        success: false,
        message: 'Entry not found.'
      };
    }

    // ==========================================
    // CURRENT STATUS
    // ==========================================

    var currentStatus =
      entryRow[8]
        ? entryRow[8].toString().trim()
        : '';

    // ==========================================
    // ONLY THESE STATUSES CAN BE EXTENDED
    // ==========================================

    var allowedStatuses = [
      '',
      'Pipe',
      'VM'
    ];

    if (
      allowedStatuses.indexOf(currentStatus) === -1
    ) {

      return {
        success: false,
        message:
          'This lead does not have an extendable grace period.'
      };
    }

    // ==========================================
    // FIND USER ROLE
    // ==========================================

    var userRole = '';

    for (var j = 1; j < agentData.length; j++) {

      if (
        String(agentData[j][0]) ===
        String(u.id)
      ) {

        userRole =
          agentData[j][7]
            ? agentData[j][7].toString().trim()
            : '';

        break;
      }
    }

    // ==========================================
    // ONLY ADMIN AND SALES PARTNER CAN EXTEND
    // ==========================================

    if (
      userRole !== 'Admin' &&
      userRole !== 'Sales Partner'
    ) {

      return {
        success: false,
        message:
          'You do not have permission to extend leads.'
      };
    }

    // ==========================================
    // SALES PARTNER ACCESS CHECK
    // ==========================================

    if (userRole === 'Sales Partner') {

      var assignedAgentId =
        entryRow[7]
          ? entryRow[7].toString().trim()
          : '';

      var minedById =
        entryRow[12]
          ? entryRow[12].toString().trim()
          : '';

      var hasAccess =
        String(assignedAgentId) === String(u.id) ||
        String(minedById) === String(u.id);

      if (!hasAccess) {

        return {
          success: false,
          message:
            'You do not have permission to extend this lead.'
        };
      }
    }

    // ==========================================
    // RESET STATUS START TIME
    // ==========================================

    var newStatusStartedAt = new Date();

    /*
     * Column N = StatusStartedAt
     *
     * Status remains unchanged.
     */

    entrySheet
      .getRange(entryRowIndex + 1, 14)
      .setValue(newStatusStartedAt);

    SpreadsheetApp.flush();

    // ==========================================
    // ACTIVITY LOG
    // ==========================================

    var statusLabel =
      currentStatus === ''
        ? 'No Status'
        : currentStatus;

    logActivity(
      u.id,
      u.name,
      'Extended grace period #' +
        entryId +
        ' (' +
        statusLabel +
        ')'
    );

    addSystemRemark(
      entryId,
      u.name,
      'Grace period extended for ' +
        statusLabel
    );

    return {
      success: true,
      message: 'Grace period extended successfully.',
      status: currentStatus,
      statusStartedAt:
        newStatusStartedAt.toISOString()
    };

  } catch (e) {

    console.error(
      'extendLeadGracePeriod error:',
      e
    );

    return {
      success: false,
      message:
        'Unable to extend the grace period.'
    };
  }
}

// ========== MARK LEAD AS WRONG NUMBER ==========
function markLeadWrongNumber(sessionId, entryId) {

  try {

    var u = getUserFromCache(sessionId);

    if (!u) {
      return {
        success: false,
        message: 'Session expired.'
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

    var entryData =
      entrySheet.getDataRange().getValues();

    var agentData =
      agentsSheet.getDataRange().getValues();

    var entryRow = null;
    var entrySheetRow = -1;

    // ==========================================
    // FIND ENTRY
    // ==========================================

    for (var i = 1; i < entryData.length; i++) {

      if (
        String(entryData[i][0]) ===
        String(entryId)
      ) {

        entryRow = entryData[i];
        entrySheetRow = i + 1;

        break;
      }
    }

    if (!entryRow) {

      return {
        success: false,
        message: 'Entry not found.'
      };
    }

    // ==========================================
    // CURRENT STATUS
    // ==========================================

    var currentStatus =
      entryRow[8]
        ? entryRow[8].toString().trim()
        : '';

    // Already Wrong Number

    if (currentStatus === 'Wrong Number') {

      return {
        success: false,
        message: 'This lead is already marked as Wrong Number.'
      };
    }

    // ==========================================
    // FIND LOGGED-IN USER ROLE
    // ==========================================

    var userRole = '';

    for (var j = 1; j < agentData.length; j++) {

      if (
        String(agentData[j][0]) ===
        String(u.id)
      ) {

        userRole =
          agentData[j][7]
            ? agentData[j][7].toString().trim()
            : '';

        break;
      }
    }

    // ==========================================
    // ONLY ADMIN / SALES PARTNER CAN USE MWN
    // ==========================================

    if (
      userRole !== 'Admin' &&
      userRole !== 'Sales Partner'
    ) {

      return {
        success: false,
        message: 'You are not authorized to mark this lead as Wrong Number.'
      };
    }

    // ==========================================
    // FIND WHO MINED THE LEAD
    // ==========================================

    var minedById =
      entryRow[12]
        ? entryRow[12].toString().trim()
        : '';

    if (!minedById) {

      return {
        success: false,
        message: 'This lead does not have a recorded Lead Gen Specialist.'
      };
    }

    var minedByRole = '';
    var minedByName = 'Unknown';

    for (var k = 1; k < agentData.length; k++) {

      if (
        String(agentData[k][0]) ===
        String(minedById)
      ) {

        minedByRole =
          agentData[k][7]
            ? agentData[k][7].toString().trim()
            : '';

        minedByName =
          agentData[k][3]
            ? agentData[k][3].toString().trim()
            : 'Unknown';

        break;
      }
    }

    // ==========================================
    // MUST HAVE BEEN MINED BY LGS
    // ==========================================

    if (minedByRole !== 'Lead Gen Specialist') {

      return {
        success: false,
        message: 'Only leads mined by a Lead Gen Specialist can be marked as Wrong Number.'
      };
    }

    // ==========================================
    // CURRENT OWNER
    // ==========================================

    var previousOwnerId =
      entryRow[7]
        ? entryRow[7].toString().trim()
        : '';

    var previousOwnerName =
      previousOwnerId
        ? (getAgentName(previousOwnerId) || 'Unknown')
        : 'Unassigned';

    // ==========================================
    // SAVE MWN STATE
    // ==========================================

    // Status
    entrySheet
      .getRange(entrySheetRow, 9)
      .setValue('Wrong Number');

    // Pause grace period
    entrySheet
      .getRange(entrySheetRow, 14)
      .clearContent();

    SpreadsheetApp.flush();

    // ==========================================
    // ACTIVITY LOG
    // ==========================================

    logActivity(
      u.id,
      u.name,
      'Marked #' +
      entryId +
      ' as Wrong Number. Previous status: ' +
      (currentStatus || 'No Status') +
      '. Previous owner: ' +
      previousOwnerName +
      '. Returned to Lead Gen Specialist: ' +
      minedByName
    );

    // ==========================================
    // SYSTEM REMARK
    // ==========================================

    addSystemRemark(
      entryId,
      u.name,
      'Wrong Number: Lead returned to Lead Gen Specialist ' +
      minedByName +
      ' for re-mining. Previous owner: ' +
      previousOwnerName +
      '. Previous status: ' +
      (currentStatus || 'No Status') +
      '. Grace period paused.'
    );

    return {
      success: true,
      message: 'Lead marked as Wrong Number and returned to Lead Gen Specialist.',
      entryId: entryId
    };

  } catch (e) {

    console.error(
      'markLeadWrongNumber error:',
      e
    );

    return {
      success: false,
      message: 'Unable to mark lead as Wrong Number.'
    };
  }
}

// ========== PROCESS LEAD GRACE PERIODS ==========
function processLeadGracePeriods() {

  var lock = LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    var ss =
      SpreadsheetApp.openById(SHEET_ID);

    var entrySheet =
      ss.getSheetByName('Entries');

    if (!entrySheet) {
      return;
    }

    var data =
      entrySheet.getDataRange().getValues();

    if (data.length <= 1) {
      return;
    }

    var now = new Date();

    var DAY_MS =
      24 * 60 * 60 * 1000;

    for (var i = 1; i < data.length; i++) {

      var row = data[i];

      var entryId = row[0];

      var status =
        row[8]
          ? row[8].toString().trim()
          : '';

      var statusStartedAt =
        row[13]
          ? row[13].toString().trim()
          : '';

      // ==========================================
      // ONLY PROCESS GRACE-PERIOD STATUSES
      // ==========================================

      if (
        status !== '' &&
        status !== 'Pipe' &&
        status !== 'VM'
      ) {
        continue;
      }

      if (!statusStartedAt) {
        continue;
      }

      var startedAt =
        new Date(statusStartedAt);

      if (isNaN(startedAt.getTime())) {
        continue;
      }

      // ==========================================
      // DETERMINE GRACE PERIOD
      // ==========================================

      var graceDays = 30;

      if (status === 'VM') {
        graceDays = 10;
      }

      var expirationTime =
        startedAt.getTime() +
        (graceDays * DAY_MS);

      // ==========================================
      // CHECK EXPIRATION
      // ==========================================

      if (
        now.getTime() >=
        expirationTime
      ) {

        expireAndRedistributeLead(
          entryId
        );
      }
    }

  } catch (e) {

    console.error(
      'processLeadGracePeriods error:',
      e
    );

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

// ========== EXPIRE AND REDISTRIBUTE LEAD ==========

function expireAndRedistributeLead(entryId) {

  try {

    var ss =
      SpreadsheetApp.openById(SHEET_ID);

    var entrySheet =
      ss.getSheetByName('Entries');

    if (!entrySheet) {
      return;
    }

    var data =
      entrySheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) !==
        String(entryId)
      ) {
        continue;
      }

      var rowNumber = i + 1;

      var currentStatus =
        data[i][8]
          ? data[i][8].toString().trim()
          : '';

      var previousAssignedAgentId =
        data[i][7]
          ? data[i][7].toString().trim()
          : '';

      // ==========================================
      // SOLD IS ALWAYS FINAL
      // ==========================================

      if (currentStatus === 'Sold') {
        return;
      }

      // ==========================================
      // ONLY EXPIRE GRACE STATUSES
      // ==========================================

      if (
        currentStatus !== '' &&
        currentStatus !== 'Pipe' &&
        currentStatus !== 'VM'
      ) {
        return;
      }

      var oldStatus =
        currentStatus || 'No Status';

      // ==========================================
      // RESET TO NO STATUS
      // ==========================================

      entrySheet
        .getRange(rowNumber, 9)
        .setValue('');

      // Clear timer temporarily.
      // distributeNewEntry() will start
      // a fresh timer after assigning it.
      entrySheet
        .getRange(rowNumber, 14)
        .clearContent();

      SpreadsheetApp.flush();

      // ==========================================
      // LOG EXPIRATION
      // ==========================================

      logActivity(
        0,
        'System',
        '#' +
        entryId +
        ' expired from ' +
        oldStatus +
        ' and is being redistributed.'
      );

      addSystemRemark(
        entryId,
        'System',
        oldStatus +
        ' grace period expired. Lead returned to No Status and will be redistributed.'
      );

      // ==========================================
      // REDISTRIBUTE
      // ==========================================

      distributeNewEntry(
        entryId,
        previousAssignedAgentId
      );

      return;
    }

  } catch (e) {

    console.error(
      'expireAndRedistributeLead error:',
      e
    );
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
    return items.slice(0, 50);
  } catch (e) {
    console.error( 'getEntryActivity failed. Entry ID: ' + entryId + ' | Error: ' + e.message + ' | Stack: ' + e.stack ); return []; }
}

function logActivity(uid, un, action) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('ActivityLog');
    if (!sheet) { console.error('logActivity: ActivityLog sheet not found.'); return; }
    var nextId = sheet.getLastRow();
    sheet.appendRow([ nextId, uid, un, action, new Date().toISOString() ]);
  } catch (e) {
    console.error( 'logActivity failed. Action: ' + action + ' | Error: ' + e.message + ' | Stack: ' + e.stack ); }
}

function addSystemRemark(eid, un, msg) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Remarks');
    if (!sheet) { console.error('addSystemRemark: Remarks sheet not found.'); return; }
    var nextId = sheet.getLastRow();
    sheet.appendRow([ nextId, eid, new Date().toISOString(), 0, un, msg ]);
  } catch (e) {
    console.error( 'addSystemRemark failed. Entry ID: ' + eid + ' | Error: ' + e.message + ' | Stack: ' + e.stack ); }
}

// // ========== Round-Robin DISTRIBUTION ==========
// function distributeNewEntry(entryId) {
//   var lock = LockService.getScriptLock();
//   try {
//     lock.waitLock(10000);
//     var ss = SpreadsheetApp.openById(SHEET_ID);
//     var agentSheet = ss.getSheetByName('Agents');
//     var entrySheet = ss.getSheetByName('Entries');
//     if (!agentSheet || !entrySheet) { return; }
//     var agentData = agentSheet.getDataRange().getValues();
//     if (agentData.length <= 1) { return; }
//     agentData.shift();
//     var eligibleAgents = [];
//     for (var i = 0; i < agentData.length; i++) {
//       var agent = agentData[i];
//       var agentId = agent[0];
//       var agentName = agent[3];
//       var role = agent[7] ? agent[7].toString().trim() : '';
//       var status = agent[8] ? agent[8].toString().trim() : '';
//       if ( status === 'Active' && (role === 'Admin' || role === 'Sales Partner') ) {
//         eligibleAgents.push({ id: agentId, name: agentName, role: role, rowIndex: i }); } }
//     if (eligibleAgents.length === 0) { return; }
//     var properties = PropertiesService.getScriptProperties();
//     var lastAgentId = properties.getProperty('LAST_DISTRIBUTED_AGENT_ID');
//     var nextAgentIndex = 0;
//     if (lastAgentId !== null) {
//       var lastIndex = -1;
//       for (var j = 0; j < eligibleAgents.length; j++) {
//         if (String(eligibleAgents[j].id) === String(lastAgentId)) { lastIndex = j; break; } }
//       if (lastIndex !== -1) { nextAgentIndex = (lastIndex + 1) % eligibleAgents.length;
//       } else {
//         nextAgentIndex = 0; } }
//     var selectedAgent = eligibleAgents[nextAgentIndex];
//     var entryData = entrySheet.getDataRange().getValues();
//     var entryFound = false;
//     for (var k = 1; k < entryData.length; k++) {
//       if (String(entryData[k][0]) === String(entryId)) {
//         entrySheet.getRange(k + 1, 8).setValue(selectedAgent.id);
//         entrySheet .getRange(k + 1, 10) .setValue(new Date().toISOString());
//         entryFound = true; break; } }
//     if (!entryFound) { return; }
//     properties.setProperty( 'LAST_DISTRIBUTED_AGENT_ID', String(selectedAgent.id) );
//     SpreadsheetApp.flush();
//     logActivity( 0, 'System', 'Auto-assigned #' + entryId + ' to ' + selectedAgent.name );
//     addSystemRemark( entryId, 'System', 'Lead auto-assigned to ' + selectedAgent.name );
//   } catch (e) {
//     console.error('Distribution error:', e);
//   } finally {
//     try { lock.releaseLock(); } catch (lockError) {} }
// }

// ========== ROUND-ROBIN DISTRIBUTION ==========

function distributeNewEntry(entryId, excludedAgentId) {

  var lock = LockService.getScriptLock();

  try {

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

    agentData.shift();

    var eligibleAgents = [];

    for (var i = 0; i < agentData.length; i++) {

      var agent = agentData[i];

      var agentId = agent[0];
      var agentName = agent[3];

      var role = agent[7]
        ? agent[7].toString().trim()
        : '';

      var status = agent[8]
        ? agent[8].toString().trim()
        : '';

      if (
        status === 'Active' &&
        (
          role === 'Admin' ||
          role === 'Sales Partner'
        )
      ) {

        eligibleAgents.push({
          id: agentId,
          name: agentName,
          role: role,
          rowIndex: i
        });

      }
    }

    if (eligibleAgents.length === 0) {
      return;
    }

    // ROUND-ROBIN SELECTION

    var properties =
      PropertiesService.getScriptProperties();

    var lastAgentId =
      properties.getProperty(
        'LAST_DISTRIBUTED_AGENT_ID'
      );

    var startIndex = 0;

    if (lastAgentId !== null) {

      var lastIndex = -1;

      for (var j = 0; j < eligibleAgents.length; j++) {

        if (
          String(eligibleAgents[j].id) ===
          String(lastAgentId)
        ) {

          lastIndex = j;
          break;
        }
      }

      if (lastIndex !== -1) {

        startIndex =
          (lastIndex + 1) %
          eligibleAgents.length;

      }
    }

    var selectedAgent = null;

    for (
      var offset = 0;
      offset < eligibleAgents.length;
      offset++
    ) {

      var candidateIndex =
        (startIndex + offset) %
        eligibleAgents.length;

      var candidate =
        eligibleAgents[candidateIndex];

      // Skip only the previous owner when
      // this function was called with an
      // excludedAgentId.
      if (
        excludedAgentId &&
        String(candidate.id) ===
        String(excludedAgentId)
      ) {
        continue;
      }

      selectedAgent = candidate;
      break;
    }

    // ==========================================
    // SAFETY CHECK
    // ==========================================

    if (!selectedAgent) {

      console.error(
        'No eligible agent available for entry #' +
        entryId +
        ' after applying exclusion.'
      );

      return;
    }

    // ==========================================
    // FIND ENTRY
    // ==========================================

    var entryData =
      entrySheet.getDataRange().getValues();

    var entryFound = false;

    for (var k = 1; k < entryData.length; k++) {

      if (
        String(entryData[k][0]) ===
        String(entryId)
      ) {

        var now =
          new Date().toISOString();

        // H = Assigned Agent
        entrySheet
          .getRange(k + 1, 8)
          .setValue(selectedAgent.id);

        // I = Status
        // System-distributed leads start
        // with No Status.
        entrySheet
          .getRange(k + 1, 9)
          .setValue('');

        // J = Created/Assigned At
        entrySheet
          .getRange(k + 1, 10)
          .setValue(now);

        // N = StatusStartedAt
        // Start a fresh 30-day No Status timer.
        entrySheet
          .getRange(k + 1, 14)
          .setValue(now);

        entryFound = true;

        break;
      }
    }

    if (!entryFound) {
      return;
    }

    // ==========================================
    // SAVE ROUND-ROBIN POSITION
    // ==========================================

    properties.setProperty(
      'LAST_DISTRIBUTED_AGENT_ID',
      String(selectedAgent.id)
    );

    SpreadsheetApp.flush();

    // ==========================================
    // LOG
    // ==========================================

    logActivity(
      0,
      'System',
      'Auto-assigned #' +
      entryId +
      ' to ' +
      selectedAgent.name
    );

    addSystemRemark(
      entryId,
      'System',
      'Lead auto-assigned to ' +
      selectedAgent.name
    );

  } catch (e) {

    console.error(
      'Distribution error:',
      e
    );

  } finally {

    try {
      lock.releaseLock();
    } catch (lockError) {}

  }
}

function getAgentName(id) { 
  try { var d=SpreadsheetApp.openById(SHEET_ID).getSheetByName('Agents').getDataRange().getValues(); 
    for(var i=1;i<d.length;i++){if(d[i][0]==id)return d[i][3];} return 'Unknown'; 
  } catch(e){return'Unknown';} 
}