/**
 * Training Ride Creator
 * Creates a new training ride form, links it to the master sheet,
 * previews a draft, and sends individually to all registrants.
 */

// ─── CONFIG ───────────────────────────────────────────────────────────────────
var MASTER_SHEET_ID      = '1LND_ZZJpalSgnyyVtOCEJgOy_N4kh-B3e9gfbguCpPE';
var TEMPLATE_FORM_ID     = '1vew47VRL4NK591rpe_w8NIEE1yjGLj_wOK_HtI-wFj4';
var FORM_FOLDER_ID       = '1fgBJisqSUYmBYIzxwR6jIRcgIHF2wtO1';
var REGISTRANT_FOLDER_ID = '11Jsj1pPf7NWYdVCzmuTrjTWj80JObSa4';
var PREVIEW_EMAIL        = 'shepherds@rfmc-mv.org';

var RIDE_LINE     = '315-ASK-RIDE (315-275-7433)';
var SCHEDULE_URL  = 'https://www.rfmc-mv.org/tr-schedule';
var CONTACT_EMAIL = 'shepherds@rfmc-mv.org';
var PO_BOX        = 'PO Box 6, Washington Mills, NY 13479';

var RIDER_SHEETS = [
  'MV Corporate Riders',
  'MV New Returning Riders',
  'MV New Returning High School St',
  'MV Reciprocal Riders',
  'Supplemental'  // skipped silently if not present
];

var EMAIL_COLUMN_NAME = 'Email';  // adjust if your header differs

// ── FB post (emailed to admin for manual posting — no FB API calls) ──────────
var FB_POST_RECIPIENT    = 'mike@pmichaeljordan.com';
var ROUTES_SHEET_NAME    = 'Routes';
var RIDE_PHOTO_FOLDER_ID = '1IdzXeaKQwlR-gN1_DhLEJQnYdSPsHx38';
var RIDE_PHOTO_USED_SUBFOLDER = 'used';

// ── Calendar invite (.ics attached to rider emails) ──────────────────────────
var RIDE_DURATION_HOURS = 2;
var EVENT_TIMEZONE      = 'America/New_York';

// ── Deliverability: receivers that frequently reject bulk mail ───────────────
// Recipients in these domains get a slower second-pass with longer delays.
var STRICT_DOMAINS = [
  'icloud.com', 'me.com', 'mac.com',
  'roadrunner.com', 'rr.com', 'twcny.rr.com', 'twc.com',
  'spectrum.net', 'charter.net'
];

var UNSUBSCRIBE_URL = 'https://www.rfmc-mv.org/unsubscribe';

var LOCATIONS = {
  'Westmoreland Upper Elementary School' : 'https://goo.gl/maps/5d8hHjof31LnTsrU6',
  'Whitesboro High School'               : 'https://www.google.com/maps/place/Whitesboro+High+School/@43.1514062,-75.2898054,2743m/data=!3m1!1e3!4m6!3m5!1s0x89d93f097d74a2dd:0x5a3d9188a3c9f65!8m2!3d43.1513431!4d-75.2892904!16s%2Fm%2F076qg69?entry=ttu',
  'Town Of Paris Park'                   : 'https://goo.gl/maps/Tc8s1yY22fdPZTjv7',
  'Clinton Arena'                        : 'https://www.google.com/maps/place/Clinton+Arena/@43.0537422,-75.3777248,466m/data=!3m1!1e3!4m6!3m5!1s0x89d969a494a3f621:0x2b17bdefc7364628!8m2!3d43.0536111!4d-75.3769444!16s%2Fm%2F026r83r?entry=ttu&g_ep=EgoyMDI1MDQyMS4wIKXMDSoASAFQAw%3D%3D',
  'Clinton Hannaford'                    : 'https://www.google.com/maps/place/Hannaford/@43.0423703,-75.3868412,241m/data=!3m1!1e3!4m5!3m4!1s0x89d969000a56f001:0xba3084e46ff4d64a!8m2!3d43.0421227!4d-75.3872056',
  'Herkimer Hannaford'                   : 'https://www.google.com/maps/place/Hannaford/@43.0272495,-74.9817993,17.25z/data=!4m5!3m4!1s0x0:0xd0b658e986c1be0d!8m2!3d43.0296732!4d-74.9812633',
  'New Hartford Recreation Center'       : 'https://www.google.com/maps/place/New+Hartford+Recreation+Center+Ice+Skating+Rink/@43.0717038,-75.2856762,346m/data=!3m1!1e3!4m5!3m4!1s0x89d9417bb8ab4d57:0x39a39051c7aa452f!8m2!3d43.0716184!4d-75.28553',
  'MVCC P1/P2 Parking Lot'              : 'https://www.google.com/maps/@43.0997532,-75.1528652,83m/data=!3m1!1e3',
  'Sauquoit Valley High School'          : 'https://goo.gl/maps/HYRQNf7R5SGGAHEw6',
  'Marcy Town Hall'                      : 'https://goo.gl/maps/vRVTq1M2ZLWdeV1Z9',
  'Kuyahoora Valley Town Park'           : 'https://www.google.com/maps/place/43%C2%B013\'39.4%22N+75%C2%B003\'27.6%22W/@43.2275287,-75.0574343,1365m/data=!3m1!1e3!4m4!3m3!8m2!3d43.2276111!4d-75.0576667?entry=ttu',
  'Byrne Dairy (Clinton)'                : 'https://www.google.com/maps/place/Byrne+Dairy+and+Deli/@43.0792994,-75.3759878,1128m/data=!3m2!1e3!4b1!4m6!3m5!1s0x89d9698bc5e95267:0xc73d5696ca77a0dc!8m2!3d43.0792994!4d-75.3734129!16s%2Fg%2F1td2hx8t?entry=ttu&g_ep=EgoyMDI1MDQyMy4wIKXMDSoASAFQAw%3D%3D'
};
// ─────────────────────────────────────────────────────────────────────────────


// ═══════════════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚴 Training Rides')
    .addItem('Create New Ride', 'createTrainingRide')
    .addSeparator()
    .addItem('Send Training Ride Emails…', 'promptSendEmails')
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════════════════════
// STEP 1 — CREATE RIDE (form + sheet tab + preview draft to self)
// ═══════════════════════════════════════════════════════════════════════════════

function createTrainingRide() {
  var html = HtmlService.createHtmlOutput(buildDialog())
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '🚴 Training Ride Setup');
}

function buildDialog() {
  var locationOptions = Object.keys(LOCATIONS).map(function(name) {
    return '<option value="' + name + '">' + name + '</option>';
  }).join('\n');

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; font-size: 13px; padding: 16px; }
    label { display: block; margin-top: 12px; font-weight: bold; }
    input, select { width: 100%; padding: 6px; margin-top: 4px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
    .sag-row { display: flex; gap: 12px; margin-top: 4px; }
    .sag-row label { font-weight: normal; display: flex; align-items: center; gap: 6px; }
    button { margin-top: 18px; width: 100%; padding: 10px; background: #4285F4; color: white; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; }
    button:hover { background: #2a6dd9; }
    button:disabled { background: #999; cursor: not-allowed; }
    #customLocationRow { display: none; margin-top: 8px; }
    .hint { font-size: 11px; color: #888; margin-top: 2px; }
  </style>
</head>
<body>
  <label>Ride Date
    <input type="text" id="rideDate" placeholder="e.g. April 2nd, 2026" />
  </label>

  <label>Departure Time
    <input type="text" id="rideTime" placeholder="e.g. 5:30 PM" />
  </label>

  <label>Departure Location
    <select id="locationSelect" onchange="handleLocationChange()">
      ${locationOptions}
      <option value="__custom__">Other (type below)...</option>
    </select>
  </label>

  <div id="customLocationRow">
    <input type="text" id="customLocation" placeholder="Enter custom location" />
    <div class="hint">This will be used in the form and email.</div>
  </div>

  <label>SAG Available?</label>
  <div class="sag-row">
    <label><input type="radio" name="sag" value="yes" checked /> Yes</label>
    <label><input type="radio" name="sag" value="no" /> No</label>
  </div>

  <button id="submitBtn" onclick="submitForm()">Create Ride →</button>

  <script>
    function handleLocationChange() {
      var sel = document.getElementById('locationSelect').value;
      document.getElementById('customLocationRow').style.display =
        sel === '__custom__' ? 'block' : 'none';
    }

    function submitForm() {
      var rideDate = document.getElementById('rideDate').value.trim();
      var rideTime = document.getElementById('rideTime').value.trim();
      var locSel   = document.getElementById('locationSelect').value;
      var location = locSel === '__custom__'
        ? document.getElementById('customLocation').value.trim()
        : locSel;
      var sag = document.querySelector('input[name="sag"]:checked').value === 'yes';

      if (!rideDate || !rideTime || !location) {
        alert('Please fill in all fields.');
        return;
      }

      var btn = document.getElementById('submitBtn');
      btn.disabled = true;
      btn.innerText = '⏳ Creating ride, please wait...';

      google.script.run
        .withSuccessHandler(function(msg) {
          document.body.innerHTML = '<p style="padding:20px;font-size:14px;">✅ ' + msg + '</p>';
          setTimeout(function(){ google.script.host.close(); }, 3000);
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.innerText = 'Create Ride →';
          btn.style.background = '#4285F4';
          alert('Error: ' + err.message);
        })
        .processRide(rideDate, rideTime, location, sag);
    }
  </script>
</body>
</html>`;
}

function processRide(rideDate, rideTime, rideLocation, sagAvailable) {
  var formTitle = rideDate + ' Training Ride';

  var sagLine = sagAvailable
    ? 'There will be limited SAG (support and gear) space available.'
    : 'There will not be SAG (support and gear) space available.';

  var description =
    'Please fill out this form if you plan on attending the ' + rideDate +
    ' training ride leaving ' + rideLocation + '. Please be prepared' +
    ' to arrive a few minutes early and be ready to roll at ' + rideTime +
    '. ' + sagLine +
    ' Please come prepared with plenty of water in a cycling-specific water bottle' +
    ' (no disposable or steel bottles), a spare tube, and the tools needed to fix a flat.';

  // ── Copy template form ────────────────────────────────────────────────────
  var templateFile = DriveApp.getFileById(TEMPLATE_FORM_ID);
  var newFormFile  = templateFile.makeCopy(formTitle);
  var newFormId    = newFormFile.getId();

  var targetFolder = DriveApp.getFolderById(FORM_FOLDER_ID);
  targetFolder.addFile(newFormFile);
  DriveApp.getRootFolder().removeFile(newFormFile);

  var newForm = FormApp.openById(newFormId);
  newForm.setPublished(true);
  newForm.setTitle(formTitle);
  newForm.setDescription(description);
  newForm.setAcceptingResponses(true);

  // ── Link to master sheet ──────────────────────────────────────────────────
  var masterSS    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var placeholder = masterSS.insertSheet(rideDate + '_tmp');

  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, MASTER_SHEET_ID);
  SpreadsheetApp.flush();

  var sheets = masterSS.getSheets();
  var renamedSheet;
  for (var i = sheets.length - 1; i >= 0; i--) {
    if (sheets[i].getName().startsWith('Form Responses')) {
      sheets[i].setName(rideDate);
      renamedSheet = sheets[i];
      break;
    }
  }
  masterSS.deleteSheet(placeholder);

  if (renamedSheet) {
    masterSS.setActiveSheet(renamedSheet);
    masterSS.moveActiveSheet(masterSS.getNumSheets());
  }

  // ── Store ride details for the Send step ──────────────────────────────────
  var subject = 'Training Ride – ' + rideDate + ' at ' + rideTime;
  var props   = PropertiesService.getScriptProperties();
  props.deleteProperty('fb_emailed');  // reset FB-email flag for the new ride
  props.setProperties({
    'pending_date'     : rideDate,
    'pending_time'     : rideTime,
    'pending_location' : rideLocation,
    'pending_sag'      : sagAvailable ? 'yes' : 'no',
    'pending_subject'  : subject,
    'pending_formUrl'  : newForm.getPublishedUrl()
  });

  // ── Preview draft to self ─────────────────────────────────────────────────
  var htmlBody = buildEmailHtml(rideDate, rideTime, rideLocation, sagAvailable, newForm.getPublishedUrl());
  GmailApp.createDraft(PREVIEW_EMAIL, '[PREVIEW] ' + subject, '', { htmlBody: htmlBody });

  return 'Done! Tab "' + rideDate + '" created. A preview draft is in your Gmail — review it, then use "Send Training Ride Emails…" from the menu.';
}


// ═══════════════════════════════════════════════════════════════════════════════
// STEP 2 — SEND DIALOG (test group or full list)
// ═══════════════════════════════════════════════════════════════════════════════

function promptSendEmails() {
  var props = PropertiesService.getScriptProperties();
  var rideDate     = props.getProperty('pending_date');
  var rideTime     = props.getProperty('pending_time');
  var rideLocation = props.getProperty('pending_location');

  if (!rideDate) {
    SpreadsheetApp.getUi().alert('No pending ride found. Create a ride first.');
    return;
  }

  var emailCount = 0;
  var testCount  = 0;
  var loadError  = '';

  try {
    emailCount = getRegistrantEmails().length;
  } catch (e) {
    loadError = e.message;
  }

  try {
    testCount = getEmailsFromSheet('Test Email').length;
  } catch (e) {
    // non-fatal — button will show as disabled
  }

  var routes    = [];
  var routesErr = '';
  try {
    routes = getRoutesFromSheet();
  } catch (e) {
    routesErr = e.message;
  }

  var photo = null;
  try {
    photo = getNextRidePhoto();
  } catch (e) {
    routesErr += (routesErr ? ' | ' : '') + 'Photo: ' + e.message;
  }

  var fbEmailed = PropertiesService.getScriptProperties().getProperty('fb_emailed') === 'yes';
  var defaultPostText = buildFbPostText(rideDate, rideTime, rideLocation, '');

  var html = HtmlService.createHtmlOutput(
    buildSendDialog(
      rideDate, rideTime, rideLocation,
      emailCount, testCount, loadError,
      routes, photo, defaultPostText, fbEmailed, routesErr
    )
  ).setWidth(560).setHeight(780);

  SpreadsheetApp.getUi().showModalDialog(html, '📧 Send Training Ride Emails');
}

function buildSendDialog(rideDate, rideTime, rideLocation, emailCount, testCount, loadError,
                         routes, photo, defaultPostText, fbEmailedAlready, routesErr) {
  var errorHtml    = loadError
    ? '<p style="color:#c0392b;font-size:12px;">⚠️ Could not load registrant list: ' + loadError + '</p>'
    : '';
  var fullDisabled = (loadError || emailCount === 0) ? 'disabled' : '';
  var testDisabled = testCount === 0 ? 'disabled' : '';
  var testLabel    = testCount > 0
    ? 'Send Test (' + testCount + ' address' + (testCount === 1 ? '' : 'es') + ')'
    : 'Send Test (no "Test Email" sheet found)';
  var fullLabel    = emailCount > 0
    ? 'Send to All ' + emailCount + ' Riders →'
    : 'Send to All Riders (none loaded)';

  var routesOptionsHtml = (routes || []).map(function(r) {
    return '<option value="' + escapeHtmlAttr(r.url) + '">' + escapeHtmlText(r.name) + '</option>';
  }).join('\n');

  var photoHtml = photo
    ? '📎 ' + escapeHtmlText(photo.name) + ' <span style="color:#888;">(will be moved to <code>' + RIDE_PHOTO_USED_SUBFOLDER + '/</code> after emailing)</span>'
    : '<span style="color:#c0392b;">No photos available in the ride photos folder.</span>';
  var photoIdJs = photo ? JSON.stringify(photo.id) : 'null';

  var routesErrHtml = routesErr
    ? '<p style="color:#c0392b;font-size:11px;margin:4px 0 0 0;">⚠️ ' + escapeHtmlText(routesErr) + '</p>'
    : '';

  var emailedBadge = fbEmailedAlready
    ? '<span style="color:#2e7d32;font-size:12px;margin-left:6px;">✅ already emailed for this ride</span>'
    : '';

  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; font-size: 13px; padding: 20px; color: #333; }
    .ride-summary { background: #f8f8f8; border-radius: 6px; padding: 12px 16px; margin-bottom: 18px; line-height: 1.8; }
    .ride-summary strong { display: inline-block; width: 70px; }
    .btn { width: 100%; padding: 11px; border: none; border-radius: 5px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 10px; }
    .btn-test { background: #f0f0f0; color: #333; }
    .btn-test:hover:not(:disabled) { background: #e0e0e0; }
    .btn-full { background: #fa65a6; color: #fff; }
    .btn-full:hover:not(:disabled) { background: #e04d90; }
    .btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .divider { border: none; border-top: 1px solid #e0e0e0; margin: 16px 0; }
    .status { font-size: 12px; color: #888; margin-top: 10px; min-height: 18px; }
  </style>
</head>
<body>

  <div class="ride-summary">
    <strong>Date:</strong> ${rideDate}<br>
    <strong>Time:</strong> ${rideTime}<br>
    <strong>Location:</strong> ${rideLocation}
  </div>

  ${errorHtml}

  <label style="display:block;font-weight:bold;margin-bottom:4px;">Additional email message (optional)</label>
  <textarea id="extraText" rows="3" style="width:100%;padding:6px;box-sizing:border-box;border:1px solid #ccc;border-radius:4px;font-family:Arial,sans-serif;font-size:13px;" placeholder="Any extra note to include in this email..."></textarea>

  <button class="btn btn-test" id="testBtn" ${testDisabled} onclick="send('test')">${testLabel}</button>

  <hr class="divider" />

  <div style="font-weight:bold;font-size:14px;margin-bottom:8px;">📘 Facebook Post (emailed to admin) ${emailedBadge}</div>

  <label style="display:block;font-size:12px;margin-top:6px;">Route</label>
  <select id="routeSelect" onchange="handleRouteChange()" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:4px;">
    <option value="">— No route link —</option>
    ${routesOptionsHtml}
    <option value="__custom__">Other (paste URL below)...</option>
  </select>
  <input type="text" id="customRoute" placeholder="Paste route URL" style="display:none;width:100%;padding:6px;margin-top:6px;box-sizing:border-box;border:1px solid #ccc;border-radius:4px;" oninput="handleCustomRouteInput()" />
  ${routesErrHtml}

  <label style="display:block;font-size:12px;margin-top:10px;">Post text (editable)</label>
  <textarea id="postText" rows="9" style="width:100%;padding:6px;box-sizing:border-box;border:1px solid #ccc;border-radius:4px;font-family:Arial,sans-serif;font-size:13px;">${escapeHtmlText(defaultPostText)}</textarea>
  <div style="font-size:11px;color:#888;margin-top:2px;">Tip: pick your route first — changing it rebuilds the text and will overwrite edits.</div>

  <label style="display:block;font-size:12px;margin-top:10px;">Photo</label>
  <div id="photoInfo" style="font-size:12px;padding:6px;background:#f8f8f8;border-radius:4px;">${photoHtml}</div>

  <button class="btn" id="fbBtn" style="background:#1877f2;color:#fff;margin-top:10px;" onclick="emailFb()">Email FB Post to Test Email List</button>

  <hr class="divider" />

  <button class="btn btn-full" id="fullBtn" ${fullDisabled} onclick="send('full')">${fullLabel}</button>

  <p class="status" id="status"></p>

  <script>
    var rideDate_js     = ${JSON.stringify(rideDate)};
    var rideTime_js     = ${JSON.stringify(rideTime)};
    var rideLocation_js = ${JSON.stringify(rideLocation)};
    var photoFileId     = ${photoIdJs};
    var fbEmailed       = ${fbEmailedAlready ? 'true' : 'false'};

    function getSelectedRouteUrl() {
      var sel = document.getElementById('routeSelect').value;
      if (sel === '__custom__') return document.getElementById('customRoute').value.trim();
      return sel;
    }

    function handleRouteChange() {
      var sel = document.getElementById('routeSelect').value;
      document.getElementById('customRoute').style.display = (sel === '__custom__') ? 'block' : 'none';
      rebuildPostText();
    }
    function handleCustomRouteInput() { rebuildPostText(); }

    function rebuildPostText() {
      var url = getSelectedRouteUrl();
      google.script.run
        .withSuccessHandler(function(text) {
          document.getElementById('postText').value = text;
        })
        .buildFbPostText(rideDate_js, rideTime_js, rideLocation_js, url);
    }

    function emailFb() {
      var fbBtn  = document.getElementById('fbBtn');
      var status = document.getElementById('status');
      var text   = document.getElementById('postText').value.trim();
      if (!text) { alert('Post text is empty.'); return; }
      if (!confirm('Email this Facebook post content to the Test Email list?')) return;

      fbBtn.disabled = true;
      status.style.color = '#888';
      status.innerText = '⏳ Sending email...';

      google.script.run
        .withSuccessHandler(function(msg) {
          fbEmailed = true;
          status.style.color = '#2e7d32';
          status.innerText = '✅ ' + msg;
          fbBtn.innerText = 'Emailed ✓';
          document.getElementById('photoInfo').innerHTML = '<span style="color:#2e7d32;">Photo moved to used/ folder.</span>';
        })
        .withFailureHandler(function(err) {
          fbBtn.disabled = false;
          status.style.color = '#c0392b';
          status.innerText = '❌ ' + err.message;
        })
        .emailFbPostToAdmin(text, photoFileId);
    }

    function send(mode) {
      if (mode === 'full' && !fbEmailed) {
        if (!confirm('You haven\\'t emailed the Facebook post for this ride. Send the live email anyway?')) return;
      }
      var testBtn = document.getElementById('testBtn');
      var fullBtn = document.getElementById('fullBtn');
      var status  = document.getElementById('status');
      var extraText = document.getElementById('extraText').value;

      testBtn.disabled = true;
      fullBtn.disabled = true;
      status.style.color = '#888';
      status.innerText = mode === 'test'
        ? '⏳ Sending test emails...'
        : '⏳ Sending to full list, please wait...';

      google.script.run
        .withSuccessHandler(function(msg) {
          status.style.color = '#2e7d32';
          status.innerText = '✅ ' + msg;
          if (mode === 'full') {
            setTimeout(function() { google.script.host.close(); }, 3000);
          } else {
            fullBtn.disabled = ${emailCount === 0 ? 'true' : 'false'};
          }
        })
        .withFailureHandler(function(err) {
          testBtn.disabled = ${testCount === 0 ? 'true' : 'false'};
          fullBtn.disabled = ${emailCount === 0 ? 'true' : 'false'};
          status.style.color = '#c0392b';
          status.innerText = '❌ ' + err.message;
        })
        .executeSend(mode, extraText);
    }
  </script>
</body>
</html>`;
}

/**
 * Called by the send dialog. mode is 'test' or 'full'.
 *
 * Uses the Gmail advanced service (Gmail.Users.Messages.send) so we can attach
 * a List-Unsubscribe header — iCloud/Spectrum reward its presence.
 * Two-pass: normal recipients first at 2s pacing; strict-domain recipients
 * (iCloud, Spectrum/Charter/TWC) get a slower 10s pacing in a second pass.
 */
function executeSend(mode, extraText) {
  if (typeof Gmail === 'undefined') {
    throw new Error('Enable the "Gmail API" advanced service in Apps Script (Services → + Gmail API), then re-run.');
  }

  var props        = PropertiesService.getScriptProperties();
  var subject      = props.getProperty('pending_subject');
  var rideDate     = props.getProperty('pending_date');
  var rideTime     = props.getProperty('pending_time');
  var rideLocation = props.getProperty('pending_location');
  var sagAvailable = props.getProperty('pending_sag') === 'yes';
  var formUrl      = props.getProperty('pending_formUrl');

  if (!rideDate) throw new Error('No pending ride found. Create a ride first.');

  var emails = mode === 'test'
    ? getEmailsFromSheet('Test Email')
    : getRegistrantEmails();

  if (emails.length === 0) throw new Error('No emails found for mode: ' + mode);

  var htmlBody = buildEmailHtml(rideDate, rideTime, rideLocation, sagAvailable, formUrl, extraText);
  var textBody = buildPlainTextBody(rideDate, rideTime, rideLocation, sagAvailable, formUrl, extraText);

  // Try to build an .ics attachment. If the date/time can't be parsed, skip the attachment
  // rather than failing the send — a rider email without a calendar invite is still useful.
  var attachments = [];
  try {
    attachments = [buildIcsBlob(rideDate, rideTime, rideLocation)];
  } catch (e) {
    Logger.log('Skipping .ics attachment: ' + e.message);
  }

  // Split into normal vs. strict-domain recipients
  var mainList = [];
  var strictList = [];
  emails.forEach(function(email) {
    if (isStrictDomain(email)) strictList.push(email);
    else mainList.push(email);
  });

  var sentCount    = 0;
  var failedEmails = [];

  function sendOneWithRetry(email, retryDelayMs) {
    try {
      sendRiderEmail(email, subject, htmlBody, textBody, attachments);
      sentCount++;
      return;
    } catch (e) {
      Logger.log('First attempt failed: ' + email + ' — ' + e.message + '. Retrying in ' + (retryDelayMs / 1000) + 's...');
    }
    Utilities.sleep(retryDelayMs);
    try {
      sendRiderEmail(email, subject, htmlBody, textBody, attachments);
      sentCount++;
    } catch (e2) {
      Logger.log('Retry failed: ' + email + ' — ' + e2.message);
      failedEmails.push(email);
    }
  }

  Logger.log('Main pass: ' + mainList.length + ' recipients (2s pacing)');
  mainList.forEach(function(email) {
    sendOneWithRetry(email, 5000);
    Utilities.sleep(2000);
  });

  if (strictList.length > 0) {
    Logger.log('Pausing 30s before strict-domain pass (' + strictList.length + ' recipients)...');
    Utilities.sleep(30000);
    strictList.forEach(function(email) {
      sendOneWithRetry(email, 15000);
      Utilities.sleep(10000);
    });
  }

  // Only clear pending state after the full send
  if (mode === 'full') {
    props.deleteAllProperties();
  }

  var summary = sentCount + ' email' + (sentCount === 1 ? '' : 's') + ' sent.';
  if (strictList.length > 0) summary += ' (incl. ' + strictList.length + ' strict-domain)';
  if (failedEmails.length > 0) summary += ' ⚠️ Failed: ' + failedEmails.join(', ');
  return summary;
}


// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Converts the latest parsed_tickets_*.xlsx to a temporary Google Sheet,
 * runs the callback with it, then trashes the temp file.
 */
function withRegistrantSpreadsheet(callback) {
  var folder = DriveApp.getFolderById(REGISTRANT_FOLDER_ID);
  var files  = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);

  var latest     = null;
  var latestDate = new Date(0);

  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().match(/^parsed_tickets_.*\.xlsx$/i)) {
      if (f.getLastUpdated() > latestDate) {
        latestDate = f.getLastUpdated();
        latest = f;
      }
    }
  }

  if (!latest) throw new Error('No parsed_tickets_*.xlsx found in the registrant folder.');
  Logger.log('Using: ' + latest.getName() + ' (updated ' + latestDate + ')');

  // Convert xlsx → Google Sheet, retrying on transient Drive errors (e.g. "Empty response")
  var convertedFile = copyXlsxToGoogleSheetWithRetry(latest.getId());

  try {
    var ss = SpreadsheetApp.openById(convertedFile.id);
    return callback(ss);
  } finally {
    // Always clean up the temp file whether or not an error occurred
    try {
      DriveApp.getFileById(convertedFile.id).setTrashed(true);
    } catch (e) {
      Logger.log('Temp file cleanup failed (non-fatal): ' + e.message);
    }
  }
}

/**
 * Calls Drive.Files.copy with up to 3 attempts and small backoff.
 * Drive occasionally returns "Empty response" or 5xx for no good reason.
 */
function copyXlsxToGoogleSheetWithRetry(sourceFileId) {
  var lastErr;
  var delayMs = 1000;
  for (var i = 0; i < 3; i++) {
    try {
      return Drive.Files.copy(
        { title: '__tmp_registrants__', mimeType: MimeType.GOOGLE_SHEETS },
        sourceFileId,
        { convert: true }
      );
    } catch (e) {
      lastErr = e;
      Logger.log('Drive.Files.copy attempt ' + (i + 1) + ' failed: ' + e.message);
      Utilities.sleep(delayMs);
      delayMs *= 2;  // 1s, 2s, 4s
    }
  }
  throw new Error('Drive.Files.copy failed after 3 attempts: ' + (lastErr && lastErr.message));
}

/**
 * Returns a deduplicated array of email addresses from all configured rider sheets.
 */
function getRegistrantEmails() {
  return withRegistrantSpreadsheet(function(ss) {
    var sheetMap = {};
    ss.getSheets().forEach(function(s) { sheetMap[s.getName()] = s; });

    var emails = {};

    RIDER_SHEETS.forEach(function(sheetName) {
      var sheet = sheetMap[sheetName];
      if (!sheet) {
        Logger.log('Sheet not found (skipping): ' + sheetName);
        return;
      }

      var data = sheet.getDataRange().getValues();
      if (data.length < 2) return;

      var headers  = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
      var emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase());
      if (emailCol === -1) {
        Logger.log('No "' + EMAIL_COLUMN_NAME + '" column in: ' + sheetName);
        return;
      }

      for (var i = 1; i < data.length; i++) {
        var email = String(data[i][emailCol]).trim().toLowerCase();
        if (email && email.indexOf('@') > 0) emails[email] = true;
      }
      Logger.log('Read ' + (data.length - 1) + ' rows from: ' + sheetName);
    });

    return Object.keys(emails);
  });
}

/**
 * Returns emails from a single named sheet in the registrant file.
 */
function getEmailsFromSheet(sheetName) {
  return withRegistrantSpreadsheet(function(ss) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers  = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var emailCol = headers.indexOf(EMAIL_COLUMN_NAME.toLowerCase());
    if (emailCol === -1) throw new Error('No "' + EMAIL_COLUMN_NAME + '" column in sheet "' + sheetName + '".');

    var emails = [];
    for (var i = 1; i < data.length; i++) {
      var email = String(data[i][emailCol]).trim().toLowerCase();
      if (email && email.indexOf('@') > 0) emails.push(email);
    }
    return emails;
  });
}

/**
 * Builds the HTML email body. Shared by the preview draft and all sends.
 */
function buildEmailHtml(rideDate, rideTime, rideLocation, sagAvailable, formUrl, extraText) {
  var extraHtml = '';
  if (extraText && String(extraText).trim()) {
    var escaped = String(extraText)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\n/g, '<br>');
    extraHtml =
      '<p style="font-size: 17px; color: #333; background: #fff6fa; border-left: 4px solid #fa65a6; padding: 12px 16px; margin: 20px 0;">' +
      escaped + '</p>';
  }
  var mapsUrl  = LOCATIONS[rideLocation] || '';
  var mapsLink = mapsUrl
    ? ' (<a href="' + mapsUrl + '" style="color: #fa65a6;">Get Directions</a>)'
    : '';

  var sagHtml = sagAvailable
    ? '&#9989; Limited SAG (support and gear) space will be available.'
    : '&#10060; SAG (support and gear) will <strong>not</strong> be available for this ride.';

  return (
    '<html><head><meta charset="UTF-8"></head><body>' +
    '<div style="font-family: Arial, sans-serif; font-size: 15px; color: #333; max-width: 600px;">' +

    '<p style="font-size: 17px;">Our next training ride is coming up! Here are the details:</p>' +

    '<table style="margin: 16px 0; border-collapse: collapse;">' +
    '<tr><td style="padding: 6px 16px 6px 0; font-weight: bold;">&#128197; Date</td><td style="padding: 6px 0;">' + rideDate + '</td></tr>' +
    '<tr><td style="padding: 6px 16px 6px 0; font-weight: bold;">&#128336; Time</td><td style="padding: 6px 0;">' + rideTime + '</td></tr>' +
    '<tr><td style="padding: 6px 16px 6px 0; font-weight: bold;">&#128205; Location</td><td style="padding: 6px 0;">' + rideLocation + mapsLink + '</td></tr>' +
    '</table>' +

    '<p>' + sagHtml + '</p>' +

    '<p>Please arrive a few minutes early and be ready to roll at <strong>' + rideTime + '</strong>. ' +
    'Come prepared with plenty of water in a cycling-specific water bottle (no disposable or steel bottles), a spare tube, and the tools needed to fix a flat.</p>' +

    '<p style="margin: 24px 0;">' +
    '<a href="' + formUrl + '" style="background-color: #fa65a6; color: #ffffff; padding: 12px 28px; ' +
    'text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px;">Sign Up for This Ride &#8594;</a>' +
    '<br><span style="font-size: 12px; color: #888; margin-top: 6px; display: inline-block;">(Signing up helps us gauge interest and plan accordingly &#8212; even if you\'re unsure, please sign up!)</span>' +
    '</p>' +

    extraHtml +

    '<p><a href="' + SCHEDULE_URL + '" style="color: #fa65a6;">&#128203; View the full training schedule</a> ' +
    '<span style="font-size: 12px; color: #888;">&#8592; Bookmark this! Same link every year.</span></p>' +

    '<hr style="border: none; border-top: 1px solid #ddd; margin: 24px 0;" />' +

    '<p style="font-size: 13px; color: #666;">' +
    '<strong>Reminders:</strong><br>' +
    '&#8226; New riders must complete <strong>two</strong> training rides. Returning riders must complete <strong>one</strong>.<br>' +
    '&#8226; Be sure to sign in at the training ride.<br>' +
    '&#8226; Routes may change at the last minute &#8212; check the Ride Line and the <a href="https://www.facebook.com/groups/theridemv" style="color: #fa65a6;">CNY Ride Family</a> Facebook group for updates.<br>' +
    '&#8226; Ride Line: ' + RIDE_LINE +
    '</p>' +

    '<hr style="border: none; border-top: 1px solid #ddd; margin: 24px 0;" />' +

    '<p style="font-size: 12px; color: #999;">' +
    'You\'re receiving this because you registered for the 2026 Ride for Missing Children &#8211; MV.<br>' +
    '&#9993; <a href="mailto:' + CONTACT_EMAIL + '" style="color: #999;">' + CONTACT_EMAIL + '</a> &#183; ' + PO_BOX + '<br>' +
    '<a href="https://www.rfmc-mv.org/unsubscribe" style="color: #999;">Unsubscribe</a>' +
    '</p>' +

    '</div></body></html>'
  );
}


// ═══════════════════════════════════════════════════════════════════════════════
// FACEBOOK POST (EMAILED TO ADMIN)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Returns [{url, name}] from the Routes tab in the master sheet.
 * Expected columns: URL | Name | Notes (header row required).
 */
function getRoutesFromSheet() {
  var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var sheet = ss.getSheetByName(ROUTES_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + ROUTES_SHEET_NAME + '" not found in master spreadsheet.');

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var urlCol  = headers.indexOf('url');
  var nameCol = headers.indexOf('name');
  if (urlCol === -1) throw new Error('No "URL" column in Routes sheet.');

  var routes = [];
  for (var i = 1; i < data.length; i++) {
    var url  = String(data[i][urlCol]).trim();
    var name = nameCol > -1 ? String(data[i][nameCol]).trim() : '';
    if (!url) continue;
    routes.push({ url: url, name: name || url });
  }
  return routes;
}

/**
 * Returns the first image file in the ride-photo folder (ignoring the "used" subfolder).
 * Returns null if none is available.
 */
function getNextRidePhoto() {
  var folder = DriveApp.getFolderById(RIDE_PHOTO_FOLDER_ID);
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var mime = f.getMimeType() || '';
    if (mime.indexOf('image/') === 0) {
      return { id: f.getId(), name: f.getName(), mime: mime };
    }
  }
  return null;
}

/**
 * Moves a file from the ride-photo folder into the "used" subfolder (auto-creating it).
 */
function moveRidePhotoToUsed(fileId) {
  var parent = DriveApp.getFolderById(RIDE_PHOTO_FOLDER_ID);
  var usedIter = parent.getFoldersByName(RIDE_PHOTO_USED_SUBFOLDER);
  var usedFolder = usedIter.hasNext() ? usedIter.next() : parent.createFolder(RIDE_PHOTO_USED_SUBFOLDER);
  var file = DriveApp.getFileById(fileId);
  usedFolder.addFile(file);
  parent.removeFile(file);
}

/**
 * Assembles the default FB post text. Called from both server (initial fill) and client (on route change).
 */
function buildFbPostText(rideDate, rideTime, rideLocation, routeUrl) {
  var lines = [
    '🚴 Training ride time!',
    '',
    'We have a training ride coming up — gear up, and come ride with us!',
    '',
    '📅 Date: ' + rideDate,
    '🕠 Time: ' + rideTime,
    '📍 Location: ' + rideLocation
  ];
  if (routeUrl) lines.push('🗺️ Route: ' + routeUrl);
  lines.push('');
  lines.push('See you out there! 🙌');
  return lines.join('\n');
}

/**
 * Returns a photo blob that's guaranteed to be <= 10MB. If the original file is already
 * under the limit, returns the raw blob. Otherwise fetches progressively smaller JPEG
 * renditions via Drive's thumbnail endpoint until one fits.
 */
function getPhotoBlobForEmail(photoFileId) {
  var MAX_BYTES = 10 * 1024 * 1024;
  var file = DriveApp.getFileById(photoFileId);
  var original = file.getBlob().copyBlob();

  if (original.getBytes().length <= MAX_BYTES) return original;

  Logger.log('Photo "' + file.getName() + '" is ' +
             Math.round(original.getBytes().length / 1024 / 1024) + 'MB — resizing.');

  var token = ScriptApp.getOAuthToken();
  var widths = [2400, 2000, 1600, 1200, 900];

  for (var i = 0; i < widths.length; i++) {
    var url = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(photoFileId) +
              '&sz=w' + widths[i];
    try {
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
        followRedirects: true
      });
      if (resp.getResponseCode() === 200) {
        var blob = resp.getBlob();
        var len  = blob.getBytes().length;
        if (len > 0 && len <= MAX_BYTES) {
          Logger.log('Using ' + widths[i] + 'px thumbnail (' + Math.round(len / 1024) + ' KB).');
          return blob.setName(file.getName()).setContentType('image/jpeg');
        }
      }
    } catch (e) {
      Logger.log('Thumbnail w' + widths[i] + ' fetch failed: ' + e.message);
    }
  }

  throw new Error('Could not produce a resized photo under 10MB for "' + file.getName() + '".');
}

/**
 * Called by the dialog. Emails the FB post content (+ photo attachment) to FB_POST_RECIPIENT,
 * marks state, and moves the photo to used/.
 */
function emailFbPostToAdmin(message, photoFileId) {
  if (!message || !String(message).trim()) throw new Error('Post text is empty.');

  var props    = PropertiesService.getScriptProperties();
  var rideDate = props.getProperty('pending_date') || 'ride';
  var subject  = 'FB Post Content – ' + rideDate;

  // Recipients: the same "Test Email" list used for test sends.
  // Falls back to FB_POST_RECIPIENT if the Test Email sheet is missing/empty.
  var recipients;
  try {
    recipients = getEmailsFromSheet('Test Email');
  } catch (e) {
    recipients = [];
  }
  if (recipients.length === 0) recipients = [FB_POST_RECIPIENT];

  var htmlBody =
    '<html><head><meta charset="UTF-8"></head><body>' +
    '<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">' +
    '<p>Here is the Facebook post content for the upcoming training ride. ' +
    'Copy and paste it into the RFMC-MV Facebook page.</p>' +
    '<hr>' +
    '<pre style="white-space: pre-wrap; font-family: inherit; font-size: 14px; ' +
    'background: #f8f8f8; padding: 12px; border-radius: 6px;">' +
    escapeHtmlTextSafe(message) + '</pre>' +
    (photoFileId ? '<p>&#128206; Photo attached.</p>' : '<p><em>No photo attached.</em></p>') +
    '</div></body></html>';

  var options = { htmlBody: htmlBody, charset: 'UTF-8' };
  var photoName = '';
  if (photoFileId) {
    var file = DriveApp.getFileById(photoFileId);
    photoName = file.getName();
    options.attachments = [getPhotoBlobForEmail(photoFileId)];
  }

  // Pass empty string as plain-text body — Gmail will use the HTML htmlBody instead,
  // and some clients render a UTF-8-broken text/plain fallback if we provide one here.
  GmailApp.sendEmail(recipients.join(','), subject, '', options);

  if (photoFileId) {
    try {
      moveRidePhotoToUsed(photoFileId);
    } catch (e) {
      Logger.log('Email sent but photo move failed: ' + e.message);
    }
  }

  props.setProperty('fb_emailed', 'yes');
  return 'Emailed to ' + recipients.length + ' address' +
         (recipients.length === 1 ? '' : 'es') +
         (photoName ? ' with ' + photoName : '') + '.';
}


// ═══════════════════════════════════════════════════════════════════════════════
// CALENDAR INVITE (.ICS)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Parses free-text rideDate ("April 23rd, 2026") + rideTime ("5:30 PM") into a Date.
 * Throws if the combination can't be parsed.
 */
function parseRideStart(rideDate, rideTime) {
  var cleanDate = String(rideDate).replace(/(\d+)(st|nd|rd|th)/ig, '$1').replace(/,/g, '');

  // If no 4-digit year is in the date string, append the current year.
  // Otherwise V8's loose date parser can misread the time ("1:00 PM") as a year (2001).
  if (!/\b\d{4}\b/.test(cleanDate)) {
    cleanDate += ' ' + new Date().getFullYear();
  }

  var combined = cleanDate + ' ' + String(rideTime);
  var d = new Date(combined);
  if (isNaN(d.getTime())) throw new Error('Could not parse "' + combined + '" as a date/time.');
  return d;
}

/**
 * Builds an .ics calendar-invite blob for the ride.
 */
function buildIcsBlob(rideDate, rideTime, rideLocation) {
  var start = parseRideStart(rideDate, rideTime);
  var end   = new Date(start.getTime() + RIDE_DURATION_HOURS * 60 * 60 * 1000);

  var fmt = function(d) {
    return Utilities.formatDate(d, EVENT_TIMEZONE, "yyyyMMdd'T'HHmmss");
  };
  var uid     = 'ride-' + start.getTime() + '@rfmc-mv.org';
  var dtstamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  var desc    = 'RFMC-MV training ride departing from ' + rideLocation +
                '. See ' + SCHEDULE_URL + ' for the full schedule.';

  var ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//RFMC-MV//Training Ride//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + dtstamp,
    'DTSTART;TZID=' + EVENT_TIMEZONE + ':' + fmt(start),
    'DTEND;TZID=' + EVENT_TIMEZONE + ':' + fmt(end),
    'SUMMARY:RFMC-MV Training Ride',
    'LOCATION:' + String(rideLocation).replace(/,/g, '\\,'),
    'DESCRIPTION:' + desc.replace(/,/g, '\\,'),
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');

  return Utilities.newBlob(ics, 'text/calendar; charset=UTF-8; method=PUBLISH', 'training-ride.ics');
}


// ═══════════════════════════════════════════════════════════════════════════════
// HTML escape helpers
// ═══════════════════════════════════════════════════════════════════════════════

function escapeHtmlText(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Same as escapeHtmlText, but also converts every non-ASCII character (including
 * surrogate-pair emojis) into a numeric HTML entity. Avoids MIME-transport encoding
 * issues where 4-byte UTF-8 sequences get garbled to replacement characters.
 */
function escapeHtmlTextSafe(s) {
  var str = String(s == null ? '' : s);
  var out = '';
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    // Combine surrogate pair into full codepoint
    if (code >= 0xD800 && code <= 0xDBFF && i + 1 < str.length) {
      var low = str.charCodeAt(i + 1);
      if (low >= 0xDC00 && low <= 0xDFFF) {
        code = 0x10000 + (code - 0xD800) * 0x400 + (low - 0xDC00);
        i++;
      }
    }
    if (code === 0x26) out += '&amp;';
    else if (code === 0x3C) out += '&lt;';
    else if (code === 0x3E) out += '&gt;';
    else if (code < 0x80) out += String.fromCharCode(code);
    else out += '&#' + code + ';';
  }
  return out;
}

function escapeHtmlAttr(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}


// ═══════════════════════════════════════════════════════════════════════════════
// RAW MIME SENDING (Gmail advanced service — enables List-Unsubscribe header)
// ═══════════════════════════════════════════════════════════════════════════════

function isStrictDomain(email) {
  var at = String(email).indexOf('@');
  if (at < 0) return false;
  var domain = String(email).substring(at + 1).toLowerCase();
  for (var i = 0; i < STRICT_DOMAINS.length; i++) {
    var d = STRICT_DOMAINS[i];
    if (domain === d || domain.endsWith('.' + d)) return true;
  }
  return false;
}

/**
 * Plain-text alternative for the rider email. Improves deliverability (HTML-only
 * messages are scored worse by some receivers).
 */
function buildPlainTextBody(rideDate, rideTime, rideLocation, sagAvailable, formUrl, extraText) {
  var lines = [
    'Our next training ride is coming up.',
    '',
    'Date: ' + rideDate,
    'Time: ' + rideTime,
    'Location: ' + rideLocation,
    '',
    sagAvailable
      ? 'Limited SAG (support and gear) space will be available.'
      : 'SAG (support and gear) will not be available for this ride.',
    '',
    'Please arrive a few minutes early. Come prepared with plenty of water in a',
    'cycling-specific water bottle (no disposable or steel bottles), a spare',
    'tube, and the tools needed to fix a flat.'
  ];
  if (extraText && String(extraText).trim()) {
    lines.push('');
    lines.push('---');
    lines.push(String(extraText).trim());
    lines.push('---');
  }
  lines.push('');
  lines.push('Sign up: ' + formUrl);
  lines.push('Full training schedule: ' + SCHEDULE_URL);
  lines.push('');
  lines.push('Reminders:');
  lines.push('- New riders must complete two training rides. Returning riders must complete one.');
  lines.push('- Be sure to sign in at the training ride.');
  lines.push('- Routes may change at the last minute. Check the Ride Line or the CNY Ride Family Facebook group.');
  lines.push('- Ride Line: ' + RIDE_LINE);
  lines.push('');
  lines.push('---');
  lines.push("You're receiving this because you registered for the 2026 Ride for Missing Children - MV.");
  lines.push('Contact: ' + CONTACT_EMAIL);
  lines.push('Unsubscribe: ' + UNSUBSCRIBE_URL);
  return lines.join('\r\n');
}

/**
 * RFC 2047 encoded-word for headers that contain non-ASCII characters.
 */
function encodeHeaderValue(value) {
  var s = String(value);
  if (/^[\x00-\x7F]*$/.test(s)) return s;
  return '=?UTF-8?B?' + Utilities.base64Encode(s, Utilities.Charset.UTF_8) + '?=';
}

/**
 * Builds a full RFC 5322 message (multipart/alternative with text+HTML, optionally
 * wrapped in multipart/mixed when attachments are present).
 */
function buildRawMimeMessage(opts) {
  var altBoundary = 'alt_' + Utilities.getUuid();
  var mixBoundary = 'mix_' + Utilities.getUuid();
  var hasAttachments = opts.attachments && opts.attachments.length > 0;

  var headers = [];
  headers.push('To: ' + opts.to);
  headers.push('Subject: ' + encodeHeaderValue(opts.subject));
  headers.push('MIME-Version: 1.0');
  if (opts.extraHeaders) {
    Object.keys(opts.extraHeaders).forEach(function(k) {
      headers.push(k + ': ' + opts.extraHeaders[k]);
    });
  }

  var body = [];
  if (hasAttachments) {
    headers.push('Content-Type: multipart/mixed; boundary="' + mixBoundary + '"');
    body.push('--' + mixBoundary);
    body.push('Content-Type: multipart/alternative; boundary="' + altBoundary + '"');
    body.push('');
  } else {
    headers.push('Content-Type: multipart/alternative; boundary="' + altBoundary + '"');
  }

  // text/plain part
  body.push('--' + altBoundary);
  body.push('Content-Type: text/plain; charset=UTF-8');
  body.push('Content-Transfer-Encoding: base64');
  body.push('');
  body.push(Utilities.base64Encode(opts.textBody || '', Utilities.Charset.UTF_8));
  body.push('');

  // text/html part
  body.push('--' + altBoundary);
  body.push('Content-Type: text/html; charset=UTF-8');
  body.push('Content-Transfer-Encoding: base64');
  body.push('');
  body.push(Utilities.base64Encode(opts.htmlBody, Utilities.Charset.UTF_8));
  body.push('');
  body.push('--' + altBoundary + '--');

  if (hasAttachments) {
    opts.attachments.forEach(function(blob) {
      body.push('');
      body.push('--' + mixBoundary);
      body.push('Content-Type: ' + blob.getContentType());
      body.push('Content-Disposition: attachment; filename="' + blob.getName() + '"');
      body.push('Content-Transfer-Encoding: base64');
      body.push('');
      body.push(Utilities.base64Encode(blob.getBytes()));
    });
    body.push('');
    body.push('--' + mixBoundary + '--');
  }

  return headers.join('\r\n') + '\r\n\r\n' + body.join('\r\n');
}

/**
 * Sends one rider email via the Gmail API with List-Unsubscribe headers.
 * Throws on failure.
 */
function sendRiderEmail(to, subject, htmlBody, textBody, attachments) {
  var raw = buildRawMimeMessage({
    to: to,
    subject: subject,
    htmlBody: htmlBody,
    textBody: textBody,
    attachments: attachments,
    extraHeaders: {
      'List-Unsubscribe': '<mailto:' + CONTACT_EMAIL + '?subject=Unsubscribe>, <' + UNSUBSCRIBE_URL + '>'
    }
  });
  var encoded = Utilities.base64EncodeWebSafe(raw).replace(/=+$/g, '');
  Gmail.Users.Messages.send({ raw: encoded }, 'me');
}