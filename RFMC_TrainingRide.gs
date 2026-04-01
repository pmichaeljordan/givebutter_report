/**
 * Training Ride Creator
 * Creates a new training ride form, links it to the master sheet, and drafts an email.
 */

// ─── CONFIG ───────────────────────────────────────────────────────────────────
var MASTER_SHEET_ID  = '1LND_ZZJpalSgnyyVtOCEJgOy_N4kh-B3e9gfbguCpPE';
var TEMPLATE_FORM_ID = '1vew47VRL4NK591rpe_w8NIEE1yjGLj_wOK_HtI-wFj4';
var FORM_FOLDER_ID   = '1fgBJisqSUYmBYIzxwR6jIRcgIHF2wtO1';

var RIDE_LINE      = '315-ASK-RIDE (315-275-7433)';
var CANCEL_GROUPS  = "the Ride Line and the 'CNY Ride Family' Facebook group";
var SCHEDULE_URL   = 'https://www.rfmc-mv.org/tr-schedule';
var CONTACT_EMAIL  = 'shepherds@rfmc-mv.org';
var PO_BOX         = 'PO Box 6, Washington Mills, NY 13479';

var LOCATIONS = {
  'Westmoreland Upper Elementary School'     : 'https://goo.gl/maps/5d8hHjof31LnTsrU6',
  'Whitesboro High School'         : 'https://www.google.com/maps/place/Whitesboro+High+School/@43.1514062,-75.2898054,2743m/data=!3m1!1e3!4m6!3m5!1s0x89d93f097d74a2dd:0x5a3d9188a3c9f65!8m2!3d43.1513431!4d-75.2892904!16s%2Fm%2F076qg69?entry=ttu',
  'Town Of Paris Park'             : 'https://goo.gl/maps/Tc8s1yY22fdPZTjv7',
  'Clinton Arena'                  : 'https://www.google.com/maps/place/Clinton+Arena/@43.0537422,-75.3777248,466m/data=!3m1!1e3!4m6!3m5!1s0x89d969a494a3f621:0x2b17bdefc7364628!8m2!3d43.0536111!4d-75.3769444!16s%2Fm%2F026r83r?entry=ttu&g_ep=EgoyMDI1MDQyMS4wIKXMDSoASAFQAw%3D%3D',
  'Clinton Hannaford'              : 'https://www.google.com/maps/place/Hannaford/@43.0423703,-75.3868412,241m/data=!3m1!1e3!4m5!3m4!1s0x89d969000a56f001:0xba3084e46ff4d64a!8m2!3d43.0421227!4d-75.3872056',
  'Herkimer Hannaford'             : 'https://www.google.com/maps/place/Hannaford/@43.0272495,-74.9817993,17.25z/data=!4m5!3m4!1s0x0:0xd0b658e986c1be0d!8m2!3d43.0296732!4d-74.9812633',
  'New Hartford Recreation Center' : 'https://www.google.com/maps/place/New+Hartford+Recreation+Center+Ice+Skating+Rink/@43.0717038,-75.2856762,346m/data=!3m1!1e3!4m5!3m4!1s0x89d9417bb8ab4d57:0x39a39051c7aa452f!8m2!3d43.0716184!4d-75.28553',
  'MVCC P1/P2 Parking Lot'        : 'https://www.google.com/maps/@43.0997532,-75.1528652,83m/data=!3m1!1e3',
  'Sauquoit Valley High School'    : 'https://goo.gl/maps/HYRQNf7R5SGGAHEw6',
  'Marcy Town Hall'                : 'https://goo.gl/maps/vRVTq1M2ZLWdeV1Z9',
  'Kuyahoora Valley Town Park'     : 'https://www.google.com/maps/place/43%C2%B013\'39.4%22N+75%C2%B003\'27.6%22W/@43.2275287,-75.0574343,1365m/data=!3m1!1e3!4m4!3m3!8m2!3d43.2276111!4d-75.0576667?entry=ttu',
  'Byrne Dairy (Clinton)'          : 'https://www.google.com/maps/place/Byrne+Dairy+and+Deli/@43.0792994,-75.3759878,1128m/data=!3m2!1e3!4b1!4m6!3m5!1s0x89d9698bc5e95267:0xc73d5696ca77a0dc!8m2!3d43.0792994!4d-75.3734129!16s%2Fg%2F1td2hx8t?entry=ttu&g_ep=EgoyMDI1MDQyMy4wIKXMDSoASAFQAw%3D%3D'
};
// ─────────────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚴 Training Rides')
    .addItem('Create New Ride', 'createTrainingRide')
    .addToUi();
}

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

  // ── Copy template form (preserves image and colors) ───────────────────────
  var templateFile = DriveApp.getFileById(TEMPLATE_FORM_ID);
  var newFormFile  = templateFile.makeCopy(formTitle);
  var newFormId    = newFormFile.getId();

  // ── Move to the correct folder ────────────────────────────────────────────
  var targetFolder = DriveApp.getFolderById(FORM_FOLDER_ID);
  targetFolder.addFile(newFormFile);
  DriveApp.getRootFolder().removeFile(newFormFile);

  // ── Publish, update title and description ─────────────────────────────────
  var newForm = FormApp.openById(newFormId);
  newForm.setPublished(true);
  newForm.setTitle(formTitle);
  newForm.setDescription(description);
  newForm.setAcceptingResponses(true);

  // ── Link to master sheet and create named tab ─────────────────────────────
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

  // ── Build Gmail draft (HTML) ──────────────────────────────────────────────
  var formUrl  = newForm.getPublishedUrl();
  var mapsUrl  = LOCATIONS[rideLocation] || '';
  var mapsLink = mapsUrl
    ? ' (<a href="' + mapsUrl + '" style="color: #fa65a6;">Get Directions</a>)'
    : '';

  var subject = 'Training Ride – ' + rideDate + ' at ' + rideTime;

  var sagHtml = sagAvailable
    ? '&#9989; Limited SAG (support and gear) space will be available.'
    : '&#10060; SAG (support and gear) will <strong>not</strong> be available for this ride.';

  var htmlBody =
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
    '&#9993; <a href="mailto:' + CONTACT_EMAIL + '" style="color: #999;">' + CONTACT_EMAIL + '</a> &#183; ' + PO_BOX +
    '</p>' +

    '</div></body></html>';

  GmailApp.createDraft('', subject, '', { htmlBody: htmlBody });

  return 'Done! Tab "' + rideDate + '" added to master sheet. Check Gmail Drafts to review and send.';
}
