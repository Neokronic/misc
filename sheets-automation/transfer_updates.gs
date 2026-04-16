/**
 * No menu on open — keeps UI clean during screen sharing.
 * Use Extensions > Macros > CopyNotes (Ctrl+Alt+Shift+1) to trigger.
 */
function onOpen() {
  // Menu hidden by default. Run showMenu() from script editor if needed.
}

/**
 * Manually show the menu (run from Extensions > Apps Script if needed).
 */
function showMenu() {
  var currentUser = Session.getActiveUser().getEmail();
  if (currentUser !== 'kiran.kamreddy@celerdata.com') return;

  SpreadsheetApp.getUi()
    .createMenu('📋 Update Tools')
    .addItem('Copy Call Notes to WA1', 'CopyNotes')
    .addToUi();
}

/**
 * Copies rows from Dashboard where Column J has content into WA1.
 * Transfers: Key (with hyperlink), Issue Type, Summary, PM, Assignee, Priority, Status, Call Notes.
 * Skips Column I (Updates) to avoid confusion.
 *
 * Same-day behavior:
 *   If today's date group already exists at the top of WA1, appends to it:
 *   - Existing keys: appends new notes to Col H (separated by " | ")
 *   - New keys: inserts new rows at the end of the date group
 *
 * New-day behavior:
 *   Inserts new rows at the top (row 2) with an auto-generated date label.
 *   Sets cell formatting to overflow (CLIP). Groups data rows under the date label.
 *   Adds empty separator row between date groups.
 *
 * Always updates Dashboard Column I and R VLOOKUP to reference the current WA1 range.
 * Then clears Column J in Dashboard.
 */
function CopyNotes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  var wa1 = ss.getSheetByName('WA1');

  if (!dashboard || !wa1) {
    SpreadsheetApp.getUi().alert('Could not find Dashboard or WA1 sheet.');
    return;
  }

  var lastRow = dashboard.getLastRow();
  if (lastRow < 4) {
    SpreadsheetApp.getUi().alert('No data rows found in Dashboard.');
    return;
  }

  // Read Dashboard data (row 4 onward, columns A through J) in one batch
  var dataRange = dashboard.getRange(4, 1, lastRow - 3, 10);
  var data = dataRange.getValues();

  var rowsToClear = [];
  var transferRows = [];
  var jiraKeys = [];

  for (var i = 0; i < data.length; i++) {
    var colJ = data[i][9]; // Column J (index 9) = Call Notes
    if (colJ && colJ.toString().trim() !== '') {
      var jiraKey = data[i][1].toString().trim(); // Column B (plain key)
      if (!jiraKey) continue;

      transferRows.push([
        jiraKey,    // A: Key (will add hyperlink after)
        data[i][2], // B: Issue Type
        data[i][3], // C: Summary
        data[i][4], // D: PM
        data[i][5], // E: Assignee
        data[i][6], // F: Priority
        data[i][7], // G: Status
        data[i][9]  // H: Call Notes (from Col J)
      ]);
      jiraKeys.push(jiraKey);

      rowsToClear.push(i + 4); // Dashboard row number
    }
  }

  if (transferRows.length === 0) {
    SpreadsheetApp.getUi().alert('No rows with Call Notes (Column J) content found.');
    return;
  }

  // Generate today's date label (e.g., "April 2 2026")
  var now = new Date();
  var months = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
  var dateLabel = months[now.getMonth()] + ' ' + now.getDate() + ' ' + now.getFullYear();

  // Check if today's date group already exists at the top of WA1 (row 2)
  var existingDateLabel = wa1.getRange(2, 1).getValue().toString().trim();
  var isSameDay = (existingDateLabel === dateLabel);

  var groupDataStartRow;
  var groupDataEndRow;
  var appendedCount = 0;
  var newCount = 0;

  if (isSameDay) {
    // --- SAME-DAY: merge into existing date group ---

    // Batch-read Column A from row 3 onward to find the group boundary
    var wa1LastRow = wa1.getLastRow();
    var scanRows = Math.min(wa1LastRow - 2, 200); // scan up to 200 rows — more than enough
    var colAValues = wa1.getRange(3, 1, scanRows, 1).getValues();

    var existingGroupStart = 3;
    var existingGroupEnd = existingGroupStart;
    for (var r = 0; r < colAValues.length; r++) {
      if (colAValues[r][0].toString().trim() === '') break;
      existingGroupEnd = existingGroupStart + r;
    }

    // Batch-read the existing group data (all 8 columns)
    var groupSize = existingGroupEnd - existingGroupStart + 1;
    var existingData = wa1.getRange(existingGroupStart, 1, groupSize, 8).getValues();

    // Build map of existing keys: key -> { row number, index in existingData }
    var existingKeys = {};
    for (var e = 0; e < existingData.length; e++) {
      var eKey = existingData[e][0].toString().trim();
      if (eKey) {
        existingKeys[eKey] = { row: existingGroupStart + e, idx: e };
      }
    }

    // Separate into updates vs new inserts
    var newRows = [];
    var newKeys = [];
    var notesToAppend = []; // { row, combinedNotes }

    for (var t = 0; t < transferRows.length; t++) {
      var tKey = transferRows[t][0];
      if (existingKeys[tKey]) {
        var info = existingKeys[tKey];
        var existingNotes = existingData[info.idx][7].toString().trim(); // Col H from batch data
        var newNotes = transferRows[t][7].toString().trim();
        var combined = existingNotes ? existingNotes + ' | ' + newNotes : newNotes;
        notesToAppend.push({ row: info.row, notes: combined });
        appendedCount++;
      } else {
        newRows.push(transferRows[t]);
        newKeys.push(tKey);
      }
    }

    // Batch-write appended notes
    for (var a = 0; a < notesToAppend.length; a++) {
      wa1.getRange(notesToAppend[a].row, 8).setValue(notesToAppend[a].notes);
    }

    // Insert new rows at the end of the existing group
    if (newRows.length > 0) {
      var insertAt = existingGroupEnd + 1;
      wa1.insertRowsBefore(insertAt, newRows.length);
      wa1.getRange(insertAt, 1, newRows.length, 8).setValues(newRows);

      // Set CLIP formatting for new rows
      wa1.getRange(insertAt, 1, newRows.length, wa1.getMaxColumns())
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      // Add hyperlinks for new rows
      for (var h = 0; h < newKeys.length; h++) {
        var link = 'https://celerdata-inc.atlassian.net/browse/' + newKeys[h];
        var richText = SpreadsheetApp.newRichTextValue()
          .setText(newKeys[h])
          .setLinkUrl(link)
          .build();
        wa1.getRange(insertAt + h, 1).setRichTextValue(richText);
      }

      // Extend the row group to include the new rows
      wa1.getRange(insertAt, 1, newRows.length, 1).shiftRowGroupDepth(1);

      existingGroupEnd = existingGroupEnd + newRows.length;
      newCount = newRows.length;
    }

    groupDataStartRow = existingGroupStart;
    groupDataEndRow = existingGroupEnd;

  } else {
    // --- NEW DAY: create new date group at the top ---

    var totalNewRows = 1 + transferRows.length + 1; // date label + data + separator
    wa1.insertRowsBefore(2, totalNewRows);

    wa1.getRange(2, 1).setValue(dateLabel);
    wa1.getRange(3, 1, transferRows.length, 8).setValues(transferRows);

    // Set text wrapping to CLIP for all newly inserted rows
    wa1.getRange(2, 1, totalNewRows, wa1.getMaxColumns())
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Add hyperlinks to Column A
    for (var h2 = 0; h2 < jiraKeys.length; h2++) {
      var link2 = 'https://celerdata-inc.atlassian.net/browse/' + jiraKeys[h2];
      var richText2 = SpreadsheetApp.newRichTextValue()
        .setText(jiraKeys[h2])
        .setLinkUrl(link2)
        .build();
      wa1.getRange(3 + h2, 1).setRichTextValue(richText2);
    }

    // Group data rows under the date label
    var gStart = 3;
    var gEnd = 2 + transferRows.length;
    wa1.getRange(gStart, 1, gEnd - gStart + 1, 1).shiftRowGroupDepth(1);

    groupDataStartRow = 3;
    groupDataEndRow = 2 + transferRows.length;
    newCount = transferRows.length;
  }

  // Clear Column J in Dashboard for transferred rows
  for (var j = 0; j < rowsToClear.length; j++) {
    dashboard.getRange(rowsToClear[j], 10).clearContent();
  }

  // Batch-update VLOOKUP formulas in Dashboard using setFormulas (columns I and R)
  var dashLastRow = dashboard.getLastRow();
  var numFormulaRows = dashLastRow - 3; // rows 4 through dashLastRow

  if (numFormulaRows > 0) {
    // Build formula arrays for Column I (col 9) and Column R (col 18)
    var formulasI = [];
    var formulasR = [];

    for (var f = 4; f <= dashLastRow; f++) {
      formulasI.push(['=IFNA(VLOOKUP(A' + f + ',\'WA1\'!$A$' + groupDataStartRow + ':$I$' + groupDataEndRow + ',9,FALSE),"")']);
      formulasR.push(['=IFNA(VLOOKUP(L' + f + ',\'WA1\'!$A$' + groupDataStartRow + ':$I$' + groupDataEndRow + ',9,FALSE),"")']);
    }

    // Two batch writes instead of 2×N individual writes
    dashboard.getRange(4, 9, numFormulaRows, 1).setFormulas(formulasI);
    dashboard.getRange(4, 18, numFormulaRows, 1).setFormulas(formulasR);
  }

  // Build summary message
  var message = 'Done! ';
  if (isSameDay) {
    message += 'Updated existing date group: ' + dateLabel + '\n' +
      '  - Appended notes to existing keys: ' + appendedCount + '\n' +
      '  - New rows added: ' + newCount + '\n';
  } else {
    message += 'Transferred ' + newCount + ' row(s) to WA1.\n' +
      'Date group: ' + dateLabel + '\n';
  }
  message += 'Column J cleared. Columns I & R formulas updated to WA1 range A' + groupDataStartRow + ':I' + groupDataEndRow + '.';

  SpreadsheetApp.getUi().alert(message);
}
