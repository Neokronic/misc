/**
 * No menu on open — keeps UI clean during screen sharing.
 * Use Extensions > Macros > copyNotesToWA1 (Ctrl+Alt+Shift+1) to trigger.
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
    .addItem('Copy Call Notes to WA1', 'copyNotesToWA1')
    .addToUi();
}

/**
 * Copies rows from Dashboard where Column J has content into WA1.
 * Transfers: Key (with hyperlink), Issue Type, Summary, PM, Assignee, Priority, Status, Call Notes.
 * Skips Column I (Updates) to avoid confusion.
 * Inserts new rows at the top (row 2) with an auto-generated date label.
 * Sets cell formatting to overflow (CLIP). Groups data rows under the date label.
 * Adds empty separator row between date groups.
 * Updates Dashboard Column I and R VLOOKUP to reference the new WA1 range.
 * Then clears Column J in Dashboard.
 */
function copyNotesToWA1() {
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

  // Read Dashboard data (row 4 onward, columns A through J)
  // Columns: A=Key(hyperlink), B=Key(plain), C=IssueType, D=Summary, E=PM, F=Assignee, G=Priority, H=Status, I=Updates(skip), J=CallNotes
  var dataRange = dashboard.getRange(4, 1, lastRow - 3, 10); // A4:J{lastRow}
  var data = dataRange.getValues();

  var rowsToClear = [];
  var transferRows = [];
  var jiraKeys = []; // store keys for hyperlink creation

  for (var i = 0; i < data.length; i++) {
    var colJ = data[i][9]; // Column J (index 9) = Call Notes
    if (colJ && colJ.toString().trim() !== '') {
      var jiraKey = data[i][1].toString().trim(); // Column B (plain key)
      if (!jiraKey) continue;

      // Transfer: Key, Issue Type, Summary, PM, Assignee, Priority, Status, Call Notes
      // Skip Column I (Updates) to avoid confusion
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

  // Generate date label (e.g., "April 2 2026")
  var now = new Date();
  var months = ['January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
  var dateLabel = months[now.getMonth()] + ' ' + now.getDate() + ' ' + now.getFullYear();

  // Insert rows at the top: 1 date row + N data rows + 1 empty separator row, starting at row 2
  var totalNewRows = 1 + transferRows.length + 1;
  wa1.insertRowsBefore(2, totalNewRows);

  // Write date label in row 2, Column A
  wa1.getRange(2, 1).setValue(dateLabel);

  // Write data rows starting at row 3 (8 columns: Key, IssueType, Summary, PM, Assignee, Priority, Status, CallNotes)
  wa1.getRange(3, 1, transferRows.length, 8).setValues(transferRows);

  // Row after last data row is the empty separator (left blank intentionally)

  // Set text wrapping to CLIP (overflow) for all newly inserted rows (date row + data rows)
  var newRowsRange = wa1.getRange(2, 1, totalNewRows, wa1.getMaxColumns());
  newRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Add hyperlinks to Column A (Key) for each transferred row
  for (var h = 0; h < jiraKeys.length; h++) {
    var link = 'https://celerdata-inc.atlassian.net/browse/' + jiraKeys[h];
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(jiraKeys[h])
      .setLinkUrl(link)
      .build();
    wa1.getRange(3 + h, 1).setRichTextValue(richText);
  }

  // Group only the data rows so the collapse toggle sits on the date row (row 2)
  // Data rows = row 3 through row (2 + transferRows.length)
  var groupStartRow = 3;
  var groupEndRow = 2 + transferRows.length;
  wa1.getRange(groupStartRow, 1, groupEndRow - groupStartRow + 1, 1).shiftRowGroupDepth(1);

  // Clear Column J (Call Notes) in Dashboard for transferred rows
  for (var j = 0; j < rowsToClear.length; j++) {
    dashboard.getRange(rowsToClear[j], 10).clearContent(); // Col J
  }

  // Update VLOOKUP formulas in Dashboard to point to the new WA1 range
  // WA1 cols: Key(A), IssueType(B), Summary(C), PM(D), Assignee(E), Priority(F), Status(G), CallNotes(H), Update(I)
  // VLOOKUP fetches column 9 = Update (I) — the refined update written manually in WA1
  var wa1StartRow = 3;
  var wa1EndRow = 3 + transferRows.length - 1;
  var dashLastRow = dashboard.getLastRow();

  for (var f = 4; f <= dashLastRow; f++) {
    // Column I (col 9) — VLOOKUP by Key in Column A
    var formulaI = '=IFNA(VLOOKUP(A' + f + ',\'WA1\'!$A$' + wa1StartRow + ':$I$' + wa1EndRow + ',9,FALSE),"")';
    dashboard.getRange(f, 9).setFormula(formulaI);

    // Column R (col 18) — same VLOOKUP but by Key in Column L (for the Done/Closed section)
    var formulaR = '=IFNA(VLOOKUP(L' + f + ',\'WA1\'!$A$' + wa1StartRow + ':$I$' + wa1EndRow + ',9,FALSE),"")';
    dashboard.getRange(f, 18).setFormula(formulaR);
  }

  SpreadsheetApp.getUi().alert(
    'Done! Transferred ' + transferRows.length + ' row(s) to WA1.\n' +
    'Date group: ' + dateLabel + '\n' +
    'Column J cleared. Columns I & R formulas updated to WA1 range A' + wa1StartRow + ':I' + wa1EndRow + '.'
  );
}
