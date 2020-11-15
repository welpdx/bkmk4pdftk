// Custom Menu to run Scripts
function onOpen() {

    var ui = SpreadsheetApp.getUi();

    ui.createMenu('Menu')
        .addItem('DataToText', 'dataToText')
        .addItem('ExportToTextFile', 'export2Text')
        .addToUi();
}

// Takes data and writes text in cells on the 1st column of new 'bookmark_output' as well as col 5 of 'bookmark_input' for easy comparison (to check for errors).
// Uncomment line_ if you want to automatically go to exporting.

function dataToText() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('bookmark_input'); // I know it's bad to hardcode but hey, it works
    var values = sheet.getDataRange().getValues();

    var startRow = 2; // Start at second row because the first row contains the data labels
    var numRows = sheet.getLastRow(); // Process all rows in 'bookmark_input'

    // Column for text to be written in 'bookmark_input'
    var col_bookmark_input = 5
    // Column for text to be written in 'bookmark_output'
    var col_bookmark_output = 1


    // Range of data (table of contents)
    var dataRange = sheet.getRange(startRow, 1, numRows - 1, 3);

    // Fetch values for each row in the Range.
    var data = dataRange.getValues();

    // Declare buffer variable for writing cells. (dest short for destination)
    var destValues = [];
    destValues[0] = [];

    // Counter for For loop
    var j = 2

    // Check if 'bookmark_output' exists
    if (ss.getSheetByName('bookmark_output') == null) {

        var bkmkSheet = ss.insertSheet('bookmark_output', 0);

    }
    // if sheet does exist
    else {
        var bkmkSheet = ss.getSheetByName('bookmark_output')

        // Prompt to clear? Yes to clear, no to end.
        var response = Browser.msgBox('UH OH', 'The tab ""bookmark_output"" is already being used. Shall I clear the contents?', Browser.Buttons.YES_NO);

        if (response == "yes") {
            bkmkSheet.clearContents();

        } else {
            Browser.msgBox('Ok. Terminating Script...');
            return;
        }
    }

    // For loop: per row, get data and add it in text, write text into column 5 cell of same row

    for (i in data) {

        var row = data[i]; // get a row of data
        var level = row[0]; // level number
        var page = row[1]; // page number
        var title = row[2]; // title

        // text to be written in to cells
        var text = "BookmarkBegin \r\nBookmarkTitle: " + title + " \r\n" + "BookmarkLevel: " + level + " \r\n" + "BookmarkPageNumber: " + page;

        // Push text into array for setValues()
        destValues[0].push(text)

        // Write in 'bookmark_output' tab
        bkmkSheet.getRange(j - 1, col_bookmark_output, 1, 1).setNumberFormat("@").setValues([destValues]);
        // Write in bookmark_input for easy comparison
        sheet.getRange(j, col_bookmark_input, 1, 1).setNumberFormat("@").setValues([destValues]);



        // Ending init
        j = j + 1
        destValues = []
        destValues[0] = []
    }

    // Automatically resize Column 1 in 'bookmark_output' tab
    bkmkSheet.autoResizeColumn(col_bookmark_output);
    // Automatically resize Column 5 in 'bookmark_input' tab
    sheet.autoResizeColumn(col_bookmark_input);

    // Uncomment this if you want to automatically go to export
    // export2Text()
}

// Export bookmark_output to textfile. Save to Drive and Download link
function export2Text() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("bookmark_output");
    var values = sheet.getDataRange().getValues();

    // get and join text from all cells
    var text = values.map(function(a) {
        return a.join('\t');
    }).join('\r\n');


    // Prompt for file name. OK for custom name, CANCEL for default name. 
    var ui = SpreadsheetApp.getUi();
    var input = ui.prompt("Enter name of file. \nOK sets file to custom input name.\nCancel sets file to default: ''Bookmarks.txt''.", ui.ButtonSet.OK_CANCEL);

    if (input.getSelectedButton() == ui.Button.OK) {
        // do something if OK
        var fileName = input.getResponseText();
        var file = DriveApp.createFile(fileName.toString().slice(0, 15), text, MimeType.PLAIN_TEXT);

    } else if (input.getSelectedButton() == ui.Button.CANCEL) {

        // Create a file with name: Bookmarks.txt.
        var file = DriveApp.createFile('Bookmarks'.toString().slice(0, 15), text, MimeType.PLAIN_TEXT);

    } else if (input.getSelectedButton() == ui.Button.CLOSE) {
        return;
    }



    // get fileId for download link
    var fileid = file.getId()
    // download link
    var URL = 'https://docs.google.com/file/d/' + fileid + '/view?usp=sharing';

    // Display a modal dialog box with download link.
    // Create dialog box content
    var string = '<br></b><b>Or copy and paste the text below<br><br></b>';
    string += '<div contenteditable="true">'
    for (var n in values) {
        string += values[n].join(' ') + '<br>';
    }
    string += '</div>'

    // Change \n to <br> for display in html
    string = string.replace(/\n/g, '<br>');

    // idk, dont ask
    var htmlOutput = HtmlService
        .createHtmlOutput('<a href="' + URL + '" target="_blank">Link to File</a>' + string)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(800)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download file');


}
