function onOpen() {
    var ui = SpreadsheetApp.getUi();
    
    ui.createMenu('Menu')
    
        .addItem("Create masked sheets", "createMaskedSheets")
        .addItem('Initialize sheets for all employees', 'createEmployeeSheetFromTemplate')
        // .addItem('Create different spreadsheet for each employee', 'createAndShareSheets')
        .addItem('Delete sheets specified in Sheets_To_Delete (col A)', 'deleteSheetsBasedOnNames')
  
        .addSeparator()      
        .addToUi();
  }
  
  
  function createAndShareSheets() {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadSheet.getSheets();
    const sheetUrl = spreadSheet.getUrl();
    const editors = ['divyansh.mishra@mathisys-india.com', 'tushar.bhartiya@mathisys-india.com'];
  
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
  
      if (isValidGmail(sheetName)) {
        try {
          var newSpreadSheet = retryFunction(() => SpreadsheetApp.create(sheetName));
          var newSheet = newSpreadSheet.getSheets()[0];
  
          var importRangeFormula = `=IMPORTRANGE("${sheetUrl}", "${sheetName}!A:J")`;
          newSheet.getRange('A1').setFormula(importRangeFormula);
          
          retryFunction(() => newSpreadSheet.addViewer(sheetName));
          retryFunction(() => newSpreadSheet.addEditors(editors));
  
          Logger.log('New spreadsheet created: ' + newSpreadSheet.getUrl());
        } catch (e) {
          Logger.log('Failed to create or share spreadsheet for ' + sheetName + ': ' + e.message);
        }
      }
    }
  }
  
  function isValidGmail(email) {
    var emailRegex = /^[a-zA-Z0-9._%+-]+@mathisys-india\.com$/;
    return emailRegex.test(email);
  }
  
  
  function retryFunction(func, retries = 3, waitTime = 1000) {
    for (var i = 0; i < retries; i++) {
      try {
        return func();
      } catch (e) {
        if (i === retries - 1) {
          throw e;
        }
        Utilities.sleep(waitTime);
      }
    }
  }
  
  function createEmployeeSheetFromTemplate() {
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = source.getSheetByName('Template');
  
    var spreadSheet_names = source.getSheetByName('Monthly_Requests');
    var csvSheetValues = spreadSheet_names.getDataRange().getValues();
    var emails = Array.from(new Set( csvSheetValues.map( value => value[1] ).filter(email => isValidGmail(email)) ))
  
    emails.forEach( (email) => {
  
      if( source.getSheetByName(email) ) 
        source.deleteSheet(source.getSheetByName(email));
      var newSheet = sheet.copyTo(source).setName(email);
      newSheet.getRange('L2').setValue(email);
    })
    createAndShareSheets();
  }
  
  function deleteSheetsBasedOnNames() {
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var sheet2 = source.getSheetByName('Sheets_To_Delete');
  
    var sheetNamesRange = sheet2.getRange('A2:A' + sheet2.getLastRow());
    var sheetNames = sheetNamesRange.getValues();
  
    for (var i = 0; i < sheetNames.length; i++) {
      var sheetName = sheetNames[i][0];
      if (sheetName) {
        var sheetToDelete = source.getSheetByName(sheetName);
        if (sheetToDelete) {
          source.deleteSheet(sheetToDelete);
        }
  
        const existingFiles = DriveApp.getFilesByName(sheetName);
        while (existingFiles.hasNext()) {
          const file = existingFiles.next();
          file.setTrashed(true);
        }
      }
    }
  }
  
  function createMaskedSheets(){
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var editorsList = listOfMails("Please enter E_mail ids for edit access: ", "Some editors Gmail id's were incorrect");
    var viewersList = listOfMails("Please enter E_mail ids for view access: ", "Some viewers Gmail id's were incorrect");
  
    var employee = ui.prompt("Enter the employee Id whoose sheet is required: ");
    const employeeId = employee.getResponseText().trim();
  
    var inputSheetName = ui.prompt("Enter the name by which you want to recognise the sheet: ");
    inputSheetName = inputSheetName.getResponseText();
  
    var sourceSheet = spreadSheet.getSheetByName(employeeId);
    var data = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), 10).getValues();
  
    var newSpreadSheet = retryFunction(() => SpreadsheetApp.create(inputSheetName));
    var destinationSheet = newSpreadSheet.getSheets()[0];
    destinationSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    retryFunction(() => newSpreadSheet.addViewers(viewersList));
    retryFunction(() => newSpreadSheet.addEditors(editorsList));
  
    SpreadsheetApp.getUi().alert('Data copied successfully from ' + employeeId + ' to ' + inputSheetName);
  }
  
  
  function listOfMails(promptMessage, errorMessage){
    var ui = SpreadsheetApp.getUi();
    var input = ui.prompt(promptMessage);
    var inputText = input.getResponseText();
    var inputList = parseEmailIds(inputText);
    var isInvalidGmail = false;
    inputList.forEach(function(email) {
      if (!(isValidGmail(email))){
        ui.alert(errorMessage);
        isInvalidGmail = true;
        return;
      }
    });
    if (isInvalidGmail){
      return;
    }
    return inputList
  }
  
  
  
  function parseEmailIds(emailString) {
    var emailArray = emailString.split(',').map(function(email){
      return email.trim();
    })
    return emailArray;
  }
    