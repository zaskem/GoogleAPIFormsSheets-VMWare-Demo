// Object to represent the input/output Google Form and response Sheet
var VMFromVM = {responseSheet: "CreateVMFromVM", formQuestionTitle:"Base VM", formInputSheet:"VMsAvailable", formInputRange:"A:A", formID:""}


// Generalized function to remove blank form response records in Google Sheets
// Please Note: This function is known to Not Work As Expected in the circumstance of the number of blank rows to be removed being > 1.
// As such, for best results (without making code changes), this should be triggered after every sheet row count change (clear, etc.).
// The argument provided to this function should be one of the variables declared above.
function deleteFormResponses(formActionInformation) {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(formActionInformation.responseSheet);
  // Delete blank responses after the header
  var currentRow = 2
  var lastRow = currentSheet.getLastRow()
  if (lastRow > currentRow) {
    var dataset = currentSheet.getSheetValues(currentRow, 1, lastRow - currentRow, 1)
    for each (var row in dataset) {
      if("" == row) {
        currentSheet.deleteRow(currentRow);
      }
      currentRow++;
    }
  }
}


// Generalized functions to roll through Google Form questions and update "matching" questions (to those specified in the input sheets) as necessary.
// The argument provided to the parent function should be one of the variables declared above

// Updates Google Form Questions (parent function) 
function updateFormQuestions(formActionInformation) {
  var form = FormApp.openById(formActionInformation.formID);
  var items = form.getItems();
  for (var i = 0; i < items.length; i += 1){
      var item = items[i]
      if (item.getTitle() === formActionInformation.formQuestionTitle){
        updateListChoices(item.asListItem(), formActionInformation.formInputSheet, formActionInformation.formInputRange);
        break;
    }
  }
}

// Updates Google Form Question Options (child function)
function updateListChoices(item, inputSheet, inputRange){
  var data = (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(inputSheet).getRange(inputRange).getValues());
  var choices = [];
  // If you want the column to have a header, make i = 0 to i = 1
  for (var i = 1; i < data.length; i+=1){
    if(data[i][0] != "") 
      choices.push(item.createChoice(data[i][0]));
  }
  if(choices.length != 0)
     item.setChoices(choices);
}