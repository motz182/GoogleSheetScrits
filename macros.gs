/** @OnlyCurrentDoc */
function addc() {
  var spreadsheet = SpreadsheetApp.getActive();  
  var qnt=spreadsheet.getRange('D3').activate().getValue();  
  var ind=spreadsheet.getRange('A10').activate().getValue();
    while(qnt-1 !=0){
      var ind=spreadsheet.getRange('A10').activate().getValue();
      ind++;
      spreadsheet.getRange('a10').activate();
      spreadsheet.getRange('10:10').activate();
      spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
      spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
      spreadsheet.getCurrentCell().setValue(ind);
      spreadsheet.getRange('B10:C10').activate();
      spreadsheet.getRange('D6:E6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      spreadsheet.getRange('B3').activate();  
      qnt--;
   };
  
  if(qnt<2){
    var ind=spreadsheet.getRange('A10').activate().getValue();
    ind++;
    spreadsheet.getRange('a10').activate();
    spreadsheet.getRange('10:10').activate();
    spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
    spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    spreadsheet.getCurrentCell().setValue(ind);
    spreadsheet.getRange('B10:C10').activate();
    spreadsheet.getRange('D6:E6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('B3').activate();  
  };
};