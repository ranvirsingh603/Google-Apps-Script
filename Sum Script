function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ui.createMenu('Sum by Matharu')
      .addItem('Sum', 'sum')
      .addToUi();
}

function sum() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();

  var first = ui.prompt('First cell', ui.ButtonSet.OK_CANCEL);
 
  if (first.getSelectedButton() == ui.Button.OK) {
   var first = first.getResponseText();
 } 
  
 var second = ui.prompt('second cell', ui.ButtonSet.OK_CANCEL);
  if (second.getSelectedButton() == ui.Button.OK) {
  var second = second.getResponseText()
  }

    ss.getCurrentCell().setFormula('= '+first+' + '+second+'');
    var currentCell = ss.getCurrentCell();
    ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();
     ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

}
