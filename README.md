# How to Create a Google Sheets App Script 📊

![Google Sheets Logo](https://www.gstatic.com/images/branding/product/1x/sheets_48dp.png)

## **Step 1:** Open Google Sheets 📝

Begin by opening Google Sheets on your computer.

## **Step 2:** Create a Drawing / Icon 🎨
1. Click on any random cell in your spreadsheet.
2. Navigate to the **Insert** menu and select **Drawing**.
3. Draw a unique and creative icon or image.
4. Once done, click **Save and Close** to add it to your sheet.
<div style="float: right; padding-left: 20px;">
  <img src="https://github.com/ranvirsingh603/Google-Apps-Script/blob/main/Screenshot%202023-10-10%20104149.png" alt="Google Sheets Logo" height="150">
</div>   

## **Step 3:** Assign a Script to the Icon 📜
1. Click on the Drawing icon you just created.
2. In the top-right corner, click on the three dots (`...`).
3. Choose **Assign Script** from the dropdown menu.
4. Name your script; let's call it "onOpen()."
<div style="float: right; padding-left: 20px;">
  <img src="https://github.com/ranvirsingh603/Google-Apps-Script/blob/main/Screenshot%202023-10-10%20105402.png" alt="Google Sheets Logo" height="150">
</div> 
That's it! You've successfully assigned the "Sarpt" script, which will run when you open the sheet. ✨

Now, you can add functionality to your Google Sheets using your custom script! 🚀

# Google Apps Script - Custom Menu Creation
First of all click on **Extensions** menu open **Apps Script**.
The following code demonstrates how to create a custom menu in Google Sheets using Google Apps Script. This menu, named "Sum by Matharu," contains an item labeled "Sum," which is linked to the `sum` function. The code is designed to automatically execute when the Google Sheets document is opened.

## Code for making **custom menu**

```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ui.createMenu('Sum by Matharu')
      .addItem('Sum', 'sum')
      .addToUi();
}

```

### Explanation:

```
function onOpen() {

This line defines a JavaScript function named onOpen. It holds a special role in Google Apps Script as it is automatically triggered when you open your Google Sheets document.
```

```
var ui = SpreadsheetApp.getUi();
Here, we create a variable called ui and associate it with the user interface of Google Sheets. The SpreadsheetApp.getUi() function grants access to various user interface elements, such as menus and dialogs, within Google Sheets.
```

```
var ss = SpreadsheetApp.getActive();
The next line establishes a variable, ss, which represents the currently active spreadsheet. With SpreadsheetApp.getActive(), we gain the ability to manipulate the active spreadsheet programmatically.
```

```

ui.createMenu('Sum by Matharu')
Now, we embark on the creation of our custom menu. The code above generates a menu bearing the name "Sum by Matharu." We employ the ui.createMenu() method to fashion this novel menu.
```

```

.addItem('Sum', 'sum')
Within our custom menu, we add an item labeled "Sum." This item is linked to a function called sum. Clicking on this menu item will trigger the execution of the sum function.
```

```
.addToUi();
To complete our customization, this line adds the custom menu, complete with the "Sum" option, to the Google Sheets user interface. Now, when you open your Google Sheets document, you'll discover a fresh menu titled "Sum by Matharu" that offers the option to "Sum."
```

This script serves as a prime example of a common technique in Google Apps Script. It demonstrates how you can enhance Google Sheets by crafting custom menus, thereby enriching your spreadsheet's functionality and user experience.


This code represents a Google Apps Script function named `sum()`. Its purpose is to prompt the user for two input values, perform addition on them, and set the result as a formula in the current cell of a Google Sheets document.
##Code for making sum function

```javascript
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
```


## Function Details

- **Function Name:** `sum()`
- **Function Behavior:**
  - Retrieves the active spreadsheet and user interface of Google Sheets.
  - Prompts the user to enter a value for the first cell and stores the input in the `first` variable. The user's response is captured if they click the "OK" button; otherwise, `first` remains unchanged.
  - Prompts the user to enter a value for the second cell and stores the input in the `second` variable in a similar manner.
  - Sets a formula in the currently selected cell in the spreadsheet. The formula calculates the sum of the values entered by the user in the first and second prompts.

 ## Select the entire colummn 

- **`ss`:** This likely represents the active spreadsheet, which was defined earlier in the script. It specifies the spreadsheet on which the following actions will be performed.

- **`getSelection()`:** This method is called on the `ss` object. It retrieves the currently selected range or cell within the spreadsheet. The selected range is the starting point for the subsequent actions.

- **`getNextDataRange(SpreadsheetApp.Direction.DOWN)`:** This method is called on the selected range. It searches for the next contiguous data range in the specified direction, which is downward (`SpreadsheetApp.Direction.DOWN`). In other words, it looks for the next block of data (cells with content) below the currently selected range.

- **`.activate()`:** After identifying the next data range, the `activate()` method is called. This activates, or selects, the identified data range. This means that the identified data range becomes the currently selected range in the spreadsheet.

- **`currentCell`:** This variable seems to represent a specific cell or range. Its purpose is not clear from the provided code snippet, but it is likely defined elsewhere in the script.

- **`activateAsCurrentCell()`:** This method is called on the `currentCell` object. It activates, or selects, the `currentCell` as the currently selected cell in the spreadsheet.


## copy the formula to entire range of column

- **`copyTo(destination, copyPasteType, transposed)`:** This method is called on the cell obtained using `getCurrentCell()`. It facilitates the copy operation and pastes the content into a designated destination within the spreadsheet.

  - **`destination`:** In this specific instance, it's `ss.getActiveRange()`. The `getActiveRange()` method retrieves the presently selected or active range (a group of cells) within the spreadsheet. So, `ss.getActiveRange()` signifies where the copied content will be pasted.

  - **`copyPasteType`:** This parameter specifies the type of copy-paste operation to be executed. In the code, it's set to `SpreadsheetApp.CopyPasteType.PASTE_NORMAL`, indicating a standard paste operation. Such an operation retains the formatting and values of the copied content.

  - **`transposed`:** This boolean value (true or false) determines whether the copied data should be transposed during the paste operation. Transposition involves switching rows and columns. In this code, it's set to `false`, meaning there is no transposition involved.

In summary, this code line copies content from the currently selected cell in the active spreadsheet, then pastes it into the active range within the same spreadsheet. The operation is a normal paste, preserving formatting and values, and no transposition is applied.

This script can be used to quickly perform addition in a Google Sheets document by prompting the user for the values to add and placing the result as a formula in the active cell.

Feel free to use and customize this script as needed for your Google Sheets automation tasks.


# CONGRATULATIONS! You have made your first google apps script
**Ranvir Singh Matharu**








