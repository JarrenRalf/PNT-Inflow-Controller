/**
 * This function is run when an html web app is launched. In our case, when the modal dialog box is produced at 
 * the point a user has downloaded inFlow Barcodes, Product Details, Purchase Order, Sales Order or Stock Levels inorder to produce the csv file.
 * 
 * @param {Event} e : The event object 
 * @return Returns the Html Output for the web app.
 */
function doGet(e)
{
  if (e.parameter['inFlowImport'] !== undefined) // The request parameter
  {
    const inFlowImportType = e.parameter['inFlowImport'];

    if (inFlowImportType === 'Barcodes')
      return downloadInflowBarcodes()
    if (inFlowImportType === 'NewItems')
      return downloadInflowNewItems()
    else if (inFlowImportType === 'Pictures')
      return downloadInflowPictures()
    else if (inFlowImportType === 'ProductDetails')
      return downloadInflowProductDetails()
    else if (inFlowImportType === 'PurchaseOrder')
      return downloadInflowPurchaseOrder()
    else if (inFlowImportType === 'SalesOrder')
      return downloadInflowSalesOrder()
    else if (inFlowImportType === 'StockLevels')
      return downloadInflowStockLevels()
      else if (inFlowImportType === 'StockLevels_fromCountsSheet')
      return downloadInflowStockLevels_fromCountsSheet()
  }
}

/**
 * This function handles all of the on change events of the spreadsheet, specifically looking for a new sheet that is being added to the spreadsheet,
 * which is assumed to be an inFlow Purchase Order import.
 * 
 * @param {Event Object} e : An instance of an event object that occurs when the spreadsheet is changed
 * @author Jarren Ralf
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function handles all of the on edit events of the spreadsheet, specifically looking for rows that need to be moved to different sheets,
 * barcodes that are scanned on the Item Scan sheet, searches that are made, and formatting issues that need to be fixed.
 * 
 * @param {Event Object} e : An instance of an event object that occurs when the spreadsheet is editted
 * @author Jarren Ralf
 */
function installed_onEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === "Item Search") // Check if the user is searching for an item or trying to marry, unmarry or add a new item to the upc database
      search(e, spreadsheet, sheet);
    else if (sheetName === "Counts") // Check if the user typed in the quantity in the wrong column
      warning(e, sheet, sheetName);
    else if (sheetName === "Scan") // Check if a barcode has been scanned
      manualScan(e, spreadsheet, sheet)
  } 
  catch (err) 
  {
    var error = err['stack'];
    Logger.log(error)
    Browser.msgBox(error)
    throw new Error(error);
  }
}

/**
 * This function handles all of the on edit events of the spreadsheet, 
 * 
 * @param {Event Object} e : An instance of an event object that occurs when the spreadsheet is editted
 * @author Jarren Ralf
 */
function installed_onOpen()
{
  //openDragAndDrop();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PNT Controls')
    .addSubMenu(ui.createMenu('Download')
      .addItem('Barcodes', 'downloadButton_Barcodes')
      .addItem('New Items', 'downloadInflowNewItems')
      .addItem('Pictures', 'downloadButton_Pictures')
      .addItem('Product Details', 'downloadButton_ProductDetails')
      .addItem('Purchase Orders', 'downloadButton_PurchaseOrder')
      .addItem('Sales Orders', 'downloadButton_SalesOrder')
      .addItem('Stock Levels', 'downloadButton_StockLevels'))
    .addSubMenu(ui.createMenu('Import')
      .addItem('Anything', 'openDragAndDrop')
      .addItem('Stock Levels (From Drive)', 'updateStockLevels'))
    // .addSubMenu(ui.createMenu('Watch Video')
    //   .addItem('How to Export from Adagio', 'openDragAndDrop'))
    .addItem('Add New Items (Selected)', 'addToNewItems')
    .addToUi();
}

/**
 * This function moves the selected items from the item search sheet to the Counts page.
 * 
 * @author Jarren Ralf
 */
function addToCounts()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('Counts'))
}

/**
 * This function moves the selected items from the item search sheet to the new items page.
 * 
 * @author Jarren Ralf
 */
function addToNewItems()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('New Items'))
}

/**
 * This function moves the selected items from the item search sheet to the new items page.
 * 
 * @author Jarren Ralf
 */
function addToNewItems_ButtonPressed()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = SpreadsheetApp.getActiveSheet();
  const checkBoxRange = sheet.getRange(3, 1);

  if (checkBoxRange.isChecked())
    copySelectedValues(spreadsheet.getSheetByName('New Items'))
  else
  {
    const rng  = sheet.getRange(1, 1, 3, 5);
    const vals = rng.getValues()
    vals[0][3] = 'Add NEW item Mode: ON'
    rng.setBackgrounds([  ['#3c78d8', 'white',   '#3c78d8', '#3c78d8', '#3c78d8'], 
                          ['#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8'], 
                          ['#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8']]).setValues(vals)
    checkBoxRange.check()

    const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                    // Clear the formatting of the range of the search box
          .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
          .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
          .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
          .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
          .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

    const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

    if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
    {
      spreadsheet.toast('Searching...')
      const startTime = new Date().getTime();
      const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
      const functionRunTimeRange = sheet.getRange(2, 1);   // The range that will display the runtimes for the search and formatting
      const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5); // The entire range of the Item Search page
      const numSearches = searches.length; // The number searches
      const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
      const uom = data[0].indexOf('Price Unit')
      const fullDescription = data[0].indexOf('Item List')
      const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
      const inflowItems = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 1);
      var output = [], numSearchWords, isInflow;

      if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
      {
        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  isInflow = inflowItems.find(item => item[0] === data[i][fullDescription])
                  output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                  break loop;
                }
              }
              else
                break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
            }
          }
        }
      }
      else // The word 'not' was found in the search string
      {
        const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  for (var l = 0; l < dontIncludeTheseWords.length; l++)
                  {
                    if (!data[i][fullDescription].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                    {
                      if (l === dontIncludeTheseWords.length - 1)
                      {
                        isInflow = inflowItems.find(item => item[0] === data[i][fullDescription])
                        output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                        break loop;
                      }
                    }
                    else
                      break;
                  }
                }
              }
              else
                break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
            }
          }
        }
      }

      const numItems = output.length;

      if (numItems === 0) // No items were found
      {
        sheet.getRange('B1').activate(); // Move the user back to the seachbox
        itemSearchFullRange.clearContent(); // Clear content
        const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#660000').build();
        const message = SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 17, textStyle).build();
        searchResultsDisplayRange.setRichTextValue(message);
      }
      else
      {
        sheet.getRange('B4').activate(); // Move the user to the top of the search items
        itemSearchFullRange.clearContent(); // Clear content and reset the text format
        sheet.getRange(4, 1, numItems, 5).setValues(output);
        (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
      }

      functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " s");
      spreadsheet.toast('Searching Complete.')
    }
  }
}

/**
 * This function moves the selected items from the item search sheet to the purchase order page.
 * 
 * @author Jarren Ralf
 */
function addToPurchaseOrder()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('Purchase Order'))
}

/**
 * This function moves the selected items from the item search sheet to the sales order page.
 * 
 * @author Jarren Ralf
 */
function addToSalesOrder()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('Sales Order'))
}

/**
 * This function clears the items on either the Sales Order, Stock Levels, or Purchase Order sheet.
 * 
 * @author Jarren Ralf
 */
function clear()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const numRows = sheet.getLastRow() - 2

  if (numRows > 0)
  {
    const sheetName = sheet.getSheetName()
    const numCols = sheet.getLastColumn();
    const colours = (sheetName === 'Sales Order' ) ? new Array(numRows).fill([...new Array(numCols - 1).fill('white'), '#d9d9d9']): 
                    (sheetName === 'Stock Levels') ? new Array(numRows).fill([...new Array(numCols - 2).fill('white'), '#d9d9d9', '#d9d9d9']) : 
                                                     new Array(numRows).fill(new Array(numCols).fill('white'));
    sheet.getRange(3, 1, numRows, numCols).setBackgrounds(colours).clearContent()
  }
}

/**
 * This function moves the selected values from the current sheet to the destination sheet.
 * 
 * @param  {Sheet}    sheet    : The sheet that the selected items are being moved to.
 * @param {Boolean} isTransfer : Whether or not the items are transfering location or not.
 * @author Jarren Ralf
 */
function copySelectedValues(sheet, isTransfer)
{
  const  activeSheet = SpreadsheetApp.getActiveSheet();
  const activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const firstRows = [], lastRows = [];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows.push(activeRanges[r].getRow());
     lastRows.push(activeRanges[r].getLastRow());
  }
  
  const     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  const lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
  const finalDataRow = activeSheet.getLastRow() + 1;
  const numHeaders = 3;
  
  if (row > numHeaders && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
  {   
    // Concatenate all of the item values as a 2-D array
    const itemValues = [].concat.apply([], firstRows.map((_, r) => activeSheet.getSheetValues(firstRows[r], 2, lastRows[r] - firstRows[r] + 1, 4))); 

    switch (sheet.getSheetName())
    {
      case 'Sales Order':
        var range = sheet.getRange(2, 1);
        var orderCounter = Number(range.getValue()) + 1;
        var items = itemValues.map(v => ['newSalesOrder_' + orderCounter, 'Richmond PNT', v[0], '', '', '', v[2]])
        var colours = items.map(_ => ['white', 'white', 'white', 'white', 'white', 'white', 'white'])
        range.setValue(orderCounter)
        break;
      case 'Stock Levels':
        if (isTransfer)
        {
          // Duplicate each item such that the transfered from location is zero and the destination location is blank
          var items = itemValues.flatMap(v => [[v[0], v[1], 0, v[3], v[2], 'T'], [v[0], '', v[2], v[3], '', 'T']]) 
          var colours = items.map((_, i) => (i % 2 === 0) ? ['#ea9999', '#ea9999', '#ea9999', '#ea9999', '#d9d9d9', '#d9d9d9'] : ['#ea9999', '#e06666', '#ea9999', '#ea9999', '#d9d9d9', '#d9d9d9'])
        }
        else // Stock Adjustment
        {
          var items = itemValues.map(v => [v[0], v[1], '', v[3], v[2], 'A'])
          var colours = items.map(_ => ['#f9cb9c', '#f9cb9c', '#f6b26b', '#f9cb9c', '#d9d9d9', '#d9d9d9'])
        }
        break;
      case 'Purchase Order':
        var range = sheet.getRange(2, 1);
        var orderCounter = Number(range.getValue()) + 1;
        var items = itemValues.map(v => ['newPurchaseOrder_' + orderCounter, 'PACIFIC NET & TWINE', v[0], '', '', 0, 0, 0])
        var colours = items.map(_ => ['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
        range.setValue(orderCounter)
        break;
      case 'New Items':
        var splitDescription;
        var items = itemValues.map(v => {
          splitDescription = v[0].split(' - ');
          splitDescription.pop();
          splitDescription.pop();
          splitDescription.pop();
          splitDescription.pop();
          return isNotBlank(v[1]) ? [v[0], '', splitDescription.join(' - '), -1] : ['', '', '', ''];
        }).filter(u => isNotBlank(u[3]))
        var colours = items.map(_ => ['white', 'white', 'white', 'white'])
        break;
      case 'Counts':
        var items = [], index = 0;
        itemValues.map(v => {
          index = items.findIndex(descrip => descrip[0] == v[0]);

          if (index !== -1)
          {
            items[index][1] += ', ' + v[1] + ': ' + v[2]
            items[index][3] += v[2]
          }
          else
            (isNotBlank(v[2])) ? items.push([v[0], v[1] + ': ' + v[2], '', Number(v[2])]) : items.push([v[0], '', '', 0])
          
        })
        var colours = items.map(_ => ['white', 'white', 'white'])
        items = items.map(item => {

          if (item[1].split(':').length > 2) // If there are more than 1 location, then report the sum
            item[1] += '; Sum: ' + item.pop()
          else
            item.pop()

          return item 
        })

        if (items.length > 0)
          formatCountsPage(sheet, sheet.getLastRow() + 1, items.length, 7)
        break;
    }

    // Move the item values to the destination sheet
    if (items.length > 0)
      sheet.getRange(sheet.getLastRow() + 1, 1, items.length, items[0].length).setNumberFormat('@').setBackgrounds(colours).setValues(items).activate(); 
    else
      SpreadsheetApp.getUi().alert('Please select an item that is NOT already in inFlow.');
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
}

/**
 * This function creates all of the triggers for the spreadsheet to function properly.
 * 
 * @author Jarren
 */
function createTriggers()
{
  const spreadsheet = SpreadsheetApp.getActive()
  ScriptApp.newTrigger('onChange').forSpreadsheet(spreadsheet).onChange().create()
  ScriptApp.newTrigger('installed_onEdit').forSpreadsheet(spreadsheet).onEdit().create()
  ScriptApp.newTrigger('installed_onOpen').forSpreadsheet(spreadsheet).onOpen().create()
  ScriptApp.newTrigger('updateUPCs').timeBased().atHour(9).everyDays(1).create()
}

/**
 * This function deletes all of the triggers associated with this project.
 * 
 * @author Jarren
 */
function deleteAllTriggers()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger))
}

/**
 * This function launches a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of the export data being downloaded.
 * 
 * @param {String} importType : The type of information that will be imported into inFlow.
 * @author Jarren Ralf
 */
function downloadButton(importType)
{
  var htmlTemplate = HtmlService.createTemplateFromFile('DownloadButton')
  htmlTemplate.inFlowImportType = importType;
  var html = htmlTemplate.evaluate().setWidth(250).setHeight(75)
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Export');
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Product Details containing barcodes to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_Barcodes()
{
  downloadButton('Barcodes')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Purchase Order to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_NewItems()
{
  downloadButton('NewItems')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Product Details to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_Pictures()
{
  downloadButton('Pictures')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Purchase Order to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_ProductDetails()
{
  downloadButton('ProductDetails')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Purchase Order to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_PurchaseOrder()
{
  downloadButton('PurchaseOrder')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Sales Order to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_SalesOrder()
{
  downloadButton('SalesOrder')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of inFlow Stock Levels for a particular set of items to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_StockLevels()
{
  downloadButton('StockLevels')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of inFlow Stock Levels for a particular set of items to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_StockLevels_fromCountsSheet()
{
  downloadButton('StockLevels_fromCountsSheet')
}

/**
 * This function takes three arguments that will be used to create a csv file that can be downloaded from the Browser.
 * 
 * @param {String} sheetName  : The name of the sheet that the data is coming from
 * @param {String} csvHeaders : The header names of the csv file
 * @param {String} fileName   : The name of the csv file that will be produced
 * @param {Number} excludeCol : The number of columns at the end of the data that provide information to the user, but do not need to be imported into inFlow
 * @param {String[]} varArgs  : The variable number of arguments which is the name of the headers from the Product Details data
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflow(sheetName, csvHeaders, fileName, excludeCol, ...varArgs)
{
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn() - excludeCol)

  if (sheetName === 'Product Details')
  {
    const header = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
    const indecies = varArgs.map(arg => header.indexOf(arg))
    var sku, googleDescription, uniqueSKUs = [], inflowData = [], index = -1;
    indecies.unshift(header.indexOf('Name')) // Add the google description column to the front of the list

    if (varArgs.includes('Barcode'))
    {
      const upcs = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString())
      const numUpcs = upcs.length
      var upcCodes = ''

      data.map(descrip => {
        googleDescription = descrip[0].toString().split(' - ');
        sku = googleDescription.pop().toString().toUpperCase();
        upcCodes = ''

        if (googleDescription.length >= 5)
        {
          index = uniqueSKUs.indexOf(sku);

          if (index === -1)
          {
            for (var i = 1; i < numUpcs; i++)
              if (upcs[i][1].toUpperCase() === sku)
                upcCodes += upcs[i][0] + ','

            uniqueSKUs.push(sku)
            inflowData.push([descrip[indecies[0]], upcCodes])
          }
          else
          {
            uniqueSKUs.splice(index, 1)
            inflowData.splice(index, 1)
          }
        }     
      })

      inflowData = inflowData.filter(barcode => isNotBlank(barcode[1]));
    }
    else if (varArgs.includes('PicturePath'))
    {
      const fromShopifySheet = SpreadsheetApp.openById('1sLhSt5xXPP5y9-9-K8kq4kMfmTuf6a9_l9Ohy0r82gI').getSheetByName('FromShopify')
      const shopifyHeader = fromShopifySheet.getSheetValues(1, 1, 1, fromShopifySheet.getLastColumn())[0];
      const numRows = fromShopifySheet.getLastRow() - 1;
      const skus  = fromShopifySheet.getSheetValues(2, shopifyHeader.indexOf('Variant SKU')   + 1, numRows, 1);
      const imgs1 = fromShopifySheet.getSheetValues(2, shopifyHeader.indexOf('Image Src')     + 1, numRows, 1);
      const imgs2 = fromShopifySheet.getSheetValues(2, shopifyHeader.indexOf('Variant Image') + 1, numRows, 1);

      data.map(descrip => {
        googleDescription = descrip[0].toString().split(' - ');
        sku = googleDescription.pop().toString().toUpperCase();

        if (googleDescription.length >= 5)
        {
          index = uniqueSKUs.indexOf(sku);

          if (index === -1)
          {
            for (var i = 0; i < numRows; i++)
            {
              if (skus[i][0].toString().toUpperCase() == sku)
              {
                if (isNotBlank(imgs1[i][0]))
                {
                  uniqueSKUs.push(sku)
                  inflowData.push([descrip[indecies[0]], imgs1[i][0]])
                }
                else if (isNotBlank(imgs2[i][0]))
                {
                  uniqueSKUs.push(sku)
                  inflowData.push([descrip[indecies[0]], imgs2[i][0]])
                }

                break;
              }
            }
          }
          else
          {
            uniqueSKUs.splice(index, 1)
            inflowData.splice(index, 1)
          }
        }     
      })
    }
    else // Regular Product Details
    {
      data.map(descrip => {
        googleDescription = descrip[0].toString().split(' - ');
        sku = googleDescription.pop().toString().toUpperCase();

        if (googleDescription.length >= 5)
        {
          index = uniqueSKUs.indexOf(sku);

          if (index === -1)
          {
            uniqueSKUs.push(sku)
            inflowData.push([...indecies.map(index => descrip[index])])
          }
          else
          {
            uniqueSKUs.splice(index, 1)
            inflowData.splice(index, 1)
          }
        }     
      })
    }

    data = inflowData;
  }
  else if (sheetName === 'Counts')
  {
    var inflowData = [];

    data.map(item => {
      loc = item[5].split('\n')
      qty = item[6].split('\n')

      if (loc.length === qty.length) // Make sure there is a location for every quantity and vice versa
        for (i = 0; i < loc.length; i++) // Loop through the number of inflow locations
          if (isNotBlank(loc[i]) && isNotBlank(qty)) // Do not add the data to the csv file if either the location or the quantity is blank
            inflowData.push([item[0], loc[i], qty[i]])
    })

    data = inflowData;
  }

  for (var row = 0, csv = csvHeaders; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
      {
        quotationMarks = data[row][col].toString().split('"')
        data[row][col] = (quotationMarks.length !== 1) ? "\"" + quotationMarks.join('""') + "\"" : data[row][col] = "\"" + data[row][col] + "\"";
      }
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile(fileName);
}

/**
 * This function takes the array of data on the Product Details page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowBarcodes()
{
  const sheetName = 'Product Details';
  const csvHeaders = "Name,Barcode\r\n";
  const fileName = 'inFlow_ProductDetails.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude, 'Barcode')
}

/**
 * This function takes the array of data on the New Items page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowNewItems()
{
  const sheetName = 'New Items';
  const csvHeaders = "Name,Category,Description,ReorderPoint\r\n";
  const fileName = 'inFlow_ProductDetails.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude, 'Category', 'Description', 'ReorderPoint')
}

/**
 * This function takes the array of data on the Product Details page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowPictures()
{
  const sheetName = 'Product Details';
  const csvHeaders = "Name,PicturePath\r\n";
  const fileName = 'inFlow_ProductDetails.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude, 'PicturePath')
}

/**
 * This function takes the array of data on the Purchase Order page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowProductDetails()
{
  const sheetName = 'Product Details';
  const csvHeaders = "Name,Category,ReorderPoint,ReorderQuantity,Remarks,NOTES,IsActive\r\n";
  const fileName = 'inFlow_ProductDetails.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude, 'Category', 'ReorderPoint', 'ReorderQuantity', 'Remarks', 'NOTES', 'IsActive')
}

/**
 * This function takes the array of data on the Purchase Order page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowPurchaseOrder()
{
  const sheetName = 'Purchase Order';
  const csvHeaders = "OrderNumber,Vendor,ItemName,ItemQuantity,OrderRemarks,AmountPaid,ItemUnitPrice,ItemSubtotal\r\n";
  const fileName = 'inFlow_PurchaseOrder.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude)
}

/**
 * This function takes the array of data on the Sales Order page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowSalesOrder()
{
  const sheetName = 'Sales Order';
  const csvHeaders = "OrderNumber,Customer,ItemName,ItemQuantity,OrderRemarks,ContactName\r\n";
  const fileName = 'inFlow_SalesOrder.csv';
  const numColsToExclude = 1; // Columns at the end of the data that provide information to the user, but does not need to be imported into inFlow
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude)
}

/**
 * This function takes the array of data on the Stock Levels page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowStockLevels()
{
  const sheetName = 'Stock Levels';
  const csvHeaders = "Item,Location,Quantity,Serial\r\n";
  const fileName = 'inFlow_StockLevels.csv';
  const numColsToExclude = 2; // Columns at the end of the data that provide information to the user, but does not need to be imported into inFlow
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude)
}

/**
 * This function takes the array of data on the Counts page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowStockLevels_fromCountsSheet()
{
  const sheetName = 'Counts';
  const csvHeaders = "Item,Location,Quantity\r\n";
  const fileName = 'inFlow_StockLevels.csv';
  const numColsToExclude = 0 // Columns at the end of the data that provide information to the user, but does not need to be imported into inFlow
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude)
}

/**
 * Apply the proper formatting to the Counts page.
 *
 * @param {Sheet}   sheet  : The Counts sheet that needs a formatting adjustment
 * @param {Number}   row   : The row that needs formating
 * @param {Number} numRows : The number of rows that needs formatting
 * @param {Number} numCols : The number of columns that needs formatting
 * @author Jarren Ralf
 */
function formatCountsPage(sheet, row, numRows, numCols)
{
  var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#', '@', '#', '@', '@']);
  sheet.getRange(row, 1, numRows, numCols).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
  sheet.getRange(row, 3, numRows         ).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                          .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(row, 5, numRows,       2).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID) 
                                          .setBorder(null, null, null, null, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
                                          .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
* This function checks if the given input is a number or not.
*
* @param {Object} num The inputted argument, assumed to be a number.
* @return {Boolean} Returns a boolean reporting whether the input paramater is a number or not
* @author Jarren Ralf
*/
function isNumber(num)
{
  return !(isNaN(Number(num)));
}

/**
 * This function returns true if the presented number is a UPC-A, false otherwise.
 * 
 * @param {Number} upcNumber : The UPC-A number
 * @returns Whether the given value is a UPC-A or not
 * @author Jarren Ralf
 */
function isUPC_A(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i])*3 : Number(upc[i])

  return upc.endsWith(Math.ceil(sum/10)*10 - sum)
}

/**
 * This function returns true if the presented number is a EAN_13, false otherwise.
 * 
 * @param {Number} upcNumber : The EAN_13 number
 * @returns Whether the given value is a EAN_13 or not
 * @author Jarren Ralf
 */
function isEAN_13(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i]) : Number(upc[i])*3

  return upc.endsWith(Math.ceil(sum/10)*10 - sum)
}

/**
 * This function updates the Inventory sheet either by handling the imported inFlow Stock Levels csv or pulling data from the inFlow Stock Levels file in the drive.
 * 
 * @param {String[][]}    values    : The values of the inFlow Stock Levels
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @param {Number}       startTime  : The time that the function began running at
 * @author Jarren Ralf
 */
function importStockLevels(values, spreadsheet, startTime)
{
  if (arguments.length !== 3)
    startTime = new Date().getTime();

  const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
  const productDetailsSheet = spreadsheet.getSheetByName('Product Details');
  const numRows_StockLevels = values.length;
  var inventory = [], itemHasZeroInventory;

  const uniqueNumProducts = productDetailsSheet.getSheetValues(3, 1, productDetailsSheet.getLastRow() - 2, 1).map(item => {
    
    itemHasZeroInventory = true;

    for (var i = 1; i < numRows_StockLevels; i++)
    {
      if (values[i][0] === item[0])
      {
        inventory.push([values[i][0], values[i][1], values[i][4], values[i][3]])
        itemHasZeroInventory = false;
      }
    }

    if (itemHasZeroInventory)
      inventory.push([item[0], '', '', ''])
  })

  const numRows_Inventory = inventory.length;
  const formats = new Array(numRows_Inventory).fill(['@', '@', '#', '@'])

  inventorySheet.getRange(1, 2, 1, 3).clearContent() // Clear number of items and timestamp
    .offset(2, -1, inventorySheet.getMaxRows(), 4).clearContent() // Clear the previous inventory
    .offset(0, 0, numRows_Inventory, 4).setNumberFormats(formats).setValues(inventory) // Set the updated inventory
    .offset(-2, 1, 1, 3).setValues([[
      uniqueNumProducts.length, 
      (new Date().getTime() - startTime)/1000 + ' s', 
      Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM HH:mm')
    ]])
}

/**
 * This function handles the imported Adagio Sales Order and converts it into an inFlow Sales Order.
 * 
 * @param {String[][]}    values    : The values of the Adagio Sales Order
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @author Jarren Ralf
 */
function importAdagioSalesOrder(values, spreadsheet)
{
  const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
  const inflowData = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 3).filter(item => item[0].split(' - ').length > 4)
  values.pop() // Remove the last row which contains summary data
  const customersSheet = spreadsheet.getSheetByName('Customers');
  const customers = customersSheet.getSheetValues(2, 1, customersSheet.getLastRow() - 1, 2);
  const header = values.shift();
  const customerNumber = header.indexOf('Cust #')
  const qty = header.indexOf('Qty Original Ordered')
  const sku = header.indexOf('Item')
  var output = [], item, customer, isOrderNumberUpdated = false;

  for (var i = 0; i < values.length; i++)
  {
    if (sku !== -1) // Found the SKU column
    {
      if (qty !== -1 && values[i][qty] !== 0) // Found the quantity column and the ordered quantity is not zero
      {
        item = inflowData.find(description => 
          description[0].split(' - ').pop() === values[i][sku].substring(0, 4) + values[i][sku].substring(5, 9) + values[i][sku].substring(10))
          
        if (item != null) // Item was found in inFlow
        {
          if (!isOrderNumberUpdated)
          {
            isOrderNumberUpdated = true;
            var salesOrderSheet = spreadsheet.getSheetByName('Sales Order');
            const range = salesOrderSheet.getRange(2, 1);
            var num = parseInt(range.getValue()) + 1
            range.setValue(num)
          }

          if (customerNumber !== -1) // Found the customer column
          {
            customer = customers.find(val => val[1] === values[i][customerNumber])
            output.push(['newCustomerSalesOrder' + num, (customer != null) ? customer[0] : 'PNT Customer', item[0], values[i][qty], item[2]])
          }
          else // Default the customer to PACIFIC NET & TWINE if not found
            output.push(['newCustomerSalesOrder' + num, 'PNT Customer', item[0], values[i][qty], item[2]])
        }
      }
    }
  }

  if (output.length !== 0)
  {
    const row = salesOrderSheet.getLastRow() + 1
    salesOrderSheet.getRange(row, 1, output.length, 5).setValues(output).activate()
  }
  else
    SpreadsheetApp.getUi().alert('The items on this Adagio Sales Order could not be placed on an inFlow Purchase Order because either the items are not found in the inFlow database or the Adagio description(s) are ambiguous.')
}

/**
 * This function handles the imported Adagio Purchase Order and converts it into an inFlow Purchase Order.
 * 
 * @param {String[][]}    values    : The values of the Adagio Purchase Order
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @author Jarren Ralf
 */
function importAdagioPurchaseOrder(values, spreadsheet)
{
  const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
  const inflowData = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 1).filter(item => item[0].split(' - ').length > 4)
  const inFlowSkus = inflowData.map(descrip => descrip[0]);
  values.pop() // Remove the last row which contains summary data
  const vendorsSheet = spreadsheet.getSheetByName('Vendors');
  const vendors = vendorsSheet.getSheetValues(2, 1, vendorsSheet.getLastRow() - 1, 2);
  const header = values.shift();
  const vendorName = header.indexOf('Vendor name')
  const qty = header.indexOf('Backordered')
  const sku = header.indexOf('Item#')
  const poNumber = header.indexOf('Doc #')
  var output = [], item, vendor;

  for (var i = 0; i < values.length; i++)
  {
    if (sku !== -1) // Found the SKU column
    {
      if (qty !== -1 && values[i][qty] !== 0) // Found the quantity column and the ordered quantity is not zero
      {
        item = inFlowSkus.find(description => 
          description.split(' - ').pop() === values[i][sku].substring(0, 4) + values[i][sku].substring(5, 9) + values[i][sku].substring(10))

        if (item != null) // Item was found in inFlow
        {
          if (vendorName !== -1) // Found the vendor column
          {
            vendor = vendors.find(val => val[0] === values[i][vendorName])
            output.push(['theNewPurchaseOrder', 
              (vendor != null) ? vendor[1] : 'PACIFIC NET & TWINE', item, values[i][qty], 
              (poNumber !== -1) ? values[i][poNumber] : '', 0, 0, 0])
          }
          else // Default the vendor to PACIFIC NET & TWINE if not found
            output.push(['theNewPurchaseOrder', 'PACIFIC NET & TWINE', item, values[i][qty], 
              (poNumber !== -1) ? values[i][poNumber] : '', 0, 0, 0])
        }
      }
    }
  }

  if (output.length !== 0)
  {
    const purchaseOrderSheet = spreadsheet.getSheetByName('Purchase Order')
    const lastRow = purchaseOrderSheet.getLastRow()

    purchaseOrderSheet.getRange(3, 1, lastRow, 8).clearContent().offset(0, 0, output.length, 8).setValues(output).activate()
  }
  else
    SpreadsheetApp.getUi().alert('The items on this Adagio Purchase Order could not be placed on an inFlow Purchase Order because either the items are not found in the inFlow database or the Adagio description(s) are ambiguous.')
}

/**
 * This function gets the time that the particular item was last counted and calculates how long it has been since now, then displays that info to the user.
 * 
 * @param {Number} lastScannedTime : The time that an item was last scanned on the Manual Scan page or inputed on the Manual Counts page, represented as a number in milliseconds (Epoche Time).
 * @returns {String} 
 * @author Jarren Ralf
 */
function getCountedSinceString(lastScannedTime)
{
  if (isNotBlank(lastScannedTime))
  {
    const countedSince = (new Date().getTime() - lastScannedTime)/(1000) // This is in seconds

    if (countedSince < 60) // Number of seconds in 1 minute
      return Math.floor(countedSince) + ' seconds ago'
    else if (countedSince < 3600) // Number of seconds in 1 hour
      return (Math.floor(countedSince/60) === 1) ? Math.floor(countedSince/60) +  ' minute ago' : Math.floor(countedSince/60) +  ' minutes ago'
    else if (countedSince < 86400) // Number of seconds in 24 hours
    {
      const numHours = Math.floor(countedSince/3600);
      const numMinutes = Math.floor((countedSince - numHours*3600)/60);

      return (numHours === 1) ? numHours + ' hour ' + ((numMinutes === 0) ? 'ago' : (numMinutes === 1) ? numMinutes +  ' minute ago' : numMinutes +  ' minutes ago') : 
        numHours + ' hours ' + ((numMinutes === 0) ? 'ago' : (numMinutes === 1) ? numMinutes +  ' minute ago' : numMinutes +  ' minutes ago');
    }
    else // Greater than 24 hours
    {
      const numDays = Math.floor(countedSince/86400);
      const numHours = Math.floor((countedSince - numDays*86400)/3600);

      return (numDays === 1) ? numDays + ' day ' + ((numHours === 0) ? 'ago' : (numHours === 1) ? numHours + ' hour ago' : numHours + ' hours ago') : 
        numDays + ' days ' + ((numHours === 0) ? 'ago' : (numHours === 1) ? numHours + ' hour ago' : numHours + ' hours ago');
    }
  }
  else
    return '1 second ago'
}

/**
 * This function watches two cells and if the left one is edited then it searches the UPC Database for the upc value (the barcode that was scanned).
 * It then checks if the item is on the Counts page and stores the relevant data in the left cell. If the right cell is edited, then the function
 * uses the data in the left cell and moves the item over to the Counts page with the updated quantity.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function manualScan(e, spreadsheet, sheet)
{
  const manualCountsPage = spreadsheet.getSheetByName("Counts");
  const barcodeInputRange = e.range;

  if (barcodeInputRange.columnEnd === 1) // Barcode is scanned
  {
    var upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
      .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
      .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
      .getValue();

    if (isNotBlank(upcCode)) // The user may have hit the delete key
    {
      if (/^\d+$/.test(upcCode))
      {
        if (isUPC_A(upcCode) || isEAN_13(upcCode))
        {
          const lastRow = manualCountsPage.getLastRow();
          const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
          const upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
          var l = 0; // Lower-bound
          var u = upcDatabase.length - 1; // Upper-bound
          var m = Math.ceil((u + l)/2) // Midpoint
          upcCode = parseInt(upcCode)

          if (lastRow <= 2) // There are no items on the Counts page
          {
            while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
            {
              if (upcCode < parseInt(upcDatabase[m][0]))
                u = m;   
              else if (upcCode > parseInt(upcDatabase[m][0]))
                l = m;
              else // UPC code was found!
              {
                const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]
                barcodeInputRange.setValue(description + '\nwill be added to the Counts page at line :\n' + 3);
                break; // Item was found, therefore stop searching
              }
                
              m = Math.ceil((u + l)/2) // Midpoint
            }
          }
          else // There are existing items on the Counts page
          {
            const row = lastRow + 1;
            const manualCountsValues = manualCountsPage.getSheetValues(3, 1, row - 3, 5);

            while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
            {
              if (upcCode < parseInt(upcDatabase[m][0]))
                u = m;   
              else if (upcCode > parseInt(upcDatabase[m][0]))
                l = m;
              else // UPC code was found!
              {
                const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]

                for (var j = 0; j < manualCountsValues.length; j++) // Loop through the Counts page
                {
                  if (manualCountsValues[j][0] === description) // The description matches
                  {
                    const countedSinceString = (isNotBlank(manualCountsValues[j][4])) ? '\nLast Counted :\n' + getCountedSinceString(manualCountsValues[j][4]) : '';
                      
                    barcodeInputRange.setValue(description  + '\nwas found on the Counts page at line :\n' + (j + 3) 
                                                            + '\ninFlow Location(s) :\n' + manualCountsValues[j][1]
                                                            + '\nManual Count :\n' + manualCountsValues[j][2] 
                                                            + '\nRunning Sum :\n' + manualCountsValues[j][3]
                                                            + countedSinceString);

                    break; // Item was found on the Counts page, therefore stop searching
                  }
                }

                if (j === manualCountsValues.length) // Item was not found on the Counts page
                  barcodeInputRange.setValue(description + '\nwill be added to the Counts page at line :\n' + row);

                break; // Item was found, therefore stop searching
              }
                
              m = Math.ceil((u + l)/2) // Midpoint
            }
          }

          if (l >= m || m >= u)
          {
            if (upcCode.toString().length > 25)
              sheet.getRange(1, 1, 1, 2).setValues([['Barcode is Not Found.', '']]);
            else
              sheet.getRange(1, 1, 1, 2).setValues([['Barcode:\n\n' + upcCode + '\n\n is NOT FOUND.', '']]);

            sheet.getRange(1, 1).activate()
          }
          else
            sheet.getRange(1, 2).setValue('').activate();
        }
        else
          barcodeInputRange.setValue('The following is not a UPC-A or EAN-13: ' + upcCode);
      }
      else
        barcodeInputRange.setValue('The following barcode contains non-numerals: ' + upcCode);
    }
  }
  else if (barcodeInputRange.columnStart !== 1) // Quantity is entered
  {
    const quantity = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
      .setFontFamily("Arial").setFontColor("black").setFontSize(25)                      // Set the font parameters
      .setVerticalAlignment("middle").setHorizontalAlignment("center")                   // Set the alignment parameters
      .getValue();

    if (isNotBlank(quantity)) // The user may have hit the delete key
    {
      const item = sheet.getRange(1, 1).getValue().split('\n');    // The information from the left cell that is used to move the item to the Counts page
      const quantity_String = quantity.toString().toLowerCase();
      const quantity_String_Split = quantity_String.split(' ');

      if (quantity_String === 'clear')
      {
        manualCountsPage.getRange(item[2], 3, 1, 5).setNumberFormat('@').setValues([['', '', '', '', '']])
        sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                        + '\ninFlow Location(s) :'
                                                        + item[4]
                                                        + '\nManual Count :\n\nRunning Sum :\n',
                                                        '']]);
      }
      else if (quantity_String_Split[0] === 'uuu') // Unmarry upc
      {
        const upc = quantity_String_Split[1];

        if (upc > 100000)
        {
          const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
          unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[upc, item[0]]]);
          barcodeInputRange.setValue('UPC Code has been added to the unmarry list.')
          spreadsheet.getSheetByName("Scan").getRange(1, 1).activate();
        }
        else
          barcodeInputRange.setValue('Please enter a valid UPC Code to unmarry.')
      }
      else if (quantity_String_Split[0] === 'mmm') // Marry upc
      {
        const upc = quantity_String_Split[1];

        if (upc > 100000)
        {
          const marriedItem = item[0].split(' - ');
          const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
          const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
          manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[marriedItem.pop(), upc, marriedItem.pop(), item[0]]]);
          upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[upc, item[0]]]); 
          barcodeInputRange.setValue('UPC Code has been added to the database temporarily.');
          spreadsheet.getSheetByName("Scan").getRange(1, 1).activate();
        }
        else
          barcodeInputRange.setValue('Please enter a valid UPC Code to marry.')
      }
      else if (isNumber(quantity_String_Split[0]) && isNotBlank(quantity_String_Split[1]) && quantity_String_Split[1] != null)
      {
        if (item.length !== 1) // The cell to the left contains valid item information
        {
          quantity_String_Split[1] = quantity_String_Split[1].toUpperCase()

          if (item[1].split(' ')[0] === 'was') // The item was already on the Counts page
          {
            if (Number(quantity_String_Split[0]) < 0)
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()[0]
              const updatedCount = Number(itemValues[0]) + Number(quantity_String_Split[0]);
              const countedSince = getCountedSinceString(itemValues[2])
              const runningSum_Split = itemValues[1].split(' + ').map(location => location.split(': '))
              const idx = runningSum_Split.findIndex(loc => loc[0] == quantity_String_Split[1])

              if (idx !== -1)
              {
                runningSum_Split[idx][1] = runningSum_Split[idx][1] + quantity_String_Split[0]
                var runningSum = runningSum_Split.map(u => u.join(': ')).join(' + ')
                const quantity_Split = itemValues[4].split('\n')
                quantity_Split[idx] = Number(quantity_Split[idx]) + Number(quantity_String_Split[0])
                itemValues[4] = quantity_Split.join('\n')
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3], itemValues[4]]]);
              }
              else
              {
                var runningSum = (isNotBlank(itemValues[1])) ? ((Math.sign(quantity_String_Split[0]) === 1 || Math.sign(quantity_String_Split[0]) === 0)  ? 
                                                                    String(itemValues[1]) + ' \+ ' + quantity_String_Split[1] + ': ' + String(   quantity_String_Split[0])  : 
                                                                    String(itemValues[1]) + ' \- ' + quantity_String_Split[1] + ': ' + String(-1*quantity_String_Split[0])) :
                                                                      ((isNotBlank(itemValues[0])) ? 
                                                                        String(itemValues[0]) + ' \+ ' + quantity_String_Split[1] + ': ' + String(quantity_String_Split[0]) : 
                                                                        quantity_String_Split[1] + ': ' + String(quantity_String_Split[0]));

                if (isNotBlank(itemValues[3]) && isNotBlank(itemValues[4]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3] + '\n' + quantity_String_Split[1], itemValues[4] + '\n' + quantity_String_Split[0].toString()]]);
                else if (isNotBlank(itemValues[3]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
                else if (isNotBlank(itemValues[4]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    quantity_String_Split[1], itemValues[4] + '\n' + quantity_String_Split[0].toString()]]);
                else
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    quantity_String_Split[1], quantity_String_Split[0].toString()]]);
              }

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                              + '\ninFlow Location(s) :\n' + item[4]
                                                              + '\nManual Count :\n' + updatedCount 
                                                              + '\nRunning Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
            else
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()[0]
              const updatedCount = Number(itemValues[0]) + Number(quantity_String_Split[0]);
              const countedSince = getCountedSinceString(itemValues[2])
              const runningSum = (isNotBlank(itemValues[1])) ? ((Math.sign(quantity_String_Split[0]) === 1 || Math.sign(quantity_String_Split[0]) === 0)  ? 
                                                                    String(itemValues[1]) + ' \+ ' + quantity_String_Split[1] + ': ' + String(   quantity_String_Split[0])  : 
                                                                    String(itemValues[1]) + ' \- ' + quantity_String_Split[1] + ': ' + String(-1*quantity_String_Split[0])) :
                                                                      ((isNotBlank(itemValues[0])) ? 
                                                                        String(itemValues[0]) + ' \+ ' + quantity_String_Split[1] + ': ' + String(quantity_String_Split[0]) : 
                                                                        quantity_String_Split[1] + ': ' + String(quantity_String_Split[0]));

              if (isNotBlank(itemValues[3]) && isNotBlank(itemValues[4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[3] + '\n' + quantity_String_Split[1], itemValues[4] + '\n' + quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[3]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[3] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[1], itemValues[4] + '\n' + quantity_String_Split[0].toString()]]);
              else
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[1], quantity_String_Split[0].toString()]]);

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                              + '\ninFlow Location(s) :\n' + item[4]
                                                              + '\nManual Count :\n' + updatedCount 
                                                              + '\nRunning Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
          }
          else
          {
            const lastRow = manualCountsPage.getLastRow();
            const row = lastRow + 1;
            const range = manualCountsPage.getRange(row, 1, 1, 7)
            const itemValues = range.getValues()[0]

            if (isNotBlank(itemValues[5]) && isNotBlank(itemValues[6]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                new Date().getTime(), itemValues[5] + '\n' + quantity_String_Split[1], itemValues[6] + '\n' + quantity_String_Split[0].toString()]]);
            else if (isNotBlank(itemValues[5]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                new Date().getTime(), itemValues[5] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
            else if (isNotBlank(itemValues[6]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                new Date().getTime(), quantity_String_Split[1], itemValues[6] + '\n' + quantity_String_Split[0].toString()]]);
            else
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                new Date().getTime(), quantity_String_Split[1], quantity_String_Split[0].toString()]]);

            formatCountsPage(manualCountsPage, row, 1, 7)
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Counts page at line :\n' + item[2] 
                                                            + '\ninFlow Location(s) :\n' + item[4]
                                                            + '\nManual Count :\n' + quantity_String_Split[0],
                                                            '']]);
          }
        }
        else // The cell to the left does not contain the necessary item information to be able to move it to the Counts page
          barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

        sheet.getRange(1, 1).activate();
      }
      else if (isNumber(quantity_String_Split[1]))
      {
        if (item.length !== 1) // The cell to the left contains valid item information
        {
          quantity_String_Split[0] = quantity_String_Split[0].toUpperCase()

          if (item[1].split(' ')[0] === 'was') // The item was already on the Counts page
          {
            if (Number(quantity_String_Split[1]) < 0)
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()[0]
              const updatedCount = Number(itemValues[0]) + Number(quantity_String_Split[1]);
              const countedSince = getCountedSinceString(itemValues[2])
              const runningSum_Split = itemValues[1].split(' + ').map(location => location.split(': '))
              const idx = runningSum_Split.findIndex(loc => loc[0] == quantity_String_Split[0])

              if (idx !== -1)
              {
                runningSum_Split[idx][1] = runningSum_Split[idx][1] + quantity_String_Split[1]
                var runningSum = runningSum_Split.map(u => u.join(': ')).join(' + ')
                const quantity_Split = itemValues[4].split('\n')
                quantity_Split[idx] = Number(quantity_Split[idx]) + Number(quantity_String_Split[1])
                itemValues[4] = quantity_Split.join('\n')
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3], itemValues[4]]])
              }
              else
              {
                var runningSum = (isNotBlank(itemValues[1])) ? ((Math.sign(quantity_String_Split[1]) === 1 || Math.sign(quantity_String_Split[1]) === 0)  ? 
                                                                    String(itemValues[1]) + ' \+ ' + quantity_String_Split[0] + ': ' + String(   quantity_String_Split[1])  : 
                                                                    String(itemValues[1]) + ' \- ' + quantity_String_Split[0] + ': ' + String(-1*quantity_String_Split[1])) :
                                                                      ((isNotBlank(itemValues[0])) ? 
                                                                        String(itemValues[0]) + ' \+ ' + quantity_String_Split[0] + ': ' + String(quantity_String_Split[1]) : 
                                                                        quantity_String_Split[0] + ': ' + String(quantity_String_Split[1]));

                if (isNotBlank(itemValues[3]) && isNotBlank(itemValues[4]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3] + '\n' + quantity_String_Split[0], itemValues[4] + '\n' + quantity_String_Split[1].toString()]]);
                else if (isNotBlank(itemValues[3]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    itemValues[3] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
                else if (isNotBlank(itemValues[4]))
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    quantity_String_Split[0], itemValues[4] + '\n' + quantity_String_Split[1].toString()]]);
                else
                  range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                    quantity_String_Split[0], quantity_String_Split[1].toString()]]);
              }

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                              + '\ninFlow Location(s) :\n' + item[4]
                                                              + '\nManual Count :\n' + updatedCount 
                                                              + '\nRunning Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
            else
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()[0]
              const updatedCount = Number(itemValues[0]) + Number(quantity_String_Split[1]);
              const countedSince = getCountedSinceString(itemValues[2])
              const runningSum = (isNotBlank(itemValues[1])) ? ((Math.sign(quantity_String_Split[1]) === 1 || Math.sign(quantity_String_Split[1]) === 0)  ? 
                                                                    String(itemValues[1]) + ' \+ ' + quantity_String_Split[0] + ': ' + String(   quantity_String_Split[1])  : 
                                                                    String(itemValues[1]) + ' \- ' + quantity_String_Split[0] + ': ' + String(-1*quantity_String_Split[1])) :
                                                                      ((isNotBlank(itemValues[0])) ? 
                                                                        String(itemValues[0]) + ' \+ ' + quantity_String_Split[0] + ': ' + String(quantity_String_Split[1]) : 
                                                                        quantity_String_Split[0] + ': ' + String(quantity_String_Split[1]));

              if (isNotBlank(itemValues[3]) && isNotBlank(itemValues[4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[3] + '\n' + quantity_String_Split[0], itemValues[4] + '\n' + quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[3]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[3] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[0], itemValues[4] + '\n' + quantity_String_Split[1].toString()]]);
              else
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[0], quantity_String_Split[1].toString()]]);

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                              + '\ninFlow Location(s) :\n' + item[4]
                                                              + '\nManual Count :\n' + updatedCount 
                                                              + '\nRunning Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
          }
          else
          {
            const lastRow = manualCountsPage.getLastRow();
            const row = lastRow + 1;
            const range = manualCountsPage.getRange(row, 1, 1, 7)
            const itemValues = range.getValues()[0]

            if (isNotBlank(itemValues[5]) && isNotBlank(itemValues[6]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                new Date().getTime(), itemValues[5] + '\n' + quantity_String_Split[0], itemValues[6] + '\n' + quantity_String_Split[1].toString()]]);
            else if (isNotBlank(itemValues[5]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                new Date().getTime(), itemValues[5] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
            else if (isNotBlank(itemValues[6]))
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                new Date().getTime(), quantity_String_Split[0], itemValues[6] + '\n' + quantity_String_Split[1].toString()]]);
            else
              range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], '', quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                new Date().getTime(), quantity_String_Split[0], quantity_String_Split[1].toString()]]);

            formatCountsPage(manualCountsPage, row, 1, 7)
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Counts page at line :\n' + item[2] 
                                                            + '\nManual Count :\n' + quantity_String_Split[1],
                                                            '']]);
          }
        }
        else // The cell to the left does not contain the necessary item information to be able to move it to the Counts page
          barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

        sheet.getRange(1, 1).activate();
      }
      else if (quantity <= 100000) // If false, Someone probably scanned a barcode in the quantity cell (not likely to have counted an inventory amount of 100 000)
      {
        if (item.length !== 1) // The cell to the left contains valid item information
        {
          if (item[1].split(' ')[0] === 'was') // The item was already on the Counts page
          {
            const range = manualCountsPage.getRange(item[2], 3, 1, 3);
            const itemValues = range.getValues()[0]
            const updatedCount = Number(itemValues[0]) + quantity;
            const countedSince = getCountedSinceString(itemValues[2])
            const runningSum = (isNotBlank(itemValues[1])) ? ((Math.sign(quantity) === 1 || Math.sign(quantity) === 0)  ? 
                                                                  String(itemValues[1]) + ' \+ ' + String(   quantity)  : 
                                                                  String(itemValues[1]) + ' \- ' + String(-1*quantity)) :
                                                                    ((isNotBlank(itemValues[0])) ? 
                                                                      String(itemValues[0]) + ' \+ ' + String(quantity) : 
                                                                      String(quantity));
            range.setNumberFormats([['#.#', '@', '#']]).setValues([[updatedCount, runningSum, new Date().getTime()]])
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Counts page at line :\n' + item[2] 
                                                            + '\ninFlow Location(s) :\n' + item[4]
                                                            + '\nManual Count :\n' + updatedCount 
                                                            + '\nRunning Sum :\n' + runningSum
                                                            + '\nLast Counted :\n' + countedSince,
                                                            '']]);
          }
          else
          {
            const lastRow = manualCountsPage.getLastRow();
            const row = lastRow + 1;
            manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']]).setValues([[item[0], '', quantity, '\'' + String(quantity), new Date().getTime()]])
            formatCountsPage(manualCountsPage, row, 1, 7)
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Counts page at line :\n' + item[2] 
                                                            + '\nManual Count :\n' + quantity,
                                                            '']]);
          }
        }
        else // The cell to the left does not contain the necessary item information to be able to move it to the Counts page
          barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

        sheet.getRange(1, 1).activate();
      }
      else 
        barcodeInputRange.setValue('Please enter a valid quantity.')
    }
  }
}

/**
 * This function opens a modal dialogue box that allows the user to drag and drop a file for import.
 */
function openDragAndDrop()
{
  const html = HtmlService.createHtmlOutputFromFile('DragAndDrop.html').setWidth(800).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload File')
}

/**
 * This function takes the information from the Item Search or Manual Counts page and the user's recently scanned barcode in the created date column and it 
 * populates the Manual Scan page with the relevant data need to update the count of the particular item.
 * 
 * @param {Spreadsheet}  ss    : The active spreadsheet.
 * @param {Sheet}       sheet  : The active sheet.
 * @param {Number}      rowNum : The row number of the current item.
 * @author Jarren Ralf
 */
function populateManualScan(ss, sheet, rowNum, newItemDescription)
{
  const barcodeInputRange = ss.getSheetByName('Scan').getRange(1, 1);
  const manualCountsPage = ss.getSheetByName("Counts");
  const currentStock = (sheet.getSheetName() === 'Item Search') ? 2 : 1;
  const lastRow = manualCountsPage.getLastRow();
  var itemValues = (sheet.getSheetName() === 'Item Search') ? sheet.getSheetValues(rowNum, 2, 1, 3)[0] : sheet.getSheetValues(rowNum, 1, 1, 2)[0];

  if (newItemDescription != null)
  {
    itemValues[0] = newItemDescription;
    itemValues[currentStock] = '';
  }

  if (lastRow <= 3) // There are no items on the manual counts page
    barcodeInputRange.setValue(itemValues[0] + '\nwill be added to the Counts page at line :\n' + 4 + '\nStock :\n' + itemValues[currentStock]);
  else // There are existing items on the manual counts page
  {
    const row = lastRow + 1;
    const manualCountsValues = manualCountsPage.getSheetValues(3, 1, row - 3, 4);

    for (var j = 0; j < manualCountsValues.length; j++) // Loop through the manual counts page
    {
      if (manualCountsValues[j][0] === itemValues[0]) // The description matches
      {
        barcodeInputRange.setValue(itemValues[0]  + '\nwas found on the Counts page at line :\n' + (j + 3) 
                                                  + '\nManual Count :\n' + manualCountsValues[j][2] 
                                                  + '\nRunning Sum :\n' + manualCountsValues[j][3]);
        break; // Item was found on the manual counts page, therefore stop searching
      }
    }

    if (j === manualCountsValues.length) // Item was not found on the manual counts page
      barcodeInputRange.setValue(itemValues[0] + '\nwill be added to the Counts page at line :\n' + row + '\nStock :\n' + itemValues[currentStock]);
  }

  barcodeInputRange.offset(0, 1).activate();
}

/**
 * This function processes the import of an InFlow Purchase Order.
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3;

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      info = [
        sheets[sheet].getLastRow(),
        sheets[sheet].getLastColumn(),
        sheets[sheet].getMaxRows(),
        sheets[sheet].getMaxColumns()
      ]

      // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
      if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
          (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0)) 
      {
        const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); // This is the order entry data

        if (values[0].includes('Vendor name'))
          importAdagioPurchaseOrder(values, spreadsheet);
        else if (values[0].includes('Cust #'))
          importAdagioSalesOrder(values, spreadsheet); // Needs to be written
        else if (values[0].includes('Sublocation')) 
          importStockLevels(values, spreadsheet);
        else if (values[0].includes('DefaultPricingScheme'))
          updateInFlowCustomerList(values, spreadsheet);
        else if (values[0].includes('ReorderQuantity')) 
          updateInflowProductDetails(values, spreadsheet);
        else if (values[0].includes('PreferredCarrier'))
          updateInFlowVendorList(values, spreadsheet);

        if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
          spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

        break;
      }
    }
  }
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the SearchData page for the items in question.
 * It also highlights the items that are already on the shipped page and already on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function search(e, spreadsheet, sheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const rowEnd = range.rowEnd;
  const colEnd = range.columnEnd;

  if ((colEnd == null || colEnd == 3 || col == colEnd)) // Check and make sure only a single column is being edited
  {
    if (row == rowEnd) // Check and make sure only a single cell is being edited
    {
      if (row === 3 && col === 1) // The check box that toggle the "Add New Items" mode
      {
        var rng  = sheet.getRange(1, 1, 3, 5);
        var vals = rng.getValues()

        if (e.value === 'FALSE') // Regular Search mode
        {
          vals[0][3] = '=COUNTIF(INVENTORY!$B$3:$B, "DOCK")&" items in Location DOCK"'

          rng.setBackgrounds([ ['#f1c232', 'white',   '#f1c232', '#f1c232', '#f1c232'], 
                              ['#f1c232', '#f1c232', '#f1c232', '#f1c232', '#f1c232'], 
                              ['#f1c232', '#f1c232', '#f1c232', '#f1c232', '#f1c232']]).setValues(vals)

          const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                    // Clear the formatting of the range of the search box
            .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
            .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
            .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
            .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
            .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

          const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

          if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
          {
            spreadsheet.toast('Searching...')
            const startTime = new Date().getTime();
            const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
            const functionRunTimeRange = sheet.getRange(2, 1);   // The range that will display the runtimes for the search and formatting
            const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5); // The entire range of the Item Search page
            const numSearches = searches.length; // The number searches
            const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
            const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);
            var output = [], numSearchWords, UoM;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (searches[0][0].substring(0, 3) === 'loc') // Search for locations
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'loc')
                        continue;

                      if (searches[j][k][0] === '_' && data[i][1].toString().toLowerCase()[data[i][1].toString().length - 1] === searches[j][k][searches[j][k].length - 1])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else if (searches[j][k][searches[j][k].length - 1] === '_' && data[i][1].toString().toLowerCase()[0] === searches[j][k][0])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
              else if (searches[0][0].substring(0, 3) === 'ser') // Search for the serial number
              {
                if (numSearches === 1 && searches[0].length == 1)
                  output.push(...data.filter(serial => isNotBlank(serial[3])).map(values => {
                    UoM = values[0].toString().split(' - ');
                    UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm
                    return [UoM, ...values]
                  }))
                else
                {
                  for (var i = 0; i < data.length; i++) // Loop through all of the serial numbers from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (searches[j][k].substring(0, 3) === 'ser')
                          continue;

                        if (data[i][3].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            UoM = data[i][0].toString().split(' - ')
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                            output.push([UoM, ...data[i]]);
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }

                output = output.sort(sortBySerial)
              }
              else // Regular search through the descriptions
              {
                if (/^\d+$/.test(searches[0][0]) && (isUPC_A(searches[0][0]) || isEAN_13(searches[0][0])) && numSearches === 1 && searches[0].length == 1) // Check if a barcode was scanned
                {
                  const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
                  const upcs = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
                  var l = 0; // Lower-bound
                  var u = upcs.length - 1; // Upper-bound
                  var m = Math.ceil((u + l)/2) // Midpoint
                  searches[0][0] = parseInt(searches[0][0])

                  while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
                  {
                    if (searches[0][0] < parseInt(upcs[m][0]))
                      u = m;   
                    else if (searches[0][0] > parseInt(upcs[m][0]))
                      l = m;
                    else // UPC code was found!
                    {
                      const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]

                      for (var i = 0; i < data.length; i++)
                      {
                        if (description === data[i][0])
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                        }
                      }
                      break; // Item was found, therefore stop searching
                    }

                    m = Math.ceil((u + l)/2) // Midpoint
                  }
                }
                else
                {
                  for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            UoM = data[i][0].toString().split(' - ')
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                            output.push([UoM, ...data[i]]);
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
            }
            else // The word 'not' was found in the search string
            {
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              if (searches[0][0].substring(0, 3) === 'loc')
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'loc')
                        continue;

                      if (searches[j][k][0] === '_' && data[i][1].toString().toLowerCase()[data[i][1].toString().length - 1] === searches[j][k][searches[j][k].length - 1])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else if (searches[j][k][searches[j][k].length - 1] === '_' && data[i][1].toString().toLowerCase()[0] === searches[j][k][0])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
              else if (searches[0][0].substring(0, 3) === 'ser')
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'ser')
                        continue;

                      if (data[i][3].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][3].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortBySerial)
              }
              else // Regular search through the descriptions
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
            }

            const numItems = output.length;

            if (numItems === 0) // No items were found
            {
              sheet.getRange('B1').activate(); // Move the user back to the seachbox
              itemSearchFullRange.clearContent(); // Clear content
              const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#660000').build();
              const message = SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 17, textStyle).build();
              searchResultsDisplayRange.setRichTextValue(message);
            }
            else
            {
              sheet.getRange('B4').activate(); // Move the user to the top of the search items
              itemSearchFullRange.clearContent(); // Clear content and reset the text format
              sheet.getRange(4, 1, numItems, 5).setValues(output);
              (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
            }

            functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " s");
            spreadsheet.toast('Searching Complete.')
          }
        }
        else if (e.value) // Add New Item Mode
        {
          vals[0][3] = 'Add NEW item Mode: ON'

          rng.setBackgrounds([ ['#3c78d8', 'white',   '#3c78d8', '#3c78d8', '#3c78d8'], 
                               ['#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8'], 
                               ['#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8', '#3c78d8']]).setValues(vals)

          const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                    // Clear the formatting of the range of the search box
            .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
            .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
            .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
            .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
            .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

          const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

          if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
          {
            spreadsheet.toast('Searching...')
            const startTime = new Date().getTime();
            const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
            const functionRunTimeRange = sheet.getRange(2, 1);   // The range that will display the runtimes for the search and formatting
            const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5); // The entire range of the Item Search page
            const numSearches = searches.length; // The number searches
            const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
            const inflowItems = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 1);
            var output = [], numSearchWords, isInflow;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (/^\d+$/.test(searches[0][0]) && (isUPC_A(searches[0][0]) || isEAN_13(searches[0][0])) && numSearches === 1 && searches[0].length == 1) // Check if a barcode was scanned in the cell
              {
                const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
                const upcs = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
                var l = 0; // Lower-bound
                var u = upcs.length - 1; // Upper-bound
                var m = Math.ceil((u + l)/2) // Midpoint
                searches[0][0] = parseInt(searches[0][0])

                while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
                {
                  if (searches[0][0] < parseInt(upcs[m][0]))
                    u = m;   
                  else if (searches[0][0] > parseInt(upcs[m][0]))
                    l = m;
                  else // UPC code was found!
                  {
                    const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]
                    const sku = description.toString().toUpperCase().split(' - ').pop();
                    isInflow = inflowItems.find(v => v[0].toString().toUpperCase().split(' - ').pop() === sku)
                    UoM = description.toString().split(' - ')
                    UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm
                    output.push([UoM, description, (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                    break; // Item was found, therefore stop searching
                  }
                    
                  m = Math.ceil((u + l)/2) // Midpoint
                }
              }
              else
              {
                const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
                const uom = data[0].indexOf('Price Unit')
                const fullDescription = data[0].indexOf('Item List')
                const itemNumber = data[0].indexOf('Item #')
                
                for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          const sku = data[i][itemNumber].toString().toUpperCase()
                          isInflow = inflowItems.find(item => item[0].toString().toUpperCase().split(' - ').pop() === sku)
                          output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }   
            }
            else // The word 'not' was found in the search string
            {
              const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
              const uom = data[0].indexOf('Price Unit')
              const fullDescription = data[0].indexOf('Item List')
              const itemNumber = data[0].indexOf('Item #')
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                      {
                        for (var l = 0; l < dontIncludeTheseWords.length; l++)
                        {
                          if (!data[i][fullDescription].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                          {
                            if (l === dontIncludeTheseWords.length - 1)
                            {
                              const sku = data[i][itemNumber].toString().toUpperCase()
                              isInflow = inflowItems.find(item => item[0].toString().toUpperCase().split(' - ').pop() === sku)
                              output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                              break loop;
                            }
                          }
                          else
                            break;
                        }
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                  }
                }
              }
            }

            const numItems = output.length;

            if (numItems === 0) // No items were found
            {
              sheet.getRange('B1').activate(); // Move the user back to the seachbox
              itemSearchFullRange.clearContent(); // Clear content
              const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#660000').build();
              const message = SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 17, textStyle).build();
              searchResultsDisplayRange.setRichTextValue(message);
            }
            else
            {
              sheet.getRange('B4').activate(); // Move the user to the top of the search items
              itemSearchFullRange.clearContent(); // Clear content and reset the text format
              sheet.getRange(4, 1, numItems, 5).setValues(output);
              (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
            }

            functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " s");
            spreadsheet.toast('Searching Complete.')
          }
        }
      }
      else if (row === 1 && col === 2) // Check if the search box is edited
      {
        const startTime = new Date().getTime();
        const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
        const functionRunTimeRange = sheet.getRange(2, 1);   // The range that will display the runtimes for the search and formatting
        const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5); // The entire range of the Item Search page
        const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                    // Clear the formatting of the range of the search box
          .setBorder(true, true, true, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
          .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
          .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
          .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
          .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

        const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

        if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
        {
          spreadsheet.toast('Searching...')
          const numSearches = searches.length; // The number searches
          var output = [];

          if (sheet.getSheetValues(3, 1, 1, 1)[0][0]) // Check if the Search page is in "Add New Item" mode
          {
            const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
            const inflowItems = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 1);
            var numSearchWords, isInflow;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (/^\d+$/.test(searches[0][0]) && (isUPC_A(searches[0][0]) || isEAN_13(searches[0][0])) && numSearches === 1 && searches[0].length == 1) // Check if a barcode was scanned in the cell
              {
                const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
                const upcs = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
                var l = 0; // Lower-bound
                var u = upcs.length - 1; // Upper-bound
                var m = Math.ceil((u + l)/2) // Midpoint
                searches[0][0] = parseInt(searches[0][0])

                while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
                {
                  if (searches[0][0] < parseInt(upcs[m][0]))
                    u = m;   
                  else if (searches[0][0] > parseInt(upcs[m][0]))
                    l = m;
                  else // UPC code was found!
                  {
                    const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]
                    const sku = description.toString().toUpperCase().split(' - ').pop();
                    isInflow = inflowItems.find(v => v[0].toString().toUpperCase().split(' - ').pop() === sku)
                    UoM = description.toString().split(' - ')
                    UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm
                    output.push([UoM, description, (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                    break; // Item was found, therefore stop searching
                  }
                    
                  m = Math.ceil((u + l)/2) // Midpoint
                }
              }
              else
              {
                const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
                const uom = data[0].indexOf('Price Unit')
                const fullDescription = data[0].indexOf('Item List')
                const itemNumber = data[0].indexOf('Item #')

                for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          const sku = data[i][itemNumber].toString().toUpperCase()
                          isInflow = inflowItems.find(item => item[0].toString().toUpperCase().split(' - ').pop() === sku)
                          output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }
            else // The word 'not' was found in the search string
            {
              const data = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
              const uom = data[0].indexOf('Price Unit')
              const fullDescription = data[0].indexOf('Item List')
              const itemNumber = data[0].indexOf('Item #')
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[i][fullDescription].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                      {
                        for (var l = 0; l < dontIncludeTheseWords.length; l++)
                        {
                          if (!data[i][fullDescription].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                          {
                            if (l === dontIncludeTheseWords.length - 1)
                            {
                              const sku = data[i][itemNumber].toString().toUpperCase()
                              isInflow = inflowItems.find(item => item[0].toString().toUpperCase().split(' - ').pop() === sku)
                              output.push([data[i][uom], data[i][fullDescription], (isInflow == null) ? 'NOT in inFlow' : '', '', '']);
                              break loop;
                            }
                          }
                          else
                            break;
                        }
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                  }
                }
              }
            }
          }
          else // Regular inFlow search mode
          {
            const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
            const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);
            var numSearchWords, UoM;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (searches[0][0].substring(0, 3) === 'loc')
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'loc')
                        continue;

                      if (searches[j][k][0] === '_' && data[i][1].toString().toLowerCase()[data[i][1].toString().length - 1] === searches[j][k][searches[j][k].length - 1])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else if (searches[j][k][searches[j][k].length - 1] === '_' && data[i][1].toString().toLowerCase()[0] === searches[j][k][0])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
              else if (searches[0][0].substring(0, 3) === 'ser')
              {
                if (numSearches === 1 && searches[0].length == 1)
                  output.push(...data.filter(serial => isNotBlank(serial[3])).map(values => {
                    UoM = values[0].toString().split(' - ');
                    UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm
                    return [UoM, ...values]
                  }))
                else
                {
                  for (var i = 0; i < data.length; i++) // Loop through all of the serial numbers from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (searches[j][k].substring(0, 3) === 'ser')
                          continue;

                        if (data[i][3].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            UoM = data[i][0].toString().split(' - ')
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                            output.push([UoM, ...data[i]]);
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }

                output = output.sort(sortBySerial)
              }
              else // Regular search through the descriptions
              {
                if (/^\d+$/.test(searches[0][0]) && (isUPC_A(searches[0][0]) || isEAN_13(searches[0][0])) && numSearches === 1 && searches[0].length == 1) // Check if a barcode was scanned
                {
                  const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database')
                  const upcs = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1)
                  var l = 0; // Lower-bound
                  var u = upcs.length - 1; // Upper-bound
                  var m = Math.ceil((u + l)/2) // Midpoint
                  searches[0][0] = parseInt(searches[0][0])

                  while (l < m && u > m) // Loop through the UPC codes using the binary search algorithm
                  {
                    if (searches[0][0] < parseInt(upcs[m][0]))
                      u = m;   
                    else if (searches[0][0] > parseInt(upcs[m][0]))
                      l = m;
                    else // UPC code was found!
                    {
                      const description = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0]

                      for (var i = 0; i < data.length; i++)
                      {
                        if (description === data[i][0])
                        {
                          UoM = data[i][0].toString().split(' - ')
                          UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                          output.push([UoM, ...data[i]]);
                        }
                      }

                      break; // Item was found, therefore stop searching
                    }
                      
                    m = Math.ceil((u + l)/2) // Midpoint
                  }
                }
                else
                {
                  for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            UoM = data[i][0].toString().split(' - ')
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                            output.push([UoM, ...data[i]]);
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
            }
            else // The word 'not' was found in the search string
            {
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              if (searches[0][0].substring(0, 3) === 'loc')
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'loc')
                        continue;

                      if (searches[j][k][0] === '_' && data[i][1].toString().toLowerCase()[data[i][1].toString().length - 1] === searches[j][k][searches[j][k].length - 1])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else if (searches[j][k][searches[j][k].length - 1] === '_' && data[i][1].toString().toLowerCase()[0] === searches[j][k][0])
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
              else if (searches[0][0].substring(0, 3) === 'ser')
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the locations from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k].substring(0, 3) === 'ser')
                        continue;

                      if (data[i][3].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][3].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortBySerial)
              }
              else // Regular search through the descriptions
              {
                for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          for (var l = 0; l < dontIncludeTheseWords.length; l++)
                          {
                            if (!data[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === dontIncludeTheseWords.length - 1)
                              {
                                UoM = data[i][0].toString().split(' - ')
                                UoM = (UoM.length >= 5) ? UoM[UoM.length - 2] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                                output.push([UoM, ...data[i]]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                    }
                  }
                }

                output = output.sort(sortByLocations)
              }
            }
          }

          const numItems = output.length;

          if (numItems === 0) // No items were found
          {
            sheet.getRange('B1').activate(); // Move the user back to the seachbox
            itemSearchFullRange.clearContent(); // Clear content
            const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#660000').build();
            const message = SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 17, textStyle).build();
            searchResultsDisplayRange.setRichTextValue(message);
          }
          else
          {
            sheet.getRange('B4').activate(); // Move the user to the top of the search items
            itemSearchFullRange.clearContent(); // Clear content and reset the text format
            sheet.getRange(4, 1, numItems, 5).setValues(output);
            (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
          }
        }
        else
        {
          itemSearchFullRange.clearContent(); // Clear content 
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#660000').build();
          const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\n\nPlease try again.").setTextStyle(0, 15, textStyle).build();
          searchResultsDisplayRange.setRichTextValue(message);
        }

        functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " s");
        spreadsheet.toast('Searching Complete.')
      }
      else if (row > 3 && col === 3)
      {
        if (userHasNotPressedDelete(e.value))
        {
          const value = e.value.split(' ', 2);
          range.setValue(e.oldValue);

          if (value[0].toLowerCase() === 'mmm')
          {
            if (value[1] > 100000)
            {
              const item = sheet.getSheetValues(row, 1, 1, 4)[0];
              const itemSplit = item[1].split(' - ');
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[value[1], item[1]]])
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[itemSplit.pop(), value[1], itemSplit.pop(), item[1]]]);
              const range = upcDatabaseSheet.getDataRange();
              range.setNumberFormat('@').setValues(range.getValues().sort(sortUPCsNumerically))
              populateManualScan(spreadsheet, sheet, row)
            }
            else
              Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          }
          else if (value[0].toLowerCase() === 'uuu')
          {
            if (value[1] > 100000)
            {
              const item = sheet.getSheetValues(row, 2, 1, 1)[0][0];
              const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
              unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[value[1], item]]);
              spreadsheet.getSheetByName('Scan').getRange(1, 1).activate()
            }
            else
              Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          }
          // else if (value[0].toLowerCase() === 'aaa')
          // {
          //   if (value[1] > 100000)
          //   {
          //     const item = sheet.getSheetValues(row, 1, 1, 2)[0]
          //     const newItem = item[1].split(' - ')
          //     newItem[newItem.length - 1] = 'MAKE_NEW_SKU'
          //     item[1] = newItem.join(' - ')
          //     const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
          //     const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
          //     const inventorySheet = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName('INVENTORY') : spreadsheet.getSheetByName('SearchData');
          //     manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([['MAKE_NEW_SKU', value[1], item[0], item[1]]]);
          //     upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[value[1], item[0], item[1]]]); 
          //     inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[item[0], item[1]]]); // Add the 'MAKE_NEW_SKU' item to the inventory sheet

          //     populateManualScan(spreadsheet, sheet, row, item[1])
          //     sheet.getRange(4, 1, MAX_NUM_ITEMS, 6).setValues(spreadsheet.getSheetByName('Recent').getSheetValues(2, 1, MAX_NUM_ITEMS, 6));
          //     sheet.getRange(1, 1, 1, 2).setValues([["The last " + MAX_NUM_ITEMS + " created items are displayed.", ""]]);
          //   }
          //   else
          //     Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          // }
          else if (value[0].toLowerCase() === 'sss')
            populateManualScan(spreadsheet, sheet, row)
        }
        else
          range.setValue(e.oldValue);
      }
    }
    else if (row != rowEnd && row > 3 & col == 2) // Multiple rows were pasted on thew search page
    {
      const values = range.getValues().filter(blank => isNotBlank(blank[0]))

      if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
      {
        const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
        const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);
        var someSKUsNotFound = false, skus;

        if (values[0][0].toString().includes(' - ')) // Strip the sku from the google description
        {
          skus = values.map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][0].split(' - ').pop().toString().toUpperCase() == item[0].toString().split(' - ').pop().toString().toUpperCase())
                return [data[i][0].split(' - ')[data[i][0].split(' - ').length - 2], ...data[i]];
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), '', '', '']
          });
        }
        else if (values[0][0].toString().includes('-'))
        {
          skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][0].split(' - ').pop().toString().toUpperCase() == item.toString().toUpperCase())
                return [data[i][0].split(' - ')[data[i][0].split(' - ').length - 2], ...data[i]];
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item, '', '', '']
          });
        }
        else
        {
          skus = values.map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][0].split(' - ').pop().toString().toUpperCase() == item[0].toString().toUpperCase())
                return [data[i][0].split(' - ')[data[i][0].split(' - ').length - 2], ...data[i]];
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0], '', '', '']
          });
        }

        if (someSKUsNotFound)
        {
          const skusNotFound = [];
          var isSkuFound;

          const skusFound = skus.filter(item => {
            isSkuFound = item[0] !== 'SKU Not Found:'

            if (!isSkuFound)
              skusNotFound.push(item)

            return isSkuFound;
          })

          const numSkusFound = skusFound.length;
          const numSkusNotFound = skusNotFound.length;
          const items = [].concat.apply([], [skusNotFound, skusFound.sort(sortByLocations)]); // Concatenate all of the item values as a 2-D array
          const numItems = items.length;
          const numItemsOutOfStock = items.reverse().findIndex(loc => loc[2] !== '');
          const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center'])
          const WHITE = new Array(5).fill('white')
          const YELLOW = new Array(5).fill('#ffe599')
          const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

          sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 5)
              .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours)
              .setBorder(false, null, false, null, false, false).setValues(items.reverse())
            .offset(numSkusNotFound, 0, numSkusFound - numItemsOutOfStock, 5).activate()
        }
        else // All SKUs were succefully found
        {
          const numItems = skus.length
          const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center'])

          sheet.getRange(4, 1, sheet.getMaxRows() - 2, 5).clearContent().setBackground('white').setFontColor('black').offset(0, 0, numItems, 5)
            .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments)
            .setBorder(false, null, false, null, false, false).setValues(skus).activate()
        }
      }
    }
  }
}

/**
* Sorts data by the customers while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCustomers(a, b)
{
  return (a[0].toLowerCase() === b[0].toLowerCase()) ? 0 : (a[0] === '') ? 1 : (b[0] === '') ? -1 : (a[0].toLowerCase() < b[0].toLowerCase()) ? -1 : 1;
}

/**
* Sorts data by the locations while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByLocations(a, b)
{
  return (a[2].toLowerCase() === b[2].toLowerCase()) ? 0 : (a[2] === '') ? 1 : (b[2] === '') ? -1 : (a[2].toLowerCase() < b[2].toLowerCase()) ? -1 : 1;
}

/**
* Sorts data by the serial number based on the PNT conentions for bales of net. The first criteria is to sort numerically by PO number, then sort numerically again by BALE numer.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortBySerial(a, b)
{
  var baleNum1 = a[4].toLowerCase().split(' bale');
  var baleNum2 = b[4].toLowerCase().split(' bale');
  var poNum1 = baleNum1[0].split('po')[1];
  var poNum2 = baleNum2[0].split('po')[1];

  // If PO number is not present in the serial, push those items to the to of the list
  if (poNum2 === null || poNum2 === undefined) 
    return 1;
  else if (poNum1 === null || poNum1 === undefined)
    return -1;
  else 
  {
    if (Number(poNum1.match(/\d+/g)) == Number(poNum2.match(/\d+/g))) // PO numbers match numerically
    {
      // If BALE number is not present in the serial, push those items to the top of the list
      if (baleNum1[1]  === undefined)
        return 1;
      else if (baleNum2[1]  === undefined)
        return -1;
      else
        return Number(baleNum1[1].match(/\d+/g)) - Number(baleNum2[1].match(/\d+/g)); // For matching PO numbers, sort numerically by BALE number
    }
    else
      return Number(poNum1.match(/\d+/g)) - Number(poNum2.match(/\d+/g)) // Sort the PO numbers numerically
  }
}

/**
 * This function sorts the UPC Codes in numerical order.
 * 
 * @author Jarren Ralf
 */
function sortUPCsNumerically(a, b)
{
  return parseInt(a[0]) - parseInt(b[0]);
}

/**
 * This function moves the selected items from the item search sheet to the stock levels page in preparation for Stock Adjustments.
 * 
 * @author Jarren Ralf
 */
function stockAdjustment()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('Stock Levels'), false)
}

/**
 * This function moves the selected items from the item search sheet to the stock levels page in preparation for Stock Transfers.
 * 
 * @author Jarren Ralf
 */
function stockTransfer()
{
  copySelectedValues(SpreadsheetApp.getActive().getSheetByName('Stock Levels'), true)
}

/**
 * This function handles either the imported inFlow Vendor or Customer List and updates the relevant data.
 * 
 * @param {String[][]}    values    : The values of the inFlow information
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @param {String}       sheetName  : The name of the sheet to be updated
 * @param {Number}    startingIndex : The value to start at while looping through the data.
 * @author Jarren Ralf
 */
function updateInFlowList_(values, spreadsheet, sheetName, startingIndex)
{
  const sheet = spreadsheet.getSheetByName(sheetName)

  if (sheetName === 'Product Details')
  {
    const lastCol = sheet.getLastColumn();
    const header_ProductDetails = values.shift();
    const columnsToKeep = sheet.getSheetValues(1, 1, 1, lastCol)[0].map(col => header_ProductDetails.indexOf(col))
    const productDetails = values.map(col => [...columnsToKeep.map(c => col[c])])
    sheet.showSheet().getRange(3, 1, sheet.getLastRow(), lastCol).clearContent().offset(0, 0, productDetails.length, lastCol).setValues(productDetails).activate()
  }
  else
  {
    const inFlowValues = sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 1).flat();
    var isThereNewValuesToAdd = false, newValues = [];

    for (var i = startingIndex; i < values.length; i++)
    {
      if (!inFlowValues.includes(values[i][0]))
      {
        isThereNewValuesToAdd = true;
        newValues.push([values[i][0], ''])
      }
    }

    if (isThereNewValuesToAdd)
    {
      const updatedValues = sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 2)
      updatedValues.push(...newValues)
      updatedValues.sort(sortByCustomers)
      sheet.getRange(2, 1, updatedValues.length, 2).setValues(updatedValues).activate()
      spreadsheet.toast('Number of new ' + sheetName.toLowerCase() + ' added: ' + newValues.length)
    }
    else
      spreadsheet.toast('No new ' + sheetName.toLowerCase() + ' to add...')
  }
}

/**
 * This function handles the imported inFlow Customer List and updates the data on the Customer tab.
 * 
 * @param {String[][]}    values    : The values of the inFlow Customer list
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @author Jarren Ralf
 */
function updateInFlowCustomerList(values, spreadsheet)
{
  updateInFlowList_(values, spreadsheet, 'Customers', 2)
}

function updateInflowProductDetails(values, spreadsheet)
{
  updateInFlowList_(values, spreadsheet, 'Product Details')
}

/**
 * This function handles the imported inFlow Vendor List and updates the data on the Vendor tab.
 * 
 * @param {String[][]}    values    : The values of the inFlow Vendor list
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @author Jarren Ralf
 */
function updateInFlowVendorList(values, spreadsheet)
{
  updateInFlowList_(values, spreadsheet, 'Vendors', 1)
}

/**
 * This function places the inFlow Stock Levels on the Inventory sheet from the source file on the google drive.
 * 
 * @author Jarren Ralf
 */
function updateStockLevels()
{
  const startTime = new Date().getTime();
  const inflowData = Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString())
  importStockLevels(inflowData, SpreadsheetApp.getActive(), startTime)
}

/**
 * This function looks at the UPC database and removes all of the barcodes that are not UPC-A. It also updates the data with the typical Google sheets description string.
 * 
 * @author Jarren Ralf
 */
function updateUPCs()
{
  var sku_upc, item;
  const adagioInventory = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = adagioInventory[0].indexOf('Item #')
  const fullDescription = adagioInventory[0].indexOf('Item List')
  const data = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString()).filter(upc => isUPC_A(upc[0]) || isEAN_13(upc[0])).map(upcs => {
    sku_upc = upcs[1].toUpperCase()
    item = adagioInventory.find(sku => sku[itemNum] === sku_upc)
    return (item != null) ? [upcs[0], item[fullDescription]] : null;
  }).filter(val => val != null).sort(sortUPCsNumerically)

  SpreadsheetApp.getActive().getSheetByName('UPC Database').clearContents().getRange(1, 1, data.length, data[0].length).setNumberFormat('@').setValues(data)
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning false if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is not-undefined or not.
* @author Jarren Ralf
*/
function userHasNotPressedDelete(value)
{
  return value !== undefined;
}

/**
 * This function checks if the user edits the item description or the Current Count column on the 
 * Counts page. If they did, then a warning appears and reverses the changes that they made.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param    {String}     sheetName  : The string that represents the name of the sheet
 * @author Jarren Ralf
 */
function warning(e, sheet, sheetName)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;

  if (row == range.rowEnd && col == range.columnEnd) // Single cell
  {
    if (col == 1)
    {
      SpreadsheetApp.getUi().alert("Please don't attempt to change the items from the Counts page.\n\nGo to the Scan page to add new products to this list.")
      range.setValue(e.oldValue); // Put the old value back in the cell
    }
    else if (col == 2)
    {
      SpreadsheetApp.getUi().alert("Please don't change values in the Current Count column.\n\nType your updated inventory quantity in the New Count column.");
      range.setValue(e.oldValue); // Put the old value back in the cell
      if (userHasNotPressedDelete(e.value)) sheet.getRange(row, 3).setValue(e.value).activate(); // Move the count the user entered to the New Count column
    }
    else if (col == 3 && sheetName === 'Counts')
    {
      if (e.oldValue !== undefined) // Old value is NOT blank
      {
        if (userHasNotPressedDelete(e.value)) // New value is NOT blank
        {
          const valueSplit = e.value.toString().split(' ');

          if (isNumber(e.value))
          {
            if (isNumber(e.oldValue))
            {
              const difference  = e.value - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 4, 1, 2);
              var runningSumValue = newCountDataRange.getValue().toString();

              if (runningSumValue === '')
                runningSumValue = Math.round(e.oldValue).toString();

              (difference > 0) ? 
                newCountDataRange.setValues([[runningSumValue.toString() + ' + ' + difference.toString(), new Date().getTime()]]) : 
                newCountDataRange.setValues([[runningSumValue.toString() + ' - ' + (-1*difference).toString(), new Date().getTime()]]);
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 4, 1, 2);
              var runningSumValue = newCountDataRange.getValue().toString();

              if (isNotBlank(runningSumValue))
                newCountDataRange.setValues([[runningSumValue + ' + ' + Math.round(e.value).toString(), new Date().getTime()]]);
              else
                newCountDataRange.setValues([[Math.round(e.value).toString(), new Date().getTime()]]);
            }
          }
          else if (valueSplit[0].toLowerCase() === 'a' || valueSplit[0].toLowerCase() === 'add') // The number is preceded by the letter 'a' and a space, in order to trigger an "add" operation
          {
            if (valueSplit.length === 3) // An add event with an inflow location
            { 
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNumber(valueSplit[1]))
              {
                newCountValues[0][0] = valueSplit[1]
                valueSplit[2] = valueSplit[2].toUpperCase()

                if (isNumber(newCountValues[0][0])) // New Count is a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                  }
                  else
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][0].toString()]]);
                    }
                  }
                }
                else // New count is Not a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], Math.round(e.oldValue).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], Math.round(e.oldValue).toString()]]);
                    }
                  }

                  SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
                }
              }
              else if (isNumber(valueSplit[2]))
              {
                newCountValues[0][0] = valueSplit[2]
                valueSplit[1] = valueSplit[1].toUpperCase()

                if (isNumber(newCountValues[0][0])) // New Count is a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                  }
                  else
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][0].toString()]]);
                    }
                  }
                }
                else // New count is Not a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], Math.round(e.oldValue).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], Math.round(e.oldValue).toString()]]);
                    }
                  }

                  SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
                }
              }
              else
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), new Date().getTime(), 
                      newCountValues[0][3], newCountValues[0][4].toString()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), new Date().getTime(),
                      newCountValues[0][3], newCountValues[0][4].toString()]])
                }

                SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
              }
            }
            else if (valueSplit.length === 2) // Just an add event with NO inflow location assosiated to the inventory
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 3);
              var newCountValues = newCountDataRange.getValues()
              newCountValues[0][0] = valueSplit[1]

              if (isNumber(newCountValues[0][0])) // New Count is a number
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1]))
                    newCountDataRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), 
                      newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), 
                      parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                }
                else
                {
                  if (isNotBlank(newCountValues[0][1]))
                    newCountDataRange.setNumberFormat('@').setValues([[newCountValues[0][0], 
                      newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[newCountValues[0][0], newCountValues[0][0].toString(), new Date().getTime()]])
                }
              }
              else // New count is Not a number
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), new Date().getTime()]])
                }

                SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
              }
            }
          }
          else if (isNumber(valueSplit[0])) // The first split value is a number and the other is an inflow location
          {
            valueSplit[1] = valueSplit[1].toUpperCase()

            if (isNumber(e.oldValue))
            {
              const difference  = valueSplit[0] - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues();

              if (newCountValues[0][1] === '')
                newCountValues[0][1] = Math.round(e.oldValue).toString();

              if (difference > 0)
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], difference.toString()]]);
              }
              else
              { 
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], difference.toString()]]);
              }
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNotBlank(newCountValues[0][1]))
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], valueSplit[0].toString()]]);
              }
              else
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], valueSplit[0].toString()]]);
              }
            }
          }
          else if (isNumber(valueSplit[1])) // The first split value is an inflow location and the other is a number
          {
            valueSplit[0] = valueSplit[0].toUpperCase()

            if (isNumber(e.oldValue))
            {
              const difference  = valueSplit[1] - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues();

              if (newCountValues[0][1] === '')
                newCountValues[0][1] = Math.round(e.oldValue).toString();

              if (difference > 0)
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], difference.toString()]]);
              }
              else
              { 
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], difference.toString()]]);
              }
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNotBlank(newCountValues[0][1]))
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], valueSplit[1].toString()]]);
              }
              else
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], valueSplit[1].toString()]]);
              }
            }
          }
          else // New value is not a number
          {
            const runningSumRange = sheet.getRange(row, 4);
            const runningSumValue = runningSumRange.getValue().toString();

            if (isNumber(e.oldValue))
            {
              if (isNotBlank(runningSumValue))
                runningSumRange.setNumberFormat('@').setValue(runningSumValue + ' + ' + NaN.toString())
              else
                runningSumRange.setNumberFormat('@').setValue(Math.round(e.oldValue).toString())
            }

            SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
          }
        }
        else // New value IS blank
          sheet.getRange(row, 4, 1, 4).setValues([['', '', '', '']]); // Clear the running sum and last counted time
      }
      else
      {
        if (isNumber(e.value))
          sheet.getRange(row, 4, 1, 2).setNumberFormats([['@', '#']]).setValues([[e.value, new Date().getTime()]])
        else
        {
          const inflowData = e.value.split(' ');

          if (isNumber(inflowData[0]))
            sheet.getRange(row, 3, 1, 5).setNumberFormats([['#', '@', '#', '@', '#']]).setValues([[inflowData[0], inflowData[0], new Date().getTime(), inflowData[1].toUpperCase(), inflowData[0]]])
          else if (isNumber(inflowData[1]))
            sheet.getRange(row, 3, 1, 5).setNumberFormats([['#', '@', '#', '@', '#']]).setValues([[inflowData[1], inflowData[1], new Date().getTime(), inflowData[0].toUpperCase(), inflowData[1]]])
          else
            SpreadsheetApp.getUi().alert("The quantity you entered is not a number.");
        }
      }
    }
  }
}