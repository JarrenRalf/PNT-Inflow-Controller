/* Maybe create an onOpen function that prompts the user to import the inFlow stock levels so that the most update values can be used!
 */

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
    else if (inFlowImportType === 'ProductDetails')
      return downloadInflowProductDetails()
    else if (inFlowImportType === 'PurchaseOrder')
      return downloadInflowPurchaseOrder()
    else if (inFlowImportType === 'SalesOrder')
      return downloadInflowSalesOrder()
    else if (inFlowImportType === 'StockLevels')
      return downloadInflowStockLevels()
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
function onEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === "Item Search") // Check if the user is searching for an item or trying to marry, unmarry or add a new item to the upc database
      search(e, spreadsheet, sheet);
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
                                                     new Array(numCols).fill(new Array(numCols).fill('white'));
    sheet.getRange(3, 1, numRows, numCols).setBackgrounds(colours).clearContent()
  }
}

/**
 * This function moves the selected values from the current sheet to the destination sheet.
 * 
 * @param  {Sheet}    sheet    : The sheet that the selected items are being moved to.
 * @param {Boolean} isTransfer : The sheet that the selected items are being moved to.
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
    }

    // Move the item values to the destination sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, items.length, items[0].length).setNumberFormat('@').setBackgrounds(colours).setValues(items).activate(); 
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
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
 * This function takes the array of data on the Moncton's inFlow Item Quantities page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowBarcodes()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Moncton's inFlow Item Quantities");
  const upcDatabase = spreadsheet.getSheetByName('UPC Database');
  const upcs = upcDatabase.getSheetValues(2, 1, upcDatabase.getLastRow() - 1, 3)
  const data = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, 1).map(item => {
    item.push('');
    upcs.map(upc => {
      if (upc[2] === item[0])
        item[1] += upc[0] + ','
    })
    return item;
  })

  for (var row = 0, csv = "Name,Barcode\r\n"; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_ProductDetails.csv');
}

/**
 * This function takes three arguments that will be used to create a csv file that can be downloaded from the Browser.
 * 
 * @param {String} sheetName  : The name of the sheet that the data is coming from
 * @param {String} csvHeaders : The header names of the csv file
 * @param {String} fileName   : The name of the csv file that will be produced
 * @param {Number} excludeCol : The number of columns at the end of the data that provide information to the user, but do not need to be imported into inFlow
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflow(sheetName, csvHeaders, fileName, excludeCol)
{
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn() - excludeCol)

  if (sheetName === 'Product Details')
  {
    const name = data[0].indexOf('Name')
    const description = data[0].indexOf('Description')
    data = data.filter(header => !header.some(element => element.toString().includes('\n'))).map(header => {
      header[name] = "\"" + header[name] + "\"";
      header[description] = "\"" + header[description] + "\"";

      return header;
    })
  }

  for (var row = 0, csv = csvHeaders; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile(fileName);
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
  const csvHeaders = "Name,Category,ItemType,Description,BarCode,ReorderPoint,ReorderQuantity,Remarks,NOTES,Barcode,IsActive,PicturePath\r\n";
  const fileName = 'inFlow_ProductDetails.csv';
  const numColsToExclude = 0;
  
  return downloadInflow(sheetName, csvHeaders, fileName, numColsToExclude)
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
 * This function handles the imported inFlow Stock Levels and converts it into 
 * 
 * @param {String[][]}    values    : The values of the inFlow Stock Levels
 * @param {Spreadsheet} spreadsheet : The active Spreadsheet
 * @author Jarren Ralf
 */
function importStockLevels(values, spreadsheet, startTime)
{
  if (arguments.length !== 3)
    startTime = new Date().getTime();

  const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
  const numRows = values.length;
  const formats = new Array(numRows - 1).fill(['@', '@', '#', '@'])
  const uniqueItems = []
  formats.unshift(['@', '@', '@', '@']) // Header row

  values = values.map(col => {

    if (!uniqueItems.includes(col[0])) // Count the unique number of items in inFlow
      uniqueItems.push(col[0])

    return [col[0], col[1], col[4], col[3]] // Item, Location, Quantity, Serial (Remove the Sublocation column)
  }); 

  inventorySheet.getRange(1, 2, 1, 3).clearContent() // Clear number of items and timestamp
    .offset(1, -1, inventorySheet.getMaxRows(), 4).clearContent() // Clear the previous inventory
    .offset(0, 0, numRows, 4).setNumberFormats(formats).setValues(values) // Set the updated inventory
    .offset(-1, 1, 1, 3).setValues([[uniqueItems.length, (new Date().getTime() - startTime)/1000 + ' s', Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM HH:mm')]])
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
  const inflowData = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 3).filter(item => item[0].split(" - ").length > 4)
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
          description[0].split(' - ', 1)[0] === values[i][sku].substring(0, 4) + values[i][sku].substring(5, 9) + values[i][sku].substring(10))
          
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
  const inflowData = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 1).filter(item => item[0].split(" - ").length > 4)
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
          description.split(' - ', 1)[0] === values[i][sku].substring(0, 4) + values[i][sku].substring(5, 9) + values[i][sku].substring(10))

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

  if (row == rowEnd && (colEnd == null || colEnd == 3 || col == colEnd)) // Check and make sure only a single cell is being edited
  {
    if (row === 1 && col === 2) // Check if the search box is edited
    {
      const startTime = new Date().getTime();
      const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
      const functionRunTimeRange = sheet.getRange(2, 1, 2);   // The range that will display the runtimes for the search and formatting
      const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 8); // The entire range of the Item Search page
      const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                      // Clear the formatting of the range of the search box
        .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
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
        var numSearchWords, UoM;

        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {
          const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
          const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);

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
                      UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                      output.push([UoM, ...data[i]]);
                      break loop;
                    }
                  }
                  else if (searches[j][k][searches[j][k].length - 1] === '_' && data[i][1].toString().toLowerCase()[0] === searches[j][k][0])
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      UoM = data[i][0].toString().split(' - ')
                      UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                      output.push([UoM, ...data[i]]);
                      break loop;
                    }
                  }
                  else if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      UoM = data[i][0].toString().split(' - ')
                      UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm
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
                        UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                      UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
        }
        else // The word 'not' was found in the search string
        {
          const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
          const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
          const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);

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
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
                            UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

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
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 16, textStyle).build();
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
        const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
        const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\n\nPlease try again.").setTextStyle(0, 14, textStyle).build();
        searchResultsDisplayRange.setRichTextValue(message);
      }

      functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " s");
      spreadsheet.toast('Searching Complete.')
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