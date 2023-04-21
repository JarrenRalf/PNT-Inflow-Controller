/* Maybe create an onOpen function that prompts the user to import the inFlow stock levels so that the most update values can be used!
 */

const inflow_conversions = {
  '10010021FT - WEB: 210/60x3-1/4"X100md X200FM Body #21 -  - Twisted Tarred Nylon - FOOT': 1200,
  '10100027 - WEB: 210/27x1-1/8"x200MDx105FMx235# -  - Twisted Tarred Nylon - POUND': 235, 
  '101021027118 - WEB: 210/27x1-1/8"x100MDx 105FMS -  - Twisted Tarred Nylon - POUND': 226, 
  '10110096 - WEB: 210/96 (6x16) x3"x100MDx50FMx230lbs -  - Cargo/Barrier - POUND': 230, 
  '10120495FOOT - WEB: 210/224x3"x100MDxfoot ) #14x16 -  - Braided Tarred Nylon - FOOT': 150, 
  '10210096 - WEB: 210/96x3-5/8"x25MDx100 FMS Braid k -  - Braided Tarred Nylon - POUND': 96,
  '10500027FT - WEB: 210/27x 2"x400MDx foot GOLF -  - Golf - FOOT': 600, 
  '10500030 - WEB: 3MM Braided  Knotted PE X4"X 100MD -  - Golf - POUND': 285,
  '10500128 - WEB: 210/128x2"x50MDx100FMx250LBS - North Pacific - Hockey/Lacrosse - POUND': 250,
  '10500144 - Black Cod Web 210/144 x 3in x 28md x 200 -  - Web - Miscellaneous - POUND': 375, 
  '10500360 - WEB: #36 x 3"x34MD BROWN HD ACRYLIC -  - Golf - POUND': 300, 
  '10501001FT - WEB: PNT BLACKBIRD 15mm Sq x 2m deep -  - Golf - FOOT': 328.084, 
  '10503000 - WEB: #30 x 2"x50MD BLACK HD ACRYLIC COAT -  - Golf - POUND': 300, 
  "10503600 - VEXAR L36 WEB for CRAB CAGE  (100'/ROLL) -  - Golf - FOOT": 100,
  '10710010FT - WEB: 210/10x1/2"x800MDx100FMx235# RACHL -  - Raschel Knotless - FOOT': 600, 
  '10782109038 - Rachel Black We 210/9 X 3/8" X465MDX 900 -  - Raschel Knotless - POUND': 235,
  '24400000 - BLACK RUBBER MATTING RIBBED    3 \' WIDE - ERIKS - Mats & Tables - FOOT': 225,
  '26014025 - GRADE 43 HIGH TEST GALV CHAIN 1/4" -  - Chain - FOOT': 500,
  '21000001 - CHAIN: PROOF COIL 1/4" Hot Dipped Galv - VANGUARD - Chain - FOOT': 500,
  '21000003 - CHAIN: PROOF COIL 3/8" Hot Dipped Galv - VANGUARD - Chain - FOOT': 400, 
  '21000004 - CHAIN: PROOF COIL 1/2" Hot Dipped Galv - VANGUARD - Chain - FOOT': 200
}

/**
 * This function is run when an html web app is launched. In our case, when the modal dialog box is produced at 
 * the point a user has clicked the Download inFlow Pick List button inorder to produce the csv file.
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
    else if (inFlowImportType === 'PurchaseOrder')
      return downloadInflowPurchaseOrder()
    else if (inFlowImportType === 'SalesOrder')
      return downloadInflowPickList()
    else if (inFlowImportType === 'StockLevels')
      return downloadInflowStockLevels()
    
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
        else if (values[0].includes('PreferredCarrier'))
          updateInFlowVendorList(values, spreadsheet);
        else if (values[0].includes('DefaultPricingScheme'))
          updateInFlowCustomerList(values, spreadsheet);
        else if (values[0].includes('Sublocation')) 
          importStockLevels(values, spreadsheet);

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
      spreadsheet.toast('Searching...')
      const startTime = new Date().getTime();
      const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
      const functionRunTimeRange = sheet.getRange(2, 1, 2);   // The range that will display the runtimes for the search and formatting
      const searchWords = sheet.getRange(1, 2, 1, 2).clearFormat()                                      // Clear the formatting of the range of the search box
        .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
        .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
        .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
        .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
        .getValue().toString().toLowerCase().split(/\s+/);                                              // Split the search string at whitespacecharacters into an array of search words

      const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 2, 8); // The entire range of the Item Search page

      if (isNotBlank(searchWords[0])) // If the value in the search box is NOT blank, then compute the search
      {
        const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
        const data = inventorySheet.getSheetValues(3, 1, inventorySheet.getLastRow() - 2, 4);
        const numSearchWords = searchWords.length - 1; // The number of search words - 1
        const output = [];
        var UoM;

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          for (var j = 0; j <= numSearchWords; j++) // Loop through each word in the User's query
          {
            if (data[i][0].toString().toLowerCase().includes(searchWords[j])) // Does the i-th item description contain the j-th search word
            {
              if (j === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                UoM = data[i][0].toString().split(' - ')
                UoM = (UoM.length >= 5) ? UoM[UoM.length - 1] : ''; // If the items is in Adagio pull out the unit of measure and put it in the first columm

                output.push([UoM, ...data[i]]);
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
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
  const inFlowValues = sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 1).flat();
  var isThereNewValuesToAdd = false, newValues = [];

  for (var i = startingIndex; i < values.length; i++) // Start at 2 because we ignore the header and the first customer which is '.'
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

/**
 * This function identifies the items in the inFlow database that are in location DOCK and puts then on the DOCK sheet.
 * 
 * @author Jarren Ralf
 */
function updateItemsOnDock()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const inventorySheet = spreadsheet.getSheetByName('Inventory')
  const sheet = spreadsheet.getSheetByName('DOCK')
  var serialNum;

  const itemsOnDock = inventorySheet.getSheetValues(4, 1, inventorySheet.getLastRow() - 3, 4).filter(location => location[1] === 'DOCK').map(item => {
    serialNum = item.pop()
    item[1] = item[2] // Move the quantity to column 2
    item[2] = '' // Make the third column blank

    if (isNotBlank(serialNum))
      item[0] = item[0].toString() + '\n\t' + serialNum;
    
    return item
  })

  if (itemsOnDock.length !== 0)
    spreadsheet.getSheetByName('DOCK').getRange(4, 1, itemsOnDock.length, 3).setValues(itemsOnDock)

  sheet.getRange(1, 3).setFormula('=' + itemsOnDock.length + '-B1')
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
 * This function allows Adrian to select items on the INVENTORY page and move them to the SearchData page which will cause 
 * them to now be available for search on Item Search sheet. This function effectively removes "No TS" for the current day.
 * 
 * @author Jarren Ralf
 */
function addItemsToSearchData()
{
  const NUM_COLS = 6;
  const spreadsheet = SpreadsheetApp.getActive();
  const searchDataSheet = spreadsheet.getSheetByName("SearchData");
  const sheet = spreadsheet.getActiveSheet(); // Assumed to be the inventory page (because that is where the button for this function lives)
  const startRow = searchDataSheet.getLastRow() + 1; // The bottom of the list
  var firstRows, row, lastRow, rows = [];
  [firstRows, numRows, row, lastRow] = copySelectedValues(searchDataSheet, startRow, NUM_COLS); // Move items to SearchData
  const totalNumRows = lastRow - row + 1;

  // Determine which rows "No TS" needs to be removed from
  for (var i = 0; i < firstRows.length; i++)
  {
    for (var j = 0; j < numRows[i]; j++)
      rows.push(firstRows[i] - row + j);
  }

  var range = sheet.getRange(row, 9, totalNumRows);
  var values = range.getValues();
  rows.map(row => values[row][0] = '');
  range.setValues(values);
}

/**
 * This function allows the user to select items on the Manual Counts page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData()
{
  const NUM_COLS = 4;
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");

  if (sheet.getSheetName() === 'Manual Scan')
  { 
    const ui = SpreadsheetApp.getUi();
    const barcodeInputRange = sheet.getRange(1, 1);
    const values = barcodeInputRange.getValue().split('\n');

    const response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + values[0] +'.', ui.ButtonSet.OK_CANCEL)
    {
      if (response.getSelectedButton() == ui.Button.OK)
      {
        const item = values[0].split(' - ');
        const upc = response.getResponseText();
        manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 5).setNumberFormat('@').setValues([[item[0], upc, item[4], values[0], '']]);
        upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, NUM_COLS).setNumberFormat('@').setValues([[upc, item[4], values[0], values[4]]]);
        barcodeInputRange.activate();
      }
    }
  }
  else if (sheet.getSheetName() === 'Manual Counts' || sheet.getSheetName() === 'Item Search')
  {
    const startRow = upcDatabaseSheet.getLastRow() + 1; // The bottom of the list
    copySelectedValues(upcDatabaseSheet, startRow, NUM_COLS, 'upc', false, [manAddedUPCsSheet]); // Move items to UPC Database and the Manually Added UPCs
    const row = sheet.getActiveRange().getRow();
    populateManualScan(spreadsheet, sheet, row);
  }
}

/**
 * This function allows the user to select items on the Manual Counts page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan. In this case, the item/s in question appear not to be found in the Adagio database.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData_ItemsNotFound()
{
  const NUM_COLS = 4;
  const spreadsheet = SpreadsheetApp.getActive();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
  const inventorySheet = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName("INVENTORY") : spreadsheet.getSheetByName("SearchData");
  const startRow = upcDatabaseSheet.getLastRow() + 1; // The bottom of the list

  copySelectedValues(upcDatabaseSheet, startRow, NUM_COLS, 'upc', false, [manAddedUPCsSheet, inventorySheet], true); // Move items to UPC Database, Manually Added UPCs, and INVENTORY sheets
  spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
}

/**
 * This function takes the user's selected items on the Item Search page of the Richmond spreadsheet and it places those items on the inFlowPick page.
 * 
 * @param {Number} qty : If an argument is passed to this function, it is the quantity that a user is entering on the Order page for the inFlow pick list
 * @author Jarren Ralf
 */
function addToInflowPickList(qty)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = (!isRichmondSpreadsheet(spreadsheet)) ? SpreadsheetApp.openById('1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk').getSheetByName('inFlowPick') : 
                                                                                                                    spreadsheet.getSheetByName('inFlowPick');
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const firstRows = [], lastRows = [], itemValues = [];

  const inflowData = Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString())
    .filter(item => item[0].split(" - ").length > 4).map(descrip => descrip[0])

  if (activeSheet.getSheetName() === 'Item Search')
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < activeRanges.length; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
        lastRows[r] = activeRanges[r].getLastRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 2, lastRows[r] - firstRows[r] + 1, 6);
    }
    
    const     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => ['newRichmondPick', 'Richmond PNT', inflowData.find(description => description === item[0]), '', item[5]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)

    if (row > 3 && lastRow <= activeSheet.getLastRow())
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
  else if (activeSheet.getSheetName() === 'Suggested inFlowPick')
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < activeRanges.length; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
        lastRows[r] = activeRanges[r].getLastRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 1, lastRows[r] - firstRows[r] + 1, 3);
    }
    
    const     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => ['newSuggestedPick', 'Richmond PNT', inflowData.find(description => description === item[2]), item[0], item[2]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)

    if (row > 1 && lastRow <= activeSheet.getLastRow())
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
  else if (activeSheet.getSheetName() === 'Order')
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < activeRanges.length; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 3, activeRanges[r].getLastRow() - firstRows[r] + 1, 7);
    }

    if (isParksvilleSpreadsheet(spreadsheet))
    {
      var inFlowOrderNumber = 'newParksvillePick';
      var inFlowCustomerName = 'Parksville PNT';
    }
    else
    {
      var inFlowOrderNumber = 'newRupertPick';
      var inFlowCustomerName = 'Rupert PNT';
    }

    const row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => [inFlowOrderNumber, inFlowCustomerName, 
                                                    inflowData.find(description => description === item[2]), (qty) ? qty : item[0], item[3].split('): ')[1]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)
    
    if (row > 3)
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
      {
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
        spreadsheet.toast('Item(s) added to inFlow Pick List on the Richmond sheet')
      }
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
}

/**
 * Apply the proper formatting to the Order, Shipped, Received, ItemsToRichmond, Manual Counts, or InfoCounts page.
 *
 * @param {Sheet}   sheet  : The current sheet that needs a formatting adjustment
 * @param {Number}   row   : The row that needs formating
 * @param {Number} numRows : The number of rows that needs formatting
 * @param {Number} numCols : The number of columns that needs formatting
 * @author Jarren Ralf
 */
function applyFullRowFormatting(sheet, row, numRows, numCols)
{
  const SHEET_NAME = sheet.getSheetName();

  if (SHEET_NAME === "InfoCounts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#']);
    sheet.getRange(row, 1, numRows, numCols).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
    sheet.getRange(row, 3, numRows         ).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  else if (SHEET_NAME === "Manual Counts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#', '@', '#', '@', '@']);
    sheet.getRange(row, 1, numRows, numCols).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
    sheet.getRange(row, 3, numRows         ).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(row, 5, numRows,       2).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID) 
                                            .setBorder(null, null, null, null, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
  }
  else if (SHEET_NAME === "Trites Counts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '#.#']);
    sheet.getRange(row, 1, numRows, 3).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
  }
  else
  {
    const   BLUE = '#c9daf8', GREEN = '#d9ead3', YELLOW = '#fff2cc';
    const isItemsToRichmondPage = (SHEET_NAME === "ItemsToRichmond") ? true : false;

    if (isItemsToRichmondPage)
    {
      var      borderRng = sheet.getRange(row, 1, numRows, 8);
      var  shippedColRng = sheet.getRange(row, 6, numRows   );
      var thickBorderRng = sheet.getRange(row, 6, numRows, 3);
      var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
      var horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
      var wrapStrategies = [...Array(numRows)].map(e => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
          SpreadsheetApp.WrapStrategy.CLIP, SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.WRAP]);
    }
    else
    {
      var         borderRng = sheet.getRange(row, 1, numRows, numCols);
      var     shippedColRng = sheet.getRange(row, 9, numRows         );
      var    thickBorderRng = sheet.getRange(row, 9, numRows,       2);
      var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
      var horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(3).fill('center'), 'left', ...new Array(6).fill('center')]);
      var wrapStrategies = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
    }

    borderRng.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
                  .setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setWrapStrategies(wrapStrategies)
                  .setBorder(true, true, true, true,  null, true, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground('white');

    thickBorderRng.setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackground(GREEN);
    shippedColRng.setBackground(YELLOW);

     if (!isItemsToRichmondPage)
       sheet.getRange(row, 7, numRows, 2).setBorder(null,  true,  null,  null,  true,  null, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground(BLUE);
  }
}

/**
 * This function sets the formatting across every sheet in this spreadsheet if called with null parameters. When this function is called by the 
 * formatActiveSheet function (see below), then a singular sheet is formatted.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param   {Sheet[]}      sheets   : The active sheet in an array.
 * @author Jarren Ralf
 */
function applyFullSpreadsheetFormatting(spreadsheet, sheets)
{
  if (arguments.length === 0) // If no arguments are sent to the 
  {
    spreadsheet = SpreadsheetApp.getActive();
    sheets = spreadsheet.getSheets();
  }

  const Store_Name = spreadsheet.getName().split(" ")[1]; // Gets the store name from the name of the spreadsheet
  const STORE_NAME = Store_Name.toUpperCase();            // Makes the store name upper case
  const sheetNames = sheets.map(sheet => sheet.getSheetName());
  const RED = '#ea9999', GREEN = '#b6d7a8', YELLOW = '#ffd666'; // The colours of the order date highlighting
  const today = new Date();
  const      YEAR = today.getFullYear();
  const     MONTH = today.getMonth()
  const       DAY = today.getDate();
  const ONE_WEEK  = new Date(YEAR, MONTH, DAY -  7); // Used to highlight the order dates correctly
  const ONE_MONTH = new Date(YEAR, MONTH, DAY - 31);
  var numHeaders, rowStart, maxRow, lastRow, lastCol, numRows, dataRange, dataValues, descriptionWithHyperlinkRange, descriptionWithHyperlink, fontSizes, fontColours, 
    numberFormats, backgroundColours, horizontalAlignments, wrapStrategies, notesRange, noteBackgroundColours, richTextValues, headerNumberFormats, headerValues, 
    headerBackgroundColours, headerFontColours, headerFontSizes, headerHorizontalAlignments, headerFonts, columnWidths;

  for (var j = 0; j < sheets.length; j++)
  {
    if(sheetNames[j] === "Order" || sheetNames[j] === "Shipped" || sheetNames[j] === "Received" )
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, 10);
        dataRange = sheets[j].getRange(rowStart, 1, numRows,    11);
       dataValues = dataRange.getValues();

      // Set the column widths and the header's row heights
      sheets[j].setRowHeights(1, 2, 65).setRowHeightsForced(3, 1, 2).setFrozenRows(2);
      sheets[j].hideRows(3);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [90, 100, 50, 75, 650, 250, 40, 40, 75, 180, 125, 25, 50][c]);

      headerValues = [['','','','','','','Current Stock','Actual Count','',''],
                      ['Order Date','Entered By:','Qty','UoM','Description','Notes','','','Shipped','Shipment Status'], 
                      ['', '', '', '', '', '', '', '', '', '']];
      headerBackgroundColours = [ '', [...new Array(8).fill('white'), '#fff2cc', '#d9ead3'], new Array(10).fill('white')];

      if (sheetNames[j] === "Order")
      {
        headerValues[0][0] = 'ITEMS ORDERED BY PNT ' + STORE_NAME;
        headerBackgroundColours[0] = new Array(10).fill('#5b95f9');
        headerFontColours = [ new Array(10).fill('white'),  new Array(10).fill('black'), new Array(10).fill('black')];
        descriptionWithHyperlinkRange = sheets[j].getRange(1, 9);
        descriptionWithHyperlink = descriptionWithHyperlinkRange.getRichTextValue();
      }
      else if (sheetNames[j] === "Shipped")
      {
        headerValues[0][0] = 'SHIPPED ITEMS IN TRANSIT TO PNT ' + STORE_NAME;
        headerBackgroundColours[0] = new Array(10).fill('#ffd666');
        headerFontColours = [...Array(numHeaders)].map(e => new Array(10).fill('black'));
      }
      else if (sheetNames[j] === "Received")
      {
        headerValues[0][0] = 'ITEMS RECEIVED INTO PNT ' + STORE_NAME;
        headerBackgroundColours[0] = new Array(10).fill('#8bc34a');
        headerFontColours = [...Array(numHeaders)].map(e => new Array(10).fill('black'));
        descriptionWithHyperlinkRange = sheets[j].getRange(2, 5);
        descriptionWithHyperlink = descriptionWithHyperlinkRange.getRichTextValue();
      }

      // Prepare and set all of the headerRange values
      headerFontSizes = [[30, ...new Array(9).fill(10)], [...new Array(8).fill(14), ...new Array(2).fill(12)], new Array(10).fill(10)];
          headerFonts = [new Array(10).fill('Verdana'), new Array(10).fill('Arial'), new Array(10).fill('Arial')];
      headerRange.setWrap(true).setNumberFormat('@').setBackgrounds(headerBackgroundColours)
        .setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamilies(headerFonts).setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setVerticalAlignment('middle').setHorizontalAlignment('center').setValues(headerValues);

      if (sheetNames[j] === "Received")
      {
        descriptionWithHyperlinkRange.setRichTextValue(descriptionWithHyperlink);
        sheets[j].getRange(3, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Back to Shipped']).build());
        sheets[j].getRange(3, 12).insertCheckboxes().check();
      }
      else
      {
        if (sheetNames[j] === "Order")
        {
          var col = 'B';
          descriptionWithHyperlinkRange.setRichTextValue(descriptionWithHyperlink.copy().setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('white').build()).build()); // White
        }
        else
        {
          var col = 'D';
          sheets[j].getRange(3, 13).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Receive ALL']).build());
        }

        var dataValidationSheet = (sheets.length === 1) ? spreadsheet.getSheetByName("Data Validation") : sheets[sheetNames.indexOf("Data Validation")];
        sheets[j].getRange(3, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dataValidationSheet.getRange('$' + col + '$1:$' + col)).build());
      }

      // Prepare all of the dataRange values and formats
      fontSizes            = [...Array(numRows)].map(e => new Array(11).fill(10));
      fontColours          = [...Array(numRows)].map(e => new Array(11).fill('black'));
      numberFormats        = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
      backgroundColours    = [...Array(numRows)].map(e => [null, 'white', 'white', 'white', 'white', null, '#c9daf8', '#c9daf8', '#fff2cc', '#d9ead3', 'white']);
      horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'center']);
      wrapStrategies       = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
                                   ...new Array(3).fill(SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
      notesRange = sheets[j].getRange(rowStart, 6, numRows); // To preserve the background and text colours of the Notes, we must store that data first
      noteBackgroundColours = notesRange.getBackgrounds();
      richTextValues = notesRange.getRichTextValues();

      if (sheetNames[j] === "Shipped")
      {
        for (var i = 0; i < numRows; i++)
        {
          backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED ); // Highlight the dates correctly

          if (dataValues[i][10] === "via") // Locate the shipping carrier banner and apply the appropriate changes
          {
            fontSizes[i] = new Array(11).fill(14);
            numberFormats[i] = new Array(11).fill('@');
            backgroundColours[i] = new Array(11).fill('#6d9eeb');
            horizontalAlignments[i] = new Array(11).fill('left');
            fontColours[i] = [...new Array(10).fill('white'), '#6d9eeb'];
            sheets[j].getRange(i + 4, 1, 1, 10).merge();
            sheets[j].setRowHeight(i + 4, 40).getRange(i + 4, 1, 1, 11).setBorder(true,true,true,true,false,false);
          }
        }

        sheets[j].getRange(3, 12).setFormula('=ArrayFormula(if(K3:K="via",A3:A,""))');
        sheets[j].hideColumns(12, 2);
      }
      else
      {
        for (var i = 0; i < numRows; i++)
          backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED ); // Highlight the dates correctly
      }

      // Set all of the dataRange values and formats
      dataRange.setFontLine('none').setFontStyle('normal').setFontWeight('bold').setFontFamily('Arial').setFontSizes(fontSizes).setFontColors(fontColours)
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
        .setNumberFormats(numberFormats).setBackgrounds(backgroundColours).setBorder(true, true, true, true, false, true);

      sheets[j].getRange(rowStart, 7, numRows, 2).setBorder(true, true, true, true, true, null);
      sheets[j].getRange(rowStart, 9, numRows, 2).setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

      if (sheetNames[j] !== "Shipped")
        sheets[j].autoResizeRows(rowStart, numRows);
      
      notesRange.setBackgrounds(noteBackgroundColours).setRichTextValues(richTextValues);
    }
    else if (sheetNames[j] === "ItemsToRichmond")
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastRow = sheets[j].getLastRow();
      lastCol = 8;
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, lastCol + 1);
        dataRange = sheets[j].getRange(rowStart, 1,    numRows, lastCol);
       dataValues = dataRange.getValues(); 

      // Set the column widths and the header's row heights
      sheets[j].setRowHeights(1, 2, 65);
      for (var c = 0; c < lastCol + 1; c++)
        sheets[j].setColumnWidth(c + 1, [90, 100, 75, 700, 250, 75, 100, 125, 25][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues = [  ['ITEMS BEING SHIPPED TO PNT RICHMOND', '', '', '', '', '', '', '', 'TRANSFERRED'],
                        ['Order Date','Entered By:','UoM','Description','Notes', 'Shipped','Carrier','Received By', ''], 
                        [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]];
      headerFonts = [new Array(9).fill('Verdana'), new Array(9).fill('Arial'), new Array(9).fill('Arial')];
      headerFontSizes = [[30, ...new Array(8).fill(10)], [...new Array(5).fill(14), ...new Array(4).fill(12)], new Array(9).fill(10)];
      headerFontColours = [new Array(9).fill('white'),  new Array(9).fill('black'), new Array(9).fill('black')];
      headerBackgroundColours = [new Array(9).fill('#5b95f9'), [...new Array(5).fill('white'), '#fff2cc', '#d9ead3', '#d9ead3', ''], new Array(9).fill('white')];
      headerRange.setFontLine('none').setFontStyle('normal').setFontFamilies(headerFonts).setFontSizes(headerFontSizes).setFontWeight('bold').setFontColors(headerFontColours)
        .setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true).setBackgrounds(headerBackgroundColours).setValues(headerValues);

      // Prepare all of the dataRange values and formats
      horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
      wrapStrategies = [...Array(numRows)].map(e => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW),  ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
                                                                          SpreadsheetApp.WrapStrategy.CLIP,       ...new Array(2).fill(SpreadsheetApp.WrapStrategy.WRAP)]);
      backgroundColours = [...Array(numRows)].map(e => [null, 'white', 'white', 'white', null, '#fff2cc', '#d9ead3', '#d9ead3']);
      numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
      notesRange = sheets[j].getRange(rowStart, 5, numRows); // To preserve the background and text colours of the Notes, we must store that data first
      noteBackgroundColours = notesRange.getBackgrounds();
      richTextValues = notesRange.getRichTextValues();

      // Apply the correct highlighting for the dates
      for (var i = 0; i < numRows; i++)
        backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED );

      // Set all of the dataRange values and formats
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('bold').setFontFamily('Arial')
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
        .setNumberFormats(numberFormats).setBackgrounds(backgroundColours).setBorder(true, true, false, false, false, true);
      
      sheets[j].getRange(rowStart, 6, numRows, 3).setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      notesRange.setBackgrounds(noteBackgroundColours).setRichTextValues(richTextValues);
      sheets[j].autoResizeRows(rowStart, numRows);
    }
    else if (sheetNames[j] === "Manual Counts" || sheetNames[j] === "InfoCounts")
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
      
      // Set the header's row heights and sheet's column widths
      for (var r = 0; r < numHeaders; r++)
        sheets[j].setRowHeightsForced(r + 1, 1, [45, 45, 2][r]);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [900, 80, 80, 500, 130, 85, 85][c]);

      numberFormats = [...Array(numHeaders)].map(e => new Array(lastCol).fill('@'));

      if (sheetNames[j] === "Manual Counts")
      {
        spreadsheet.setNamedRange('Completed_ManualCounts', sheets[j].getRange('B1'));
        spreadsheet.setNamedRange('Progress_ManualCounts', sheets[j].getRange('B3'));
        spreadsheet.setNamedRange('Remaining_ManualCounts', sheets[j].getRange('C1'));
        headerValues = [[  Store_Name + ' Manual Counts', '=COUNTA($C$4:$C)', '=COUNTA($A$4:$A)-Completed_ManualCounts', 'Scanning Data', '', 'Inflow Data', ''], 
                        ['Sku# - Description - Category - UoM', 'Current Count', 'New Count', 'Running Sum', 'Last Scan Time (ms)', 'Location', 'Quantity'], 
                        ['', '=Completed_ManualCounts&\"/\"&(Completed_ManualCounts + Remaining_ManualCounts)', '', '', '', '', '']];
        sheets[j].hideColumns(4, 4);
        numberFormats[2][6] = '#';
      }
      else
      {
        spreadsheet.setNamedRange('Completed_InfoCounts', sheets[j].getRange('B1'));
        spreadsheet.setNamedRange('Progress_InfoCounts', sheets[j].getRange('B3'));
        spreadsheet.setNamedRange('Remaining_InfoCounts', sheets[j].getRange('C1'));
        headerValues = [['New ' + Store_Name + ' Counts', '=COUNTA($C$4:$C$' + lastRow + ')', '=' + numRows + '-Completed_InfoCounts'], 
                        ['Sku# - Description - Category - UoM', 'Current Count', 'New Count'], 
                        ['', '=Completed_InfoCounts&\"/\"&(Completed_InfoCounts+Remaining_InfoCounts)', '']];
      }

      // Prepare and set all of the headerRange formatting
      headerFontSizes = [new Array(lastCol).fill(18), new Array(lastCol).fill(12), new Array(lastCol).fill(2)];
      headerFontColours = [['black', '#b7b7b7', ...new Array(lastCol - 2).fill('black')], new Array(lastCol).fill('black'), new Array(lastCol).fill('black')]
      headerHorizontalAlignments = [['right', ...new Array(lastCol - 1).fill('center')], ['right', ...new Array(lastCol - 1).fill('center')], ['right', ...new Array(lastCol - 1).fill('center')]];
      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Verdana').setFontColors(headerFontColours).setFontSizes(headerFontSizes)
        .setWrap(true).setNumberFormats(numberFormats).setVerticalAlignment('middle').setHorizontalAlignments(headerHorizontalAlignments).setBackground('white')
        .setBorder(null, null, null, null, null, true).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setValues(headerValues)

      sheets[j].hideRows(3);

      if (sheetNames[j] === "Manual Counts")
        sheets[j].getRange(3, lastCol).insertCheckboxes().uncheck() // This is the "Add-One" mode for the Manual Scanner
        
      if (numRows > 0)
      {
        // Prepare and set all of the dataRange values and formats
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
        fontColours = [...Array(numRows)].map(e => ['black', '#b7b7b7', ...new Array(lastCol - 2).fill('black')]);
        horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(lastCol - 1).fill('center')]);
        numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#', ...new Array(lastCol - 3).fill('@')]);
        wrapStrategies = [...Array(numRows)].map(e => [SpreadsheetApp.WrapStrategy.OVERFLOW, ...new Array(lastCol - 1).fill(SpreadsheetApp.WrapStrategy.CLIP)]);
        dataRange.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Verdana').setFontColors(fontColours)
          .setBackground('white').setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
          .setBorder(true, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

        if (sheetNames[j] === "Manual Counts")
          sheets[j].getRange(rowStart, 5, numRows, 2).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID) 
                                                     .setBorder(null, null, null, null, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
                                                     .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
          
        sheets[j].getRange(1, 3, numHeaders + numRows).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                                      .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

        if (maxRow > lastRow)
          sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

        sheets[j].autoResizeRows(rowStart, numRows);
      }
      else if (maxRow >= 5)
        sheets[j].deleteRows(5, maxRow - 4) // Leave 1 blank row
    }
    else if (sheetNames[j] === "Item Search")
    {
      if (sheets.length > 1) recentlyCreatedItems(spreadsheet, sheets[j]); // If the full spreadsheet is being formatted, then put the recently created items on the search page
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastCol = sheets[j].getMaxColumns();
      const MAX_NUM_ITEMS = 500;
      numRows = MAX_NUM_ITEMS;
      maxRow = sheets[j].getMaxRows();
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      headerValues = headerRange.getValues();
      
      sheets[j].setRowHeight(1, 150);
      sheets[j].setRowHeight(2,  32);
      sheets[j].setRowHeight(3,  22);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [160, 725, 85, 60, 60, 60, 60, 180][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues[1][3] = 'Current Stock In Each Location';
      headerValues[1][7] = 'Items last updated on';
      headerValues[2][3] = (isRichmondSpreadsheet(spreadsheet)) ? 'Rich' : ((isParksvilleSpreadsheet(spreadsheet)) ? 'Parks' : 'Rupert');
      headerValues[2][4] = (isRichmondSpreadsheet(spreadsheet)) ? 'Parks' : 'Rich';
      headerValues[2][5] = (isRichmondSpreadsheet(spreadsheet) || isParksvilleSpreadsheet(spreadsheet)) ? 'Rupert' : 'Parks';
      headerValues[2][6] = 'Trites';
      headerNumberFormats = [new Array(8).fill('@'), new Array(8).fill('@'), [...new Array(7).fill('@'), 'dd MMM yyyy']]
      headerFontSizes = [[16, 14, 14, ...new Array(5).fill(28)], [12, ...new Array(7).fill(11)], [12, ...new Array(7).fill(11)]];
      headerBackgroundColours = [['#4a86e8', 'white', 'white', ...new Array(5).fill('#4a86e8')], new Array(8).fill('4a86e8'), new Array(8).fill('4a86e8')];
      headerFontColours = [['white', 'black', 'black', ...new Array(5).fill('white')], new Array(8).fill('white'), new Array(8).fill('white')];
      headerRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setBackgrounds(headerBackgroundColours).setNumberFormats(headerNumberFormats).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true, true, true, true, null, null).setValues(headerValues);

      // Format the search box
      sheets[j].getRange(1, 2, 1, 2).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).merge();

      fontSizes = [...Array(numRows)].map(e => [...new Array(lastCol - 1).fill(10), 12]);
      horizontalAlignments = [...Array(numRows)].map(e => ['center', 'left', ...new Array(lastCol - 2).fill('center')]);
      numberFormats = [...Array(numRows)].map(e => ['@', '@', 'dd MMM yyyy', '@', '@', '@', '@', '@']);
      dataRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSizes(fontSizes).setNumberFormats(numberFormats)
        .setVerticalAlignment('middle').setHorizontalAlignments(horizontalAlignments).setWrap(true).setBorder(true, true, true, true, false, false);

      // Apply all of the different borders
      sheets[j].getRange(2, 4, 2, 5).setBorder(null, null, null, null, true, true, '#1155cc', SpreadsheetApp.BorderStyle.SOLID);
      sheets[j].getRange(2, 4, 2, 5).setBorder(true, true, null, null, null, null, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheets[j].getRange(2, 4, 2, 4).setBorder(null, null, null, true, null, null, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheets[j].getRange(1, 4).setFormula('=Remaining_InfoCounts&\" Items left to count on the InfoCounts page.\"')

      if (maxRow > MAX_NUM_ITEMS + 3)
        sheets[j].deleteRows(MAX_NUM_ITEMS + 4, maxRow - MAX_NUM_ITEMS - 3); // Delete the blank rows
    }
    else if (sheetNames[j] === "INVENTORY" || sheetNames[j] === "SearchData" || sheetNames[j] === "Recent")
    {
      numHeaders = (sheetNames[j] === "INVENTORY") ? (isRichmondSpreadsheet(spreadsheet) ? 7 : 9) : 1;
      rowStart = numHeaders + 1;
      maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, lastCol);
        dataRange = sheets[j].getRange(rowStart, 1,    numRows, lastCol);
      columnWidths = isRichmondSpreadsheet(spreadsheet) ? [100, 675, 100, 100, 100, 120, 120, 100, 100] : [100, 675, 100, 100, 100, 100, 120, 100, 100];

      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, columnWidths[c]);

      // Prepare and set all of the headerRange values and formats
      if (sheetNames[j] === "INVENTORY")
      {
        const date = sheets[j].getRange(4, 1).getValue().split(' on ')[1];
        const upcDatabaseLastUpdated = new Date(date);

        if (isRichmondSpreadsheet(spreadsheet))
        {
          sheets[j].setRowHeights(1, numHeaders - 1, 28);
          headerValues = headerRange.getValues();
          headerFontSizes = [ [30, 11, 13, 13, 13, 13, 20, 11, 12], 
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              new Array(lastCol).fill(8)];
          headerFontWeights = [ ['bold',   'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                new Array(lastCol).fill('bold')];
          headerFontColours = [ ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                new Array(lastCol).fill('black')];
          if (upcDatabaseLastUpdated <= ONE_WEEK) headerFontColours[1][3] = 'red';
          headerBackgroundColours = [ ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      new Array(lastCol).fill('#f0f0f0')];
          headerHorizontalAlignments = [['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        new Array(lastCol).fill('center')];
          sheets[j].getRange(3, 3, 3, 5).setFormulas([  ['=Remaining_InfoCounts&\" items on the infoCounts\npage that haven\'t been counted\"', '', '', '', '=Progress_InfoCounts'],
                                                        ['' , '', '', '', ''],
                                                        ['=Remaining_ManualCounts&\" items on the Manual Counts\npage that haven\'t been counted\"', '', '', '', '=Progress_ManualCounts']]);
        }
        else
        {
          sheets[j].setRowHeights(1, numHeaders - 1, 40);
          headerValues = headerRange.getValues();
          headerFontSizes = [ [30, 12, 12, 12, 12, 12, 22, 11, 12], 
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              new Array(lastCol).fill(8)];
          headerFontWeights = [ ['bold',   'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                new Array(lastCol).fill('bold')];
          headerFontColours = [ ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                new Array(lastCol).fill('black')];
          if (upcDatabaseLastUpdated <= ONE_WEEK) headerFontColours[3][0] = 'red';
          headerBackgroundColours = [ ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      new Array(lastCol).fill('#f0f0f0')];
          headerHorizontalAlignments = [['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        new Array(lastCol).fill('center')];
          sheets[j].getRange(3, 7, 4).setFormulas([ ['=COUNTIF(Received_Checkbox,FALSE)'],
                                                    ['=COUNTIF(ItemsToRichmond_Checkbox,FALSE)'],
                                                    ['=COUNTIF(Order_ActualCounts,">=0")'],
                                                    ['=COUNTIF(Shipped_ActualCounts,">=0")']]);
          sheets[j].getRange(7, 1, 2, 7).setFormulas([['=Remaining_InfoCounts&\" items on the infoCounts page that haven\'t been counted\"',      '', '', '', '', '', '=Progress_InfoCounts'],
                                                      ['=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', '', '', '', '', '', '=Progress_ManualCounts']]);
        }
      }
      else
      {
        headerFontWeights = [new Array(lastCol).fill('bold')]
        headerHorizontalAlignments = [new Array(lastCol).fill('center')]
        headerBackgroundColours = [new Array(lastCol).fill('#f0f0f0')];
        headerFontColours = [new Array(lastCol).fill('black')];
        headerFontSizes = [new Array(lastCol).fill(8)];
        sheets[j].hideSheet();
      }
      
      headerRange.setFontFamily('Arial').setFontWeights(headerFontWeights).setFontLine('none').setFontStyle('normal').setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setBackgrounds(headerBackgroundColours).setNumberFormat('@').setHorizontalAlignments(headerHorizontalAlignments).setVerticalAlignment('middle');

      // Prepare and set all of the dataRange values and formats
      horizontalAlignments = [...Array(numRows)].map(e => ['center', 'right', ...new Array(lastCol - 2).fill('center')]);
      dataRange.setFontSize(8).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setBackground('white')
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setNumberFormat('@');

      if (maxRow > lastRow)
        sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

      sheets[j].setFrozenRows(numHeaders);
      sheets[j].autoResizeRows(numHeaders, numRows + 1);
    }
    else if (sheetNames[j] === "UPC Database" || sheetNames[j] === "Manually Added UPCs" || sheetNames[j] === 'UPCs to Unmarry')
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
      maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);

      if (sheetNames[j] === "UPC Database")
      {
        columnWidths = [125, 100, 600, 100]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'center', 'left', 'center']);
      }
      else if (sheetNames[j] === "Manually Added UPCs")
      {
        columnWidths = [125, 125, 100, 600, 125]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left', 'center', 'left', 'center']);
        sheets[j].hideColumns(5);
      }
      else
      {
        columnWidths = [125, 600]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left']);
      }

      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, columnWidths[c]);

      // Prepare and set all of the headerRange values and formats
      headerRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSize(18).setFontColor('black')
        .setBackground('white').setNumberFormat('@').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

      if (numRows > 0)
      {
        // Prepare and set all of the dataRange values and formats
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol).setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial')
          .setBackground('white').setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setNumberFormat('@');

        if (maxRow > lastRow)
          sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows
      }

      sheets[j].hideSheet();
    }
    else if (sheetNames[j] === "Count Log")
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
      
      // Set the header's row heights and sheet's column widths
      sheets[j].setRowHeight(1, 30);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [150, 1000, 100, 100][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues = [["SKU", "Description", "Sheet", "Date"]];
      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black').setFontSize(14)
        .setWrap(true).setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('white').setBorder(false, false, false, false, false, false);

      // Prepare and set all of the dataRange values and formats
      dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      numberFormats = [...Array(numRows)].map(e => [...new Array(lastCol - 1).fill('@'), 'dd MMM yyyy']);
      horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left', 'center', 'right']);
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setBorder(false, false, false, false, false, false);

      if (maxRow > lastRow)
        sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

      sheets[j].autoResizeRows(rowStart, numRows).hideSheet();
    }
    else if (sheetNames[j] === "Data Validation")
      sheets[j].hideSheet().getDataRange().setFontSize(20).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    else if (sheetNames[j] === 'Manual Scan' || sheetNames[j] === 'Item Scan')
      (sheetNames[j] === 'Manual Scan') ? 
        sheets[j].getRange(1, 1, 1, 2).setNumberFormats([['@', '#.#']]).setFontSize(25).setFontLine('none').setFontStyle('none').setFontWeight('normal').setFontFamily('Arial')
          .setFontColor('black').setBackground('white').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) :
        sheets[j].getRange(1, 1).setNumberFormat('@').setFontSize(25).setFontLine('none').setFontStyle('none').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
          .setBackground('white').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    else if (sheetNames[j] === "inFlowPick" || sheetNames[j] === "Suggested inFlowPick" || sheetNames[j] === "Moncton's inFlow Item Quantities")
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);

      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('white').setFontSize(16)
        .setWrap(true).setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('#f1c232')

      if (sheetNames[j] === "Suggested inFlowPick")
      {
        var richTextValue = SpreadsheetApp.newRichTextValue().setText('Adagio Quantity (Trites + Moncton)')
          .setTextStyle(0, 16, SpreadsheetApp.newTextStyle().setFontSize(16).build())
          .setTextStyle(16, 34, SpreadsheetApp.newTextStyle().setFontSize(14).build())
          .build()
        headerRange.offset(0, 4, 1, 1).setRichTextValues([[richTextValue]])
      }

      // Prepare and set all of the dataRange values and formats
      dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    }
    else 
      sheets[j].hideSheet()
  }
}

/**
 * This function clears the inFlow pick list.
 * 
 * @author Jarren Ralf
 */
function clearInflowPickList()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const numRows = sheet.getLastRow() - 2

  if (numRows > 0)
    SpreadsheetApp.getActiveSheet().getRange(3, 1, numRows, 5).clearContent()
}

/**
 * This function clears all of the manual counts that have been completed, but leaves the ones that have a blank in the counts column.
 * 
 * @author Jarren Ralf
 */
function clearManualCounts()
{
  const startTime = new Date().getTime();
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Manual Counts');
  const numHeaders = 3;
  const numItems = sheet.getLastRow() - numHeaders;

  if (numItems > 0) // If there are items on the manual counts page
  {
    const numCols = sheet.getLastColumn();
    const rowStart = numHeaders + 1;
    const items = sheet.getSheetValues(rowStart, 1, numItems, numCols);
    const nonCountedItems = items.filter(count => count[2] === '' || count[0].split(' - ', 1)[0] === 'MAKE_NEW_SKU'); // These are the items that have not been counted
    const numRemainingItems = nonCountedItems.length;

    if (numItems !== numRemainingItems) // If there are some items that have been counted, enter this code block
    {
      const numRows = sheet.getMaxRows() - numHeaders;
      sheet.getRange(rowStart, 1, numRows, numCols).clearContent();

      if (numRemainingItems !== 0) // There are some remaining items to count
      {
        sheet.getRange(rowStart, 1, numRemainingItems, numCols).setValues(nonCountedItems);
        sheet.deleteRows(numRemainingItems + rowStart, numRows - numRemainingItems);
      }
      else if (numRows - 1 !== 0) // There are no more items to count
        sheet.deleteRows(rowStart + 1, numRows - 1);
    }
  }

  if (isRichmondSpreadsheet(spreadsheet))
    spreadsheet.getSheetByName('INVENTORY').getRange(5, 3, 1, 7)
      .setValues([[ '=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', null, null, null, 
                    '=Progress_ManualCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
  else
    spreadsheet.getSheetByName('INVENTORY').getRange(8, 1, 1, 9)
      .setValues([[ '=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', null, null, null, null, null, 
                    '=Progress_ManualCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);    
}

/**
 * This function moves the selected values on the sheet to the desired output page (Order, ItemsToRichmond, Manual Counts, and SearchData sheets).
 * 
 * @param {Sheet}   sheet   : The sheet that the selected items are being moved to.
 * @param {Number} startRow : The first row of the target sheet where the selected items will be moved to.
 * @param {Number}  numCols : The number of columns to grab from the item search page and move to the target sheet.
 * @param {Number}  qtyCol  : The column number of which the sheet in particular has the quantity value inputed.
 * @param {Boolean} isInfoCountsPage : A boolean that represents whether the user is on the infoCounts page or not.
 * @param {Sheet[]}  sheets : The additional sheets to post information to.
 * @param {Boolean} isNotFound : Whether the item value being copied to another sheet appears to exist in Adagio or not.
 * @return {[Number[], Number[], Number, Number]} The firstRows and numRows of the active ranges as well as the first and last row that the set of active ranges span.
 * @author Jarren Ralf
 */
function copySelectedValues(sheet, startRow, numCols, qtyCol, isInfoCountsPage, sheets, isNotFound)
{
  if (arguments.length !== 5) isInfoCountsPage = false;
  const isInventoryPage    = qtyCol == undefined;
  const isUpcPage          = qtyCol == 'upc';
  const isOrderPage        = qtyCol == 9;
  const isItemsToRichPage  = qtyCol == 6;
  const isManualCountsPage = qtyCol == 3 && !isInfoCountsPage;
  
  var  activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  var firstRows = [], lastRows = [], numRows = [];
  var itemValues = [[[]]];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
     lastRows[r] = activeRanges[r].getLastRow()
  }
  
  var     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  var lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
  var finalDataRow = activeSheet.getLastRow() + 1;
  var numHeaders = (isInventoryPage) ? 9: 3;

  var col = (isManualCountsPage) ? 2 : 1; // Set the column of the item selection based on whether we are doing manual counts or item transfers
  var startCol = (isOrderPage) ? 4 : ( (isItemsToRichPage) ? 3 : 1 ); // Set the start column of the range destination based on whether we are doing manual counts or item transfers

  if (row > numHeaders && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
  {   
    for (var r = 0; r < activeRanges.length; r++)
    {
         numRows[r] = lastRows[r] - firstRows[r] + 1;
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], col, numRows[r], numCols);
    }
    
    var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
    var numItems = itemVals.length;

    if (isInventoryPage) // Removing the "No TS" from items on the inventory page and moving them to the SearchData
    {
      itemVals.map(u => u.splice(2, 0, null)); // Add a null column to the items, where the Last Counted Date goes
      numCols++;
    }
    else if (isManualCountsPage) // Moving items from the search page to the manual counts page
    {
      itemVals.map(u => u.splice(1, 1)); // Remove the column that contains the last counted on date
      numCols--;
    }
    else if (isOrderPage) // Items that are being transfered from one location to another
      itemVals = itemVals.map(u => [u[0], u[1], (u[6] > 0) ? 'Trites Stock (As of Order Date): ' + u[6] : '', u[3]]); // Replace the column that has the last counted date with Trites Stock
    else if (isItemsToRichPage) // Items that are being transfered from one location to another
      itemVals.map(u => u.splice(2, 1, null)) // Replace the column that has the last counted date with a blank
    else if (isUpcPage)
    {
      const ui = SpreadsheetApp.getUi();
      var response, response2, item, itemJoined, upc, upcTemporaryValues = [], itemTemporaryValues = [];

      if (activeSheet.getSheetName() === 'Manual Counts')
      {
        if (isNotFound)
        {
          itemVals = itemVals.map(() => {

            response = ui.prompt('Item Not Found', 'Please enter a new description:', ui.ButtonSet.OK_CANCEL)
            
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = response.getResponseText().split(' - ');
              item[0] = 'MAKE_NEW_SKU';
              itemJoined = item.join(' - ')
              response2 = ui.prompt('Item Not Found', 'Please scan the barcode for:\n\n' + itemJoined +'.', ui.ButtonSet.OK_CANCEL)

              if (ui.Button.OK === response2.getSelectedButton())
              {
                upc = response2.getResponseText();
                upcTemporaryValues.push(['MAKE_NEW_SKU', upc, item[4], itemJoined])
                itemTemporaryValues.push([item[4], itemJoined, '', ''])
                return [upc, item[4], itemJoined, '']
              }
              else
              {
                itemTemporaryValues.push([null, null, null, null])
                upcTemporaryValues.push([null, null, null, null])
                return [null, null, null, null]
              }
            }
            else
            {
              itemTemporaryValues.push([null, null, null, null])
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });

          sheets[1].getRange(sheets[1].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(itemTemporaryValues);
        }
        else
        {
          itemVals = itemVals.map(u => {
            response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + u[0] +'.', ui.ButtonSet.OK_CANCEL)
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = u[0].split(' - ');
              upc = response.getResponseText();
              upcTemporaryValues.push([item[0], upc, item[4], u[0]])
              return [upc, item[4], u[0], u[1]]
            }
            else
            {
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });
        }
      }
      else // Item Search Page
      {
        if (isNotFound)
        {
          itemVals = itemVals.map(() => {

            response = ui.prompt('Item Not Found', 'Please enter a new description:', ui.ButtonSet.OK_CANCEL)
            
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = response.getResponseText().split(' - ');
              item[0] = 'MAKE_NEW_SKU';
              itemJoined = item.join(' - ')
              response2 = ui.prompt('Item Not Found', 'Please scan the barcode for:\n\n' + itemJoined +'.', ui.ButtonSet.OK_CANCEL)

              if (ui.Button.OK === response2.getSelectedButton())
              {
                upc = response2.getResponseText();
                itemTemporaryValues.push([item[4], itemJoined, '', ''])
                upcTemporaryValues.push(['MAKE_NEW_SKU', upc, item[4], itemJoined])
                return [upc, item[4], itemJoined, '']
              }
              else
              {
                itemTemporaryValues.push([null, null, null, null])
                upcTemporaryValues.push([null, null, null, null])
                return [null, null, null, null]
              }
            }
            else
            {
              itemTemporaryValues.push([null, null, null, null])
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });

          sheets[1].getRange(sheets[1].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(itemTemporaryValues);
        }
        else
        {
          itemVals = itemVals.map(u => {
            response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + u[1] +'.', ui.ButtonSet.OK_CANCEL)
            if (ui.Button.OK === response.getSelectedButton())
            {
              upc = response.getResponseText();
              upcTemporaryValues.push([u[1].split(' - ', 1)[0], upc, u[0], u[1]])
              return [upc, u[0], u[1], u[3]]
            }
            else
            {
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });
        }
      }

      sheets[0].getRange(sheets[0].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(upcTemporaryValues);
    }

    sheet.getRange(startRow, startCol, numItems, itemVals[0].length).setNumberFormat('@').setValues(itemVals); // Move the item values to the destination sheet

    if (!isInventoryPage && !isUpcPage) 
    {
      const nCols = (sheet.getSheetName() === 'Manual Counts') ? 7 : 11;
      applyFullRowFormatting(sheet, startRow, numItems, nCols); // Apply the proper formatting
      sheet.getRange(startRow, qtyCol).activate();              // Go to the quantity column on the destination sheet
      
      // If we are moving items onto the transfer pages, set the ordered date
      if (isOrderPage || isItemsToRichPage)
        dateStamp(startRow, 1, numItems); // Set the ordered date
    }
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');

  return [firstRows, numRows, row, lastRow];
}

/**
* This function creates a dateStamp and places it on the chosen row/s for the give column.
*
* @param {Number}     row      : The  row   number
* @param {Number}     col      : The column number
* @param {Number}   numRows    : *OPTIONAL* The number of rows
* @param {Sheet}     sheet     : *OPTIONAL* The destination sheet
* @param {String} customFormat : *OPTIONAL* The date / time format
* @return {Date}  timeNow : Returns the formated date dateStamp
* @author Jarren Ralf
*/
function dateStamp(row, col, numRows, sheet, customFormat)
{
  // If the function is sent only two arguments, namely the row and column, then set the dateStampRange appropriately
  var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();             // set timezone
  var dateStampFormat = (arguments.length === 5) ? customFormat : 'dd MMM yyyy';  // set dateStamp format
  var today = new Date();                                                         // Date object representing today's date
  var timeNow = Utilities.formatDate(today, timeZone, dateStampFormat);           // Set variable for current time string

  if (row !== undefined) // If the row value is defined, then print the timestamp in the appropriate place
  {
    if (arguments.length !== 4) sheet = SpreadsheetApp.getActiveSheet();
    var dateStampRange = (arguments.length == 2) ? sheet.getRange(row, col) : sheet.getRange(row, col, numRows); 
    (col === 1) ? dateStampRange.setBackground('#b6d7a8').setValue(timeNow) : dateStampRange.setValue(timeNow);
  }

  return timeNow;
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
 * This function takes the array of data on the inFlowPick page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowPickList()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName('Sales Order');
  const data = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn() - 1)

  for (var row = 0, csv = "OrderNumber,Customer,ItemName,ItemQuantity\r\n"; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_SalesOrder.csv');
}

/**
 * This function takes the array of data on the inFlowPick page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowPurchaseOrder()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Purchase Order');
  const data = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn())

  for (var row = 0, csv = "OrderNumber,Vendor,ItemName,ItemQuantity,OrderRemarks,AmountPaid,ItemUnitPrice,ItemSubtotal\r\n"; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_PurchaseOrder.csv');
}

/**
 * This function takes the array of data on the Manual Counts page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowStockLevels()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName('Manual Counts');
  const data = [];
  var loc, qty, i;

  sheet.getSheetValues(4, 1, sheet.getLastRow() - 3, sheet.getLastColumn()).map(item => {
    loc = item[5].split('\n')
    qty = item[6].split('\n')

    if (loc.length === qty.length) // Make sure there is a location for every quantity and vice versa
      for (i = 0; i < loc.length; i++) // Loop through the number of inflow locations
        if (isNotBlank(loc[i]) && isNotBlank(qty)) // Do not add the data to the csv file if either the location or the quantity is blank
          data.push([item[0], loc[i], qty[i]])

  })

  for (var row = 0, csv = "Item,Location,Quantity\r\n"; row < data.length; row++)
  {
    for (var col = 0; col < data[row].length; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < data.length - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_StockLevels.csv');
}

/**
 * This function formats the active sheet only by calling the applyFullSpreadsheetFormatting function with the active sheet as a parameter.
 * 
 * @author Jarren Ralf
 */
function formatActiveSheet()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetArray = [spreadsheet.getActiveSheet()];
  applyFullSpreadsheetFormatting(spreadsheet, sheetArray);
}

/**
 * This function generates a list of items in the inFlow inventory system that based on the corresponding Adagio inventory values, should be picked and 
 * brought to Moncton street.
 * 
 * @author Jarren Ralf
 */
function generateSuggestedInflowPick()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const suggestedValuesSheet = spreadsheet.getSheetByName("Moncton's inFlow Item Quantities");
  const suggestInflowPickSheet = spreadsheet.getSheetByName('Suggested inFlowPick');
  const numSuggestedItems = suggestedValuesSheet.getLastRow() - 1;
  const suggestedValues = suggestedValuesSheet.getSheetValues(2, 1, numSuggestedItems, 3);
  const inventorySheet = spreadsheet.getSheetByName("INVENTORY");

  Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString()).map(item =>{
    if (item[0].split(" - ").length > 4) // If there are more than 4 "space-dash-space" strings within the inFlow description, then that item is recognized in Adagio 
    {
      for (var i = 0; i < suggestedValues.length; i++)
        if (suggestedValues[i][0] == item[0]) // The ith item of the suggested inFlowPick page was found in the inFlow csv, therefore break the for loop
          break;

      if (i === suggestedValues.length)
        suggestedValues.push([item[0], '', '']) // If there is an item in inFlow but not on the suggested inFlowPick page, then add it
    }
  })

  if (suggestedValues.length > numSuggestedItems) // Items from the inFlow csv have been added to the suggested inFlowPick page
  {
    suggestedValues.sort((a, b) => a[0].localeCompare(b[0])); // Sort the items by the description
    suggestedValuesSheet.getRange(2, 1, suggestedValues.length, 3).setValues(suggestedValues)
  }
  
  const output = inventorySheet.getSheetValues(8, 2, inventorySheet.getLastRow() - 7, 6).map(e => {

    if (isNotBlank(e[5]) && Number(e[2]) >= Number(e[5])) // Trites Inventory Column is not blank and the Adagio inventory is greater than or equal to inFlow inventory 
    {
      for (var i = 0; i < suggestedValues.length; i++)
      {
        if (suggestedValues[i][0] == e[0]) // Match the SKUs of the suggestValues list and the available inFlow inventory
        {
          const monctonStock = Number(e[2] - e[5]); // The stock levels in moncton street (Adagio - inFlow)

          if (Number(e[2]) <= Number(suggestedValues[i][1])) // If Moncton plus Trites less than or equal to the suggested quantity, then bring back everything from Trites to Moncton
            return [e[0], e[5], e[5], monctonStock, e[2]] // Bring back ALL trties stock
          else if (monctonStock < Number(suggestedValues[i][1])) // Moncton stock is less than the suggest amount for Moncton
          {
            const orderQty = Number(suggestedValues[i][1] - monctonStock);

            if (suggestedValues[i][2]) // If we try and pick this item in multiples of 'n' items, such as picking bait jars by the case and hence as multiples of 100 pcs
            {
              if (orderQty > Number(suggestedValues[i][2])) // Order quantity is greater then the number of items that we want to bring this SKU back in mutiples of
              {
                const suggestedAmount = Math.floor(orderQty/Number(suggestedValues[i][2]))*Number(suggestedValues[i][2])

                // If the suggestedAmount is greater than the Trites inventory, then bring back all of the Trites inventory, otherwise bring back the suggestedAmount
                return (suggestedAmount >= Number(e[5])) ? [e[0], e[5], e[5], monctonStock, e[2]] : [e[0], suggestedAmount, e[5], monctonStock, e[2]]
              }
            }
            else // If the orderQty is greater than the Trites inventory, then bring back all of the Trites inventory, otherwise bring back the orderQty
              return (orderQty >= Number(e[5])) ? [e[0], e[5], e[5], monctonStock, e[2]] : [e[0], orderQty, e[5], monctonStock, e[2]]
          }
        }
      }
    }

    return false // Not an available item at Trites
  }).filter(f => f) // Remove the unavailable items

  const numItems = output.length;
  const range = suggestInflowPickSheet.getRange(2, 1, suggestInflowPickSheet.getMaxRows(), 5).clearContent()
  
  if (numItems > 0)
  {
    output.sort((a,b) => a[3] - b[3]) // Sort list by the quantity in Moncton street because if Moncton has 0, then those items are the most important to pick from Trites
    range.offset(0, 0, output.length, 5).setValues(output)
  }
}

/**
* This function calculates the day that New Years Day, Canada Day, Remembrance Day, and Christmas Day, is observed on for the giving year and month. 
*
* @param  {Number}  year The given year
* @param  {Number} month The given month
* @return {Number}   day The day of the Holiday for the particular year and month
* @author Jarren Ralf
*/
function getDay(year, month)
{
  const JANUARY  =  0;
  const JULY     =  6;
  const NOVEMBER = 10;
  const DECEMBER = 11;
  const SUNDAY   =  0;
  const SATURDAY =  6;
  
  if (month == JANUARY || month == JULY || month == DECEMBER) // New Years Day or Canada Day or Christmas Day
  {
    var holiday = (month == DECEMBER) ? new Date(year, month, 25) : new Date(year, month);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() + 2 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls over to the following Monday
  }
  else if (month == NOVEMBER) // Remembrance Day
  {
    var holiday = new Date(year, month, 11);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() - 1 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls back to Friday, or over to Monday
  }
  
  return day;
}

/**
* Gets the last row number based on a selected column range values
*
* @param {Object[][]} range Takes a 2d array of a single column's values
* @returns {Number} The last row number with a value. 
*/
function getLastRowSpecial(range)
{
  var rowNum = 0;
  var blank = false;
  
  for (var row = 0; row < range.length; row++)
  {
    if(range[row][0] === "" && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (isNotBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
* This function calculates what the nth Monday in the given month is for the given year. This function is used for determining the holidays in a given year.
* Victoria Day is an exception to the rule since it is defined to be the preceding Monday before May 25th. The fourth Boolean parameter handles this scenario.
*
* @param  {Number}              n : The nth Monday the user wants to be calculated
* @param  {Number}          month : The given month
* @param  {Number}           year : The given year
* @param  {Boolean} isVictoriaDay : Whether it is Victoria Day or not
* @return {Number} The day of the month that the nth Monday is on (or that Victoria Day is on)
* @author Jarren Ralf
*/
function getMonday(n, month, year, isVictoriaDay)
{
  const NUM_DAYS_IN_WEEK = 7;
  var firstDayOfMonth = new Date(year, month).getDay();
  
  if (isVictoriaDay)
    n = (firstDayOfMonth % (NUM_DAYS_IN_WEEK - 1) < 2) ? 4 : 3; // Corresponds to the Monday preceding May 25th 
  
  return ((NUM_DAYS_IN_WEEK - firstDayOfMonth + 1) % NUM_DAYS_IN_WEEK) + NUM_DAYS_IN_WEEK*n - 6;
}

/**
* This function calculated and returns the runtime of a particular script.
*
* @param  {Number} startTime : The start time that the script began running at represented by a number in milliseconds
* @return {String}  runTime  : The runtime of the script represented by a number followed by the unit abbreviation for seconds.
* @author Jarren Ralf
*/
function getRunTime(startTime)
{
  return (new Date().getTime() - startTime)/1000 + ' s';
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
 * This function checks if every value in the import multi-array is blank, which means that the user has
 * highlighted and deleted all of the data.
 * 
 * @param {Object[][]} values : The import data
 * @return {Boolean} Whether the import data is deleted or not
 * @author Jarren Ralf
 */
function isEveryValueBlank(values)
{
  return values.every(arr => arr.every(val => val == '') === true);
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
* This function checks if today's date is a stat holiday or not.
*
* @param {Date} today : Today's date
* @return {Boolean} Returns a true boolean if today is not a stat and false otherwise.
* @author Jarren Ralf
*/
function isNotStatHoliday(today)
{
  today = today.getTime();
  const JAN =  0, FEB =  1, MAY =  4, JUL =  6, AUG =  7, SEP =  8, OCT =  9, NOV = 10, DEC = 11;
  const YEAR = new Date().getFullYear(); // An integer corresponding to today's year
  const ONE_DAY = 24*60*60*1000;
  var MMM, DD;
  [MMM, DD] = calculateGoodFriday(YEAR);

  const statHolidays = [new Date(YEAR, JAN, getDay(YEAR, JAN)),          // New Year's Day
                        new Date(YEAR, FEB, getMonday(3, FEB, YEAR)),    // Family Day
                        new Date(YEAR, MMM, DD),                         // Good Friday
                        new Date(YEAR, MAY, getMonday(0, MAY, YEAR, 1)), // Victoria Day
                        new Date(YEAR, JUL, getDay(YEAR, JUL)),          // Canada Day
                        new Date(YEAR, AUG, getMonday(1, AUG, YEAR)),    // BC Day
                        new Date(YEAR, SEP, getMonday(1, SEP, YEAR)),    // Labour Day
                        new Date(YEAR, OCT, getMonday(2, OCT, YEAR)),    // Thanksgiving Day
                        new Date(YEAR, NOV, getDay(YEAR, NOV)),          // Remembrance Day
                        new Date(YEAR, DEC, getDay(YEAR, DEC))];         // Christmas Day

  const isStat = statHolidays.reduce((a, holiday) => {if (0 < today - holiday && today - holiday < ONE_DAY) return true})

  return !isStat;
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page
*
* @author Jarren Ralf
*/
function manualCounts()
{
  const QTY_COL = 3;
  const NUM_COLS = 3;
  
  var manualCountsSheet = SpreadsheetApp.getActive().getSheetByName("Manual Counts");
  var lastRow = manualCountsSheet.getLastRow();
  var startRow = (lastRow < 3) ? 4 : lastRow + 1;

  copySelectedValues(manualCountsSheet, startRow, NUM_COLS, QTY_COL);
}

/**
* This function moves all of the selected values on the info counts page to the Manual Counts page
*
* @author Jarren Ralf
*/
function manualCounts_FromInfoCounts()
{
  const QTY_COL = 3;
  const NUM_COLS = 3;
  
  var manualCountsSheet = SpreadsheetApp.getActive().getSheetByName("Manual Counts");
  var lastRow = manualCountsSheet.getLastRow();
  var startRow = (lastRow < 3) ? 4 : lastRow + 1;

  copySelectedValues(manualCountsSheet, startRow, NUM_COLS, QTY_COL, true);
}

/**
 * This function watches two cells and if the left one is edited then it searches the UPC Database for the upc value (the barcode that was scanned).
 * It then checks if the item is on the manual counts page and stores the relevant data in the left cell. If the right cell is edited, then the function
 * uses the data in the left cell and moves the item over to the manual counts page with the updated quantity.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function manualScan(e, spreadsheet, sheet)
{
  const manualCountsPage = spreadsheet.getSheetByName("Manual Counts");
  const barcodeInputRange = e.range;

  if (manualCountsPage.getRange(3, 7).isChecked()) // Manual Scanner is in "Add-One" mode
  {
    const upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
      .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
      .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
      .getValue();
    
    if (isNotBlank(upcCode)) // The user may have hit the delete key
    {
      const upcString = upcCode.toString().toLowerCase()
      const lastRow = manualCountsPage.getLastRow();
      const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();

      if (upcString == 'clear')
      {
        var item = e.oldValue;

        if (item === undefined)
          item = barcodeInputRange.offset(0, -1).getValue();

        item = item.split('\n');
        
        if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
          sheet.getRange(1, 1, 1, 2).setValues([['Item Not Found on Manual Counts page.', '']]);
        else
        {
          manualCountsPage.getRange(item[2], 3, 1, 3).setNumberFormat('@').setValues([['', '', '']])
          sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                          + '\nCurrent Stock :\n' + item[4] 
                                                          + '\nCurrent Manual Count :\n\nCurrent Running Sum :\n',
                                                          '']]);
        }
      }
      else if (upcString == 'undo')
      {
        var item = e.oldValue;

        if (item === undefined)
          item = barcodeInputRange.offset(0, -1).getValue();

        item = item.split('\n');

        if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
          sheet.getRange(1, 1, 1, 2).setValues([['Item Not Found on Manual Counts page.', '']]);
        else
        {
          var range = manualCountsPage.getRange(item[2], 3, 1, 3);
          var manualCountsValues = range.getValues()
          
          if (isNotBlank(manualCountsValues[0][1]))
          {
            var runningSumSplit = manualCountsValues[0][1].split(' ');

            if (runningSumSplit.length === 1)
            {
              range.setNumberFormat('@').setValues([['', '', '']])
              manualCountsValues[0][0] = ''
              manualCountsValues[0][1] = ''
              manualCountsValues[0][2] = ''
              var countedSince = ''
            }
            else if (runningSumSplit[runningSumSplit.length - 2] === '+')
            {
              manualCountsValues[0][0] -= Number(runningSumSplit[runningSumSplit.length - 1])
              runningSumSplit.pop();
              runningSumSplit.pop();
              manualCountsValues[0][1] = runningSumSplit.join(' ')
              manualCountsValues[0][2] = new Date().getTime()
              var countedSince = getCountedSinceString(manualCountsValues[0][2])
            }
            else if (runningSumSplit[runningSumSplit.length - 2] === '-')
            {
              manualCountsValues[0][0] += Number(runningSumSplit[runningSumSplit.length - 1])
              runningSumSplit.pop();
              runningSumSplit.pop();
              manualCountsValues[0][1] = runningSumSplit.join(' ')
              manualCountsValues[0][2] = new Date().getTime()
              var countedSince = getCountedSinceString(manualCountsValues[0][2])
            }
          }

          manualCountsValues[0][2] = new Date().getTime()
          range.setNumberFormats([['#.#', '@', '#']]).setValues(manualCountsValues)
          sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + (item[2]) 
                                                          + '\nCurrent Stock :\n' + item[4]
                                                          + '\nCurrent Manual Count :\n' + manualCountsValues[0][0] 
                                                          + '\nCurrent Running Sum :\n' + manualCountsValues[0][1]
                                                          + '\nLast Counted :\n' + countedSince,
                                                          '']]);
        }
      }
      else if (upcCode <= 100000) // In this case, variable name: upcCode is assumed to be the quantity
      {
        var item = e.oldValue;

        if (item === undefined)
          item = barcodeInputRange.offset(0, -1).getValue();

        item = item.split('\n');

        if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
          sheet.getRange(1, 1, 1, 2).setValues([['Item Not Found on Manual Counts page.', '']]);
        else
        {
          const range = manualCountsPage.getRange(item[2], 3, 1, 3);
          const manualCountsValues = range.getValues()
          manualCountsValues[0][2] = new Date().getTime()
          manualCountsValues[0][1] = (isNotBlank(manualCountsValues[0][1])) ? ((Math.sign(upcCode) === 1 || Math.sign(upcCode) === 0)  ? 
                                                                              String(manualCountsValues[0][1]) + ' \+ ' + String(   upcCode)  : 
                                                                              String(manualCountsValues[0][1]) + ' \- ' + String(-1*upcCode)) :
                                                                                ((isNotBlank(manualCountsValues[0][0])) ? 
                                                                                  String(manualCountsValues[0][0]) + ' \+ ' + String(upcCode) : 
                                                                                  String(upcCode));
          manualCountsValues[0][0] = Number(manualCountsValues[0][0]) + upcCode;
          range.setNumberFormats([['#.#', '@', '#']]).setValues(manualCountsValues)
          sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                          + '\nCurrent Stock :\n' + item[4] 
                                                          + '\nCurrent Manual Count :\n' + manualCountsValues[0][0] 
                                                          + '\nCurrent Running Sum :\n' + manualCountsValues[0][1]
                                                          + '\nLast Counted :\n' + getCountedSinceString(manualCountsValues[0][2]),
                                                          '']]);
        }
      }
      else
      {
        if (lastRow <= 3) // There are no items on the manual counts page
        {
          for (var i = upcDatabase.length - 1; i >= 1; i--) // Loop through the UPC values
          {
            if (upcDatabase[i][0] == upcCode) // UPC found
            {
              const row = lastRow + 1;
              manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']]).setValues([[upcDatabase[i][2], upcDatabase[i][3], 1, '\'' + String(1), new Date().getTime()]])
              applyFullRowFormatting(manualCountsPage, row, 1, 7)
              sheet.getRange(1, 1, 1, 2).setValues([[upcDatabase[i][2]  + '\nwas added to the Manual Counts page at line :\n' + row 
                                                                        + '\nCurrent Stock :\n' + upcDatabase[i][3]
                                                                        + '\nCurrent Manual Count :\n1',
                                                                        '']]);
            }
          }
        }
        else // There are existing items on the manual counts page
        {
          const row = lastRow + 1;
          const manualCountsValues = manualCountsPage.getSheetValues(4, 1, row - 3, 5);

          for (var i = upcDatabase.length - 1; i >= 1; i--) // Loop through the UPC values
          {
            if (upcDatabase[i][0] == upcCode)
            {
              for (var j = 0; j < manualCountsValues.length; j++) // Loop through the manual counts page
              {
                if (manualCountsValues[j][0] === upcDatabase[i][2]) // The description matches
                {
                  if (isNotBlank(manualCountsValues[j][4]))
                  {
                    const updatedCount = Number(manualCountsValues[j][2]) + 1;
                    const countedSince = getCountedSinceString(manualCountsValues[j][4])
                    const runningSum = (isNotBlank(manualCountsValues[j][3])) ? (String(manualCountsValues[j][3]) + ' \+ 1') : ((isNotBlank(manualCountsValues[j][2])) ? 
                                                                                                                                String(manualCountsValues[j][2]) + ' \+ 1' : 
                                                                                                                                String(1));
                    manualCountsPage.getRange(j + 4, 3, 1, 3).setNumberFormats([['#.#', '@', '#']]).setValues([[updatedCount, runningSum, new Date().getTime()]])
                    sheet.getRange(1, 1, 1, 2).setValues([[manualCountsValues[j][0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                                    + '\nCurrent Stock :\n' + manualCountsValues[j][1]
                                                                                    + '\nCurrent Manual Count :\n' + updatedCount 
                                                                                    + '\nCurrent Running Sum :\n' + runningSum
                                                                                    + '\nLast Counted :\n' + countedSince,
                                                                                    '']]);
                  }
                  else
                  {
                    manualCountsPage.getRange(j + 4, 3, 1, 3).setNumberFormats([['#.#', '@', '#']]).setValues([[1, '1', new Date().getTime()]])
                    sheet.getRange(1, 1, 1, 2).setValues([[manualCountsValues[j][0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                                    + '\nCurrent Stock :\n' + manualCountsValues[j][1]
                                                                                    + '\nCurrent Manual Count :\n1',
                                                                                    '']]);
                  }
                  break; // Item was found on the manual counts page, therefore stop searching
                } 
              }

              if (j === manualCountsValues.length) // Item was not found on the manual counts page
              {
                manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']])
                  .setValues([[upcDatabase[i][2], upcDatabase[i][3], 1, '\'' + String(1), new Date().getTime()]])
                applyFullRowFormatting(manualCountsPage, row, 1, 7)
                sheet.getRange(1, 1, 1, 2).setValues([[upcDatabase[i][2]  + '\nwas added to the Manual Counts page at line :\n' + row 
                                                                          + '\nCurrent Stock :\n' + upcDatabase[i][3]
                                                                          + '\nCurrent Manual Count :\n1',
                                                                          '']]);
              }

              break;
            }
          }
        }

        if (i === 0)
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
    }
  }
  else
  {
    if (barcodeInputRange.columnEnd === 1) // Barcode is scanned
    {
      const upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
        .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
        .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
        .getValue();

      if (isNotBlank(upcCode)) // The user may have hit the delete key
      {
        const lastRow = manualCountsPage.getLastRow();
        const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();

        if (lastRow <= 3) // There are no items on the manual counts page
        {
          for (var i = upcDatabase.length - 1; i >= 1; i--) // Loop through the UPC values
          {
            if (upcDatabase[i][0] == upcCode) // UPC found
            {
              barcodeInputRange.setValue(upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + upcDatabase[i][3]);
              break; // Item was found, therefore stop searching
            }
          }
        }
        else // There are existing items on the manual counts page
        {
          const row = lastRow + 1;
          const manualCountsValues = manualCountsPage.getSheetValues(4, 1, row - 3, 5);

          for (var i = upcDatabase.length - 1; i >= 1; i--) // Loop through the UPC values
          {
            if (upcDatabase[i][0] == upcCode)
            {
              for (var j = 0; j < manualCountsValues.length; j++) // Loop through the manual counts page
              {
                if (manualCountsValues[j][0] === upcDatabase[i][2]) // The description matches
                {
                  const countedSince = getCountedSinceString(manualCountsValues[j][4])
                    
                  barcodeInputRange.setValue(upcDatabase[i][2]  + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                + '\nCurrent Stock :\n' + upcDatabase[i][3] 
                                                                + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                                                + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]
                                                                + '\nLast Counted :\n' + countedSince);
                  break; // Item was found on the manual counts page, therefore stop searching
                }
              }

              if (j === manualCountsValues.length) // Item was not found on the manual counts page
                barcodeInputRange.setValue(upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + row + '\nCurrent Stock :\n' + upcDatabase[i][3]);

              break;
            }
          }
        }

        if (i === 0)
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
    }
    else if (barcodeInputRange.columnStart !== 1) // Quantity is entered
    {
      const quantity = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
        .setFontFamily("Arial").setFontColor("black").setFontSize(25)                      // Set the font parameters
        .setVerticalAlignment("middle").setHorizontalAlignment("center")                   // Set the alignment parameters
        .getValue();

      if (isNotBlank(quantity)) // The user may have hit the delete key
      {
        const item = sheet.getRange(1, 1).getValue().split('\n');    // The information from the left cell that is used to move the item to the manual counts page
        const quantity_String = quantity.toString().toLowerCase();
        const quantity_String_Split = quantity_String.split(' ');

        if (quantity_String === 'clear')
        {
          manualCountsPage.getRange(item[2], 3, 1, 3).setNumberFormat('@').setValues([['', '', '']])
          sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                          + '\nCurrent Stock :\n' + item[4] 
                                                          + '\nCurrent Manual Count :\n\nCurrent Running Sum :\n',
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
            spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
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
            manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[marriedItem[0], upc, marriedItem[4], item[0]]]);
            upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[upc, marriedItem[4], item[0], item[4]]]); 
            barcodeInputRange.setValue('UPC Code has been added to the database temporarily.')
            spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
          }
          else
            barcodeInputRange.setValue('Please enter a valid UPC Code to marry.')
        }
        else if (isNumber(quantity_String_Split[0]) && isNotBlank(quantity_String_Split[1]) && quantity_String_Split[1] != null)
        {
          if (item.length !== 1) // The cell to the left contains valid item information
          {
            quantity_String_Split[1] = quantity_String_Split[1].toUpperCase()

            if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()
              const updatedCount = Number(itemValues[0][0]) + Number(quantity_String_Split[0]);
              const countedSince = getCountedSinceString(itemValues[0][2])
              const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity_String_Split[0]) === 1 || Math.sign(quantity_String_Split[0]) === 0)  ? 
                                                                    String(itemValues[0][1]) + ' \+ ' + String(   quantity_String_Split[0])  : 
                                                                    String(itemValues[0][1]) + ' \- ' + String(-1*quantity_String_Split[0])) :
                                                                      ((isNotBlank(itemValues[0][0])) ? 
                                                                        String(itemValues[0][0]) + ' \+ ' + String(quantity_String_Split[0]) : 
                                                                        String(quantity_String_Split[0]));

              if (isNotBlank(itemValues[0][3]) && isNotBlank(itemValues[0][4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[0][3] + '\n' + quantity_String_Split[1], itemValues[0][4] + '\n' + quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[0][3]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[0][3] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[0][4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[1], itemValues[0][4] + '\n' + quantity_String_Split[0].toString()]]);
              else
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[1], quantity_String_Split[0].toString()]]);

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + updatedCount 
                                                              + '\nCurrent Running Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
            else
            {
              const lastRow = manualCountsPage.getLastRow();
              const row = lastRow + 1;
              const range = manualCountsPage.getRange(row, 1, 1, 7)
              const itemValues = range.getValues()

              if (isNotBlank(itemValues[0][5]) && isNotBlank(itemValues[0][6]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                  new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[1], itemValues[0][6] + '\n' + quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[0][5]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                  new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
              else if (isNotBlank(itemValues[0][6]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                  new Date().getTime(), quantity_String_Split[1], itemValues[0][6] + '\n' + quantity_String_Split[0].toString()]]);
              else
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                  new Date().getTime(), quantity_String_Split[1], quantity_String_Split[0].toString()]]);

              applyFullRowFormatting(manualCountsPage, row, 1, 7)
              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + quantity_String_Split[0],
                                                              '']]);
            }
          }
          else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
            barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

          sheet.getRange(1, 1).activate();
        }
        else if (isNumber(quantity_String_Split[1]))
        {
          if (item.length !== 1) // The cell to the left contains valid item information
          {
            quantity_String_Split[0] = quantity_String_Split[0].toUpperCase()

            if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 5);
              const itemValues = range.getValues()
              const updatedCount = Number(itemValues[0][0]) + Number(quantity_String_Split[1]);
              const countedSince = getCountedSinceString(itemValues[0][2])
              const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity_String_Split[1]) === 1 || Math.sign(quantity_String_Split[1]) === 0)  ? 
                                                                    String(itemValues[0][1]) + ' \+ ' + String(   quantity_String_Split[1])  : 
                                                                    String(itemValues[0][1]) + ' \- ' + String(-1*quantity_String_Split[1])) :
                                                                      ((isNotBlank(itemValues[0][0])) ? 
                                                                        String(itemValues[0][0]) + ' \+ ' + String(quantity_String_Split[1]) : 
                                                                        String(quantity_String_Split[1]));

              if (isNotBlank(itemValues[0][3]) && isNotBlank(itemValues[0][4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[0][3] + '\n' + quantity_String_Split[0], itemValues[0][4] + '\n' + quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[0][3]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  itemValues[0][3] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[0][4]))
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[0], itemValues[0][4] + '\n' + quantity_String_Split[1].toString()]]);
              else
                range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                  quantity_String_Split[0], quantity_String_Split[1].toString()]]);

              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + updatedCount 
                                                              + '\nCurrent Running Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
            else
            {
              const lastRow = manualCountsPage.getLastRow();
              const row = lastRow + 1;
              const range = manualCountsPage.getRange(row, 1, 1, 7)
              const itemValues = range.getValues()

              if (isNotBlank(itemValues[0][5]) && isNotBlank(itemValues[0][6]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                  new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[0], itemValues[0][6] + '\n' + quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[0][5]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                  new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
              else if (isNotBlank(itemValues[0][6]))
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                  new Date().getTime(), quantity_String_Split[0], itemValues[0][6] + '\n' + quantity_String_Split[1].toString()]]);
              else
                range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                  new Date().getTime(), quantity_String_Split[0], quantity_String_Split[1].toString()]]);

              applyFullRowFormatting(manualCountsPage, row, 1, 7)
              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + quantity_String_Split[1],
                                                              '']]);
            }
          }
          else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
            barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

          sheet.getRange(1, 1).activate();
        }
        else if (quantity <= 100000) // If false, Someone probably scanned a barcode in the quantity cell (not likely to have counted an inventory amount of 100 000)
        {
          if (item.length !== 1) // The cell to the left contains valid item information
          {
            if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 3);
              const itemValues = range.getValues()
              const updatedCount = Number(itemValues[0][0]) + quantity;
              const countedSince = getCountedSinceString(itemValues[0][2])
              const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity) === 1 || Math.sign(quantity) === 0)  ? 
                                                                    String(itemValues[0][1]) + ' \+ ' + String(   quantity)  : 
                                                                    String(itemValues[0][1]) + ' \- ' + String(-1*quantity)) :
                                                                      ((isNotBlank(itemValues[0][0])) ? 
                                                                        String(itemValues[0][0]) + ' \+ ' + String(quantity) : 
                                                                        String(quantity));
              range.setNumberFormats([['#.#', '@', '#']]).setValues([[updatedCount, runningSum, new Date().getTime()]])
              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + updatedCount 
                                                              + '\nCurrent Running Sum :\n' + runningSum
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
            else
            {
              const lastRow = manualCountsPage.getLastRow();
              const row = lastRow + 1;
              manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']]).setValues([[item[0], item[4], quantity, '\'' + String(quantity), new Date().getTime()]])
              applyFullRowFormatting(manualCountsPage, row, 1, 7)
              sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + quantity,
                                                              '']]);
            }
          }
          else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
            barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

          sheet.getRange(1, 1).activate();
        }
        else 
          barcodeInputRange.setValue('Please enter a valid quantity.')
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
* Sorts data by the categories while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCategories(a, b)
{
  return (a[9].toLowerCase() === b[9].toLowerCase()) ? 0 : (a[9] === '') ? 1 : (b[9] === '') ? -1 : (a[9].toLowerCase() < b[9].toLowerCase()) ? -1 : 1;
}

/**
* Sorts data by the created date of the product for the parksville and rupert spreadsheets.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate(a, b)
{
  return (a[7] === b[7]) ? 0 : (a[7] < b[7]) ? 1 : -1;
}

/**
* Sorts data by the created date of the product for the richmond spreadsheet.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate_Richmond(a, b)
{
  return (a[8] === b[8]) ? 0 : (a[8] < b[8]) ? 1 : -1;
}


/**
* Sorts data by the created date of the product
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCountedDate(a, b)
{
  return (a[3] === b[3]) ? 0 : (a[3] < b[3]) ? -1 : 1;
}

/**
* This is a function I found and modified to keep the last instance of an item in a muli-array based on the uniqueness of one of the values.
*
* @param      {Object[][]}    arr : The given array
* @param  {Callback Function} key : A function that chooses one of the elements of the object or array
* @return     {Object[][]}    The reduced array containing only unique items based on the key
*/
function uniqByKeepLast(arr, key) {
    return [...new Map(arr.map(x => [key(x), x])).values()]
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
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}

/**
 * This function checks if the user edits the item description or the Current Count column on the 
 * Manual Counts page. If they did, then a warning appears and reverses the changes that they made.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The active spreadsheet
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param    {String}     sheetName  : The string that represents the name of the sheet
 * @author Jarren Ralf
 */
function warning(e, spreadsheet, sheet, sheetName)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;

  if (row == range.rowEnd && col == range.columnEnd) // Single cell
  {
    if (col == 1)
    {
      if (!isRichmondSpreadsheet(spreadsheet))
      {
        (sheetName === 'Manual Counts') ? // sheetName === 'TitesCounts'
          SpreadsheetApp.getUi().alert("Please don't attempt to change the items from the Manual Counts page.\n\nGo to the Item Search or Manual Scan page to add new products to this list.") :
          SpreadsheetApp.getUi().alert("Please don't attempt to change the items on the InfoCounts page.");

        range.setValue(e.oldValue); // Put the old value back in the cell
      }
    }
    else if (col == 2)
    {
      SpreadsheetApp.getUi().alert("Please don't change values in the Current Count column.\n\nType your updated inventory quantity in the New Count column.");
      range.setValue(e.oldValue); // Put the old value back in the cell
      if (userHasNotPressedDelete(e.value)) sheet.getRange(row, 3).setValue(e.value).activate(); // Move the count the user entered to the New Count column
    }
    else if (col == 3 && sheetName === 'Manual Counts')
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