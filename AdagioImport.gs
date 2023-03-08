function newFunc_runAll()
{
  
  const spreadsheet = SpreadsheetApp.getActive()
  importData(spreadsheet)
  const dataSheet = spreadsheet.getSheetByName('Imported Data')
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet')
  const numRows_data = dataSheet.getLastRow() - 3;
  const skus = getSKUs(spreadsheet);

  if (numRows_data > 0)
  {
    const data = dataSheet.getSheetValues(4, 1, numRows_data, dataSheet.getLastColumn());

    var richCounts  = [], richCounts_invalid  = [],
        parksCounts = [], parksCounts_invalid = [],
        ruptCounts  = [], ruptCounts_invalid  = [],
        transfers   = [], transfers_invalid   = [];

    for (var sheet = 0; sheet < 14; sheet++)
    {
      switch (sheet)
      {
        case 0:

          for (var i = 1; i < data[0][0]; i++)
          {
            if (isNotBlank(data[i][1]))
            {
              if (isNumber(data[i][1]))
                richCounts.push([data[i][0], data[i][1], '', 'InfoCounts'])
              else
                richCounts_invalid.push([data[i][0], data[i][1], '', 'InfoCounts'])
            }
          }
          break;
        case 1:

          for (var i = 1; i < data[0][2]; i++)
          {
            if (isNotBlank(data[i][3]))
            {
              if (isNumber(data[i][3]))
                richCounts.push([data[i][2], data[i][3], '', 'Manual Counts'])
              else
                richCounts_invalid.push([data[i][2], data[i][3], '', 'Manual Counts'])
            }
          }
          break;
        case 2:

          for (var i = 1; i < data[0][4]; i++)
          {
            if (isNotBlank(data[i][7]))
            {
              if (isNumber(data[i][7]))
              {
                if (data[i][7] != data[i][6] && data[i][5] != 'B/O')
                  parksCounts.push([data[i][4], data[i][7], '', 'Order'])
              }
              else
                parksCounts_invalid.push([data[i][4], data[i][7], '', 'Order'])
            }
          }
          break;

        case 3:

          for (var i = 1; i < data[0][8]; i++)
          {
            if (isNotBlank(data[i][10]))
            {
              if (isNumber(data[i][10]))
              {
                if (data[i][10] != data[i][9])
                  parksCounts.push([data[i][8], data[i][10], '', 'Shipped'])
              }
              else
                parksCounts_invalid.push([data[i][8], data[i][10], '', 'Shipped'])
            }
          }
          break;
        case 4:

          for (var i = 1; i < data[0][11]; i++)
          {
            if (isNotBlank(data[i][12]))
            {
              if (isNumber(data[i][12]))
                parksCounts.push([data[i][11], data[i][12], '', 'InfoCounts'])
              else
                parksCounts_invalid.push([data[i][11], data[i][12], '', 'InfoCounts'])
            }
          }
          break;
        case 5:

          for (var i = 1; i < data[0][13]; i++)
          {
            if (isNotBlank(data[i][14]))
            {
              if (isNumber(data[i][14]))
                parksCounts.push([data[i][13], data[i][14], '', 'Manual Counts'])
              else
                parksCounts_invalid.push([data[i][13], data[i][14], '', 'Manual Counts'])
            }
          }
          break;
        case 6:

          for (var i = 1; i < data[0][22]; i++)
          {
            if (isNotBlank(data[i][25]))
            {
              if (isNumber(data[i][25]))
              {
                if (data[i][25] != data[i][24] && data[i][23] != 'B/O')
                  ruptCounts.push([data[i][4], data[i][7]], '', 'Order')
              }
              else
                ruptCounts_invalid.push([data[i][22], data[i][25], '', 'Order'])
            }
          }
          break;

        case 7:

          for (var i = 1; i < data[0][26]; i++)
          {
            if (isNotBlank(data[i][28]))
            {
              if (isNumber(data[i][28]))
              {
                if (data[i][28] != data[i][27])
                  ruptCounts.push([data[i][26], data[i][28], '', 'Shipped'])
              }
              else
                ruptCounts_invalid.push([data[i][26], data[i][28], '', 'Shipped'])
            }
          }
          break;
        case 8:

          for (var i = 1; i < data[0][29]; i++)
          {
            if (isNotBlank(data[i][30]))
            {
              if (isNumber(data[i][30]))
                ruptCounts.push([data[i][29], data[i][30], '', 'InfoCounts'])
              else
                ruptCounts_invalid.push([data[i][29], data[i][30], '', 'InfoCounts'])
            }
          }
          break;
        case 9:

          for (var i = 1; i < data[0][31]; i++)
          {
            if (isNotBlank(data[i][32]))
            {
              if (isNumber(data[i][32]))
                ruptCounts.push([data[i][31], data[i][32], '', 'Manual Counts'])
              else
                ruptCounts_invalid.push([data[i][31], data[i][32], '', 'Manual Counts'])
            }
          }
          break;
        case 10:

          for (var i = 1; i < data[0][15]; i++)
          {
            if (data[i][17] == 'false')
            {
              if (isNumber(data[i][16]) && skus.includes(data[i][15]))
                transfers.push([data[i][15], data[i][16], '100', '200'])
              else
                transfers_invalid.push([data[i][15], data[i][16], '100', '200'])
            }
          }
          break;

        case 11:

          for (var i = 1; i < data[0][18]; i++)
          {
            if (data[i][21] == 'false' && isNotBlank(data[i][20]))
            {
              if (isNumber(data[i][19]) && skus.includes(data[i][18]))
                transfers.push([data[i][18], data[i][19], '200', '100'])
              else
                transfers_invalid.push([data[i][18], data[i][19], '200', '100'])
            }
          }
          break;
        case 12:

          for (var i = 1; i < data[0][33]; i++)
          {
            if (data[i][35] == 'false')
            {
              Logger.log(data[i][33])

              if (isNumber(data[i][34]) && skus.includes(data[i][33]))
                transfers.push([data[i][33], data[i][34], '100', '300'])
              else
                transfers_invalid.push([data[i][33], data[i][34], '100', '300'])
            }
          }
          break;
        case 13:

          for (var i = 1; i < data[0][36]; i++)
          {
            if (data[i][39] == 'false' && isNotBlank(data[i][38]))
            {
              if (isNumber(data[i][37]) && skus.includes(data[i][37]))
                transfers.push([data[i][36], data[i][37], '300', '100'])
              else
                transfers_invalid.push([data[i][36], data[i][37], '300', '100'])
            }
          }
          break;
      }
    }

    Logger.log(richCounts)
    Logger.log(richCounts.length)
    Logger.log(richCounts_invalid)
    Logger.log(richCounts_invalid.length)

    Logger.log(parksCounts)
    Logger.log(parksCounts.length)
    Logger.log(parksCounts_invalid)
    Logger.log(parksCounts_invalid.length)

    Logger.log(ruptCounts)
    Logger.log(ruptCounts.length)
    Logger.log(ruptCounts_invalid)
    Logger.log(ruptCounts_invalid.length)

    Logger.log(transfers)
    Logger.log(transfers.length)
    Logger.log(transfers_invalid)
    Logger.log(transfers_invalid.length)
  } 
  else
    SpreadsheetApp.getUi().alert('Imported Data is blank.')
}

function isNotBlank(str)
{
  return str !== ''
}

/**
* This function gets some of the physical counts done by a particular location based on which google sheet is being analyzed.
*
* @param  {Sheet}      sheet      The google sheet with the imported data
* @param  {Number}     descripCol The column number for the description column
* @return {String[][]} The list of SKUs, quantities, and the sheet name the data comes from
* @author Jarren Ralf
*/
function getPhysicalCounted(sheet, descripCol)
{
  const DATA_START_ROW = 4;
  const SHEET_NAME_ROW = 2;
  const    DESCRIP_COL = 0; // Of the 2-dimensional data array defined below 
  const BACK_ORDER_COL = 1; // Of the 2-dimensional data array defined below 
  const KEEP_FIRST_STRING_ONLY = 1; // Used as an argument past to the split() function below
  
  var    numRows = sheet.getLastRow() - DATA_START_ROW + 1;
  var  sheetName = sheet.getSheetName();
  var whichSheet = sheet.getRange(SHEET_NAME_ROW, descripCol).getValue(); // What the original sheet name is called that the data is being pulled from
  var sku            = [], qty         = []; // Initialize the sku  and qty  arrays
  var sku_InvalidQty = [], qty_Invalid = []; // Initialize the sku_InvalidQty and qty_invalid arrays for the non-numeric qty values
  
  // Check if there is any data
  if (numRows > 0)
  {
    /* Determine the number of columns of data needed to compile the physical counts of a particular location.
     * 
     * For instance, any physical counts for the Richmond or Trites location is only based on 2 columns of data in the "Imported .. " sheets of the Adagio Update spreadsheet.
     * As for Parksville and Rupert, they may need to look at 2, 3, or 4 columns depending on which sheet the data was imported from.
     */
    var numCols = (descripCol == 5) ? 3 : ((sheetName == "Imported Parksville Data (Loc: 200)" || sheetName == "Imported Rupert Data (Loc: 300)") && descripCol == 1) ? 4 : 2;
    
    var          qtyCol = numCols - 1; // Actual Count or Counts column in the data array defined below
    var currentStockCol = numCols - 2; // The column in the data array representing the current stock in the Adagio inventory system
    var data = sheet.getRange(DATA_START_ROW, descripCol, numRows, numCols).getValues(); // Get the whole data Range

    for (var i = 0; i < data.length; i++)
    { 
      if (data[i][qtyCol] !== '' && isNumber(data[i][qtyCol])) // Check if the entry is a number
      {
        // If we are on the Order or Shipped sheet, then check that the actual count and current stock are different, and don't include Back Orders if on the Order sheet
        if ((numCols < 3) || ((data[i][qtyCol] != data[i][currentStockCol]) && (descripCol != 1 || data[i][BACK_ORDER_COL] != "B/O"))) 
        {
          sku.push(data[i][DESCRIP_COL].toString().split(" - ", KEEP_FIRST_STRING_ONLY)[0]); // Strip off the sku
          qty.push(data[i][qtyCol]);
        }
      }
      else if (numCols == 2 && data[i][qtyCol] !== "") // If we are on the InfoCounts sheet for one of the four locations and the qty is precisely a nonempty string
      {
        sku_InvalidQty.push(data[i][DESCRIP_COL].toString().split(" - ", KEEP_FIRST_STRING_ONLY)); // Strip off the sku
        qty_Invalid.push(data[i][qtyCol]);             // This picks up quantities that are strings instead of numbers
      }
    }
    
    if (numCols == 2) // If we are on the InfoCounts page then we need to concatenate our two sets of skus and quantities
    {
      var numInvalidQty = qty_Invalid.length;
      var fromSheet_InvalidQty = new Array(numInvalidQty).fill(whichSheet);
      
      // This puts a space after the non-numeric qty entry (which still allows for the ctrl+A command to work on the rest of the data)
      if (numInvalidQty != 0)
      {
        sku_InvalidQty.push(null);
        qty_Invalid.push(null);
        fromSheet_InvalidQty.push(null);
      }
      
      // Combine (or concatenate) the non-numeric qantities with all of the original quantites and report which sheet the data is coming from
      var numSku = sku.length;
      var fromSheet = new Array(numSku).fill(whichSheet);
      var       sku_Combined = sku_InvalidQty.concat(sku);
      var       qty_Combined = qty_Invalid.concat(qty);
      var fromSheet_Combined = fromSheet_InvalidQty.concat(fromSheet);
      var    numSku_Combined = sku_Combined.length;
      var blank = new Array(numSku_Combined); 
      
      return transpose([sku_Combined, qty_Combined, blank, fromSheet_Combined]);
    }
    else
    {
      var numSku = sku.length;
      var     blank = new Array(numSku);  
      var fromSheet = new Array(numSku).fill(whichSheet);
      
      return transpose([sku, qty, blank, fromSheet]);
    }
  }
  
  var blank = [], fromSheet = [];                // There wasn't any data and hence these arrays have not been initiated
  return transpose([sku, qty, blank, fromSheet]) // Return an array of four empty arrays
}

/**
* This function gets all of the SKUs so that it only has to be done once in a single runAll execution.
*
* @return {String[][]} The list of all SKUs
* @author Jarren Ralf
*/
function getSKUs(spreadsheet)
{
  const inventorySheet = spreadsheet.getSheetByName('DataImport');

  return inventorySheet.getSheetValues(2, 7, inventorySheet.getLastRow() - 1, 1).flat();
}

/**
* This function gets the stock transfers of inventory that have been shipped between locations.
*
* @param  {Sheet}      sheet            The google sheet with the imported data
* @param  {Number}     descripCol       The column number for the description column
* @param  {Number}     fromLocation     The location of the shipper  (eg. 100, 200, 300, 400)
* @param  {Number}     toLocation       The location of the receiver (eg. 100, 200, 300, 400)
* @param  {Boolean}    isTritesTransfer A variable that represents whether to process the transfers to (and from) Trites or not
* @return {String[][]} The stock transfers and the stock transfers with invalid quantities (non-numeric quantities)
* @author Jarren Ralf
*/
function getStockTransfers(sheet, descripCol, fromLocation, toLocation, isTritesTransfer)
{
  const DESCRIPTION = 0; // The following FOUR constants represent the columns of the data array instantiated below
  const SHIPPED_QTY = 1;
  const RECEIVED_BY = 2;
  const  TRANSFERED = 3;
  const       NUM_COLS = 4; // There are four columns to take data from, namely the Description, Shipped, Entered By / Received By and the Transfered columns
  const DATA_START_ROW = 4;
  const KEEP_FIRST_STRING_ONLY = 1; // Used as an argument past to the split() function below

  var numRows = sheet.getLastRow() - DATA_START_ROW + 1;
  var sku            = [], qty         = []; // Initialize the sku and qty arrays
  var sku_InvalidQty = [], qty_Invalid = []; // Initialize the sku2 and qty_invalid arrays for the non-numeric qty values
  var data = sheet.getRange(DATA_START_ROW, descripCol, numRows, NUM_COLS).getValues(); // Get the whole data Range

  // Get the SKUs and quantities of the items that haven't been transfered in Adagio yet
  for (var i = 0; i < numRows; i++)
  { 
    /* Firstly, for all stock transfers, the TRANSFERED column must be uncheck (== 0) and the items must be received, i.e. the RECEIVED BY column must not be blank.
     * In a addition, either one of the following must also be true:
     * 1) If it is a TRITES transfer, then RECEIVED BY (or ENTERED BY) must say "Trites", otherwise if it is not, then
     * 2) isTritesTransfer must be false and RECEIVED BY (or ENTERED BY) cannot say "Trites".
     */
    if (data[i][TRANSFERED] == false && data[i][RECEIVED_BY] !== '')
    {
      if (isNumber(data[i][SHIPPED_QTY]) ) // If the qty is a valid number
      {
        sku.push(data[i][DESCRIPTION].toString().split(" - ", KEEP_FIRST_STRING_ONLY)[0]); // Strip off the sku
        qty.push(data[i][SHIPPED_QTY]);
      }
      else
      {
        sku_InvalidQty.push(data[i][DESCRIPTION].toString().split(" - ", KEEP_FIRST_STRING_ONLY)[0]); // Strip off the sku
        qty_Invalid.push(data[i][SHIPPED_QTY]); // This picks up quantities that are letters instead of numbers
      }
    }
  }
  
  // Fill the to and from locations for the stock transfers
  var numSku = sku.length;
  var   from = new Array(numSku).fill(fromLocation);  
  var     to = new Array(numSku).fill(toLocation);
  
  // Fill the to and from locations for the stock transfers that contain invalid quantities (non-numeric quantities)
  var  num_InvalidQty = qty_Invalid.length;
  var from_InvalidQty = new Array(num_InvalidQty).fill(fromLocation);
  var   to_InvalidQty = new Array(num_InvalidQty).fill(toLocation);
  
  return [transpose([sku, qty, from, to]), transpose([sku_InvalidQty, qty_Invalid, from_InvalidQty, to_InvalidQty])];
}

/**
* This function imports the Richmond data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountRich()
{
  var startTime = new Date().getTime(); // For the run time
  var richmondSheet = SpreadsheetApp.getActive().getSheetByName('Imported Richmond Data (Loc: 100)');
  var      dataCols = [1, 3];
  var isDataImported = false;
  const RICH_PHYS_COUNT_COL = 2;
  physicalCount(richmondSheet, RICH_PHYS_COUNT_COL, dataCols, isDataImported);
  setElapsedTime(startTime);
}

/**
* This function imports the Parksville data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountPark()
{  
  var startTime = new Date().getTime(); // For the run time  
  var parksvilleSheet = SpreadsheetApp.getActive().getSheetByName('Imported Parksville Data (Loc: 200)');
  var        dataCols = [1, 5, 8, 10];
  var  isDataImported = false;
  const PARKS_PHYS_COUNT_COL =  7;
  physicalCount(parksvilleSheet, PARKS_PHYS_COUNT_COL, dataCols, isDataImported);
  setElapsedTime(startTime);// To check the ellapsed times
}

/**
* This function imports the Rupert data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountRupt()
{
  var startTime = new Date().getTime(); // For the run time
  var  rupertSheet = SpreadsheetApp.getActive().getSheetByName('Imported Rupert Data (Loc: 300)');
  var     dataCols = [1, 5, 8, 10];
  var isDataImported = false;
  const RUPERT_PHYS_COUNT_COL =  12;
  physicalCount(rupertSheet, RUPERT_PHYS_COUNT_COL, dataCols, isDataImported); 
  setElapsedTime(startTime);// To check the ellapsed times
}

/**
* This function imports the Trites data and gathers the phyical counts. 
*
* @author Jarren Ralf
*/
function physCountTrit()
{
  var      startTime = new Date().getTime(); // For the run time
  var    tritesSheet = SpreadsheetApp.getActive().getSheetByName('Imported Richmond Data (Loc: 100)');
  var       dataCols = [13, 15];
  var isDataImported = false;
  const TRITES_PHYS_COUNT_COL = 17;
  physicalCount(tritesSheet, TRITES_PHYS_COUNT_COL, dataCols, isDataImported);
  setElapsedTime(startTime); // To check the ellapsed times
}

/**
* This gathers the phyical counts, and if needed imports the data in order to do so.
*
* @param {Sheet}    dataSheet      The sheet that represents the location with which the counts are being determined
* @param {Number}   physCountCol   The column that the physical data will be posted to on the Adagio page
* @param {Number[]} dataCols       An array of the data columns that the counts are based on
* @param {Boolean}  isDataImported Determines whether to import the data or not
* @author Jarren Ralf
*/
function physicalCount(dataSheet, physCountCol, dataCols, isDataImported)
{
  const TIME_STAMP_ROW = 2;
  
  // If data is already imported (because runAll is being called) skip the additional import
  if (!isDataImported)
  {
    importData(dataSheet);
  }
  
  var counts   = getPhysicalCounted(dataSheet, dataCols[0]); // dataCols[0] = 1 
  var counts_2 = getPhysicalCounted(dataSheet, dataCols[1]); // dataCols[1] = 5 
  
  // If there is more than 1 column of data compute the additional counts and set the values
  if (dataCols.length > 2)
  {
    var counts_3 = getPhysicalCounted(dataSheet, dataCols[2]); // dataCols[2] =  8 
    var counts_4 = getPhysicalCounted(dataSheet, dataCols[3]); // dataCols[3] = 10
    setPhysicalCounted(physCountCol, counts_4, counts_3, counts_2, counts);
  }
  else 
    setPhysicalCounted(physCountCol, counts_2, counts);
  
  timeStamp(TIME_STAMP_ROW, physCountCol);
}

/**
* This function imports all of the data, gathers the counts for all locations, and reports all of the stock transfers.
*
* @author Jarren Ralf
*/
function runAll()
{ 
  var startTime = new Date().getTime(); // For the run time
  
  const PHYS_COUNT_COLS = [2, 7, 12]; // Richmond (100), Parksville (200), Rupert (300), Trites (400)
  
  var allSKUs = getSKUs();   // Get all of the SKUs
  var isDataImported = true; // Set isDataImported to true because the importAllData function will be run precisely once each time this function is run
  var dataCols = [];
  var richmondImportSheet, parksImportSheet, rupertImportSheet;
  [richmondImportSheet, parksImportSheet, rupertImportSheet] = importAllData();
  
  // Richmond and Trites physical counts
  dataCols[0] = 1;
  dataCols[1] = 3;
  physicalCount(richmondImportSheet, PHYS_COUNT_COLS[0], dataCols, isDataImported);
  
  // Parksville and Prince Rupert physical counts 
  dataCols[0] =  1;
  dataCols[1] =  5;
  dataCols[2] =  8;
  dataCols[3] = 10;
  physicalCount( parksImportSheet, PHYS_COUNT_COLS[1], dataCols, isDataImported);
  physicalCount(rupertImportSheet, PHYS_COUNT_COLS[2], dataCols, isDataImported);
  
  stockTransfersAll(parksImportSheet, rupertImportSheet, allSKUs)  // Get and set all of the stock transfers
  timeStamp(2, 19);          // Run all timestamp
  setElapsedTime(startTime); // To check the ellapsed times
}

/**
* This function sets the ellapsed time of a function and prints it on the Adagio page.
*
* @param {Number} startTime The start time that the script began running at represented by a number in milliseconds
* @author Jarren Ralf
*/
function setElapsedTime(startTime)
{
  const ELAPSED_TIME_ROW = 1;
  const ELAPSED_TIME_COL = 11;
  
  var adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
  var timeNow = new Date().getTime(); // Get milliseconds from a date in past
  var elapsedTime = (timeNow - startTime)/1000;
  
  adagioSheet.getRange(ELAPSED_TIME_ROW, ELAPSED_TIME_COL).setValue(elapsedTime);
}

/**
* This function sets the IMPORTRANGE fromulas and then pastes the values on the page, which removes the formula.
*
* @param {Sheet} sheet The google sheet with the relevant information
* @author Jarren Ralf
*/
function importData(spreadsheet)
{
  // When clicking the Import Data buttons, the function will be run with 0 arguments
  if (arguments.length == 0)
    var spreadsheet = SpreadsheetApp.getActive()

  const dataSheets = [  spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'), 
                        spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'),
                        spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)')
  ]

  const formulas = [[[
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\")))))",                                                            // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!C4:C\")",     // Richmond, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\")))))",                                                         // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!C4:C\")"]],  // Richmond, Counts
    [["=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\")))))",                                                                                                                 // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!J4:J\")",          // Parksville, BO
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!G4:G\")",          // Parksville, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!H4:H\")",          // Parksville, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\")))))",                                                                                                               // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!G4:G\")",        // Parksville, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!H4:H\")",        // Parksville, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\")))))",                                                            // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!C4:C\")",     // Parksville, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\")))))",                                                         // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!C4:C\")",  // Parksville, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\")))))",                                                              // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!I4:I\")",       // Parksville, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!L4:L\")",       // Parksville, Transfered
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\")))))",                                                       // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!F4:F\")",// Parksville, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!H4:H\")",// Parksville, Received By
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!I4:I\")"]],// Parksville, Transfered
    [["=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\")))))",                                                                                                                 // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!J4:J\")",          // Rupert, BO
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!G4:G\")",          // Rupert, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!H4:H\")",          // Rupert, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\")))))",                                                                                                               // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!G4:G\")",        // Rupert, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!H4:H\")",        // Rupert, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\")))))",                                                            // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!C4:C\")",     // Rupert, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\")))))",                                                         // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!C4:C\")",  // Rupert, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\")))))",                                                              // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!I4:I\")",       // Rupert, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!L4:L\")",       // Rupert, Transfered
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"),\" - \")),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"),LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\")))))",                                                       // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!F4:F\")",// Rupert, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!H4:H\")",// Rupert, Received By
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!I4:I\")" // Rupert, Transfered
  ]]]

  for (var sheet = 0; sheet < dataSheets.length; sheet++)
  {
    var numCols = dataSheets[sheet].getLastColumn();
    dataSheets[sheet].getRange(4, 1, dataSheets[sheet].getLastRow() - 3, numCols).clearContent().offset(0, 0, 1, numCols).setValues(formulas[sheet])
    var range = dataSheets[sheet].getDataRange();
    var newValues = range.getValues()
    range.setNumberFormat('@').setValues(newValues)
  }
}

/**
* This function sets the physical counts reported by one of the locations.
*
* @param {Number}     skuCol
* @param {String[][]} order      The data from the Order page
* @param {String[][]} shipped    The data from the Shipped page
* @param {String[][]} infoCounts The data from the InfoCounts page
* @author Jarren Ralf
*/
function setPhysicalCounted(skuCol, order, shipped, infoCounts, manualCounts)
{
  const START_ROW_ADAGIO_SHEET = 24;
  const NUM_COLS = 4
  
  // If there are only three aguments sent to this function, no need to concatenate, otherwise, concatenate all of the counted items into one array
  var items_ = (arguments.length == 3) ? order.concat(shipped) : order.concat(shipped, infoCounts, manualCounts);
  var items  = items_.map(row => row.map((value, index, arr) => (Array.isArray(value)) ? "\'" + value[0] : value )); // Place an apostrophe infront of every SKU to force the datatype to be String
  var numItems = items.length;
  var adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
  var lastRowOfData = adagioSheet.getLastRow();
  
  // Clear the content of the counted items unless there isn't any data on the adagio page
  if (lastRowOfData > START_ROW_ADAGIO_SHEET) 
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, skuCol, lastRowOfData, NUM_COLS).clearContent();
  
  // Set all of the counted unless there aren't any
  if (numItems != 0)
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, skuCol, numItems, NUM_COLS).setNumberFormat("@").setValues(items);
}

/**
* This function sets the stock transfers of inventory that have been shipped between all locations.
* It removes transfers that don't contain valid skus.
*
* @param {Object[][][]} transfers An array of double arrays representing all of the stock transfers between all combinations of locations
* @param {Object[][]}   allSKUs   A set of all the active SKUs in the PNT database
* @author Jarren Ralf
*/
function setStockTransfersAll(transfers, transfers_InvalidQty, allSKUs)
{
  const START_ROW_ADAGIO_SHEET = 24;
  const STOCK_TRANSFER_COL = 17;
  const NUM_COLS = 4;
  
  var adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
  var lastRowOfData = adagioSheet.getLastRow();
  var stockTransfers            = [].concat.apply([], transfers);            // Concatenate all of the transfers as a 2-D array
  var stockTransfers_InvalidQty = [].concat.apply([], transfers_InvalidQty); // Concatenate all of the transfers with invalid quantities as a 2-D array

  // This code block removes the entries that are not in the SKU database
  var stockTransfers_Valid_SKUs_Only_ = stockTransfers.filter(e => allSKUs.filter(f => e[0] == f[0]).length != 0);
  // Place an apostrophe infront of every SKU to force the datatype to be String
  var stockTransfers_Valid_SKUs_Only = stockTransfers_Valid_SKUs_Only_.map(row => row.map(value => (Array.isArray(value)) ? "\'" + value[0] : value ));
  var num_InvalidQtyTransfers = stockTransfers_InvalidQty.length;
  var num_Valid_SKU_Transfers = stockTransfers_Valid_SKUs_Only.length;
  
  // Clear the content of the stock transfers unless there isn't any data on the adagio page
  if (lastRowOfData > START_ROW_ADAGIO_SHEET) 
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, STOCK_TRANSFER_COL, lastRowOfData, NUM_COLS).clearContent();
  
  // Set all of the transfers unless there aren't any
  if (num_InvalidQtyTransfers != 0 && num_Valid_SKU_Transfers)
  {
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, STOCK_TRANSFER_COL, num_InvalidQtyTransfers, NUM_COLS).setNumberFormat("@").setValues(stockTransfers_InvalidQty);
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET + num_InvalidQtyTransfers + 1, STOCK_TRANSFER_COL, num_Valid_SKU_Transfers, NUM_COLS).setNumberFormat("@").setValues(stockTransfers_Valid_SKUs_Only);
  }
  else if (num_InvalidQtyTransfers != 0)
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, STOCK_TRANSFER_COL, num_Valid_SKU_Transfers, NUM_COLS).setNumberFormat("@").setValues(stockTransfers_InvalidQty);
  else if (num_Valid_SKU_Transfers != 0)
    adagioSheet.getRange(START_ROW_ADAGIO_SHEET, STOCK_TRANSFER_COL, num_Valid_SKU_Transfers, NUM_COLS).setNumberFormat("@").setValues(stockTransfers_Valid_SKUs_Only);
}

/**
* This function imports all of the data and gathers all the stock transfers.
*
* @param {Sheet}      richImportSheet   The richmond import data sheet
* @param {Sheet}      parksImportSheet  The parksville import data sheet
* @param {Sheet}      rupertImportSheet The rupert import data sheet
* @param {String[][]} allSKUs           The set of all SKUs (optional argument)
* @author Jarren Ralf
*/
function stockTransfersAll(parksImportSheet, rupertImportSheet, allSKUs)
{
  var startTime = new Date().getTime(); // For the run time
  var isTritesTransfer = false; // When computing the stock transfers, ignore the transfers to and from Trites (400) locations
  var transfers = [], transfers_InvalidQty = []; // richImportSheet, parksImportSheet, rupertImportSheet;
  
  const             TIME_STAMP_ROW =  2;
  const       RECEIVED_DESCRIP_COL = 12; 
  const  ITEMS_TO_RICH_DESCRIP_COL = 16;
  const         STOCK_TRANSFER_COL = 17;
  const RICHMOND = 100; // Represents the locations
  const PARKS    = 200;
  const RUPERT   = 300;
  
  if (arguments.length === 0)
  {
    var [_, parksImportSheet, rupertImportSheet] = importAllData();
    var allSKUs = getSKUs(); 
  }
  
  // Generate all of the transfers between Richmond Store (Location: 100) and the other locations
  [transfers[0], transfers_InvalidQty[0]] = getStockTransfers( parksImportSheet,       RECEIVED_DESCRIP_COL, RICHMOND,    PARKS, isTritesTransfer);
  [transfers[1], transfers_InvalidQty[1]] = getStockTransfers( parksImportSheet,  ITEMS_TO_RICH_DESCRIP_COL,    PARKS, RICHMOND, isTritesTransfer);
  [transfers[2], transfers_InvalidQty[2]] = getStockTransfers(rupertImportSheet,       RECEIVED_DESCRIP_COL, RICHMOND,   RUPERT, isTritesTransfer);
  [transfers[3], transfers_InvalidQty[3]] = getStockTransfers(rupertImportSheet,  ITEMS_TO_RICH_DESCRIP_COL,   RUPERT, RICHMOND, isTritesTransfer);

  setStockTransfersAll(transfers, transfers_InvalidQty, allSKUs); // Set all stock transfers
  timeStamp(TIME_STAMP_ROW, STOCK_TRANSFER_COL);                  // Stock Transfer time stamp
  setElapsedTime(startTime);                                      // To check the ellapsed times
}

/**
* This function creates a formatted date string for the current time and places the timestamp on the Adagio page.
*
* @param {Number} row The   row  number of the timestamp
* @param {Number} col The column number of the timestamp
* @author Jarren Ralf
*/
function timeStamp(row, col)
{
  var   spreadsheet = SpreadsheetApp.getActive()
  var   adagioSheet = spreadsheet.getActiveSheet()
  var      timeZone = spreadsheet.getSpreadsheetTimeZone();
  var         today = new Date();
  var        format = "EEE, d MMM yyyy HH:mm:ss";
  var formattedDate = Utilities.formatDate(today, timeZone, format);
  
  if (arguments.length !== 0) adagioSheet.getRange(row, col).setValue(formattedDate);
  return formattedDate;
}

/**
* Transpose an array.
*
* @param {Object[][]} a - A two dimensional array
* @return The transpose of the imputted two dimensional array
*/
function transpose(a)
{
  return Object.keys(a[0]).map(c => a.map(r => r[c]));
}