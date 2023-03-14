/**
* This function gets some of the physical counts done by a particular location based on which google sheet is being analyzed.
*
* @param  {String} sheetName : The name of the imported data sheet
* @return {String[][]} The list of SKUs, quantities, and the sheet name the data comes from
* @author Jarren Ralf
*/
function getPhysicalCounted(sheetName)
{
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const numRows = sheet.getLastRow() - 2;
  const invalidCounts = [];  
  var values_infoCounts = [], values_manualCounts = [], values_order = [], values_shipped = []; 

  if (sheetName !== 'Imported Richmond Data (Loc: 100)')
  {
    const numRows_order = getLastRowSpecial(sheet.getSheetValues(4, 1, numRows, 1));
    const numRows_shipped = getLastRowSpecial(sheet.getSheetValues(5, 1, numRows, 1));
    
    var col = 9;

    var values_order = sheet.getSheetValues(4, 1, numRows_order, 4).filter(v => {
      if (isNotBlank(v[3]))
      {
        if (isNumber(v[3]))
        {
          if ((v[1] !== 'B/O') && (v[3] != v[2]))
            return true
          else
          {
            invalidCounts.push([v[0], v[3], '', 'Order'])
            return false;
          }
        }
        else if (v[3] !== 'x')
        {
          invalidCounts.push([v[0], v[3], '', 'Order'])
          return false;
        }
        else
          return false
      }
      else
        return false;
    }).map(w => [w[0], w[3], '', 'Order']);

    var values_shipped = sheet.getSheetValues(4, 5, numRows_shipped, 3).filter(v => {
      if (isNotBlank(v[2]))
      {
        if (isNumber(v[2]))
        {
          if (v[2] != v[1])
            return true
          else
          {
            invalidCounts.push([v[0], v[2], '', 'Shipped'])
            return false;
          }
        }
        else if (v[2] !== 'x')
        {
          invalidCounts.push([v[0], v[2], '', 'Shipped'])
          return false;
        }
        else
          return false
      }
      else
        return false;
    }).map(w => [w[0], w[2], '', 'Shipped']);
  }
  else
    var col = 2;

  const infoCounts_qty = sheet.getSheetValues(4, col, numRows, 1);
  const manualCounts_qty = sheet.getSheetValues(4, col + 2, numRows, 1);

  if (getLastRowSpecial(infoCounts_qty))
  { 
    const numRows_infoCounts = getLastRowSpecial(sheet.getSheetValues(4, col - 1, numRows, 1));
    values_infoCounts = sheet.getSheetValues(4, col - 1, numRows_infoCounts, 2).filter(v => {
      if (isNotBlank(v[1]))
      {
        if (isNumber(v[1]))
          return true
        else
        {
          invalidCounts.push([v[0], v[1], '', 'InfoCounts'])
          return false
        }
      }
      else
        return false
    }).map(w => [w[0], w[1], '', 'InfoCounts']);
  }

  if (getLastRowSpecial(manualCounts_qty))
  {
    const numRows_manualCounts = getLastRowSpecial(sheet.getSheetValues(4, col + 1, numRows, 1));
    values_manualCounts = sheet.getSheetValues(4, col + 1, numRows_manualCounts, 2).filter(v => {
      if (isNotBlank(v[1]))
      {
        if (isNumber(v[1]))
        {
          if (v[0] !== 'MAKE_NEW_SKU')
            return true
          else
          {
            invalidCounts.push([v[0], v[1], '', 'InfoCounts'])
            return false
          }
        }
        else
        {
          invalidCounts.push([v[0], v[1], '', 'InfoCounts'])
          return false
        }
      }
      else
        return false
    }).map(w => [w[0], w[1], '', 'Manual Counts']);
  }

  if (invalidCounts.length)
  {
    invalidCounts.push(['', '', '', ''])
    return invalidCounts.concat(values_order, values_shipped, values_infoCounts, values_manualCounts)
  }
  else
  {  
    const counts = values_order.concat(values_shipped, values_infoCounts, values_manualCounts)

    if (counts.length)
      return counts
  }
}

/**
* This function gets all of the SKUs so that it only has to be done once in a single runAll execution.
*
* @param {Spreadsheet}spreadsheet : The active spreadsheet.
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
function getStockTransfers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetNames = ['Imported Parksville Data (Loc: 200)', 'Imported Rupert Data (Loc: 300)'];
  const store_location = ['200', '300']
  const all_skus = getSKUs(spreadsheet)
  const invalidTransfers = [];
  var sheet, numRows, numRows_Received, numRows_ItemsToRichmond, values_Received = [[], []], values_ItemsToRichmond = [[], []];

  for (var s = 0; s < sheetNames.length; s++)
  {
    sheet = spreadsheet.getSheetByName(sheetNames[s]);
    numRows = sheet.getLastRow() - 2;
    numRows_Received = getLastRowSpecial(sheet.getSheetValues(4, 12, numRows, 1));
    numRows_ItemsToRichmond = getLastRowSpecial(sheet.getSheetValues(4, 15, numRows, 1));
    
    values_Received[s] = sheet.getSheetValues(4, 12, numRows_Received, 3).filter(v => {
      if (v[2] !== 'true')
      {
        if (isNotBlank(v[1]) && isNumber(v[1]))
        {
          if (all_skus.includes(v[0]))
            return true
          else
          {
            invalidTransfers.push([v[0], v[1], '100', store_location[s]])
            return false;
          }
        }
        else
        {
          invalidTransfers.push([v[0], v[1], '100', store_location[s]])
          return false;
        }
      }
      else
        return false
    }).map(w => [w[0], w[1], '100', store_location[s]]);

    values_ItemsToRichmond[s] = sheet.getSheetValues(4, 15, numRows_ItemsToRichmond, 4).filter(v => {
      if (v[3] !== 'true' && isNotBlank(v[2]))
      {
        if (isNotBlank(v[1]) && isNumber(v[1]))
        {
          if (all_skus.includes(v[0]))
            return true
          else
          {
            invalidTransfers.push([v[0], v[1], store_location[s], '100'])
            return false;
          }
        }
        else
        {
          invalidTransfers.push([v[0], v[1], store_location[s], '100'])
          return false;
        }
      }
      else
        return false
    }).map(w => [w[0], w[1], store_location[s], '100']);
  }

  if (invalidTransfers.length)
  {
    invalidTransfers.push(['', '', '', ''])
    return invalidTransfers.concat(values_Received[0], values_ItemsToRichmond[0], values_Received[1], values_ItemsToRichmond[1])
  }
  else
  {  
    const transfers = values_Received[0].concat(values_ItemsToRichmond[0], values_Received[1], values_ItemsToRichmond[1])
    
    if (transfers.length)
      return transfers
  }
}

/**
* This function sets the IMPORTRANGE fromulas and then pastes the values on the page, which removes the formula.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @return Returns all of the data pages
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
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"))))))",                                                            // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!C4:C\")",     // Richmond, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"))))))",                                                         // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!C4:C\")"]],  // Richmond, Counts
    [["=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!E4:E\"))))))",                                                                                                                 // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!J4:J\")",          // Parksville, BO
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!G4:G\")",          // Parksville, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Order!H4:H\")",          // Parksville, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!E4:E\"))))))",                                                                                                               // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!G4:G\")",        // Parksville, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Shipped!H4:H\")",        // Parksville, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!A4:A\"))))))",                                                            // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"InfoCounts!C4:C\")",     // Parksville, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!A4:A\"))))))",                                                         // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Manual Counts!C4:C\")",  // Parksville, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!E4:E\"))))))",                                                              // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!I4:I\")",       // Parksville, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"Received!L4:L\")",       // Parksville, Transfered
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!D4:D\"))))))",                                                       // Parksville, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!F4:F\")",// Parksville, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!H4:H\")",// Parksville, Received By
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM\", \"ItemsToRichmond!I4:I\")"]],// Parksville, Transfered
    [["=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!E4:E\"))))))",                                                                                                                 // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!J4:J\")",          // Rupert, BO
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!G4:G\")",          // Rupert, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Order!H4:H\")",          // Rupert, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!E4:E\"))))))",                                                                                                               // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!G4:G\")",        // Rupert, Current Stock
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Shipped!H4:H\")",        // Rupert, Actual Stock
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!A4:A\"))))))",                                                            // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"InfoCounts!C4:C\")",     // Rupert, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!A4:A\"))))))",                                                         // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Manual Counts!C4:C\")",  // Rupert, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!E4:E\"))))))",                                                              // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!I4:I\")",       // Rupert, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"Received!L4:L\")",       // Rupert, Transfered
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!D4:D\"))))))",                                                       // Rupert, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!F4:F\")",// Rupert, Shipped
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!H4:H\")",// Rupert, Received By
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM\", \"ItemsToRichmond!I4:I\")" // Rupert, Transfered
  ]]]

  var numCols, range, newValues;

  for (var sheet = 0; sheet < dataSheets.length; sheet++)
  {
    numCols = dataSheets[sheet].getLastColumn();
    dataSheets[sheet].getRange(4, 1, dataSheets[sheet].getLastRow() - 3, numCols).clearContent().offset(0, 0, 1, numCols).setValues(formulas[sheet])
    range = dataSheets[sheet].getDataRange();
    newValues = range.getValues()
    range.setNumberFormat('@').setValues(newValues)
  }

  return dataSheets;
}

/**
 * This function checks if the given string is not blank.
 * 
 * @param {String} str : The given string.
 * @return {Boolean} Returns true if the given string is blank, false otherwise.
 */
function isNotBlank(str)
{
  return str !== ''
}

/**
* This function imports all of the data, gathers the counts for all locations, and reports all of the stock transfers.
*
* @author Jarren Ralf
*/
function testingNew_runAll()
{ 
  var startTime = new Date().getTime(); // For the run time

  getPhysicalCounted('Imported Richmond Data (Loc: 100)')
  getPhysicalCounted('Imported Parksville Data (Loc: 200)')
  getPhysicalCounted('Imported Rupert Data (Loc: 300)')
  getStockTransfers()

  timeStamp(2, 19);          // Run all timestamp
  setElapsedTime(startTime); // To check the ellapsed times
  
  const PHYS_COUNT_COLS = [2, 7, 12]; // Richmond (100), Parksville (200), Rupert (300), Trites (400)
  

  var richmondImportSheet, parksImportSheet, rupertImportSheet;
  [richmondImportSheet, parksImportSheet, rupertImportSheet] = importAllData();
  
  stockTransfersAll(parksImportSheet, rupertImportSheet, allSKUs)  // Get and set all of the stock transfers
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
* @param {Number} startTime : The start time that the script began running at represented by a number in milliseconds
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @author Jarren Ralf
*/
function setElapsedTime(startTime, spreadsheet)
{
  spreadsheet.getSheetByName('Adagio Transfer Sheet').getRange(1, 11).setValue((new Date().getTime() - startTime)/1000);
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
* @param {Number} row : The   row  number of the timestamp
* @param {Number} col : The column number of the timestamp
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