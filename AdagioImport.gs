/**
* This function gets the of the physical counts done by a particular location based on which google sheet is being analyzed.
*
* @param  {Sheet} sheet : The imported data sheet
* @return {String[][]} The list of SKUs, quantities, and the original sheet name the data comes from
* @author Jarren Ralf
*/
function getPhysicalCounted(sheet)
{
  const numRows = sheet.getLastRow() - 2;
  const invalidCounts = [];  
  var values_infoCounts = [], values_manualCounts = [], values_order = [], values_shipped = [], values_UoM_Conversion = [], values_assembly = [];

  if (sheet.getSheetName() !== 'Imported Richmond Data (Loc: 100)') // Parksville and Rupert spreadsheets also have an Order and Shipped page that need to be analyzed
  {
    const numRows_order = getLastRowSpecial(sheet.getSheetValues(4, 1, numRows, 1));
    const numRows_shipped = getLastRowSpecial(sheet.getSheetValues(5, 1, numRows, 1));
    var col = 9;

    var values_order = sheet.getSheetValues(4, 1, numRows_order, 4).filter(v => {
      if (isNotBlank(v[3])) // Check if the quantity is blank
      {
        if (isNumber(v[3])) // Check if the quantity is a valid numeral
        {
          if ((v[1] !== 'B/O') && (v[3] != v[2])) // Check if the current stock equals the counted stock (No change) and if it's B/O, because this line will be duplicated on the shipped page
            return true
          else
          {
            invalidCounts.push([v[0], v[3], '', 'Order'])
            return false;
          }
        }
        else if (v[3] !== 'x') // An 'x' is placed in the actual counts column when the information is imported into Adagio
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
      if (isNotBlank(v[2])) // Check if the quantity is blank
      {
        if (isNumber(v[2])) // Check if the quantity is a valid numeral
        {
          if (v[2] != v[1]) // Check if the current stock equals the counted stock (No change)
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
  {
    const numRows_UoM_Conversion = getLastRowSpecial(sheet.getSheetValues(4, 6, numRows, 1));
    const numRows_assembly = getLastRowSpecial(sheet.getSheetValues(4, 8, numRows, 1));
    var col = 2;

    if (numRows_UoM_Conversion) // If false, that would mean that lastRow is zero and hence there are no counts to worry about
    { 
      values_UoM_Conversion = sheet.getSheetValues(4, 5, numRows_UoM_Conversion, 2).filter(v => {
        if (Number(v[1]) >= 0) // Check if the quantity non negative
          return true
        else
        {
          invalidCounts.push([v[0], v[1], '', 'UoM Conversion'])
          return false
        }
      }).map(w => [w[0], w[1], '', 'UoM Conversion']);
    }

    if (numRows_assembly) // If false, that would mean that lastRow is zero and hence there are no counts to worry about
    { 
      values_assembly = sheet.getSheetValues(4, 7, numRows_assembly, 2).filter(v => {
        if (Number(v[1]) >= 0) // Check if the quantity non negative
          return true
        else
        {
          invalidCounts.push([v[0], v[1], '', 'Assembly'])
          return false
        }
      }).map(w => [w[0], w[1], '', 'Assembly']);
    }
  }

  const infoCounts_qty = sheet.getSheetValues(4, col, numRows, 1);
  const manualCounts_qty = sheet.getSheetValues(4, col + 2, numRows, 1);

  if (getLastRowSpecial(infoCounts_qty)) // If false, that would mean that lastRow is zero and hence there are no counts to worry about
  { 
    const numRows_infoCounts = getLastRowSpecial(sheet.getSheetValues(4, col - 1, numRows, 1));
    values_infoCounts = sheet.getSheetValues(4, col - 1, numRows_infoCounts, 2).filter(v => {
      if (isNotBlank(v[1])) // Check if the quantity is blank
      {
        if (isNumber(v[1])) // Check if the quantity is a valid numeral
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

  if (getLastRowSpecial(manualCounts_qty)) // If false, that would mean that lastRow is zero and hence there are no counts to worry about
  {
    const numRows_manualCounts = getLastRowSpecial(sheet.getSheetValues(4, col + 1, numRows, 1));
    values_manualCounts = sheet.getSheetValues(4, col + 1, numRows_manualCounts, 2).filter(v => {
      if (isNotBlank(v[1])) // Check if the quantity is blank
      {
        if (isNumber(v[1])) // Check if the quantity is a valid numeral
        {
          if (v[0] !== 'MAKE_NEW_SKU') // Check if the sku is 'MAKE_NEW_SKU' and ignore
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

  if (invalidCounts.length) // Check if there are any invalid counts
  {
    invalidCounts.push(['', '', '', '']) // This row is used to separated valid counts from the invalid and therefore it allows the user use the ctrl + A command to select the valid counts
    return invalidCounts.concat(values_UoM_Conversion, values_assembly, values_order, values_shipped, values_infoCounts, values_manualCounts)
  }
  else
  {  
    const counts = values_UoM_Conversion.concat(values_assembly, values_order, values_shipped, values_infoCounts, values_manualCounts)

    if (counts.length)
      return counts
  }
}

/**
* This function gets all of the SKUs and puts them into a single array.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @return {String[]} The list of all SKUs
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
* @param  {Sheet[]} sheets : The sheets with the imported data
* @param  {Spreadsheet} spreadsheet : The active spreadsheet
* @return {String[][]} The stock transfers with skus, qtys, and locations
* @author Jarren Ralf
*/
function getStockTransfers(sheets, spreadsheet)
{
  const store_location = ['200', '300']
  const all_skus = getSKUs(spreadsheet) // An array of every sku in the Adagio database
  const invalidTransfers = [];
  var numRows, numRows_Received, numRows_ItemsToRichmond, values_Received = [[], []], values_ItemsToRichmond = [[], []];

  for (var s = 0; s < sheets.length; s++) // Loop through 2 sheets, for parksville and rupert data
  {
    numRows = sheets[s].getLastRow() - 2;
    numRows_Received = getLastRowSpecial(sheets[s].getSheetValues(4, 12, numRows, 1));
    numRows_ItemsToRichmond = getLastRowSpecial(sheets[s].getSheetValues(4, 15, numRows, 1));
    
    values_Received[s] = sheets[s].getSheetValues(4, 12, numRows_Received, 3).filter(v => {
      if (v[2] !== 'true') // Check if the transfer has already been updated in Adagio inventory, false := not updated
      {
        if (isNotBlank(v[1]) && isNumber(v[1])) // Check if the quantity is not blank and a valid numeral
        {
          if (all_skus.includes(v[0])) // Check if the particular sku is found in the Adagio database
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

    values_ItemsToRichmond[s] = sheets[s].getSheetValues(4, 15, numRows_ItemsToRichmond, 4).filter(v => {
      if (v[3] !== 'true' && isNotBlank(v[2])) // Check if the transfer has already been updated in Adagio inventory, false := not updated and check if the item has been received
      {
        if (isNotBlank(v[1]) && isNumber(v[1])) // Check if the quantity is not blank and a valid numeral
        {
          if (all_skus.includes(v[0])) // Check if the particular sku is found in the Adagio database
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
    invalidTransfers.push(['', '', '', '']) // This row is used to separated valid transfers from the invalid and therefore it allows the user use the ctrl + A command to select the valid transfers
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
* @return {Sheet[]} Returns all of the data sheets
* @author Jarren Ralf
*/
function importData(spreadsheet)
{
  // When clicking the Import Data buttons, the function will be run with 0 arguments
  if (arguments.length == 0)
    var spreadsheet = SpreadsheetApp.getActive()

  const dataSheets = [spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'), 
                      spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'),
                      spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)')
  ]

  const formulas = [[[
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!A4:A\"))))))",                                                            // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"InfoCounts!C4:C\")",     // Richmond, Counts
    "=ARRAYFORMULA(IF(NOT(REGEXMATCH(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"),\" - \")),TRIM(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\")),TRIM(LEFT(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"), FIND(\" - \", IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!A4:A\"))))))",                                                         // Richmond, SKU
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Manual Counts!C4:C\")",  // Richmond, Counts
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"UoM Conversion!A1:B\")",  // Richmond, SKU
    "",
    "=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk\", \"Assembly!A1:B\")", ""]], // Richmond, SKU
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

  for (var sheet = 0; sheet < dataSheets.length; sheet++) // Loop through the data sheets
  {
    numCols = dataSheets[sheet].getLastColumn();
    dataSheets[sheet].getRange(4, 1, dataSheets[sheet].getLastRow() - 2, numCols).clearContent().offset(0, 0, 1, numCols).setValues(formulas[sheet]) // Clear the content and set the formulas
    range = dataSheets[sheet].getDataRange();
    newValues = range.getValues()
    range.setNumberFormat('@').setValues(newValues) // Paste the data values, hence remove the formulas 
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
* This function imports the Richmond data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountRich()
{
  const startTime = new Date().getTime(); // For the run time
  const spreadsheet = SpreadsheetApp.getActive()
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet')
  const numRows = adagioSheet.getLastRow() - 23;
  var richmondImportSheet;
  [richmondImportSheet,,] = importData(spreadsheet)

  if (numRows > 0)
    adagioSheet.getRange(24, 2, numRows, 4).clearContent()

  const richCounts = getPhysicalCounted(richmondImportSheet)
  if (richCounts != null)
    adagioSheet.getRange(24, 2, richCounts.length, 4).setValues(richCounts)

  timeStamp(spreadsheet, 2, 2, adagioSheet);
  setElapsedTime(startTime, adagioSheet);
}

/**
* This function imports the Parksville data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountPark()
{  
  const startTime = new Date().getTime(); // For the run time
  const spreadsheet = SpreadsheetApp.getActive()
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet')
  const numRows = adagioSheet.getLastRow() - 23;
  var parksImportSheet;
  [, parksImportSheet,] = importData(spreadsheet)

  if (numRows > 0)
    adagioSheet.getRange(24, 7, numRows, 4).clearContent()

  const parksCounts = getPhysicalCounted(parksImportSheet)
  if (parksCounts != null)
    adagioSheet.getRange(24, 7, parksCounts.length, 4).setValues(parksCounts)

  timeStamp(spreadsheet, 2, 7, adagioSheet);
  setElapsedTime(startTime, adagioSheet);// To check the ellapsed times
}

/**
* This function imports the Rupert data and gathers the phyical counts.
*
* @author Jarren Ralf
*/
function physCountRupt()
{
  const startTime = new Date().getTime(); // For the run time
  const spreadsheet = SpreadsheetApp.getActive()
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet')
  const numRows = adagioSheet.getLastRow() - 23;
  var rupertImportSheet;
  [,, rupertImportSheet] = importData(spreadsheet)

  if (numRows > 0)
    adagioSheet.getRange(24, 12, numRows, 4).clearContent()

  const ruptCounts = getPhysicalCounted(rupertImportSheet)
  if (ruptCounts != null)
    adagioSheet.getRange(24, 12, ruptCounts.length, 4).setValues(ruptCounts)

  timeStamp(spreadsheet, 2, 12, adagioSheet);
  setElapsedTime(startTime, adagioSheet);// To check the ellapsed times
}

/**
* This function imports all of the data, displays the counts for all locations, and displays all of the stock transfers.
*
* @author Jarren Ralf
*/
function runAll()
{ 
  const startTime = new Date().getTime(); // For the run time
  const spreadsheet = SpreadsheetApp.getActive()
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
  const numRows = adagioSheet.getLastRow() - 23;
  var richmondImportSheet, parksImportSheet, rupertImportSheet;
  [richmondImportSheet, parksImportSheet, rupertImportSheet] = importData(spreadsheet)

  if (numRows > 0)
    adagioSheet.getRange(24, 2, numRows, adagioSheet.getLastColumn() - 1).clearContent()

  const richCounts = getPhysicalCounted(richmondImportSheet)
  if (richCounts != null)
    adagioSheet.getRange(24, 2, richCounts.length, 4).setValues(richCounts)
  timeStamp(spreadsheet, 2, 2, adagioSheet);

  const parksCounts = getPhysicalCounted(parksImportSheet)
  if (parksCounts != null)
    adagioSheet.getRange(24, 7, parksCounts.length, 4).setValues(parksCounts)
  timeStamp(spreadsheet, 2, 7, adagioSheet);

  const ruptCounts = getPhysicalCounted(rupertImportSheet)
  if (ruptCounts != null)
    adagioSheet.getRange(24, 12, ruptCounts.length, 4).setValues(ruptCounts)
  timeStamp(spreadsheet, 2, 12, adagioSheet);

  const transfers = getStockTransfers([parksImportSheet, rupertImportSheet], spreadsheet)
  if (transfers != null)
    adagioSheet.getRange(24, 17, transfers.length, 4).setValues(transfers)
  timeStamp(spreadsheet, 2, 17, adagioSheet);

  timeStamp(spreadsheet, 2, 19, adagioSheet); // Run all timestamp
  setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
}

/**
* This function sets the ellapsed time of a function and prints it on the Adagio page.
*
* @param {Number} startTime : The start time that the script began running at represented by a number in milliseconds
* @param {Sheet} adagioSheet : The Adagio Transfer Sheet
* @author Jarren Ralf
*/
function setElapsedTime(startTime, adagioSheet)
{
  adagioSheet.getRange(1, 9).setValue((new Date().getTime() - startTime)/1000 + ' seconds');
}

/**
* This function imports all of the data and gathers all the stock transfers.
*
* @author Jarren Ralf
*/
function stockTransfers()
{
  const startTime = new Date().getTime(); // For the run time
  const spreadsheet = SpreadsheetApp.getActive()
  const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
  const numRows = adagioSheet.getLastRow() - 23;
  var parksImportSheet, rupertImportSheet;
  [, parksImportSheet, rupertImportSheet] = importData(spreadsheet)

  if (numRows > 0)
    adagioSheet.getRange(24, 17, numRows, 4).clearContent()

  const transfers = getStockTransfers([parksImportSheet, rupertImportSheet], spreadsheet)
  if (transfers != null)
    adagioSheet.getRange(24, 17, transfers.length, 4).setValues(transfers)
  timeStamp(spreadsheet, 2, 17, adagioSheet);
  setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
}

/**
* This function creates a formatted date string for the current time and places the timestamp on the Adagio page.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @param   {Number}        row     : The   row  number of the timestamp
* @param   {Number}        col     : The column number of the timestamp
* @param    {Sheet}       sheet    : The sheet to place the timestamp on
* @returns {String} Returns the formatted date string.
* @author Jarren Ralf
*/
function timeStamp(spreadsheet, row, col, sheet, format)
{
  if (arguments.length < 5)
    format = "EEE, d MMM yyyy HH:mm:ss";

  var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), format);
  
  if (arguments.length > 3) sheet.getRange(row, col).setValue(formattedDate);
  
  return formattedDate;
}