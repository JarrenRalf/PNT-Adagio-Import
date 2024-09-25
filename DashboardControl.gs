var ss; // Global variable used to control the richmond, parksville and prince rupert spreadsheets

function installedOnEdit(e)
{
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  if (sheetName === 'Search')
  {
    if (isSingleColumn && ((row == 1 && col == 2 && (rowEnd == 6 || rowEnd == 1)) || (isSingleRow && (
      (row == 1 && col == 4) || (row == 1 && col == 5) || (row == 2 && col == 5) || (row == 3 && col == 5) || (row == 4 && col == 3) || 
      (row == 4 && col == 5) || (row == 5 && col == 3) || (row == 5 && col == 5) || (row == 6 && col == 3) || (row == 6 && col == 5)))))
        searchV2(e, spreadsheet, sheet, row, col)
    else if (row != rowEnd && row > 8)
      pasteMultipleSKUsOnSearchPage(range, sheet, spreadsheet)
  }
  else if (sheetName === 'Adagio Transfer Sheet')
  {
    if (isSingleRow && isSingleColumn && range.isChecked())
    {
      spreadsheet.toast('Function Running...')

      if (col == 6) // Richmond
      {
        if (row == 15)
          richmond_applyFullSpreadsheetFormatting(range);
        else if (row == 16)
          richmond_updateRecentlyCreatedItems(range);
        else if (row == 17)
          richmond_countLog(range);
        else if (row == 18)
          richmond_updateSearchData(range);
      }
      else if (col == 11) // Parksville
      {
        if (row == 15)
          parksville_applyFullSpreadsheetFormatting(range);
        else if (row == 16)
          parksville_updateRecentlyCreatedItems(range);
        else if (row == 17)
          parksville_countLog(range);
        else if (row == 18)
          parksville_updateSearchData(range);
      }
      else if (col == 16) // Rupert
      {
        if (row == 15)
          rupert_applyFullSpreadsheetFormatting(range);
        else if (row == 16)
          rupert_updateRecentlyCreatedItems(range);
        else if (row == 17)
          rupert_countLog(range);
        else if (row == 18)
          rupert_updateSearchData(range);
      }
      
      spreadsheet.toast('Function Complete')
      range.uncheck();
    }
  }
  else if (sheetName === 'SKUsToWatch_ASSEMBLY')
  {
    if (isSingleRow && isSingleColumn && row === 1 && col === 7)
      scanAssemblies(range, sheet, spreadsheet)
    else if (isSingleColumn && row > 2 && (col === 2 || col === 8))
      stripDescriptions(range)
  }
  else if (sheetName === 'SKUsToWatch')
  {
    if (isSingleRow && isSingleColumn && row === 1 && col === 7)
      scanConversions(range, sheet, spreadsheet, col)
    else if (isSingleColumn && row > 2 && (col === 2 || col === 8))
      stripDescriptions(range)
  }
}

/**
 * This function adds a menu item to the spreadsheet with a quick way run one conversion based on the item that the user has selected on either 
 * the SKUsToWatch or SKUsToWatch_ASSEMBLY sheet.
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Conversion').addItem('Convert Selected Item', 'computeOneConversion').addToUi();
}

/**
 * This function provides a way for the user to run one conversion based on the item that the user has selected on either 
 * the SKUsToWatch or SKUsToWatch_ASSEMBLY sheet.
 * 
 * @author Jarren Ralf
 */
function computeOneConversion()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() === 'SKUsToWatch')
  {
    const range = sheet.getActiveRange()
    const row = range.getRow()
    const col = range.getColumn();
    const lastRow = sheet.getLastRow();

    if (row > 2 && row <= lastRow)
    {
      if (col < 4) // Converting into a smaller pack size
      {
        const packagedItems = sheet.getSheetValues(row, 1, 1, 9)
        const quantity = sheet.getSheetValues(1, 3, 1, 1)[0][0]

        if (isNotBlank(packagedItems[0][4])) // The conversion factor is NOT blank
        {
          if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
          {
            // Prompt the user to input how many packages were created
            const ui = SpreadsheetApp.getUi();
            const response = ui.prompt('How many ' + packagedItems[0][2] + ' -  ' + packagedItems[0][1] + ' have you created?')

            if (response.getSelectedButton() == ui.Button.OK)
            {
              const qty = Number(response.getResponseText()); // The user inputted quantity

              if (isNotBlank(qty) || !isNaN(qty)) // Valid number
              {
                const inventorySheet = spreadsheet.getSheetByName('DataImport');
                const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
                var convertedItems = []

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][0].toString().toUpperCase(), inventory[j][0] + qty])
                    break;
                  }
                }

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][6].toString().toUpperCase(), inventory[j][0] - qty/packagedItems[0][4]])
                    break;
                  }
                }

                const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
                const row = UoMConversionPage.getLastRow() + 1;
                const numRows = convertedItems.length;
                UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
                sheet.getRange(1, 1).setValue('Scan Here').offset(0, 2).setValue('0').activate()
                spreadsheet.toast('added to Richmond UoM Conversion sheet',  + qty + ' units of - ' + packagedItems[0][2] + ' - ' + packagedItems[0][1] + ' package conversions', 120)
              }
              else
              {
                ui.alert('Please scan item again and enter a valid quantity.')
                sheet.getRange(1, 1).setValue('Scan Here').offset(0, 2).activate()
              }
            }
            else
              sheet.getRange(1, 1).setValue('Scan Here').offset(0, 2).activate()
          }
          else
          {
            const inventorySheet = spreadsheet.getSheetByName('DataImport');
            const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
            var convertedItems = [];

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][0].toString().toUpperCase(), inventory[j][0] + quantity])
                break;
              }
            }

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][6].toString().toUpperCase(), inventory[j][0] - quantity/packagedItems[0][4]])
                break;
              }
            }

            const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
            const row = UoMConversionPage.getLastRow() + 1;
            const numRows = convertedItems.length;
            UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
            sheet.getRange(1, 1).setValue('Scan Here').offset(0, 2).setValue('0').activate()
            spreadsheet.toast('added to Richmond UoM Conversion sheet',  + quantity + ' units of - ' + packagedItems[0][2] + ' - ' + packagedItems[0][1] + ' package conversions', 120)
          }
        }
        else // The conversion factor is blank
        {
          sheet.getRange(1, 1).setValue('Scan Here').activate()
          spreadsheet.toast('', 'The SKU for the packaging you are converting from is missing.')
        }
      }
      else if (col > 6) // Converting into a larger pack size
      {
        const packagedItems = sheet.getSheetValues(row, 1, 1, 9)
        const quantity = sheet.getSheetValues(1, 9, 1, 1)[0][0]
        
        if (isNotBlank(packagedItems[0][4])) // The conversion factor is NOT blank
        {
          if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
          {
            // Prompt the user to input how many packages were created
            const ui = SpreadsheetApp.getUi();
            const response = ui.prompt('How many ' + packagedItems[0][8] + ' -  ' + packagedItems[0][7] + ' have you created?')

            if (response.getSelectedButton() == ui.Button.OK)
            {
              const qty = Number(response.getResponseText()); // The user inputted quantity

              if (isNotBlank(qty) || !isNaN(qty)) // Valid number
              {
                const inventorySheet = spreadsheet.getSheetByName('DataImport');
                const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
                var convertedItems = []

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][6].toString().toUpperCase(), inventory[j][0] + qty])
                    break;
                  }
                }

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][0].toString().toUpperCase(), inventory[j][0] - qty*packagedItems[0][4]])
                    break;
                  }
                }

                const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
                const row = UoMConversionPage.getLastRow() + 1;
                const numRows = convertedItems.length;
                UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
                sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).setValue('0').activate()
                spreadsheet.toast('added to Richmond UoM Conversion sheet',  + qty + ' units of - ' + packagedItems[0][8] + ' - ' + packagedItems[0][7] + ' package conversions', 120)
              }
              else
              {
                ui.alert('Please scan item again and enter a valid quantity.')
                sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).activate()
              }
            }
            else
              sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).activate()
          }
          else
          {
            const inventorySheet = spreadsheet.getSheetByName('DataImport');
            const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
            var convertedItems = [];

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][6].toString().toUpperCase(), inventory[j][0] + quantity])
                break;
              }
            }

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][0].toString().toUpperCase(), inventory[j][0] - quantity*packagedItems[0][4]])
                break;
              }
            }

            const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
            const row = UoMConversionPage.getLastRow() + 1;
            const numRows = convertedItems.length;
            UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
            sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).setValue('0').activate()
            spreadsheet.toast('added to Richmond UoM Conversion sheet',  + quantity + ' units of - ' + packagedItems[0][8] + ' - ' + packagedItems[0][7] + ' package conversions', 120)
          }
        }
        else // The conversion factor is blank
        {
          sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).activate()
          spreadsheet.toast('', 'The SKU for the packaging you are converting from is missing.')
        }
      }
      else
        Browser.msgBox('Please select either the smaller package or larger package item.')
    }
    else
      Browser.msgBox('Please select an item that you want to do a conversion with.')
  }
  else if (sheet.getSheetName() === 'SKUsToWatch_ASSEMBLY')
  {
    const range = sheet.getActiveRange()
    const row = range.getRow()
    const lastRow = sheet.getLastRow();

    if (row > 2 && row <= lastRow)
    {
      const sku = sheet.getSheetValues(row, 7, 1, 1)[0][0].toString().toUpperCase()

      if (isNotBlank(sku))
      {
        const assemblyItems = sheet.getSheetValues(3, 1, lastRow - 2, 8).filter(item => item[6] == sku)
        const quantity = sheet.getSheetValues(1, 9, 1, 1)[0][0]
      
        if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
        {
          // Prompt the user to input how many packages were created
          const ui = SpreadsheetApp.getUi();
          const response = ui.prompt('How many ' + assemblyItems[0][7] + ' assemblies have you created?')

          if (response.getSelectedButton() == ui.Button.OK)
          {
            const qty = Number(response.getResponseText()); // The user inputted quantity

            if (isNotBlank(qty) || !isNaN(qty)) // Valid number
            {
              const inventorySheet = spreadsheet.getSheetByName('DataImport');
              const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
              var assembledItems = []

              for (var j = 0; j < inventory.length; j++)
              {
                if (assemblyItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  assembledItems.push([assemblyItems[0][6].toString().toUpperCase(), inventory[j][0] + qty*assemblyItems[0][4]])
                  break;
                }
              }

              for (var j = 0; j < inventory.length; j++)
              {
                if (assemblyItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  assembledItems.push([assemblyItems[0][0].toString().toUpperCase(), inventory[j][0] - qty*assemblyItems[0][4]])
                  break;
                }
              }

              for (var i = 1; i < assemblyItems.length; i++)
              {
                for (var j = 0; j < inventory.length; j++)
                {
                  if (assemblyItems[i][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    assembledItems.push([assemblyItems[i][0].toString().toUpperCase(), inventory[j][0] - qty*assemblyItems[i][4]])
                    break;
                  }
                }
              }

              const assemblyPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('Assembly')
              const row = assemblyPage.getLastRow() + 1;
              const numRows = assembledItems.length;
              assemblyPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(assembledItems)
              sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).setValue('0').activate()
              spreadsheet.toast('added to Richmond Assembly sheet',  + qty + ' ' + assemblyItems[0][7] + ' assemblies', 120)
            }
            else
            {
              ui.alert('Please scan item again and enter a valid quantity.')
              sheet.getRange(1, 7).setValue('Scan Here').activate()
            }
          }
          else
            sheet.getRange(1, 7).setValue('Scan Here').activate()
        }
        else
        {
          const inventorySheet = spreadsheet.getSheetByName('DataImport');
          const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
          var assembledItems = []

          for (var j = 0; j < inventory.length; j++)
          {
            if (assemblyItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
            {
              assembledItems.push([assemblyItems[0][6].toString().toUpperCase(), inventory[j][0] + quantity*assemblyItems[0][4]])
              break;
            }
          }

          for (var j = 0; j < inventory.length; j++)
          {
            if (assemblyItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
            {
              assembledItems.push([assemblyItems[0][0].toString().toUpperCase(), inventory[j][0] - quantity*assemblyItems[0][4]])
              break;
            }
          }

          for (var i = 1; i < assemblyItems.length; i++)
          {
            for (var j = 0; j < inventory.length; j++)
            {
              if (assemblyItems[i][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                assembledItems.push([assemblyItems[i][0].toString().toUpperCase(), inventory[j][0] - quantity*assemblyItems[i][4]])
                break;
              }
            }
          }

          Logger.log(assembledItems)

          const assemblyPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('Assembly')
          const row = assemblyPage.getLastRow() + 1;
          const numRows = assembledItems.length;
          assemblyPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(assembledItems)
          sheet.getRange(1, 7).setValue('Scan Here').offset(0, 2).setValue('0').activate()
          spreadsheet.toast('added to Richmond Assembly sheet',  + quantity + ' ' + assemblyItems[0][7] + ' assemblies', 120)
        }
      }
      else
        Browser.msgBox('Please select an item that you want to do an assembly with.')
    }
    else
      Browser.msgBox('Please select an item that you want to do an assembly with.')
  }
  else
    Browser.msgBox('You must either be on the SKUsToWatch or SKUsToWatch_ASSEMBLY sheet to run this function.')
}

/**
 * This function moves the selected values on the item search sheet to the desired manual counts page.
 * 
 * @param {Sheet}   sheet   : The sheet that the selected items are being moved to.
 * @param {Number} startRow : The first row of the target sheet where the selected items will be moved to.
 * @param {Number}  numCols : The number of columns to grab from the item search page and move to the target sheet.
 * @author Jarren Ralf
 */
function copySelectedValuesV2(sheet, startRow, numCols)
{
  var  activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  var firstRows = [], lastRows = [], numRows = [], itemValues = [[[]]];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
     lastRows[r] = activeRanges[r].getLastRow()
  }
  
  var     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  var lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
  var finalDataRow = activeSheet.getLastRow() + 1;

  if (row > 8 && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
  {   
    for (var r = 0; r < activeRanges.length; r++)
    {
         numRows[r] = lastRows[r] - firstRows[r] + 1;
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 2, numRows[r], numCols);
    }
    
    var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
    var numItems = itemVals.length;

    if (numCols === 4) // Rupert
    {
      itemVals.map(u => u.splice(1, 2)); // Remove the richmond and parksville counts
      numCols -= 2;
    }
    else if (numCols === 3) // Parksville
    {
      itemVals.map(u => u.splice(1, 1)); // Remove the richmond counts column
      numCols--;
    }

    sheet.getRange(startRow, 1, numItems, numCols).setNumberFormat('@').setValues(itemVals); // Move the item values to the destination sheet
    applyFullRowFormatting(sheet, startRow, numItems, 7); // Apply the proper formatting
    sheet.getRange(startRow, 3).activate();            // Go to the quantity column on the destination sheet
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
}

/**
 * This function extends or compresses the dashboard based on whether the particular rows are hidden or not.
 * 
 * @author Jarren Ralf
 */
function extendDashboard()
{
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const range = sheet.getRange('O2');

  if (sheet.isRowHiddenByUser(15))
  {
    sheet.unhideRow(sheet.getRange('A15:19'))
    range.setValue('Compress Dashboard')
  }
  else
  {
    sheet.hideRows(15, 5);
    range.setValue('Extend Dashboard')
  } 
}

/**
 * This function gets the items that have a equal or less inventory in Richmond (Adagio inventory system **combines Moncton Street and Trites) 
 * than Trites (inFlow inventory system) and sets it on the TritesCounts page.
 * 
 * @author Jarren Ralf
 */
function getTritesCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const inventorySheet = spreadsheet.getSheetByName("DataImport");
    const tritesCountsSheet = spreadsheet.getSheetByName("Trites Counts");
    const adagioSheet = spreadsheet.getSheetByName("Adagio Transfer Sheet");
    const output = inventorySheet.getSheetValues(2, 2, inventorySheet.getLastRow() - 1, 5).filter(e => (isNotBlank(e[4])) ? e[1] < e[4] : false).map(f => [f[0], f[1], f[4]])
    const numItems = output.length;

    tritesCountsSheet.getRange('A4:C').clearContent()
      .offset(-3, 1, 1, 1).setValues([[numItems]])
      .offset(3, -1, numItems, 3).setValues(output)

    applyFullRowFormatting(tritesCountsSheet, 4, numItems);
    timeStamp(spreadsheet, 10, 5, adagioSheet, "dd MMM HH:mm")
    
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
* This function moves all of the selected values on the tritesCounts page to the Manual Counts page
*
* @author Jarren Ralf
*/
function manualCounts_FromTritesCounts()
{
  const QTY_COL = 4;
  const NUM_COLS = 2;
  
  var manualCountsSheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561').getSheetByName("Manual Counts");
  var lastRow = manualCountsSheet.getLastRow();
  var startRow = (lastRow < 3) ? 4 : lastRow + 1;

  copySelectedValues(manualCountsSheet, startRow, NUM_COLS, QTY_COL, true);
}

/**
 * This function allows the user to paste a list of SKUs in the Item Search and the full description and inventory for each item is populated in the sheet.
 * 
 * @param {Range} range : The active range.
 * @param {Sheet} sheet : The active sheet (Item Search)
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function pasteMultipleSKUsOnSearchPage(range, sheet, spreadsheet)
{
  const startTime = new Date().getTime();
  const values = range.getValues().filter(blank => isNotBlank(blank[0]))

  if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
  {
    const inventorySheet = spreadsheet.getSheetByName('DataImport');
    const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 7);
    var someSKUsNotFound = false, skus;

    if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
    {
      skus = values.map(item => {
      
        for (var i = 0; i < data.length; i++)
        {
          if (data[i][6] == item[0].toString().split(' - ').pop().toUpperCase())
            return data[i]
        }

        someSKUsNotFound = true;

        return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), '', '', '', '', '']
      });
    }
    else if (values[0][0].toString().includes('-'))
    {
      skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
      
        for (var i = 0; i < data.length; i++)
        {
          if (data[i][6] == item.toString().toUpperCase())
            return data[i]
        }

        someSKUsNotFound = true;

        return ['SKU Not Found:', item, '', '', '', '', '']
      });
    }
    else
    {
      skus = values.map(item => {
      
        for (var i = 0; i < data.length; i++)
        {
          if (data[i][6] == item[0].toString().toUpperCase())
            return data[i]
        }

        someSKUsNotFound = true;

        return ['SKU Not Found:', item[0], '', '', '', '', '']
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
      const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
      const numItems = items.length
      const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center'])
      const WHITE = new Array(7).fill('white')
      const YELLOW = new Array(7).fill('#ffe599')
      const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

      sheet.getRange(9, 1, sheet.getMaxRows() - 2, 7).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
        .offset(0, 0, numItems, 7)
          .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours)
          .setBorder(false, null, false, null, false, false).setValues(items)
        .offset(-8, 0, 1, 1).setValue((numSkusFound !== 1) ? numSkusFound + " results found." : numSkusFound + " result found.")
        .offset( 6, 0, 1, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds")
        .offset(numSkusNotFound + 2, 0, numSkusFound, 7).activate()
    }
    else // All SKUs were succefully found
    {
      const numItems = skus.length
      const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center'])

      sheet.getRange(9, 1, sheet.getMaxRows() - 2, 7).clearContent().setBackground('white').setFontColor('black')
        .offset(0, 0, numItems, 7)
          .setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments)
          .setBorder(false, null, false, null, false, false).setValues(skus).activate()
        .offset(-8, 0, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
        .offset( 6, 0, 1, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds");
    }
  }
}

/**
 * This function is run when a barcode scan is detected on the SKUsToWatch_ASSEMBLY sheet and the user is intending to increase the inventory of an item that we assemble while 
 * decreasing the inventory of the component SKUs we used to build the product. 
 * 
 * @param    {Range}       range    : The active range that the user has scanned a barcode into.
 * @param    {Sheet}       sheet    : The sheet that the user has scanned a barcode into.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function scanAssemblies(range, sheet, spreadsheet)
{
  spreadsheet.toast('Searching UPC database...', '', 30)
  const item = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString()).find(item => item[0] == range.getValue())

  if (item != null)
  {
    spreadsheet.toast('', 'Item Found in UPC Database...')
    const quantity = range.offset(0, 2).getValue()
    const assemblyItems = range.offset(2, -6, sheet.getLastRow() - 2, 8).getValues().filter(sku => sku[6] == item[1].toString().toUpperCase())
    
    if (assemblyItems != null)
    {
      if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
      {
        // Prompt the user to input how many packages were created
        const ui = SpreadsheetApp.getUi();
        const response = ui.prompt('How many ' + assemblyItems[0][7] + ' assemblies have you created?')

        if (response.getSelectedButton() == ui.Button.OK)
        {
          const qty = Number(response.getResponseText()); // The user inputted quantity

          if (isNotBlank(qty) || !isNaN(qty)) // Valid number
          {
            const inventorySheet = spreadsheet.getSheetByName('DataImport');
            const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
            var assembledItems = []

            for (var j = 0; j < inventory.length; j++)
            {
              if (assemblyItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                assembledItems.push([assemblyItems[0][6].toString().toUpperCase(), inventory[j][0] + qty*assemblyItems[0][4]])
                break;
              }
            }

            for (var j = 0; j < inventory.length; j++)
            {
              if (assemblyItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                assembledItems.push([assemblyItems[0][0].toString().toUpperCase(), inventory[j][0] - qty*assemblyItems[0][4]])
                break;
              }
            }

            for (var i = 1; i < assemblyItems.length; i++)
            {
              for (var j = 0; j < inventory.length; j++)
              {
                if (assemblyItems[i][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  assembledItems.push([assemblyItems[i][0].toString().toUpperCase(), inventory[j][0] - qty*assemblyItems[i][4]])
                  break;
                }
              }
            }

            const assemblyPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('Assembly')
            const row = assemblyPage.getLastRow() + 1;
            const numRows = assembledItems.length;
            assemblyPage.getRange(row, 1, numRows, 2).setValues(assembledItems)
            range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
            spreadsheet.toast('added to Richmond Assembly sheet',  + qty + ' ' + assemblyItems[0][7] + ' assemblies', 120)
          }
          else
          {
            ui.alert('Please scan item again and enter a valid quantity.')
            range.setValue('Scan Here').activate()
          }
        }
        else
          range.setValue('Scan Here').activate()
      }
      else
      {
        const inventorySheet = spreadsheet.getSheetByName('DataImport');
        const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
        var assembledItems = []

        for (var j = 0; j < inventory.length; j++)
        {
          if (assemblyItems[0][6].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
          {
            assembledItems.push([assemblyItems[0][6].toString().toUpperCase(), inventory[j][0] + quantity*assemblyItems[0][4]])
            break;
          }
        }

        for (var j = 0; j < inventory.length; j++)
        {
          if (assemblyItems[0][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
          {
            assembledItems.push([assemblyItems[0][0].toString().toUpperCase(), inventory[j][0] - quantity*assemblyItems[0][4]])
            break;
          }
        }

        for (var i = 1; i < assemblyItems.length; i++)
        {
          for (var j = 0; j < inventory.length; j++)
          {
            if (assemblyItems[i][0].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
            {
              assembledItems.push([assemblyItems[i][0].toString().toUpperCase(), inventory[j][0] - quantity*assemblyItems[i][4]])
              break;
            }
          }
        }

        const assemblyPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('Assembly')
        const row = assemblyPage.getLastRow() + 1;
        const numRows = assembledItems.length;
        assemblyPage.getRange(row, 1, numRows, 2).setValues(assembledItems)
        range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
        spreadsheet.toast('added to Richmond Assembly sheet',  + quantity + ' ' + assemblyItems[0][7] + ' assemblies', 120)
      }
    }
    else
    {
      range.setValue('Assembled Item Not Found').activate()
      spreadsheet.toast('', 'Assembled Item Not Found')
    }
  }
  else
  {
    range.setValue('Barcode Not Found').activate()
    spreadsheet.toast('', 'Barcode Not Found')
  }
}

/**
 * This function is run when a barcode scan is detected on the SKUsToWatch sheet and the user is intending to convert inventory of the same item from one
 * packaging size to another. The user is able to create packages of smaller sizes or larger sizes, depending on which column they scan the barcode into.
 * 
 * @param    {Range}       range    : The active range that the user has scanned a barcode into.
 * @param    {Sheet}       sheet    : The sheet that the user has scanned a barcode into.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param    {Number}       col     : The column number that the scan took place in.
 * @author Jarren Ralf
 */
function scanConversions(range, sheet, spreadsheet, col)
{
  spreadsheet.toast('Searching UPC database...', '', 30)
  const item = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString()).find(item => item[0] == range.getValue())

  if (item != null)
  {
    spreadsheet.toast('', 'Item Found in UPC Database...')
    const quantity = range.offset(0, 2).getValue()
    const packagedItems = range.offset(2, 1 - col, sheet.getLastRow() - 2, 9).getValues().filter(sku => sku[col - 1] == item[1].toString().toUpperCase()) // Values of the row(s) the sku belongs to
    
    if (packagedItems != null) // The skus is found on the SKUsToWatch sheet
    {
      if (packagedItems.length > 1) // There is more than one instance of the same SKU found
      {
        // The user must be prompted which package size they are converting inventory from
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert('Which package size are you converting ' + packagedItems[0][col + 1] + ' - '+ packagedItems[0][col] 
          + ' from?\n\n\nSelect Yes for: ' + packagedItems[0][9 - col] + '\n\nSelect No for: ' + packagedItems[1][9 - col], ui.ButtonSet.YES_NO)

        if (response === ui.Button.YES)
        {
          if (isNotBlank(packagedItems[0][4])) // The conversion factor is NOT blank
          {
            if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
            {
              // Prompt the user to input how many packages were created
              const ui = SpreadsheetApp.getUi();
              const response = ui.prompt('How many ' + packagedItems[0][col + 1] + ' -  ' + packagedItems[0][col] + ' have you created?')

              if (response.getSelectedButton() == ui.Button.OK)
              {
                const qty = Number(response.getResponseText()); // The user inputted quantity

                if (isNotBlank(qty) || !isNaN(qty)) // Valid number
                {
                  const inventorySheet = spreadsheet.getSheetByName('DataImport');
                  const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
                  const sku1 = col - 1, sku2 = 7 - col; // Select the appropriate column numbers
                  const numPackages = (sku2 !== 6) ? qty*packagedItems[0][4] : qty/packagedItems[0][4]; // The appropriate conversion value
                  var convertedItems = []

                  for (var j = 0; j < inventory.length; j++)
                  {
                    if (packagedItems[0][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                    {
                      convertedItems.push([packagedItems[0][sku1].toString().toUpperCase(), inventory[j][0] + qty])
                      break;
                    }
                  }

                  for (var j = 0; j < inventory.length; j++)
                  {
                    if (packagedItems[0][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                    {
                      convertedItems.push([packagedItems[0][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                      break;
                    }
                  }

                  const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
                  const row = UoMConversionPage.getLastRow() + 1;
                  const numRows = convertedItems.length;
                  UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
                  range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
                  spreadsheet.toast('added to Richmond UoM Conversion sheet',  + qty + ' units of - ' + packagedItems[0][col + 1] + ' - ' + packagedItems[0][col] + ' package conversions', 120)
                }
                else
                {
                  ui.alert('Please scan item again and enter a valid quantity.')
                  range.setValue('Scan Here').activate()
                }
              }
              else
                range.setValue('Scan Here').activate()
            }
            else
            {
              const inventorySheet = spreadsheet.getSheetByName('DataImport');
              const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
              const sku1 = col - 1, sku2 = 7 - col;
              const numPackages = (sku2 !== 6) ? quantity*packagedItems[0][4] : quantity/packagedItems[0][4]; // The appropriate conversion value
              var convertedItems = [];

              for (var j = 0; j < inventory.length; j++)
              {
                if (packagedItems[0][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  convertedItems.push([packagedItems[0][sku1].toString().toUpperCase(), inventory[j][0] + quantity])
                  break;
                }
              }

              for (var j = 0; j < inventory.length; j++)
              {
                if (packagedItems[0][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  convertedItems.push([packagedItems[0][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                  break;
                }
              }

              const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
              const row = UoMConversionPage.getLastRow() + 1;
              const numRows = convertedItems.length;
              UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
              range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
              spreadsheet.toast('added to Richmond UoM Conversion sheet',  + quantity + ' units of - ' + packagedItems[0][col + 1] + ' - ' + packagedItems[0][col] + ' package conversions', 120)
            }
          }
          else // The conversion factor is blank
          {
            range.setValue('Scan Here').activate()
            spreadsheet.toast('', 'The SKU for the packaging you are converting from is missing.')
          }
        }
        else if (response === ui.Button.NO)
        {
          if (isNotBlank(packagedItems[1][4])) // The conversion factor is NOT blank
          {
            if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
            {
              // Prompt the user to input how many packages were created
              const ui = SpreadsheetApp.getUi();
              const response = ui.prompt('How many ' + packagedItems[1][col + 1] + ' -  ' + packagedItems[1][col] + ' have you created?')

              if (response.getSelectedButton() == ui.Button.OK)
              {
                const qty = Number(response.getResponseText()); // The user inputted quantity

                if (isNotBlank(qty) || !isNaN(qty)) // Valid number
                {
                  const inventorySheet = spreadsheet.getSheetByName('DataImport');
                  const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
                  const sku1 = col - 1, sku2 = 7 - col;
                  const numPackages = (sku2 !== 6) ? qty*packagedItems[1][4] : qty/packagedItems[1][4];
                  var convertedItems = []

                  for (var j = 0; j < inventory.length; j++)
                  {
                    if (packagedItems[1][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                    {
                      convertedItems.push([packagedItems[1][sku1].toString().toUpperCase(), inventory[j][0] + qty])
                      break;
                    }
                  }

                  for (var j = 0; j < inventory.length; j++)
                  {
                    if (packagedItems[1][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                    {
                      convertedItems.push([packagedItems[1][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                      break;
                    }
                  }

                  const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
                  const row = UoMConversionPage.getLastRow() + 1;
                  const numRows = convertedItems.length;
                  UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
                  range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
                  spreadsheet.toast('added to Richmond UoM Conversion sheet',  + qty + ' units of - ' + packagedItems[1][col + 1] + ' - ' + packagedItems[1][col] + ' package conversions', 120)
                }
                else
                {
                  ui.alert('Please scan item again and enter a valid quantity.')
                  range.setValue('Scan Here').activate()
                }
              }
              else
                range.setValue('Scan Here').activate()
            }
            else
            {
              const inventorySheet = spreadsheet.getSheetByName('DataImport');
              const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
              const sku1 = col - 1, sku2 = 7 - col;
              const numPackages = (sku2 !== 6) ? quantity*packagedItems[1][4] : quantity/packagedItems[1][4];
              var convertedItems = [];

              for (var j = 0; j < inventory.length; j++)
              {
                if (packagedItems[1][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  convertedItems.push([packagedItems[1][sku1].toString().toUpperCase(), inventory[j][0] + quantity])
                  break;
                }
              }

              for (var j = 0; j < inventory.length; j++)
              {
                if (packagedItems[1][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                {
                  convertedItems.push([packagedItems[1][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                  break;
                }
              }

              const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
              const row = UoMConversionPage.getLastRow() + 1;
              const numRows = convertedItems.length;
              UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
              range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
              spreadsheet.toast('added to Richmond UoM Conversion sheet',  + quantity + ' units of - ' + packagedItems[1][col + 1] + ' - ' + packagedItems[1][col] + ' package conversions', 120)
            }
          }
          else // The conversion factor is blank
          {
            range.setValue('Scan Here').activate()
            spreadsheet.toast('', 'The SKU for the packaging you are converting from is missing.')
          }
        }
        else // The user has pressed closed on the dialogue box
        {
          range.setValue('Scan Here').activate()
          spreadsheet.toast('Please Scan Again', 'Conversion Cancelled.')
        }
      }
      else // There is only 1 option for conversion
      {
        if (isNotBlank(packagedItems[0][4])) // The conversion factor is NOT blank
        {
          if (quantity == 0) // The user didn't not enter the quantity in the cell at the top of the page
          {
            // Prompt the user to input how many packages were created
            const ui = SpreadsheetApp.getUi();
            const response = ui.prompt('How many ' + packagedItems[0][col + 1] + ' -  ' + packagedItems[0][col] + ' have you created?')

            if (response.getSelectedButton() == ui.Button.OK)
            {
              const qty = Number(response.getResponseText()); // The user inputted quantity

              if (isNotBlank(qty) || !isNaN(qty)) // Valid number
              {
                const inventorySheet = spreadsheet.getSheetByName('DataImport');
                const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
                const sku1 = col - 1, sku2 = 7 - col;
                const numPackages = (sku2 !== 6) ? qty*packagedItems[0][4] : qty/packagedItems[0][4]; // The appropriate conversion value
                var convertedItems = []

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][sku1].toString().toUpperCase(), inventory[j][0] + qty])
                    break;
                  }
                }

                for (var j = 0; j < inventory.length; j++)
                {
                  if (packagedItems[0][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
                  {
                    convertedItems.push([packagedItems[0][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                    break;
                  }
                }

                const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
                const row = UoMConversionPage.getLastRow() + 1;
                const numRows = convertedItems.length;
                UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
                range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
                spreadsheet.toast('added to Richmond UoM Conversion sheet',  + qty + ' units of - ' + packagedItems[0][col + 1] + ' - ' + packagedItems[0][col] + ' package conversions', 120)
              }
              else
              {
                ui.alert('Please scan item again and enter a valid quantity.')
                range.setValue('Scan Here').activate()
              }
            }
            else
              range.setValue('Scan Here').activate()
          }
          else
          {
            const inventorySheet = spreadsheet.getSheetByName('DataImport');
            const inventory = inventorySheet.getSheetValues(2, 3, inventorySheet.getLastRow() - 1, 5);
            const sku1 = col - 1, sku2 = 7 - col;
            const numPackages = (sku2 !== 6) ? quantity*packagedItems[0][4] : quantity/packagedItems[0][4]; // The appropriate conversion value
            var convertedItems = [];

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][sku1].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][sku1].toString().toUpperCase(), inventory[j][0] + quantity])
                break;
              }
            }

            for (var j = 0; j < inventory.length; j++)
            {
              if (packagedItems[0][sku2].toString().toUpperCase() == inventory[j][4].toString().toUpperCase())
              {
                convertedItems.push([packagedItems[0][sku2].toString().toUpperCase(), inventory[j][0] - numPackages])
                break;
              }
            }

            const UoMConversionPage = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit').getSheetByName('UoM Conversion')
            const row = UoMConversionPage.getLastRow() + 1;
            const numRows = convertedItems.length;
            UoMConversionPage.getRange(row, 1, numRows, 2).setNumberFormat('@').setValues(convertedItems)
            range.setValue('Scan Here').activate().offset(0, 2).setValue('0')
            spreadsheet.toast('added to Richmond UoM Conversion sheet',  + quantity + ' units of - ' + packagedItems[0][col + 1] + ' - ' + packagedItems[0][col] + ' package conversions', 120)
          }
        }
        else // The conversion factor is blank
        {
          range.setValue('Scan Here').activate()
          spreadsheet.toast('', 'The SKU for the packaging you are converting from is missing.')
        }
      }
    }
    else
    {
      range.setValue('Item Missing from SKUsToWatch').activate()
      spreadsheet.toast('', 'Item Missing from SKUsToWatch')
    }
  }
  else
  {
    range.setValue('Barcode Not Found').activate()
    spreadsheet.toast('', 'Barcode Not Found')
  }
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the SearchData page for the items in question.
 * It also highlights the items that are already on the shipped page and already on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param   {Number}         row     : The row number that was just edited
 * @param   {Number}         col     : The column number that was just edited
 * @author Jarren Ralf 
 */
function searchV2(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(7, 1);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(9, 1, sheet.getMaxRows() - 8, 7); // The entire range of the Item Search page
  const checkBoxes = sheet.getSheetValues(1, 3, 6, 3);
  const output = [];
  const searchesOrNot = sheet.getRange(1, 2, 6).clearFormat()                                       // Clear the formatting of the range of the search box
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

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      if (checkBoxes[0][1] == true) // Compute the sheet search
      {
        const dataSheets = [];

        if (checkBoxes[3][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'))
        if (checkBoxes[4][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'))
        if (checkBoxes[5][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)'))

        const data = dataSheets.map(sheet => sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn()));
        const numSearches = searches.length; // The number of searches
        var numSearchWords, numRich = 0, numParks = 0, numRupt = 0;

        for (var loc = 0; loc < data.length; loc++) // Loop through the locations
        {
          for (var i = 3; i < data[loc].length; i++) // Loop through all of the data for the given location
          {
            for (var c = 0; c < data[loc][0].length; c++) // Loop through all of the columns for the given location data
            {
              if (data[loc][2][c] === 'Description') // Only check the description columns for the given searches and search words
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[loc][i][c].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item of the jth search
                      {
                        if (data[loc][0][0] === 'PNT Richmond Transfer Spreadsheet (Location: 100)')
                        {
                          if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                        }
                        else
                        {
                          if (data[loc][1][c] === 'Order')
                          {
                            if (checkBoxes[0][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 3], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 3], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Shipped')
                          {
                            if (checkBoxes[1][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 2], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 2], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Received')
                          {
                            if (checkBoxes[2][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'ItemsToRichmond')
                          {
                            if (checkBoxes[5][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                        }
                        
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
              else continue;
            }
          }
        }
      }
      else
      {
        const inventorySheet = spreadsheet.getSheetByName('DataImport');
        const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 7);
        const numSearches = searches.length; // The number searches
        var numSearchWords;

        if (searches[0][0].toLowerCase() === 'trites')
        {
          if (numSearches === 1 && searches[0].length == 1)
            output.push(...data.filter(item => item[5] > 0))
          else
          {
            const tritesData = data.filter(item => item[5] > 0);

            for (var i = 0; i < tritesData.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (searches[j][k] === 'trites')
                    continue;

                  if (tritesData[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      output.push(tritesData[i]);
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
        else if (searches[0][0].toLowerCase() === 'all' && searches[0][1].toLowerCase() === 'hoochies')
        {
          const hoochieData = data.filter(item => item[6].toString().substring(0, 2) == '16');
          const hoochiePrefixes = ['16060005', '16010005', '16050005', '16020000', '16020010', '16060065', '16060010', '16070000', '16075300', '16070975',
                                   '16030000', '16060175', '16200030', '16200000', '16200025', '16200065', '16200021', '16200022', '16200061'];
          const numTypesOfHoochies = hoochiePrefixes.length;
          var hoochies = new Array(numTypesOfHoochies).fill('').map(() => []);

          for (var j = 0; j < numTypesOfHoochies; j++) // Loop through the number of searches
          {
            for (var i = 0; i < hoochieData.length; i++) // Loop through all of the descriptions from the search data
              if (hoochieData[i][6].toString().substring(0, 8) === hoochiePrefixes[j] && !hoochieData[i][1].toString().toLowerCase().includes('rig')) // Does the i-th sku contain begin with the j-th hoochie prefix 
                hoochies[j].push(hoochieData[i]); // The description also does not contain the word "rig"

            hoochies[j] = sortHoochies(hoochies[j], 1, hoochiePrefixes[j])
          }

          output.push(...[].concat(...hoochies));
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
                if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                {
                  if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                  {
                    output.push(data[i]);
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
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      if (checkBoxes[0][1] == true) // Compute the sheet search
      {
        const dataSheets = [];

        if (checkBoxes[3][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'))
        if (checkBoxes[4][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'))
        if (checkBoxes[5][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)'))

        const data = dataSheets.map(sheet => sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn()));
        const numSearches = searches.length; // The number of searches
        var numSearchWords, numRich = 0, numParks = 0, numRupt = 0;

        for (var loc = 0; loc < data.length; loc++) // Loop through the locations
        {
          for (var i = 3; i < data[loc].length; i++) // Loop through all of the data for the given location
          {
            for (var c = 0; c < data[loc][0].length; c++) // Loop through all of the columns for the given location data
            {
              if (data[loc][2][c] === 'Description') // Only check the description columns for the given searches and search words
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[loc][i][c].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item of the jth search
                      {
                        if (data[loc][0][0] === 'PNT Richmond Transfer Spreadsheet (Location: 100)')
                        {
                          if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                        }
                        else
                        {
                          if (data[loc][1][c] === 'Order')
                          {
                            if (checkBoxes[0][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 3], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 3], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Shipped')
                          {
                            if (checkBoxes[1][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 2], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 2], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Received')
                          {
                            if (checkBoxes[2][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'ItemsToRichmond')
                          {
                            if (checkBoxes[5][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                        }
                        
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
              else continue;
            }
          }
        }
      }
      else
      {
        const inventorySheet = spreadsheet.getSheetByName('DataImport');
        const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 7);
        const numSearches = searches.length; // The number searches
        var numSearchWords;

        if (searches[0][0].toLowerCase() === 'trites')
        {
          if (numSearches === 1 && searches[0].length == 1)
            output.push(...data.filter(item => item[5] > 0))
          else
          {
            const tritesData = data.filter(item => item[5] > 0);

            for (var i = 0; i < tritesData.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (searches[j][k] === 'trites')
                    continue;

                  if (tritesData[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      for (var l = 0; l < dontIncludeTheseWords.length; l++)
                      {
                        if (!tritesData[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                        {
                          if (l === dontIncludeTheseWords.length - 1)
                          {
                            output.push(tritesData[i]);
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
        else
        {
          for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
          {
            loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
            {
              numSearchWords = searches[j].length - 1;

              for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
              {
                if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                {
                  if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                  {
                    for (var l = 0; l < dontIncludeTheseWords.length; l++)
                    {
                      if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                      {
                        if (l === dontIncludeTheseWords.length - 1)
                        {
                          output.push(data[i]);
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
    }

    const numItems = output.length;

    if (numItems === 0) // No items were found
    {
      sheet.getRange('B1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent(); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);

      (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([[numRich + ' on richmond'],[numParks + ' on parksville'],[numRupt + ' on rupert']]) : 
                                  sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
    }
    else
    {
      var colours = (checkBoxes[0][1] == true) ?
                      [ [...Array(numRich)].map(e => new Array(7).fill('#d9ead3')), 
                      [...Array(numParks)].map(e => new Array(7).fill('#c9daf8')), 
                      [...Array(numRupt)].map(e => new Array(7).fill('#f4cccc'))] : 
                      [ [...Array(numItems)].map(e => new Array(7).fill('white'))];
      var backgroundColours = [].concat.apply([], colours); // Concatenate all of the item values as a 2-D array
      sheet.getRange('B9').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(9, 1, numItems, 7).setBackgrounds(backgroundColours).setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");

      (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([[numRich + ' on richmond'],[numParks + ' on parksville'],[numRupt + ' on rupert']]) : 
                                  sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
    }

    spreadsheet.toast('Searching Complete.')
  }
  else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
  {
    itemSearchFullRange.setBackground('white').setValue('');
    searchResultsDisplayRange.setValue('');
    (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([['0 on richmond'],['0 on parksville'],['0 on rupert']]) : 
                                 sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);

    (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([['0 on richmond'],['0 on parksville'],['0 on rupert']]) : 
                                 sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the SearchData page for the items in question.
 * It also highlights the items that are already on the shipped page and already on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param   {Number}         row     : The row number that was just edited
 * @param   {Number}         col     : The column number that was just edited
 * @author Jarren Ralf 
 */
function searchV2V2(e, spreadsheet, sheet, row, col)
{
  const startTime = new Date().getTime();
  const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(7, 1);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(9, 1, sheet.getMaxRows() - 8, 6); // The entire range of the Item Search page
  const checkBoxes = sheet.getSheetValues(1, 3, 6, 3);
  const output = [];
  const searchesOrNot = sheet.getRange(1, 2, 6).clearFormat()                                               // Clear the formatting of the range of the search box
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

    if (checkBoxes[0][1] == true) // Compute the sheet search
    {
      if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
      {
        const dataSheets = [];

        if (checkBoxes[3][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'))
        if (checkBoxes[4][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'))
        if (checkBoxes[5][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)'))

        const data = dataSheets.map(sheet => sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn()));
        const numSearches = searches.length; // The number of searches
        var numSearchWords, numRich = 0, numParks = 0, numRupt = 0;

        for (var loc = 0; loc < data.length; loc++) // Loop through the locations
        {
          for (var i = 3; i < data[loc].length; i++) // Loop through all of the data for the given location
          {
            for (var c = 0; c < data[loc][0].length; c++) // Loop through all of the columns for the given location data
            {
              if (data[loc][2][c] === 'Description') // Only check the description columns for the given searches and search words
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[loc][i][c].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item of the jth search
                      {
                        if (data[loc][0][0] === 'PNT Richmond Transfer Spreadsheet (Location: 100)')
                        {
                          if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                        }
                        else
                        {
                          if (data[loc][1][c] === 'Order')
                          {
                            if (checkBoxes[0][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 3], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 3], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Shipped')
                          {
                            if (checkBoxes[1][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 2], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 2], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Received')
                          {
                            if (checkBoxes[2][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'ItemsToRichmond')
                          {
                            if (checkBoxes[5][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                        }
                        
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
              else continue;
            }
          }
        }
      }
      else // The word 'not' was found in the search string
      {
        var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
        const dataSheets = [];

        if (checkBoxes[3][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Richmond Data (Loc: 100)'))
        if (checkBoxes[4][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Parksville Data (Loc: 200)'))
        if (checkBoxes[5][0] == true)
          dataSheets.push(spreadsheet.getSheetByName('Imported Rupert Data (Loc: 300)'))

        const data = dataSheets.map(sheet => sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn()));
        const numSearches = searches.length; // The number of searches
        var numSearchWords, numRich = 0, numParks = 0, numRupt = 0;

        for (var loc = 0; loc < data.length; loc++) // Loop through the locations
        {
          for (var i = 3; i < data[loc].length; i++) // Loop through all of the data for the given location
          {
            for (var c = 0; c < data[loc][0].length; c++) // Loop through all of the columns for the given location data
            {
              if (data[loc][2][c] === 'Description') // Only check the description columns for the given searches and search words
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[loc][i][c].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item of the jth search
                      {
                        if (data[loc][0][0] === 'PNT Richmond Transfer Spreadsheet (Location: 100)')
                        {
                          if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], data[loc][i][c + 1], '', '', '', data[loc][1][c]]);
                              numRich++;
                            }
                          }
                        }
                        else
                        {
                          if (data[loc][1][c] === 'Order')
                          {
                            if (checkBoxes[0][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 3], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 3], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Shipped')
                          {
                            if (checkBoxes[1][2] == true)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 2], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 2], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Received')
                          {
                            if (checkBoxes[2][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'InfoCounts')
                          {
                            if (checkBoxes[3][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'Manual Counts')
                          {
                            if (checkBoxes[4][2] == true && isNotBlank(data[loc][i][c + 1]))
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                          else if (data[loc][1][c] === 'ItemsToRichmond')
                          {
                            if (checkBoxes[5][2] == true && data[loc][i][c + 3] == false)
                            {
                              if (data[loc][0][0] === 'PNT Parksville Transfer Spreadsheet (Location: 200)')
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', data[loc][i][c + 1], '', '', data[loc][1][c]]);
                                numParks++;
                              }
                              else
                              {
                                output.push([data[loc][i][c].split(' - ')[data[loc][i][c].split(' - ').length - 2], data[loc][i][c], '', '', data[loc][i][c + 1], '', data[loc][1][c]]);
                                numRupt++;
                              }
                            }
                          }
                        }
                        
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
              else continue;
            }
          }
        }
      }
    }
    else
    {
      if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
      {
        const inventorySheet = spreadsheet.getSheetByName('DataImport');
        const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 7);
        const numSearches = searches.length; // The number searches
        var numSearchWords;

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  output.push(data[i]);
                  break loop;
                }
              }
              else
                break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
            }
          }
        }
      }
      else
      {
        const inventorySheet = spreadsheet.getSheetByName('DataImport');
        const data = inventorySheet.getSheetValues(2, 1, inventorySheet.getLastRow() - 1, 7);
        const numSearches = searches.length; // The number searches
        var numSearchWords;

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
          {
            numSearchWords = searches[j].length - 1;

            for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
            {
              if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
              {
                if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                {
                  for (var l = 0; l < dontIncludeTheseWords.length; l++)
                  {
                    if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                    {
                      if (l === dontIncludeTheseWords.length - 1)
                      {
                        output.push(data[i]);
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

    const numItems = output.length;

    if (numItems === 0) // No items were found
    {
      sheet.getRange('B1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent(); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);

      (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([[numRich + ' on richmond'],[numParks + ' on parksville'],[numRupt + ' on rupert']]) : 
                                  sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
    }
    else
    {
      var colours = (checkBoxes[0][1] == true) ?
                      [ [...Array(numRich)].map(e => new Array(7).fill('#d9ead3')), 
                      [...Array(numParks)].map(e => new Array(7).fill('#c9daf8')), 
                      [...Array(numRupt)].map(e => new Array(7).fill('#f4cccc'))] : 
                      [ [...Array(numItems)].map(e => new Array(7).fill('white'))];
      var backgroundColours = [].concat.apply([], colours); // Concatenate all of the item values as a 2-D array
      sheet.getRange('B9').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      Logger.log(output)
      sheet.getRange(9, 1, numItems, 7).setBackgrounds(backgroundColours).setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");

      (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([[numRich + ' on richmond'],[numParks + ' on parksville'],[numRupt + ' on rupert']]) : 
                                  sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
    }

    spreadsheet.toast('Searching Complete.')
  }
  else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
  {
    itemSearchFullRange.setBackground('white').setValue('');
    searchResultsDisplayRange.setValue('');
    (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([['0 on richmond'],['0 on parksville'],['0 on rupert']]) : 
                                 sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);

    (checkBoxes[0][1] == true) ? sheet.getRange(4, 1, 3).setValues([['0 on richmond'],['0 on parksville'],['0 on rupert']]) : 
                                 sheet.getRange(4, 1, 3).setValues([[''],[''],['']])
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function sends an email to Adrian and Jarren to give a heads up that a function in apps script has failed to run.
 * 
 * @param {String} error : The property of the error object that displays the functions and linenumbers that the error occurs at.
 * @author Jarren Ralf
 */
function sendErrorEmail(error)
{
  if (MailApp.getRemainingDailyQuota() > 3) // Don't try and send an email if the daily quota of emails has been sent
  {
    var today = new Date()
    var formattedError = '<p>' + error.replaceAll(' at ', '<br /> &emsp;&emsp;&emsp;') + '</p>';
    var templateHtml = HtmlService.createTemplateFromFile('FunctionFailedToRun');
    templateHtml.dateAndTime = today.toLocaleTimeString() + ' on ' + today.toDateString();
    templateHtml.scriptURL   = "https://script.google.com/home/projects/178jXC1SLz1GQpIOiNLgRAzE4j4A-F1jt4OatEQ3BLLwaO3nH4rZrRDRm/edit";
    var emailBody = templateHtml.evaluate().append(formattedError).getContent();
    
    MailApp.sendEmail({      to: 'lb_blitz_allstar@hotmail.com',
                        subject: 'Adrian\'s Adagio Update Sheet Script Failure', 
                       htmlBody: emailBody
    });
  }
  else
    Logger.log('No email sent because it appears that the daily quota of emails has been met!')
}

/**
 * This function is run when their are values pasted in either of the description columns on the SKUsToWatch or SKUsToWatch_ASSEMBLY. 
 * It removes the SKU from the google description and moves it to the left cell, as well as the UoM and moves it to the right cell.
 * 
 * @param {Range} range : The active range
 */
function stripDescriptions(range)
{
  var splitDescription, sku, uom;

  const descriptions = range.getValues().map(descrip => {
    splitDescription = descrip[0].split(' - ')
    if (splitDescription.length > 2)
    {
      sku = splitDescription.pop()
      uom = splitDescription.pop()
      splitDescription.pop()
      splitDescription.pop()

      return [sku, splitDescription.join(' - '), uom]
    }
    else
      return descrip
  })

  const numCols = descriptions[0].length

  if (numCols === 3)
    range.offset(0, -1, descriptions.length, numCols).setNumberFormat('@').setHorizontalAlignment('center').setFontColor('black').setFontFamily('Arial').setFontSize(10)
      .setFontWeight('normal').setBackground('white').setVerticalAlignment('middle').setBorder(false, null, false, null, false, false).setValues(descriptions)
}

/**
 * This function looks at the dates on the dashboard that represent when the UPC database was last updated on each transfer sheet and changes the font colour
 * to red if it hasn't been updated in a week. Also, it replaces the timestamps with the word "BUTTON".
 * 
 * @author Jarren Ralf
 */
function updateDashboard()
{
  try
  {
    const today = new Date();
    const spreadsheet = SpreadsheetApp.getActive();
    const sheets = spreadsheet.getSheets();
    const adagioSheet = sheets.shift();
    const conversionSheet = sheets.shift();
    const ONE_WEEK  = new Date(today.getFullYear(), today.getMonth(), today.getDate() -  7);
    const range = adagioSheet.getRange(11, 18, 3);
    const fontColours = range.getValues().map(date => [(new Date(date[0].split(' on ')[1]) <= ONE_WEEK) ? 'red' : 'black']);
    range.setFontColors(fontColours);
    adagioSheet.getRangeList(['E4:E10', 'J4:J11', 'O4:O11', 'E15', 'J15', 'O15']).getRanges().map(range => range.setValue("BUTTON"));
    adagioSheet.hideRows(15, 5);
    adagioSheet.getRange('O2').setValue('Extend Dashboard')

    conversionSheet.getRange(2, 7).uncheck() // Uncheck the checkbox on the ConvertedExport page

    for (var j = 0; j < sheets.length; j++)
    {
      if (sheets[j].getSheetName() == 'Imported Richmond Data (Loc: 100)' || sheets[j].getSheetName() == 'Imported Parksville Data (Loc: 200)' || 
          sheets[j].getSheetName() == 'Imported Rupert Data (Loc: 300)'   || sheets[j].getSheetName() == 'DataImport')
        sheets[j].hideSheet();
    }
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * This function updates the search data with the date which particular items were last counted.
 * 
 * @author Jarren Ralf
 */
function updateSearchData_withToast()
{
  const today = new Date();
  const spreadsheet = ss;
  const adagioSpreadsheet = SpreadsheetApp.getActive();

  adagioSpreadsheet.toast('Accessing count log and sorting by date...', '', -1)

  const countLog = spreadsheet.getSheetByName("Count Log");
  const numOldCounts = countLog.getLastRow() - 1;
  const countLogRange = countLog.getRange(2, 1, numOldCounts, 4);
  const countLogData = countLogRange.getValues().sort(sortByCountedDate);

  adagioSpreadsheet.toast('Sorting complete. Removing duplicates in count log...', '', -1)

  const mostRecentCounts = uniqByKeepLast(countLogData, sku => sku[0]); // Remove duplicates
  const numNewCounts = mostRecentCounts.length;
  
  adagioSpreadsheet.toast('Duplicates removed. Clearing count log...', '', -1)

  countLogRange.clearContent();

  adagioSpreadsheet.toast('Count log cleared. Setting values to count log...', '', -1)

  countLog.getRange(2, 1, numNewCounts, 4).setValues(mostRecentCounts);

  adagioSpreadsheet.toast('Count log update complete. Retrieving search data and updating Counted On date...', '', -1)

  const searchDataRng = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName("INVENTORY").getRange('B7:C') : spreadsheet.getSheetByName("SearchData").getRange('B1:C');
  const searchData = searchDataRng.getValues().map(skuAndCountDate => {
    skuAndCountDate[1] = mostRecentCounts.find(sku => sku[0] == skuAndCountDate[0].split(' - ').pop().toString())?.[3]

    return skuAndCountDate;
  });

  const numItems = searchData.length;
  const numberFormats = [...Array(numItems)].map(e => ['@', 'dd MMM yyyy']);
  searchData[0][1] = "Last Counted On";
  numberFormats[0][1] = '@';

  adagioSpreadsheet.toast('Search data has been updated with Counted On dates. Writing dates on the search data sheet...', '', -1)
  
  searchDataRng.setNumberFormats(numberFormats).setValues(searchData);
  spreadsheet.getSheetByName("INVENTORY").getRange(6, 1).setValue('The Recent Counts were last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString());

  if (numOldCounts > numNewCounts)
    countLog.deleteRows(numNewCounts + 2, numOldCounts - numNewCounts); // Delete the blank rows

  adagioSpreadsheet.toast('Updating search data complete.', 'COMPLETE', 60)
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////// RICHMOND //////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * @author Jarren Ralf
 */
function richmond_applyFullSpreadsheetFormatting(range)
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    applyFullSpreadsheetFormatting();
    timeStamp(spreadsheet, 15, 5, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_clearInventory()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    clearInventory();
    timeStamp(spreadsheet, 4, 5, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_getCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    getCounts();
    timeStamp(spreadsheet, 6, 5, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_clearManualCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    clearManualCounts();
    timeStamp(spreadsheet, 8, 5, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_updateUPC_Database_ButtonClicked()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    updateUPC_Database(true);
    SpreadsheetApp.getActiveSheet().getRange(11, 18).setFontColor('black');
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_updateRecentlyCreatedItems(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    updateRecentlyCreatedItems();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    range.uncheck();
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_updateUPC_Database()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    updateUPC_Database();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_countLog(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    countLog();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_logCountsOnWorkdays()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    logCountsOnWorkdays();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function richmond_updateSearchData(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    updateSearchData_withToast();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page
*
* @author Jarren Ralf
*/
function richmond_manualCounts()
{
  const NUM_COLS = 2;

  try
  {
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    var manualCountsSheet = ss.getSheetByName("Manual Counts");
    var lastRow = manualCountsSheet.getLastRow();
    var startRow = (lastRow < 3) ? 4 : lastRow + 1;

    copySelectedValuesV2(manualCountsSheet, startRow, NUM_COLS);
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);    
  }
}

/**
* This function generate the a list of suggested items that should be pick from Trites and brought back to Moncton street
*
* @author Jarren Ralf
*/
function richmond_generateSuggestedInflowPick()
{
  try
  {
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk/edit#gid=592450561');
    generateSuggestedInflowPick()
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);    
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////// PARKSVILLE /////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * @author Jarren Ralf
 */
function parksville_applyFullSpreadsheetFormatting(range)
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    applyFullSpreadsheetFormatting();
    timeStamp(spreadsheet, 15, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_completeReceived()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    completeReceived();
    timeStamp(spreadsheet, 6, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_completeToRichmond()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    completeToRichmond();
    timeStamp(spreadsheet, 7, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_print_X_Order()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    print_X_Order();
    timeStamp(spreadsheet, 8, 10, adagioSheet, "dd MMM HH:mm"),
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_print_X_Shipped()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    print_X_Shipped();
    timeStamp(spreadsheet, 9, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_clearInventory()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    clearInventory();
    timeStamp(spreadsheet, 4, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_getCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    getCounts();
    timeStamp(spreadsheet, 10, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_clearManualCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    clearManualCounts();
    timeStamp(spreadsheet, 11, 10, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_updateUPC_Database_ButtonClicked()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    updateUPC_Database(true);
    SpreadsheetApp.getActiveSheet().getRange(12, 18).setFontColor('black');
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_updateRecentlyCreatedItems(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    updateRecentlyCreatedItems();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_updateUPC_Database()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    updateUPC_Database();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_countLog(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    countLog();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_logCountsOnWorkdays()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    logCountsOnWorkdays();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function parksville_updateSearchData(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    updateSearchData_withToast();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page
*
* @author Jarren Ralf
*/
function parksville_manualCounts()
{
  const NUM_COLS = 3;

  try
  {
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049');
    var manualCountsSheet = ss.getSheetByName("Manual Counts");
    var lastRow = manualCountsSheet.getLastRow();
    var startRow = (lastRow < 3) ? 4 : lastRow + 1;

    copySelectedValuesV2(manualCountsSheet, startRow, NUM_COLS);
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);   
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////// PRINCE RUPERT ////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * @author Jarren Ralf
 */
function rupert_applyFullSpreadsheetFormatting(range)
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    applyFullSpreadsheetFormatting();
    timeStamp(spreadsheet, 15, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_completeReceived()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    completeReceived();
    timeStamp(spreadsheet, 6, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_completeToRichmond()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    completeToRichmond();
    timeStamp(spreadsheet, 7, 15, adagioSheet, "dd MMM HH:mm")  
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_print_X_Order()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    print_X_Order();
    timeStamp(spreadsheet, 8, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_print_X_Shipped()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    print_X_Shipped();
    timeStamp(spreadsheet, 9, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_clearInventory()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    clearInventory();
    timeStamp(spreadsheet, 4, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_getCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    getCounts();
    timeStamp(spreadsheet, 10, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_clearManualCounts()
{
  var startTime = new Date().getTime();

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const adagioSheet = spreadsheet.getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    clearManualCounts();
    timeStamp(spreadsheet, 11, 15, adagioSheet, "dd MMM HH:mm")
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_updateUPC_Database_ButtonClicked()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    updateUPC_Database(true);
    SpreadsheetApp.getActiveSheet().getRange(13, 18).setFontColor('black');
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_updateRecentlyCreatedItems()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    updateRecentlyCreatedItems();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_updateUPC_Database()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    updateUPC_Database();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_countLog(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    countLog();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_logCountsOnWorkdays()
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    logCountsOnWorkdays();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);
  }
}

/**
 * @author Jarren Ralf
 */
function rupert_updateSearchData(range)
{
  var startTime = new Date().getTime();

  try
  {
    const adagioSheet = SpreadsheetApp.getActive().getSheetByName('Adagio Transfer Sheet');
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    updateSearchData_withToast();
    setElapsedTime(startTime, adagioSheet); // To check the ellapsed times
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    if (arguments.length !== 0) range.uncheck();
    throw new Error(error);
  }
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page
*
* @author Jarren Ralf
*/
function rupert_manualCounts()
{
  const NUM_COLS = 4;

  try
  {
    ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit#gid=407280159');
    var manualCountsSheet = ss.getSheetByName("Manual Counts");
    var lastRow = manualCountsSheet.getLastRow();
    var startRow = (lastRow < 3) ? 4 : lastRow + 1;

    copySelectedValuesV2(manualCountsSheet, startRow, NUM_COLS);
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error)
    Logger.log(error)
    throw new Error(error);   
  }
}