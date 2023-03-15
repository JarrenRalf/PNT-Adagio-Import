/**
* This function clears the scan log on the Richmond Tablet 1 spreadsheet.
*
* @author Jarren Ralf
*/
function clearTablet1()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet();
  clearScanLog(9, 5, sheet)
  displayRuntime(startTime, sheet)
}

/**
* This function clears the scan log on the Richmond Tablet 2 spreadsheet.
*
* @author Jarren Ralf
*/
function clearTablet2()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet();
  clearScanLog(9, 8, sheet)
  displayRuntime(startTime, sheet)
}

/**
* This function clears the scan log on the Richmond Tablet 3 spreadsheet.
*
* @author Jarren Ralf
*/
function clearTablet3()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet();
  clearScanLog(9, 11, sheet)
  displayRuntime(startTime, sheet)
}

/**
* This function clears the scan log on the Richmond Tablet 4 spreadsheet.
*
* @author Jarren Ralf
*/
function clearTablet4()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet();
  clearScanLog(9, 14, sheet)
  displayRuntime(startTime, sheet)
}

/**
*  This function clears the scan log on the Richmond Tablet 5 spreadsheet.
*
* @author Jarren Ralf
*/
function clearTablet5()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet();
  clearScanLog(9, 17, sheet)
  displayRuntime(startTime, sheet)
}

/**
* This function generically clears the scan log of the Richmond Tablet spreadsheets based on the row and column which contain the url link on the Count Sheets Tally page.
*
* @param {Number}  row : The
* @param {Number}  col : The
* @param {Sheet} sheet : The
* @author Jarren Ralf
*/
function clearScanLog(row, col, sheet)
{
  SpreadsheetApp.openByUrl(sheet.getRange(row, col).getRichTextValue().getLinkUrl()).getSheetByName('Scan Log').clearContents()
}

/**
 * This function concatenates the manually added UPCs with the imported list for Richmond Tablet 5.
 * 
 * @author Jarren Ralf
 */
function concatManuallyAddedUPCs()
{
  const spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1pH7jZ1Ecu7mOGHPSqb0E773jmcM_v-yqskekAXnhVwk/edit');
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  var numUpcs = upcDatabaseSheet.getLastRow()
  var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);
  var l, u, m;
  
  if (manAddedUPCsSheet.getLastRow() > 1)
  {
    var manuallyAddedUPCs = manAddedUPCsSheet.getSheetValues(2, 1, manAddedUPCsSheet.getLastRow() - 1, 3).filter(a => isNotBlank(a[0]));
    manAddedUPCsSheet.getRange(2, 1, manuallyAddedUPCs.length, 3).clearContent().setValues(manuallyAddedUPCs);

    manuallyAddedUPCs.map(v => {

      l = 0; // Lower-bound
      u = numUpcs - 1; // Upper-bound
      m = Math.ceil((u + l)/2) // Midpoint

      while (l < m && u > m)
      {
        if (v[1] < upcDatabase[m][0])
          u = m;   
        else if (v[1] > upcDatabase[m][0])
          l = m;
        else
          break;

        m = Math.ceil((u + l)/2) // Midpoint
      }

      if (v[1] < upcDatabase[0][0])
        upcDatabase.splice(0, 0, [v[1], v[2]])
      else if (v[1] < upcDatabase[m][0])
        upcDatabase.splice(m, 0, [v[1], v[2]])
      else if (v[1] > upcDatabase[m][0])
        upcDatabase.splice(m + 1, 0, [v[1], v[2]])

      numUpcs = upcDatabase.length
    })

    upcDatabaseSheet.getRange(1, 1, numUpcs, 2).setValues(upcDatabase)
  }
}

/**
* This function checks if the given string contains only numbers and turns a true boolena if it does, and false otherwise.
*
* @param {String} str : The given string.
* @return {Boolean} Returns true if the given string contains only numbers.
* @author Jarren Ralf
*/
function containsOnlyNumbers(str)
{
  return /^\d+$/.test(str);
}

/**
* This function sets the ellapsed time of a function and prints it on the Count Sheets Tally page.
*
* @param {Number} startTime : The start time that the script began running at represented by a number in milliseconds
* @param {Sheet}    sheet   : The Count Sheets Tally sheet.
* @author Jarren Ralf
*/
function displayRuntime(startTime, sheet)
{
  sheet.getRange(1, 2).setValue((new Date().getTime() - startTime)/1000 + '\nseconds')
}

/**
* This function checks all of the tablet spreadsheets and collects the data on the Scan Log pages, during the process it strips off and 
* returns only the SKU number and quantity. In addition, if multiple counts for one sku number exist, it sums the counts and returns the sku
* and the sum of all counts for that particular item. Lastly, it concatenates the full set of data and returns a master list.
*
* @author Jarren Ralf
*/
function getAllCounts()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet()
  const isTablet5_Included = sheet.getSheetValues(7, 19, 1, 1)[0][0];
  const numCols = (isTablet5_Included) ? 13 : 10;
  const urls = sheet.getRange(9, 5, 1, numCols).getRichTextValues()[0].map(v => v.getLinkUrl()).filter(u => u);
  const allCounts = [];
  var index, index_AllCounts, log, counts = [], uniqueSKUs = [], uniqueSKUs_AllCounts = [];

  if (sheet.getLastRow() > 12)
    sheet.getRange(13, 2, sheet.getLastRow() - 12, 14).clearContent()

  urls.map((url, i) => {
    log = SpreadsheetApp.openByUrl(url).getSheetByName('Scan Log').getDataRange().getValues().map(v => (v[1]) ? [v[0].split(' - ', 1)[0], v[1]] : null);

    if (log[0] !== null)
    {
      counts = log.filter(sku => {
        index = uniqueSKUs.indexOf(sku[0]); // Get the index position of the sku for the i-th item in the unique array
        index_AllCounts = uniqueSKUs_AllCounts.indexOf(sku[0]);

        if (isItemUnique(index))
        {
          uniqueSKUs.push(sku[0]); // Put the unique sku into the unique array
          counts.push(sku);        // Save the entire row of data in the output array
        }
        else if (sku[1] !== 0) // sku already in list
          counts[index][1] += sku[1]; // Add to the previous quantity

        if (isItemUnique(index_AllCounts))
        {
          uniqueSKUs_AllCounts.push(sku[0]); // Put the unique sku into the unique array for all counts
          allCounts.push(sku);               // Save the entire row of data in the output array for all counts
        }
        else if (sku[1] !== 0) // sku already in list
          allCounts[index_AllCounts][1] += sku[1]; // Add to the previous quantity

        return isItemUnique(index)
      })

      sheet.getRange(13, 5 + i*3, counts.length, 2).setNumberFormat('@').setValues(counts)
      counts.length = 0
      uniqueSKUs.length = 0;
    }
  })

  if (allCounts.length !== 0)
    sheet.getRange(13, 2, allCounts.length, 2).setNumberFormat('@').setValues(allCounts)

  displayRuntime(startTime, sheet)
}

/**
* This function...
*
* @author Jarren Ralf
*/
function importInventory()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet()
  const urls = sheet.getRange(9, 5, 1, 13).getRichTextValues()[0].map(v => v.getLinkUrl()).filter(u => u);
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()).map(c => [c[1]]);
  csvData.shift()

  urls.map((url, i) => {
    SpreadsheetApp.openByUrl(url).getSheetByName('Inventory').clearContents().getRange(1, 1, csvData.length).setNumberFormat('@').setValues(csvData);
    timeStamp(2, 5 + i*3, SpreadsheetApp.getActive(), sheet) 
  })

  displayRuntime(startTime, sheet)
}

/**
* This function...
*
* @author Jarren Ralf
*/
function importUpcs()
{
  const startTime = new Date();
  const sheet = SpreadsheetApp.getActiveSheet()
  const urls = sheet.getRange(9, 5, 1, 13).getRichTextValues()[0].map(v => v.getLinkUrl()).filter(u => u);
  const inventory = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());

  // Replace the csvData with the Adagio descriptions and current stock values
  const upcData = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString()).filter(v => {
    return inventory.filter(u => {
      isInAdagioDatabase = ((u[6] == v[1].toString().toUpperCase()) && containsOnlyNumbers(v[0])); // Match the SKU and remove strings
      if (!isInAdagioDatabase) return isInAdagioDatabase; // If the SKU isn't found in the Adagio database, return false
      // Move the Description to column 2 as well as concatenate the current stock
      v[1] = (u[10] !== 'A') ? 'This Item is Not Active in Adagio: ' + u[1] + ' - Current Stock:' + u[2] : u[1] + ' - Current Stock:' + u[2] ; 
      v.splice(2) // Remove the last two columns from the UPC data
      return isInAdagioDatabase;
    }).length != 0; // Keep only the items in the UPC database that have found a matching sku in Adagio
  }).sort((a,b) => a[0] - b[0]) // Sort values numerically

  urls.map((url, i) => {
    SpreadsheetApp.openByUrl(url).getSheetByName('UPC Database').clearContents().getRange(1, 1, upcData.length, 2).setValues(upcData);
    timeStamp(3, 5 + i*3, SpreadsheetApp.getActive(), sheet) 
  })

  ScriptApp.newTrigger('concatManuallyAddedUPCs').timeBased().after(30000).create() // wait 30 seconds before attempting to concatenate manually added UPCs
  displayRuntime(startTime, sheet)
}

/**
* This function checks if a value is contained in an array by checking the index position, and if it equals -1, it's not in the array.
*
* @param  {Number}  index : The index position of a certain value in an array
* @return {Boolean}       : Whether the item is in the array or not 
* @author Jarren Ralf
*/ 
function isItemUnique(index)
{
  return index === -1;
}