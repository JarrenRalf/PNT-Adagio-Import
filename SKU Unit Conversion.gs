/**
* This function checks if the j-th pair of conversion SKUs have been found in the Adagio database or not.
*
* @param  {Object[]} arr : The given array of two SKUs
* @return {Boolean}      : Whether both SKUs have been found in the Adagio database or not
* @author Jarren Ralf
*/
function bothSKUsAreFound(arr)
{
  return arr[0] != null && arr.length === 2;
}

/**
 * This function runs when the compute conversions button is clicked on the ConvertedExport page.
 * 
 * @author Jarren Ralf
 */
function computeConversions()
{
  SpreadsheetApp.getActiveSheet().getRange(2, 7).isChecked() ? computeConversions_Yeti() : computeConversions_Assembly_Then_PackageSize();
}

/**
* This function searches the Adagio database for items that have negative inventory AND are assembled from one or more components. If it is possible to convert inventory
* of certain components into an assembled good in order to reduce the number of occurences of negative inventory, then that computation is done for each location.
*
* @author Jarren Ralf
*/
function computeConversions_Assembly()
{
  const START_TIME = new Date().getTime(); // To calculate the elapsed time
  const    COMPONENT_SKU  = 0;             // For the jth row of the SKUsToWatch_ASSEMBLY, the SKU on the LEFT  (Used as an index for the conversionData array)
  const    ASSEMBLED_SKU  = 6;             // For the jth row of the SKUsToWatch_ASSEMBLY, the SKU on the RIGHT (Used as an index for the conversionData array)
  const CONVERSION_FACTOR = 4;             // For the jth row of the SKUsToWatch_ASSEMBLY, the conversion factor

  try
  {
    const   spreadsheet = SpreadsheetApp.getActive();
    const  assemblyData = spreadsheet.getSheetByName("SKUsToWatch_ASSEMBLY").getDataRange().getValues();
    const   exportSheet = spreadsheet.getSheetByName("ConvertedExport");
    const errorLogSheet = spreadsheet.getSheetByName("ErrorLog_Assembly");
    
    var data = spreadsheet.getSheetByName("DataImport").getDataRange().getValues();
    var assembledSkusAndQtys = {}, skusNotFound = [], exportData = [[],[],[]]

    // Set the appropriate indices based on the position of the following Strings in the header of the data 
    const locations = [data[0].indexOf('Richmond'), data[0].indexOf('Parksville'), data[0].indexOf('Rupert')];
    const SKU = data[0].indexOf('Item #');

    for (var i = 2; i < assemblyData.length; i++)
    {
      if (assembledSkusAndQtys.hasOwnProperty(assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase())) // The assembly SKU is already in the list
      {
        for (var l = 0; l < locations.length; l++)
        {
          if (assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l] < 0) // Need to compute an assembly
          {
            for (var k = 1; k < data.length; k++)
            {
              if (data[k][SKU] == assemblyData[i][COMPONENT_SKU])
              {
                data[k][locations[l]] = Number(data[k][locations[l]])
                data[k][locations[l]] += assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l]*Number(assemblyData[i][CONVERSION_FACTOR])
                exportData[l].push([data[k][SKU], data[k][locations[l]]])
                break;
              }
            }

            if (skuIsNotFound(k, data))
              skusNotFound.push([assemblyData[i][COMPONENT_SKU], '']);
          }
        }
      }
      else if (isNotBlank(assemblyData[i][ASSEMBLED_SKU])) // New assembly
      {
        for (var j = 1; j < data.length; j++)
        {
          if (data[j][SKU] == assemblyData[i][ASSEMBLED_SKU])
          {
            assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()] = [Number(data[j][locations[0]]), Number(data[j][locations[1]]), Number(data[j][locations[2]])]

            for (var l = 0; l < locations.length; l++)
            {
              if (data[j][locations[l]] < 0) // Need to compute an assembly
              {
                for (var k = 1; k < data.length; k++)
                {
                  if (data[k][SKU] == assemblyData[i][COMPONENT_SKU])
                  {
                    data[j][locations[l]] = 0
                    data[k][locations[l]] = Number(data[k][locations[l]])
                    data[k][locations[l]] += assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l]*Number(assemblyData[i][CONVERSION_FACTOR])
                    exportData[l].push([data[j][SKU], 0], [data[k][SKU], data[k][locations[l]]])
                    break;
                  }
                }

                if (skuIsNotFound(k, data))
                  skusNotFound.push([assemblyData[i][COMPONENT_SKU], '']);
              }
            }

            break;
          }
        }

        if (skuIsNotFound(j, data))
          skusNotFound.push(['', assemblyData[i][ASSEMBLED_SKU]]);
      }
    }

    exportData = exportData.map(storeAssemblies => uniqByKeepLast(storeAssemblies, sku => sku[0]));

    skusNotFound.unshift(['The following SKUs have not been found in the Adagio Database', ''],['Component Pieces', 'Assembled Goods']); // Header
    errorLogSheet.clearContents().getRange(1, 1, skusNotFound.length, 2).setValues(skusNotFound);
    exportSheet.getRange(8, 2, exportSheet.getLastRow(), 8).clearContent();
    exportData.map((loc, i) => {if (loc.length !== 0) exportSheet.getRange(8, 2 + 3*i, loc.length, 2).setValues(loc);})
    
    const headerRange = exportSheet.getRange(1, 1, 2, 8);
    const headerValues = headerRange.getValues();
    headerValues[0][0] = "Function Run Time";
    headerValues[0][4] = "Conversion Export";
    headerValues[1][7] = timeStamp(spreadsheet);
    headerValues[0][2] = elapsedTime(START_TIME);
    headerRange.setValues(headerValues)
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
* This function searches the Adagio database for items that have negative inventory AND are assembled from one or more components. If it is possible to convert inventory
* of certain components into an assembled good in order to reduce the number of occurences of negative inventory, then that computation is done for each location. Once complete,
* this function also searches the Adagio database for items that have negative inventory AND have multiple packaging sizes. If it is possible to convert inventoryof a certain 
* sized packaging into another in order to reduce the number of occurences of negative inventory, then that computation is done for each location.
*
* @author Jarren Ralf
*/
function computeConversions_Assembly_Then_PackageSize()
{
  const START_TIME = new Date().getTime(); // To calculate the elapsed time
  const CONVERSION_FACTOR = 4;             // For the jth row of the SKUsToWatch, the conversion factor
  const SMALLER_PACK_SKU  = 0;             // For the jth row of the SKUsToWatch, the SKU on the LEFT  (Used as an index for the conversionData array)
  const  LARGER_PACK_SKU  = 6;             // For the jth row of the SKUsToWatch, the SKU on the RIGHT (Used as an index for the conversionData array)
  const    COMPONENT_SKU  = 0;             // For the jth row of the SKUsToWatch_ASSEMBLY, the SKU on the LEFT  (Used as an index for the conversionData array)
  const    ASSEMBLED_SKU  = 6;             // For the jth row of the SKUsToWatch_ASSEMBLY, the SKU on the RIGHT (Used as an index for the conversionData array)

  try
  {
    const    spreadsheet = SpreadsheetApp.getActive();
    const conversionData = spreadsheet.getSheetByName("SKUsToWatch").getDataRange().getValues();
    const   assemblyData = spreadsheet.getSheetByName("SKUsToWatch_ASSEMBLY").getDataRange().getValues();
    const    exportSheet = spreadsheet.getSheetByName("ConvertedExport");
    const  errorLogSheet = spreadsheet.getSheetByName("ErrorLog");
    const  errorLogAssemblySheet = spreadsheet.getSheetByName("ErrorLog_Assembly");
    
    var data = spreadsheet.getSheetByName("DataImport").getDataRange().getValues();
    var assembledSkusAndQtys = {}, pairOfSKUs = [], skusNotFound = [], skusNotFound_Assembly = [], exportData = [[],[],[]], nonNegativeQty = [[],[],[]];

    // Set the appropriate indices based on the position of the following Strings in the header of the data 
    const locations = [data[0].indexOf('Richmond'), data[0].indexOf('Parksville'), data[0].indexOf('Rupert')];
    const SKU = data[0].indexOf('Item #');

    for (var i = 2; i < assemblyData.length; i++)
    {
      if (assembledSkusAndQtys.hasOwnProperty(assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase())) // The assembly SKU is already in the list
      {
        for (var l = 0; l < locations.length; l++)
        {
          if (assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l] < 0) // Need to compute an assembly
          {
            for (var k = 1; k < data.length; k++)
            {
              if (data[k][SKU] == assemblyData[i][COMPONENT_SKU])
              {
                data[k][locations[l]] = Number(data[k][locations[l]])
                data[k][locations[l]] += assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l]*Number(assemblyData[i][CONVERSION_FACTOR])
                exportData[l].push([data[k][SKU], data[k][locations[l]]])
                break;
              }
            }

            if (skuIsNotFound(k, data))
              skusNotFound_Assembly.push([assemblyData[i][COMPONENT_SKU], '']);
          }
        }
      }
      else if (isNotBlank(assemblyData[i][ASSEMBLED_SKU])) // New assembly
      {
        for (var j = 1; j < data.length; j++)
        {
          if (data[j][SKU] == assemblyData[i][ASSEMBLED_SKU])
          {
            assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()] = [Number(data[j][locations[0]]), Number(data[j][locations[1]]), Number(data[j][locations[2]])]

            for (var l = 0; l < locations.length; l++)
            {
              if (data[j][locations[l]] < 0) // Need to compute an assembly
              {
                for (var k = 1; k < data.length; k++)
                {
                  if (data[k][SKU] == assemblyData[i][COMPONENT_SKU])
                  {
                    data[j][locations[l]] = 0
                    data[k][locations[l]] = Number(data[k][locations[l]])
                    data[k][locations[l]] += assembledSkusAndQtys[assemblyData[i][ASSEMBLED_SKU].toString().toUpperCase()][l]*Number(assemblyData[i][CONVERSION_FACTOR])
                    exportData[l].push([data[j][SKU], 0], [data[k][SKU], data[k][locations[l]]])
                    break;
                  }
                }

                if (skuIsNotFound(k, data))
                  skusNotFound_Assembly.push([assemblyData[i][COMPONENT_SKU], '']);
              }
            }

            break;
          }
        }

        if (skuIsNotFound(j, data))
          skusNotFound_Assembly.push(['', assemblyData[i][ASSEMBLED_SKU]]);
      }
    }
    
    for (var j = 2; j < conversionData.length; j++)
    {
      // If one of the SKUs or the conversion factor is blank, then skip this iterate
      if (isBlank(conversionData[j][SMALLER_PACK_SKU]) || isBlank(conversionData[j][LARGER_PACK_SKU]) || isBlank(conversionData[j][CONVERSION_FACTOR]))
        continue;
      
      for (var i = 1; i < data.length; i++)
      {
        if (data[i][SKU] == conversionData[j][SMALLER_PACK_SKU] || data[i][SKU] == conversionData[j][LARGER_PACK_SKU]) // Locate both SKUs from the conversion sheet
           (data[i][SKU] == conversionData[j][SMALLER_PACK_SKU]) ? pairOfSKUs[0] = i : pairOfSKUs[1] = i;              // Contol the orientation of the pairOfSKUs array
        
        if (bothSKUsAreFound(pairOfSKUs)) // Once both of the SKUs are found, then get the conversions for all four locations
        {
          [data, exportData] = getConversions(data, exportData, conversionData[j], pairOfSKUs, SKU, CONVERSION_FACTOR, locations);
          pairOfSKUs.length = 0; // Empty the array by setting its length equal to 0
          break;
        }
      }
      if (skuIsNotFound(i, data))
        skusNotFound.push([conversionData[j][SMALLER_PACK_SKU], conversionData[j][LARGER_PACK_SKU]]);
    }

    exportData = exportData.map((storeAssemblies, s) => uniqByKeepLast(storeAssemblies, sku => sku[0]).filter(qty => {
      if (qty[1] < 0)
        return true
      else
      {
        nonNegativeQty[s].push(qty)
        return false
      }
    })).map((locationConversions, loc) => (locationConversions.length !== 0) ? locationConversions.concat([['' ,'']], nonNegativeQty[loc]) : nonNegativeQty[loc])

    skusNotFound.unshift(['One (or both) of the following SKUs have not been found in the Adagio Database', ''],['Smaller Pack SKU', 'Larger Pack SKU']); // Header
    skusNotFound_Assembly.unshift(['The following SKUs have not been found in the Adagio Database', ''],['Component Pieces', 'Assembled Goods']); // Header
    errorLogSheet.clearContents().getRange(1, 1, skusNotFound.length, 2).setValues(skusNotFound);
    errorLogAssemblySheet.clearContents().getRange(1, 1, skusNotFound_Assembly.length, 2).setValues(skusNotFound_Assembly);
    exportSheet.getRange(8, 2, exportSheet.getLastRow(), 8).clearContent();
    exportData.map((loc, i) => {if (loc.length !== 0) exportSheet.getRange(8, 2 + 3*i, loc.length, 2).setValues(loc);})
    
    const headerRange = exportSheet.getRange(1, 1, 2, 8);
    const headerValues = headerRange.getValues();
    headerValues[0][0] = "Function Run Time";
    headerValues[0][4] = "Conversion Export";
    headerValues[1][7] = timeStamp(spreadsheet);
    headerValues[0][2] = elapsedTime(START_TIME);
    headerRange.setValues(headerValues)
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
* This function searches the Adagio database for items that have negative inventory AND have multiple packaging sizes. If it is possible to convert inventory
* of a certain sized packaging into another in order to reduce the number of occurences of negative inventory, then that computation is done for each location.
*
* @author Jarren Ralf
*/
function computeConversions_PackageSize()
{
  const START_TIME = new Date().getTime(); // To calculate the elapsed time
  const SMALLER_PACK_SKU  = 0;             // For the jth row of the SKUsToWatch, the SKU on the LEFT  (Used as an index for the conversionData array)
  const  LARGER_PACK_SKU  = 6;             // For the jth row of the SKUsToWatch, the SKU on the RIGHT (Used as an index for the conversionData array)
  const CONVERSION_FACTOR = 4;             // For the jth row of the SKUsToWatch, the conversion factor

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const conversionData = spreadsheet.getSheetByName("SKUsToWatch").getDataRange().getValues();
    const    exportSheet = spreadsheet.getSheetByName("ConvertedExport");
    const  errorLogSheet = spreadsheet.getSheetByName("ErrorLog");
    
    var data = spreadsheet.getSheetByName("DataImport").getDataRange().getValues();
    var  pairOfSKUs = [], skusNotFound = [], exportData = [[],[],[]];

    // Set the appropriate indices based on the position of the following Strings in the header of the data 
    const locations = [data[0].indexOf('Richmond'), data[0].indexOf('Parksville'), data[0].indexOf('Rupert')];
    const SKU = data[0].indexOf('Item #');
    
    for (var j = 2; j < conversionData.length; j++)
    {
      // If one of the SKUs or the conversion factor is blank, then skip this iterate
      if (isBlank(conversionData[j][SMALLER_PACK_SKU]) || isBlank(conversionData[j][LARGER_PACK_SKU]) || isBlank(conversionData[j][CONVERSION_FACTOR]))
        continue;
      
      for (var i = 1; i < data.length; i++)
      {
        if (data[i][SKU] == conversionData[j][SMALLER_PACK_SKU] || data[i][SKU] == conversionData[j][LARGER_PACK_SKU]) // Locate both SKUs from the conversion sheet
           (data[i][SKU] == conversionData[j][SMALLER_PACK_SKU]) ? pairOfSKUs[0] = i : pairOfSKUs[1] = i;              // Contol the orientation of the pairOfSKUs array
        
        if (bothSKUsAreFound(pairOfSKUs)) // Once both of the SKUs are found, then get the conversions for all four locations
        {
          [data, exportData] = getConversions(data, exportData, conversionData[j], pairOfSKUs, SKU, CONVERSION_FACTOR, locations);
          pairOfSKUs.length = 0; // Empty the array by setting its length equal to 0
          break;
        }
      }
      if (skuIsNotFound(i, data))
        skusNotFound.push([conversionData[j][SMALLER_PACK_SKU], conversionData[j][LARGER_PACK_SKU]]);
    }

    exportData = exportData.map(storeConversions => uniqByKeepLast(storeConversions, sku => sku[0]));

    skusNotFound.unshift(['One (or both) of the following SKUs have not been found in the Adagio Database', ''],['Smaller Pack SKU', 'Larger Pack SKU']); // Header
    errorLogSheet.clearContents().getRange(1, 1, skusNotFound.length, 2).setValues(skusNotFound);
    exportSheet.getRange(8, 2, exportSheet.getLastRow(), 8).clearContent();
    exportData.map((loc, i) => {if (loc.length !== 0) exportSheet.getRange(8, 2 + 3*i, loc.length, 2).setValues(loc);})
    
    const headerRange = exportSheet.getRange(1, 1, 2, 8);
    const headerValues = headerRange.getValues();
    headerValues[0][0] = "Function Run Time";
    headerValues[0][4] = "Conversion Export";
    headerValues[1][7] = timeStamp(spreadsheet);
    headerValues[0][2] = elapsedTime(START_TIME);
    headerRange.setValues(headerValues)
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
* This function searches the Adagio database for seasonal yeti products that need to have their inventory transfered to the seasonal discontinued sku.
*
* @author Jarren Ralf
*/
function computeConversions_Yeti()
{
  const START_TIME = new Date().getTime(); // To calculate the elapsed time
  const SEASONAL_SKU  = 0;                 // For the jth row of the Yeti SKUsToWatch, the SKU on the LEFT  (Used as an index for the conversionData array)
  const  DISCONT_SKU  = 5;                 // For the jth row of the Yeti SKUsToWatch, the SKU on the RIGHT (Used as an index for the conversionData array)

  try
  {
    const spreadsheet = SpreadsheetApp.getActive();
    const conversionData = spreadsheet.getSheetByName("Yeti SKUsToWatch").getDataRange().getValues();
    const    exportSheet = spreadsheet.getSheetByName("ConvertedExport");
    const  errorLogSheet = spreadsheet.getSheetByName("Yeti ErrorLog");
    
    var data = spreadsheet.getSheetByName("DataImport").getDataRange().getValues();
    var  pairOfSKUs = [], skusNotFound = [], exportData = [[],[],[]];

    // Set the appropriate indices based on the position of the following Strings in the header of the data 
    const locations = [data[0].indexOf('Richmond'), data[0].indexOf('Parksville'), data[0].indexOf('Rupert')];
    const SKU  = data[0].indexOf('Item #');

    for (var j = 2; j < conversionData.length; j++)
    {
      // If one of the SKUs is blank, then skip this iterate
      if (isBlank(conversionData[j][SEASONAL_SKU]) || isBlank(conversionData[j][DISCONT_SKU]))
        continue;
      
      for (var i = 1; i < data.length; i++)
      {
        if (data[i][SKU] == conversionData[j][SEASONAL_SKU] || data[i][SKU] == conversionData[j][DISCONT_SKU]) // Locate both SKUs from the conversation sheet
          (data[i][SKU] == conversionData[j][SEASONAL_SKU]) ? pairOfSKUs[0] = i : pairOfSKUs[1] = i;           // Contol the orientation of the pairOfSKUs array
        
        if (bothSKUsAreFound(pairOfSKUs)) // Once both of the SKUs are found, then get the conversions for all four locations
        {
          [data, exportData] = getConversions_Yeti(data, exportData, pairOfSKUs, SKU, locations);
          pairOfSKUs.length = 0; // Empty the array by setting its length equal to 0
          break;
        }
      }
      if (skuIsNotFound(i, data))
        skusNotFound.push([conversionData[j][SEASONAL_SKU], conversionData[j][DISCONT_SKU]]);
    }

    skusNotFound.unshift(['One (or both) of the following SKUs have not been found in the Adagio Database', ''],['Active Seasonal SKU', 'Discontinued Seasonal SKU']); // Header
    errorLogSheet.clearContents().getRange(1, 1, skusNotFound.length, 2).setValues(skusNotFound);
    exportSheet.getRange(8, 2, exportSheet.getLastRow(), 8).clearContent();
    exportData.map((loc_, i) => {
      var loc = uniqByKeepLast(loc_, sku => sku[0]);
      if (loc.length !== 0) exportSheet.getRange(8, 2 + 3*i, loc.length, 2).setValues(loc);
    })
    
    const headerRange = exportSheet.getRange(1, 1, 2, 8);
    const headerValues = headerRange.getValues();
    headerValues[0][0] = "Function Run Time";
    headerValues[0][4] = "Conversion Export";
    headerValues[1][7] = timeStamp(spreadsheet);
    headerValues[0][2] = elapsedTime(START_TIME);
    headerRange.setValues(headerValues)
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
* This function calculates the ellapsed time of the computeConversions function.
*
* @param  {Number} startTime   : The start time that the script began running at represented by a number in milliseconds
* @return {Number} The total time taken for the function to run
* @author Jarren Ralf
*/
function elapsedTime(startTime)
{
  return (new Date().getTime() - startTime)/1000;
}

/**
* This function retrieves the conversion data for the j-th conversion at the given store location.
*
* @param {Object[][]}           data : The Adagio data
* @param {Object[][]}     exportData : The set of SKUs and QTYs for the completed conversions of the given location so far
* @param {Object[]}   conversionData : The conversion data for the j-th item
* @param {Object[]}       pairOfSKUs : The indices of the Adagio data of which the j-th conversion is based on
* @param {Number}                SKU : The index position of the SKU in the data array
* @param {Number}  CONVERSION_FACTOR : The index position of the conversion factor in the data array
* @param {Number}           location : An array index for the Adagio data which represents which location is being checked for possible conversions
* @return {Object[][], Object[][]} [data, exportData] : The Adagio data and the set of SKUs and QTYs for the completed conversions of the given location so far (both possibly updated)
* @author Jarren Ralf
*/
function getConversions(data, exportData, conversionData, pairOfSKUs, SKU, CONVERSION_FACTOR, locations)
{
  const SMALLER_PACK = 0; // For the jth row of the SKUsToWatch, the SKU on the LEFT  (Used as an index for the pairOfSKUs array)
  const  LARGER_PACK = 1; // For the jth row of the SKUsToWatch, the SKU on the RIGHT (Used as an index for the pairOfSKUs array)

  for (l = 0; l < locations.length; l++)
  {
    data[pairOfSKUs[SMALLER_PACK]][locations[l]] = Number(data[pairOfSKUs[SMALLER_PACK]][locations[l]]) // Make sure this variable is a NUMBER and not a string
    data[pairOfSKUs [LARGER_PACK]][locations[l]] = Number(data[pairOfSKUs [LARGER_PACK]][locations[l]]) // Make sure this variable is a NUMBER and not a string

    if (!isNonPositive(data[pairOfSKUs[SMALLER_PACK]][locations[l]]) || !isNonPositive(data[pairOfSKUs[LARGER_PACK]][locations[l]])) // If both items have a non-positive stock, don't enter the loop
    {
      if (isNegative(data[pairOfSKUs[SMALLER_PACK]][locations[l]])) // The item with the smaller pack is negative
      {  
        var numPacksNeeded = -1*Math.floor(data[pairOfSKUs[SMALLER_PACK]][locations[l]]/conversionData[CONVERSION_FACTOR]);
        
        if (numPacksNeeded < data[pairOfSKUs[LARGER_PACK]][locations[l]]) // Compute the relevant conversions
        {
          data[pairOfSKUs[LARGER_PACK]][locations[l]] -= numPacksNeeded;
          exportData[l].push([data[pairOfSKUs[LARGER_PACK]][SKU], data[pairOfSKUs[LARGER_PACK]][locations[l]]]);
          data[pairOfSKUs[SMALLER_PACK]][locations[l]] += numPacksNeeded*conversionData[CONVERSION_FACTOR];
          exportData[l].push([data[pairOfSKUs[SMALLER_PACK]][SKU], data[pairOfSKUs[SMALLER_PACK]][locations[l]]]);
        }
        else // There aren't enough of the LARGER_PACK_SKU items to make the inventory of the SMALLER_PACK_SKU positive, but conversion is done nonetheless
        {
          data[pairOfSKUs[SMALLER_PACK]][locations[l]] += data[pairOfSKUs[LARGER_PACK]][locations[l]]*conversionData[CONVERSION_FACTOR];
          exportData[l].push([data[pairOfSKUs[SMALLER_PACK]][SKU], data[pairOfSKUs[SMALLER_PACK]][locations[l]]]);
          data[pairOfSKUs[LARGER_PACK]][locations[l]] = 0;
          exportData[l].push([data[pairOfSKUs[LARGER_PACK]][SKU], data[pairOfSKUs[LARGER_PACK]][locations[l]]]);
        }
      }
      else if (isNegative(data[pairOfSKUs[LARGER_PACK]][locations[l]])) // The item with the larger pack is negative
      {
        if (-1*data[pairOfSKUs[LARGER_PACK]][locations[l]]*conversionData[CONVERSION_FACTOR] > data[pairOfSKUs[SMALLER_PACK]][locations[l]]) // Compute the relevant conversions
        {
          data[pairOfSKUs[LARGER_PACK]][locations[l]] += Math.floor(data[pairOfSKUs[SMALLER_PACK]][locations[l]]/conversionData[CONVERSION_FACTOR]); 
          exportData[l].push([data[pairOfSKUs[LARGER_PACK]][SKU], data[pairOfSKUs[LARGER_PACK]][locations[l]]]);
          data[pairOfSKUs[SMALLER_PACK]][locations[l]] %= conversionData[CONVERSION_FACTOR];
          exportData[l].push([data[pairOfSKUs[SMALLER_PACK]][SKU], data[pairOfSKUs[SMALLER_PACK]][locations[l]]]);
        }
        else // There aren't enough of the SMALLER_PACK_SKU items to make the inventory of the LARGER_PACK_SKU positive, but conversion is done nonetheless
        {
          data[pairOfSKUs[SMALLER_PACK]][locations[l]] += data[pairOfSKUs[LARGER_PACK]][locations[l]]*conversionData[CONVERSION_FACTOR];
          exportData[l].push([data[pairOfSKUs[SMALLER_PACK]][SKU], data[pairOfSKUs[SMALLER_PACK]][locations[l]]]);
          data[pairOfSKUs[LARGER_PACK]][locations[l]] = 0;
          exportData[l].push([data[pairOfSKUs[LARGER_PACK]][SKU], data[pairOfSKUs[LARGER_PACK]][locations[l]]]);
        }
      }
    }
  }

  return [data, exportData];
}

/**
* This function retrieves the yeti conversion data for the j-th conversion at the given location.
*
* @param {Object[][]}           data : The Adagio data
* @param {Object[][]}     exportData : The set of SKUs and QTYs for the completed yeti conversions of the given location so far
* @param {Object[]}       pairOfSKUs : The indices of the Adagio data of which the j-th yeti conversion is based on
* @param {Number}                SKU : The index position of the SKU in the data array
* @param {Number}           location : An array index for the Adagio data which represents which location is being checked for possible yeti conversions
* @return {Object[][], Object[][]} [data, exportData] : The Adagio data and the set of SKUs and QTYs for the completed yeti conversions of the given location so far (both possibly updated)
* @author Jarren Ralf
*/
function getConversions_Yeti(data, exportData, pairOfSKUs, SKU, locations)
{
  const  ACTIVE = 0; // For the jth row of the SKUsToWatch, the SKU on the LEFT  (Used as an index for the pairOfSKUs array)
  const DISCONT = 1; // For the jth row of the SKUsToWatch, the SKU on the RIGHT (Used as an index for the pairOfSKUs array)

  for (l = 0; l < locations.length; l++)
  {
    data[pairOfSKUs[ACTIVE]][locations[l]] = Number(data[pairOfSKUs[ACTIVE]][locations[l]]); // Make sure this variable is a NUMBER and not a string

    if (!isNonPositive(data[pairOfSKUs[ACTIVE]][locations[l]])) // If the active seasonal yeti has zero or negative stock, then don't compute
    {
      data[pairOfSKUs[DISCONT]][locations[l]] = Number(data[pairOfSKUs[DISCONT]][locations[l]]) // Make sure this variable is a NUMBER and not a string
      data[pairOfSKUs[DISCONT]][locations[l]] += Number(data[pairOfSKUs[ACTIVE]][locations[l]]);
      data[pairOfSKUs[ACTIVE]][locations[l]] = 0;
      exportData[l].push([data[pairOfSKUs[ACTIVE]][SKU], data[pairOfSKUs[ACTIVE]][locations[l]]]);
      exportData[l].push([data[pairOfSKUs[DISCONT]][SKU], data[pairOfSKUs[DISCONT]][locations[l]]]);
    }
  }

  return [data, exportData];
}

/**
* This function checks if a given string is blank.
*
* @param  {String} str : The given string
* @return {Boolean}    : Whether the string is blank or not
* @author Jarren Ralf
*/
function isBlank(str)
{
  return str == '';                  
}

/**
* This function checks if a given number is negative.
*
* @param  {Number} num : The given number
* @return {Boolean}    : Whether the number is negative or not
* @author Jarren Ralf
*/
function isNegative(num)
{
  return num < 0;                  
}

/**
* This function checks if a given number is negative or zero.
*
* @param  {Number} num : The given number
* @return {Boolean}    : Whether the number is non-positive or not
* @author Jarren Ralf
*/
function isNonPositive(num)
{
  return num <= 0;                  
}

/**
 * This function finds possible SKUs To Watch, by looking for groups of sku numbers that are full contained in other sku numbers
 * 
 * @author Jarren Ralf
 */
function possibleSKUsToWatch()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const possibleSkusSheet = spreadsheet.getSheetByName('Second Possible SKUsToWatch')
  const currentSKUstoWatchSheet = spreadsheet.getSheetByName('SKUsToWatch')
  const currentSKUstoWatch = currentSKUstoWatchSheet.getSheetValues(3, 1, currentSKUstoWatchSheet.getLastRow() - 2, 9)
  const numSkus = currentSKUstoWatch.length;
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()).map(u => u.map(v => v.toString().toUpperCase()))
  csvData.shift(); // Remove the hot spot item which has a sku of '0'
  const numItems = csvData.length
  var data = [], sku1, sku11, sku111, sku2, sku22, sku222;

  for (var i = 0; i < numItems; i++)
  {
    sku1 = csvData[i][6].substring(0, 8);
    sku11 = csvData[i][6].substring(0, 9);
    sku111 = csvData[i][6].substring(0, 11)

    for (var ii = 0; ii < numItems; ii++)
    {
      sku2 = csvData[ii][6].substring(0, 8);
      sku22 = csvData[ii][6].substring(0, 9);
      sku222 = csvData[ii][6].substring(0, 11);

      if (csvData[i][6].length == 1 || csvData[ii][6].length == 1 || csvData[i][6].length == 2 || csvData[ii][6].length == 2 ||
          csvData[i][6].length == 3 || csvData[ii][6].length == 3 || csvData[i][6].length == 4 || csvData[ii][6].length == 4 ||
          csvData[i][6].length == 5 || csvData[ii][6].length == 5 || csvData[i][6].length == 6 || csvData[ii][6].length == 6 || // Ignore skus less than 7 letter
          sku1 == '16010001' || sku2 == '16010001' || // Rigged cuttlefish
          sku1 == '16010005' || sku2 == '16010005' || // cuttlefish
          sku1 == '16070001' || sku2 == '16070001' || // rigged squid
          sku1 == '16020000' || sku2 == '16020000' || // Michael Bait
          sku1 == '16020001' || sku2 == '16020001' || // Michael Bait w/ Swivel
          sku1 == '16020011' || sku2 == '16020011' || // Michael Bait w/ Swivel
          sku1 == '16030000' || sku2 == '16030000' || // Mini Mackeral
          sku1 == '16030001' || sku2 == '16030001' || // Mini Mackeral Rigged
          sku1 == '16050001' || sku2 == '16050001' || // Needlefish Rigged
          sku1 == '16050005' || sku2 == '16050005' || // Needlefish
          sku1 == '16060001' || sku2 == '16060001' || // Octopus
          sku1 == '16060005' || sku2 == '16060005' || // Octopus
          sku1 == '16060010' || sku2 == '16060010' || // Octopus
          sku1 == '16060065' || sku2 == '16060065' || // Octopus
          sku1 == '16060100' || sku2 == '16060100' || // Octopus
          sku1 == '16060175' || sku2 == '16060175' || // Octopus
          sku1 == '16061000' || sku2 == '16061000' || // Octopus
          sku1 == '16070000' || sku2 == '16070000' || // Squid
          sku1 == '16200000' || sku2 == '16200000' || // GB Cuttlefish
          sku1 == '16200021' || sku2 == '16200021' || // GB Mini Sardine
          sku1 == '16200025' || sku2 == '16200025' || // GB Needlefish
          sku1 == '16200030' || sku2 == '16200030' || // GB Octopus
          sku1 == '16200061' || sku2 == '16200061' || // GB Michael Bait
          sku1 == '16200065' || sku2 == '16200065' || // GB Squid
          sku1 == '17020001' || sku2 == '17020001' || // Mylar Insert
          sku1 == '17020002' || sku2 == '17020002' || // Mylar Insert
          sku1 == '17020004' || sku2 == '17020004' || // Pline Squid
          sku1 == '170810ZG1' || sku2 == '170810ZG1' || // Zukers
          sku1 == '18702000' || sku2 == '18702000' || // LIGHTHOUSE BIG EYE DERBY WINNER 3.5
          sku1 == '18702001' || sku2 == '18702001' || // LIGHTHOUSE BIG EYE MINI 3.0 DERBY WINNER
          sku1 == '1MLS1010' || sku2 == '1MLS1010' || // PELAGIC MLS1010 DELUXE LS TEE
          sku1 == '80002000' || sku2 == '80002000' || // SUPERSTAR200LB 200YDS
          sku1 == '80003000' || sku2 == '80003000' || // KINGFISHER #3 RIGGED
          sku1 == '80003000' || sku2 == '80003000' || // KINGFISHER #3 RIGGED
          sku1 == '80003100' || sku2 == '80003100' || // COHO KILLER RIGGED ALL COLOURS
          sku1 == '80003500' || sku2 == '80003500' || // KINGFISHER #3.5 RIGGED
          sku1 == '80004001' || sku2 == '80004001' || // KINGFISHER #4 RIGGED ALL COLOURS
          sku11 == '170901001' || sku22 == '170901001' || // zuke
          sku111 == '1708XRMAG20' || sku222 == '1708XRMAG20' || // RAPALA X-RAP MAGNUM G20 DIVEBAIT 20 FT
          sku111 == '1708XRMAG30' || sku222 == '1708XRMAG30' || // RAPALA X-RAP MAGNUM G20 DIVEBAIT 30 FT
          sku111 == '1708XRMAG40' || sku222 == '1708XRMAG40' || // RAPALA X-RAP MAGNUM G20 DIVEBAIT 40 FT
          (csvData[i][6] == csvData[ii][6])) // Ignore the same SKUs
        continue;
      // Sku includes another sku, their UoM are not the same, and both are active in Adagio
      else if (csvData[i][6].includes(csvData[ii][6]) && (csvData[i][0] != csvData[ii][0]) &&csvData[i][10] == 'A' && csvData[ii][10] == 'A')
      {
        for (var j = 0; j < numSkus; j++)
        {
          if ((csvData[ i][6] == currentSKUstoWatch[j][0] && csvData[ii][6] == currentSKUstoWatch[j][6]) ||
              (csvData[ii][6] == currentSKUstoWatch[j][0] && csvData[ i][6] == currentSKUstoWatch[j][6]))
            break;
        } 

        if (j === numSkus)
          data.push([csvData[i][6], csvData[i][1], csvData[i][0], null, null, null, csvData[ii][6], csvData[ii][1], csvData[ii][0]])
      }
    }
  }

  possibleSkusSheet.getRange(3, 1, data.length, 9).setValues(data)
}

/**
* This function imports the data.
*
* @author Jarren Ralf
*/
function resetData()
{
  const START_TIME = new Date().getTime(); // To calculate the elapsed time

  try
  {
    var isActive, numberFormats = [];
    const spreadsheet = SpreadsheetApp.getActive();
    const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
    const inflowData = Object.values(Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString()).reduce((acc, val) => {
      // Sum the quantities if item is in multiple locations
      if (acc[val[0]]) acc[val[0]][1] = (inflow_conversions.hasOwnProperty(val[0])) ? Number(acc[val[0]][1]) + Number(val[4])*inflow_conversions[val[0]] : Number(acc[val[0]][1]) + Number(val[4]); 
      // Add the item to the new list if it contains the typical google sheets item format with "space - space"
      else if (val[0].split(" - ").length > 4) acc[val[0]] = [val[0], (inflow_conversions.hasOwnProperty(val[0])) ? Number(val[4])*inflow_conversions[val[0]] : Number(val[4])]; 
      return acc;
    }, {}));
    var isInFlowItem;
    const header = csvData.shift();               // Remove the header
    const active = header.indexOf('Active Item'); // Index of the Active Item column
    const activeItems = csvData.filter(item => {
      isActive = item[active] === 'A'; // These are the active items

      if (isActive)
      {
        numberFormats.push(['@', '@', '#.#', '#.#', '#.#', '#.#', '@']); // Ensure that the inventory values are Numbers (for math operations) and the rest are Strings
        item.splice(6, 7, item[6]); // Remove the unnecessary columns, starting with trites inventory while keeping sku
        isInFlowItem = inflowData.find(description => description[0].split(" - ", 1)[0] == item[6])
        item[5] = (isInFlowItem) ? isInFlowItem[1] : ''; // Add Trites inventory values if they are found in inFlow
      }
      return isActive // Remove the inactive Items
    }); 

    header.splice(6, 7, 'Item #');
    numberFormats.unshift(['@', '@', '@', '@', '@', '@', '@']) // Headers are Strings
    const numRows = activeItems.unshift(header); // Put the header back
    spreadsheet.getSheetByName("DataImport").clearContents().getRange(1, 1, numRows, activeItems[0].length).setNumberFormats(numberFormats).setValues(activeItems);
    
    const runtime = elapsedTime(START_TIME);
    spreadsheet.getSheetByName('ConvertedExport').getRange(1, 3, 2).setValues([[runtime], [timeStamp(spreadsheet)]]); // Elapsed time and timestamp

    const timeStampRng = spreadsheet.getSheetByName('Adagio Transfer Sheet').getRange(1, 4, 2, 6);
    const timeStampValues = timeStampRng.getValues();
    timeStampValues[0][5] = runtime + ' seconds';
    timeStampValues[1][0] = timeStamp(spreadsheet);
    timeStampRng.setValues(timeStampValues); // Elapsed time and timestamp
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
* This function calls the resetData function and passes the function a true value which ensure that a timestamp is placed on the ConvertedExport page
*
* @author Jarren Ralf
*/
function resetData_ButtonClicked_On_ConvertedExport()
{
  Logger.log('This function should not run!')
  resetData(true);
}

/**
* This function checks if the current index is the same as the length of the Adagio database, because then this means the SKU that was being looked for was not found.
*
* @param    {Number}     i  : The current index through the Adagio data
* @param  {Object[][]} data : The Adagio data
* @return   {Boolean}       : Whether one (or both) of the j-th SKUs in the conversion data has been found in the Adagio database
* @author Jarren Ralf
*/
function skuIsNotFound(i, data)
{
  return i == data.length;
}