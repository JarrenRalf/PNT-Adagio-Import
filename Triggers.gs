function deleteTriggers()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger));
}

function createTriggers()
{
  // This is an installable onEdit trigger
  ScriptApp.newTrigger('installedOnEdit').forSpreadsheet('130AZ1zZ0ImnIB-uA8juMGUAfOswfxD7rH0ldVNk1Zq0').onEdit().create()

  // AdagioImport
  ScriptApp.newTrigger("runAll").timeBased().atHour(5).everyDays(1).create();

  // SKU Unit Conversion
  ScriptApp.newTrigger("resetData").timeBased().atHour(4).everyDays(1).create(); // resetData must run before runAll 

  // Dashboard Control (Transfer sheets)
  ScriptApp.newTrigger('updateDashboard').timeBased().atHour(5).everyDays(1).create();

  // Richmond
  ScriptApp.newTrigger('richmond_applyFullSpreadsheetFormatting').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('richmond_updateRecentlyCreatedItems').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('richmond_updateUPC_Database').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('richmond_logCountsOnWorkdays').timeBased().atHour(23).everyDays(1).create();
  ScriptApp.newTrigger('richmond_updateSearchData').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('richmond_generateSuggestedInflowPick').timeBased().atHour(8).everyDays(1).create();

  // Parksville
  ScriptApp.newTrigger('parksville_applyFullSpreadsheetFormatting').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('parksville_updateRecentlyCreatedItems').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('parksville_updateUPC_Database').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('parksville_logCountsOnWorkdays').timeBased().atHour(23).everyDays(1).create();
  ScriptApp.newTrigger('parksville_updateSearchData').timeBased().atHour(9).everyDays(1).create();

  // Rupert
  ScriptApp.newTrigger('rupert_applyFullSpreadsheetFormatting').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('rupert_updateRecentlyCreatedItems').timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger('rupert_updateUPC_Database').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('rupert_logCountsOnWorkdays').timeBased().atHour(23).everyDays(1).create();
  ScriptApp.newTrigger('rupert_updateSearchData').timeBased().atHour(9).everyDays(1).create();
}