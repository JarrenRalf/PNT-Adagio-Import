function myFunction() {
  for (var key in this) {
    if (typeof this[key] == "function") {
      Logger.log(this[key])
    }
  }
}

function myFunction2() {
  for (var key in this) {
    console.log(key)
  }
}

/**
 *
 * 
 * So I made custom function based on feedback from https://stackoverflow.com/questions/42887569/using-google-apps-script-how-to-convert-export-a-drive-file
 *
 * @author Stephane Giron https://medium.com/@stephane.giron/quick-and-not-so-dirty-backup-solution-for-google-apps-script-code-b7b402b3d89
 */
function backupScripts()
{
  const json = {files:[
    {"scriptId" : "1xH-bHVDzmYD3OLWQx85zgynEe21JhKABiTA8lmqyRRbzLGPUYoD29Qkr",  "docId" : '1YwJwScU7xNHVqEzJi104hESKMSU6DVwr4HoHMEHQKXo'}
  ]}

  for(let i = 0 ; i < json.files.length ; i++)
  {
    const file = json.files[i];
    const content = getExportFile(file.scriptId, "application/vnd.google-apps.script+json")

    if(!content)
      continue ;

    const script = JSON.parse(content);
    const doc = DocumentApp.openById(file.docId);
    const body = doc.getBody();
    body.clear();

    script.files.forEach(function (item)
    {
      body.appendParagraph(item.name+'.'+item.type).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(item.source)
    })
  }
}


/**
 *
 * Direct use of Drive.Files.export() does not work
 * Error: GoogleJsonResponseException: API call to drive.files.export failed with error: Export requires alt=media to download the exported content.
 * 
 * So I made custom function based on feedback from https://stackoverflow.com/questions/42887569/using-google-apps-script-how-to-convert-export-a-drive-file
 * @author Stephane Giron https://medium.com/@stephane.giron/quick-and-not-so-dirty-backup-solution-for-google-apps-script-code-b7b402b3d89
 */
function getExportFile(id,mimeType)
{
  let file = Drive.Files.get(id, {supportsAllDrives: true})
  // We set 31 because the trigger will run each 30 minutes
  //if(new Date(file.modifiedDate).getTime() > (new Date().getTime() - (31*60*1000)))
  //{
    let url = file.exportLinks[mimeType]

    var response = UrlFetchApp.fetch(url, { 
      headers: { 
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() 
      } 
    }); 
    return response.getContentText(); 
  //}
  return false;
}