var ui = SpreadsheetApp.getUi()

function getToken(client_id, client_secret) {
  
  var token = null;
  
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : JSON.stringify({
      "grant_type": "client_credentials",
      "client_id": client_id	 ,
      "client_secret": client_secret
    })
  };
  
  try {
    token = UrlFetchApp.fetch('https://five.epicollect.net/api/oauth/token', options);  
  } catch(e) {}
  
  return token;
}

function extract(entries, iteration) {
  var temp, array = [], e;
  for(var i = 0; i < entries.length; i++) {
      temp = [];
      e = entries[i];
    
      //header
      if(i === 0 && iteration === 1) {
        Object.keys(e).forEach(function(key,index) {
          temp.push(key);
        });
        
        if(temp != null && temp.length > 0 && entries[0] != null) {
          array.push(adjustHeaderForGPS(temp, entries[0]));
        }
        
        temp = [];
      } 
    
        Object.keys(e).forEach(function(key,index) {     
          if(e[key]["latitude"] != null && e[key]["longitude"] != null && e[key]["accuracy"] != null){
            if(e[key]["longitude"] != "" && e[key]["latitude"] != "" && e[key]["accuracy"] != "") {
              //temp.push("");
              temp.push(e[key]["accuracy"].toString());
              temp.push(e[key]["latitude"].toString());
              temp.push(e[key]["longitude"].toString());
            } else {
              //temp.push("");
              temp.push("");
              temp.push("");
              temp.push("");
            }
          } else {
            temp.push(e[key]);          
          }
        });
      
        
      array.push(temp);
  }
  
  return array;
}

function adjustHeaderForGPS(header, random_row) {
  var counter = 0, so_far = 0;
  Object.keys(random_row).forEach(function(key,index) {   
    if(random_row[key]["latitude"] != null && random_row[key]["longitude"] != null && random_row[key]["accuracy"] != null){
      index = header.indexOf(key);
      header.splice(1 + counter + so_far * 2, 0, key + "_Longitude");
      header.splice(1 + counter + so_far * 2, 0, key + "_Latitude");
      header.splice(1 + counter + so_far * 2, 0, key + "_Accuracy_m");
      header.splice(index, 1);
      so_far++;
    }
    
    counter++;
  });
  
  return header;
}

function onOpen() {
  var menu = ui.createMenu("Epicollect 5");
  
  menu.addItem("Get Survey Data", "getData");
  menu.addItem("Get Survey Media", "getMediaFiles");
  menu.addItem("Reset settings", "resetSettings");

  
  menu.addToUi();
}

function resetSettings(){
  var prop = PropertiesService.getScriptProperties();
  
  prop.setProperty("Is Public Data", false);
  prop.setProperty("Client Id", "");
  prop.setProperty("Client Secret", "");
  prop.setProperty("Survey Name", "");
  
}

function getData(){
  var prop = PropertiesService.getScriptProperties()
  var isPublicData = prop.getProperty("Is Public Data") === "true"
  var client_id = prop.getProperty("Client Id")
  var client_secret = prop.getProperty("Client Secret")
  var survey_name = prop.getProperty("Survey Name")
  
  if(client_id == null || 
     client_secret == null || 
     survey_name == null ||
     client_id == "" ||
     client_secret == "" ||
     survey_name == "") {
    
    ui.alert("Please enter your Client ID, Client Secret, and the name of your survey").OK
    
    let is_public = ui.alert('Is Public Data?', ui.ButtonSet.YES_NO)
    
    if(is_public == ui.Button.NO){
      isPublicData = false;
      client_id = ui.prompt("Please enter your Client Id").getResponseText();
      client_secret = ui.prompt("Please enter your Client Secret").getResponseText();
    }else{
      isPublicData = true;
      client_id = "n/a";
      client_secret = "n/a";
    }
   
    survey_name = ui.prompt("Please enter the name of the survey as it appears in the http:// bar \n(ex. ec5-demo-project not EC5 DEMO PROJECT) ").getResponseText()
    
    prop.setProperty("Client Id",client_id)
    prop.setProperty("Client Secret",client_secret)
    prop.setProperty("Survey Name",survey_name)
    prop.setProperty("Is Public Data", isPublicData)
  }
  
  let options = {
    'method' : 'get',
    'headers' : {
      'Content-Type' : 'application/json',
    }
  }
  
  if(!isPublicData){
    var token = getToken(client_id, client_secret);
    
    if(token == null)
      return;
    
    options = {
      'method' : 'get',
      'headers' : {
        'Content-Type' : 'application/json',
        'Authorization' : 'Bearer ' + JSON.parse(token).access_token
      }
    };
  }
  
  var DAYS = 120;
  var cut_off_date = new Date( (new Date()).getTime() - 1000 * 60 * 60 * 24 * DAYS);
  var body = JSON.parse(UrlFetchApp.fetch('https://five.epicollect.net/api/export/entries/'+ survey_name +'?sort_by=created_at&sort_order=DESC&filter_by=created_at&filter_from=' + cut_off_date.toISOString(), options));
  var current = body.meta.current_page;
  var last = body.meta.last_page;
  var array = [], temp, res;
  
  array = array.concat(extract(body.data.entries, 1));
  
  if(current < last) {
    for(var j = 2; j <= last; j++) {
      Utilities.sleep(1500);
      res = JSON.parse(UrlFetchApp.fetch('https://five.epicollect.net/api/export/entries/'+ survey_name +'?sort_by=created_at&sort_order=DESC&filter_by=created_at&filter_from=' + cut_off_date.toISOString() + '&page=' + j, options));
      array = array.concat(extract(res.data.entries, j));
    }      
  }
  
  var last_row = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow();
  var last_column = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastColumn();
  
  if(last_row != 0 && last_column != 0) {
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1, last_row, last_column).clear({contentsOnly: true});
  }
  
  if(array != null && array.length > 0) {
    if (last_row != 0 || last_column != 0){
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1, array.length, array[0].length).setValues(array);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1, array.length, array[0].length).setValues(array);
    }
  }
}

function getMediaFiles(){
  
  let folder = createProjectFolder()
  
  let files = folder.getFiles()
  
  // Clean all exist files
  while(files.hasNext()){
    files.next().setTrashed(true)
  }

  let rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
  
  let headers = rows[0]
  // Skip header row and process all records rows
  for(row = 1; row < rows.length; row++){
    
    let fields = rows[row]
    let id = fields[0]
   
    let metadata = generateMetadata(headers, fields)
    
    for(column = 0; column < fields.length; column++){
      let value = fields[column]
      let blob = null
      if(isImage(value)){
        let blob = getImage(value)
        let driveUrl = createFile(folder, blob, value, metadata)
        SpreadsheetApp.getActiveSheet().getRange(row + 1, column + 1).setFormula(`=HYPERLINK("${driveUrl}", "${value}")`)
      }else if(isAudio(value)){
        let blob = getAudio(value)
        let driveUrl = createFile(folder, blob, value, metadata)
        SpreadsheetApp.getActiveSheet().getRange(row + 1, column + 1).setFormula(`=HYPERLINK("${driveUrl}", "${value}")`)
      }
    }
  }
}

function createProjectFolder(){
  var prop = PropertiesService.getScriptProperties()
  var survey_name = prop.getProperty("Survey Name") 
 
  return createFolder(createFolder(null, "Epicollect Images"), survey_name)
}

function createFolder(parent, name){
  
  if(parent === null){
    parent = DriveApp.getRootFolder()
  }
  
  let folders = parent.getFoldersByName(name)
  if(folders.hasNext()){
    return folders.next()
  }else{
    return parent.createFolder(name)
  }
}

function isImage(data){
  if(typeof data === "string" || data instanceof String){
    return data.endsWith(".jpg") || data.endsWith(".jpeg") || data.endsWith(".png")
  }else{
    return false
  }
}

function isAudio(data){
  if(typeof data === "string" || data instanceof String){
    return data.endsWith(".mp4") || data.endsWith(".wav")
  }else{
    return false
  }
}

function getAudio(name){
  Logger.log(`getAudio(${name})`)
  var prop = PropertiesService.getScriptProperties()
  var survey_name = prop.getProperty("Survey Name")
  let url = name
  if(!url.startsWith("https")){
    url = "https://five.epicollect.net/api/export/media/" + survey_name + "?type=audio&format=audio&name=" + name
  }
  return getMedia(url, name)
}

function getImage(name){
  Logger.log(`getImage(${name})`)
  var prop = PropertiesService.getScriptProperties()
  var survey_name = prop.getProperty("Survey Name")
  let url = name
  if(!url.startsWith("https")){
    url = "https://five.epicollect.net/api/export/media/" + survey_name + "?type=photo&format=entry_original&name=" + name
  }
  return getMedia(url, name)
}

function getMedia(url, name){
  Logger.log(`getMedia(${url}, ${name})`)
  
  let prop = PropertiesService.getScriptProperties();
  let isPublicData = prop.getProperty("Is Public Data") === "true"
  let client_id = prop.getProperty("Client Id")
  let client_secret = prop.getProperty("Client Secret")
  
  let options = {}
  if(!isPublicData){
    var token = getToken(client_id, client_secret);
    if(token == null)
      return;
    options = {
      'method' : 'GET',
      'headers' : {
        'Authorization' : 'Bearer ' + JSON.parse(token).access_token
      }
    }
  }
  
  let resp = UrlFetchApp.fetch(url, options)
  let contentType = resp.getHeaders()["Content-Type"]
  
  return Utilities.newBlob(resp.getContent())
}

function createFile(folder, blob, fileName, metadata){
  let file = folder.createFile(blob.setName(fileName));
  file.setDescription(JSON.stringify(metadata, null, 4))
  return file.getUrl()
}

function generateMetadata(headers, fields){
  return Object.assign({}, ...fields.map((data, idx) => ({[headers[idx]]: data})))
}
