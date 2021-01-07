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
  menu.addItem("Reset settings", "resetSettings");
  
  menu.addToUi();
}

function resetSettings(){
  var prop = PropertiesService.getScriptProperties();
  
  prop.setProperty("Client Id", "");
  prop.setProperty("Client Secret", "");
  prop.setProperty("Survey Name", "");
}

function getData(){
  var prop = PropertiesService.getScriptProperties()
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
    client_id = ui.prompt("Please enter your Client Id").getResponseText();
    client_secret = ui.prompt("Please enter your Client Secret").getResponseText();
    survey_name = ui.prompt("Please enter the name of the survey as it appears in the http:// bar \n(ex. ec5-demo-project not EC5 DEMO PROJECT) ").getResponseText()
    
    prop.setProperty("Client Id",client_id )
    prop.setProperty("Client Secret",client_secret )
    prop.setProperty("Survey Name",survey_name )
  }
    
  var token = getToken(client_id, client_secret);
  
  if(token == null)
    return;
  
  var options = {
    'method' : 'get',
    'headers' : {
      'Content-Type' : 'application/json',
      'Authorization' : 'Bearer ' + JSON.parse(token).access_token
    }
  };
  
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
