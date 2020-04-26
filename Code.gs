function cell_op_D5_p1(){ cell_op("D5",    1)} // K5 1 -1 0

function cells_get(mycell) 
{
  var ss    = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Sheet1');
  var val   = sheet.getRange(mycell).getValues();
  
  return val
}

function cell_get(mycell) 
{
  var ss    = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Sheet1');
  var val   = sheet.getRange(mycell).getValue();
  
  return val
}

function log_set(type, txt, val)
{
  var ss    = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Log_'+type);
  var d = new Date()
  sheet.appendRow([d.toISOString(), txt, val]);
}

function cell_op(mycell, op) 
{
  
  var ss    = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Sheet1');
  var val   = sheet.getRange(mycell).getValue();
  
  if (val == null)
  {
    val = 0
  }
  
  if (op == 0)
  {
    sheet.getRange(mycell).setValue(0)
  }
  else
  {
    sheet.getRange(mycell).setValue(val+op);
  }
}

function count_effective(what) 
{
  
  var ss          = SpreadsheetApp.getActiveSpreadsheet()
  var sheet       = ss.getSheetByName('Log_effectif');
  var rangeData   = sheet.getDataRange();
  var lastRow     = rangeData.getLastRow();
  var searchRange = sheet.getRange(1,1, lastRow, 3);
  var rangeValues = searchRange.getValues();
  
  var n = 0;
  
    for ( j = 0 ; j < lastRow; j++)
      if (rangeValues[j][1] === what)
        if (rangeValues[j][2] > 0)
        {
          var c = new Date()
          var d = new Date(Date.parse(rangeValues[j][0]))
          var e = new Date()
     
          var cc = c.setUTCHours(0,0,0,0);
          var ee = e.setUTCHours(23,59,59,0);
          var dd = d.getTime()
          
          if ((cc < dd) & (dd < ee))
          {
            n += rangeValues[j][2];
          }
        }
    
  return n

};

function doGet(e) {
  var queryString = e.queryString;

  var name = getQueryStringValue(queryString, "name")

  var htmlTemplate = HtmlService.createTemplateFromFile('index');

  htmlTemplate.qsName = name; //setting a variable in html template using the query string value

  var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return htmlOutput;
}

// Utility function to fetch key values from query string
function getQueryStringValue(query, key){
  var queryParts = query.split("&");
  if(queryParts && queryParts.length > 0){
    for(var i=0; i<queryParts.length; i++){
      var k = queryParts[i].split("=")[0];
      if(k == key) return queryParts[i].split("=")[1];
    }
  }
}
