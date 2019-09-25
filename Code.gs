function QBO_Entry_CY() {
  // Oauth2 token access
  var a = SpreadsheetApp;
  var aS = a.getActiveSpreadsheet().getSheetByName("Refresh_Key");   
  var refresh_k = aS.getRange(1, 1).getValue();   
  var token_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer";
  var token_options = 
      {
        "headers": {
          "Content-Type": "application/x-www-form-urlencoded",
          "Accept": "application/json",
          "Authorization": "Basic AUTH KEY"
        },
        "payload": {
          "grant_type" : "refresh_token",
          "refresh_token" : refresh_k,
        }
      };
  var token_response = UrlFetchApp.fetch(token_url,token_options); 
  var access_token = JSON.parse(token_response);
  var token_key = access_token.access_token;
  var refresh_key = access_token.refresh_token;
  var a = SpreadsheetApp;
  var aS = a.getActiveSpreadsheet().getSheetByName("Refresh_Key");
  var refreshKeyStorage = aS.getRange(1, 1).setValue(refresh_key);  
  
  
  
  
  Logger.log(token_key);
  
  
  
  
  var url = "https://quickbooks.api.intuit.com/v3/company/COMPANYNUMBER/journalentry";
  var Auth_Token = "bearer " + token_key;
  
  
  var t = a.getActiveSpreadsheet().getSheetByName("QBO Entry");  
  
  var t_rows = t.getLastRow();
  var t_cols = t.getLastColumn();
  var t_headers = t.getRange(1,1,1,t_cols).getValues();
  t_headers = t_headers[0];
  Logger.log(t_headers);
  var date_loc = t_headers.indexOf("Date")+1
  var pt_loc = t_headers.indexOf("PostingType")+1
  var gln_loc = t_headers.indexOf("GL name")+1
  var glv_loc = t_headers.indexOf("GL value")+1
  var amount_loc = t_headers.indexOf("AMOUNT")+1
  Logger.log(date_loc);
  
  
  
  
  var a = SpreadsheetApp;
  var as = a.getActiveSpreadsheet().getSheetByName("Raw Data");   
  
  var as_rows = as.getLastRow();
  
  var as_cols = as.getLastColumn();
  
  var as_headers = as.getRange(1,1,1,as_cols).getValues();
  as_headers = as_headers[0];
  
  
  
  
  
  var as_entered_loc = as_headers.indexOf("Entered") +1;
  
  
  
  var as_orderNo_loc = as_headers.indexOf("Order Number") +1;
  
  var as_orders = as.getRange(1,as_orderNo_loc,as_rows,1).getValues()
  
  // flatten embedded array
  as_orders = flatten(as_orders);
  
  // filter out empty cells
  as_orders = as_orders.filter(Boolean);
  Logger.log(as_orders);
  
  
  var f = SpreadsheetApp;
  var f = f.getActiveSpreadsheet().getSheetByName("Filtered");   
  
  var f_rows = f.getLastRow();
  
  var f_cols = f.getLastColumn();
  
  var f_headers = f.getRange(1,1,1,f_cols).getValues();
  f_headers = f_headers[0]; 
  
  var f_orderNo_loc = f_headers.indexOf("Order Number") +1;
  
  
  var f_orders = f.getRange(1,f_orderNo_loc,f_rows,1).getValues()
  
  // flatten embedded array
  f_orders = flatten(f_orders);
  
  // filter out empty cells
  f_orders = f_orders.filter(Boolean);
  f_orders = f_orders.reverse();
  Logger.log(f_orders);
  
   
  
  var j = 0;
  if(j < f_orders.length-1){
    while (j <f_orders.length-1){
      var Txndate = t.getRange(2, date_loc).getValue();
      var i = 0;
      // start of while Loop
      // while (rowNum <= Alast){
      
      // Changing Variables
      
      var Line_array = []
      
      if(i < t_rows-1){
        while ( i < t_rows-1){
          var pt = t.getRange(2+i, pt_loc).getValue();
          var gln = t.getRange(2+i, gln_loc).getValue();
          var glv = t.getRange(2+i, glv_loc).getValue();
          
          var Amount = t.getRange(2+i, amount_loc).getValue();  
          var dict =             
              {
                "Id": String(i),
                "Amount": Amount,
                "Description":"",
                "DetailType": "JournalEntryLineDetail",
                "JournalEntryLineDetail": {
                  "PostingType": pt,
                  "AccountRef": {
                    "value": String(glv),
                    "name": gln
                  }
                }
              }  
          
          Line_array.push(dict);
          i++;
        }
        
      }
      Logger.log(Line_array);
      
      
      var payload =
          {
            
            "TxnDate": Txndate,
            "Line": Line_array,
            "TxnTaxDetail": {} 
          };
      
      Logger.log(JSON.stringify(payload));
      
      var options = 
          { 
            "method" : "POST",
            "headers": {
              "Content-Type": "application/json",
              "Accept": "application/json",
              "Authorization": Auth_Token,
              "muteHttpExceptions": true
            },
            "payload": JSON.stringify(payload),
          }; 
      var response = UrlFetchApp.fetch(url,options);
      Logger.log(response); 
      
      // set raw data as entered
      as.getRange(as_orders.indexOf(f_orders[j])+1,as_entered_loc).setValue("x")
      j = j + 1
      Utilities.sleep(3000)
    }
  }
}

