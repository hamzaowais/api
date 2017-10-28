function myFunction() {

  var client_id='';
  var secret_id='';
  
  try{
    var authorizationObj = getAuthorizationToken(client_id,secret_id);    
    if(authorizationObj.error==true){
    //Send Error to admin 
    throw(new Error(authorizationObj.message));
    }
    var authorizationToken=authorizationObj.access_token;
    unpaidInvoiceObj=getUnpaidInvoice(authorizationToken);
    if(unpaidInvoiceObj.error){
      throw(new Error(unpaidInvoiceObj.message));
    }
    


    if(unpaidInvoiceObj.invoices.length>0){
      var htmlBody = htmlifyObject(unpaidInvoiceObj.invoices);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var range = sheet.getRange(5, 1, htmlBody.length, 5);
    var values = range.setValues(htmlBody);
    }
     
    
    
    
    
    
    
  }catch(e){
    Logger.log(e.message);  
    MailApp.sendEmail("hamza.agreed@gmail.com","Contact Admin- Automated invoicing Api not working",e.message);
  }
}





function getUnpaidInvoice(authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
   params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true, 
     payload : JSON.stringify({
       status: ["UNPAID","SENT"],
       page:0,
       page_size:50
     })
   }
  tokenEndpoint='https://api.paypal.com/v1/invoicing/search';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
   var unpaidInvoiceResponse={};
  if (responseCode === 200) {
    var responseJson = JSON.parse(responseBody);
    unpaidInvoiceResponse.error=false;
    unpaidInvoiceResponse.invoices=responseJson.invoices;
    return unpaidInvoiceResponse;

    } else {
      unpaidInvoiceResponse.error=true;
      unpaidInvoiceResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
      return unpaidInvoiceResponse;
    }   

}


function htmlifyObject(invoices){
  

  var outputValues=[];
  invoices.forEach(function(invoice){
    var row=[];

    // add invoice number

    if(invoice.id){
          row.push(invoice.id);
    }else{
          row.push('N/A');     
    }
    //add merchant email
    if(invoice.billing_info&&invoice.billing_info.length>0&&invoice.billing_info[0].email){
      row.push(invoice.billing_info[0].email);
    }else{
          row.push('N/A'); 
    }
    //add invoice DAte
    if(invoice.invoice_date){
      row.push(invoice.invoice_date);
    }else{
          row.push('N/A'); 
    }
    //add total amount
    if(invoice.total_amount&&invoice.total_amount.value){
      row.push(invoice.total_amount.value);      
    }else{
          row.push('N/A'); 
    }
    // add status
    if(invoice.status){
      row.push(invoice.status);   
    }else{
          row.push('N/A'); 
    }
    outputValues.push(row);
  });  
  return outputValues;
}




function getAuthorizationToken(client_id,secret_id){
  var tokenEndpoint = "https://api.paypal.com/v1/oauth2/token";
    var head = {
      'Authorization':"Basic "+ Utilities.base64Encode(client_id+':'+secret_id),
      'Accept': 'application/json',
      'Content-Type': 'application/x-www-form-urlencoded'
    }
    var postPayload = {
        "grant_type" : "client_credentials"
    }
    var params = {
        headers:  head,
        contentType: 'application/x-www-form-urlencoded',
        method : "post",
        muteHttpExceptions: true,
        payload : postPayload      
    }
    var request = UrlFetchApp.getRequest(tokenEndpoint, params); 
    var response = UrlFetchApp.fetch(tokenEndpoint, params); 
    var responseCode = response.getResponseCode()
    var responseBody = response.getContentText()

    if (responseCode === 200) {
      var tokenResponse={};
      var responseJson = JSON.parse(responseBody);
      if(responseJson&&responseJson.error){
        tokenResponse.error=true;
        tokenResponse.message=responseJson.error;
        return tokenResponse;
      }
      if(responseJson.access_token){
        tokenResponse.error=false;
        tokenResponse.access_token=responseJson.access_token;
        return tokenResponse;
      }
     tokenResponse.error=true;
     tokenResponse.message='Access Token not found';
     return tokenResponse;
    } else {
        var tokenResponse={};
        tokenResponse.error=true;
        tokenResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
        return tokenResponse;
      }
}

