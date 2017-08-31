

function sendInvoice() {
  Logger.log('Starting the Script to check Email email for invoices and sending them');
  //Iterate through all emails 
  
  var client_id='';
  var secret_id='';
  
  var threads = GmailApp.getInboxThreads();
  var invoiceThreads=[];
  var flagInvoiceEmail=0;
  
  for (var j = 0; j < threads.length; j++) {
    if(threads[j].isUnread()){
      var subject = threads[j].getFirstMessageSubject();
      subject = subject.trim();
      subject=subject.toLowerCase()
      if(subject=='send paypal invoice now!'){
        
        invoiceThreads.push(threads[j]);
        var emailmessages= threads[j].getMessages();
        var emailmessage=emailmessages[0].getPlainBody();
        var data=getInvoiceDetailsFromMail(emailmessage);
      } 
    }
 }
  
  if(invoiceThreads.length>0){
    // get authorizationToken
    var authorizationObj = getAuthorizationToken(client_id,secret_id);
    if(authorizationObj.error==true){
      
    }
    Logger.log(JSON.stringify(authorizationObj));
  }
  
 
   if(flagInvoiceEmail==1){
   var tokenEndpoint = "https://api.sandbox.paypal.com/v1/oauth2/token";
    var head = {
      'Authorization':"Basic "+ Utilities.base64Encode("AbvrSkKmXtvhRe9WFxumBOQsL-tkhZPXtLzYTEoZ-tu7UmkKwwlJd3QyIjpJEv6iolklSYNiVCEbP8gz:ENQKkopSBOPJx0UmKmc89fRN_2UDupiAhVnh36SD0bOZ8U-hOLXtiQH9QDKSO4bIZkcG7k2EDoi8DTud"),
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
        payload : postPayload
    }
    var request = UrlFetchApp.getRequest(tokenEndpoint, params); 
    var response = UrlFetchApp.fetch(tokenEndpoint, params); 
    var result = response.getContentText();
    Logger.log(result);
    var resultObject = JSON.parse(result);
    var authorizationToken=resultObject.access_token;
  
    
      head = {
      'Authorization':"Bearer "+ authorizationToken,
      'Content-Type': 'application/json'
    }
      var create_invoice_json = {
    "merchant_info": {
        "email": "hamza.agreed-facilitator@gmail.com",
        "first_name": "Dennis",
        "last_name": "Doctor",
        "business_name": "Medical Professionals, LLC",
        "phone": {
            "country_code": "001",
            "national_number": "5032141716"
        },
        "address": {
            "line1": "1234 Main St.",
            "city": "Portland",
            "state": "OR",
            "postal_code": "97217",
            "country_code": "US"
        }
    },
    
    "items": [{
        "name": "Sutures",
        "quantity": 100.0,
        "unit_price": {
            "currency": "USD",
            "value": 5
        }
    }]
};
  
      params = {
        headers:  head,
        method : "post",
        muteHttpExceptions: true,
        payload:JSON.stringify(create_invoice_json)
        
    }
      tokenEndpoint='https://api.sandbox.paypal.com/v1/invoicing/invoices';
           request = UrlFetchApp.getRequest(tokenEndpoint, params); 
     response = UrlFetchApp.fetch(tokenEndpoint, params); 
  var responseCode = response.getResponseCode()
var responseBody = response.getContentText()

if (responseCode === 200) {
  var responseJson = JSON.parse(responseBody)
  // ...
} else {
  Logger.log(Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody))
  // ...
}
  
  }    
  

  
}

function createInvoiceDraft(invoiceData, authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
  params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true,
    payload:JSON.stringify(invoiceData)    
  }
  tokenEndpoint='https://api.sandbox.paypal.com/v1/invoicing/invoices';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
  if (responseCode === 201) {
    var responseJson = JSON.parse(responseBody);
    invoiceResponse.error=false;
    invoiceResponse.id=responseJson.id;
    } else {
      invoiceResponse.error=true;
      invoiceResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
    }
  
  return invoiceResponse;
}

function getAuthorizationToken(client_id,secret_id){
  var tokenEndpoint = "https://api.sandbox.paypal.com/v1/oauth2/token";
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

function getInvoiceDetailsFromMail(emailmessage){
        emailmessage = emailmessage.split("\n");
        var i =0;
        var data={};
        while(i<emailmessage.length){
          
          var currentEmailLine=emailmessage[i];
          currentEmailLine=currentEmailLine.trim();
              currentEmailLine=currentEmailLine.split(':=');
          if(currentEmailLine.length==2){
            currentEmailLine[0]=currentEmailLine[0]
            if(currentEmailLine[0]=='merchant_email'){
              if(!data.merchant_info){
                data.merchant_info={};
              }
              data.merchant_info.email=currentEmailLine[1];
            }
            
            if(currentEmailLine[0]=='first_name'){
              if(!data.merchant_info){
                data.merchant_info={};
                
              }
              data.merchant_info.first_name=currentEmailLine[1];
            }
            if(currentEmailLine[0]=='last_name'){
              if(!data.merchant_info){
                data.merchant_info={};   
              }
              data.merchant_info.last_name=currentEmailLine[1];
            }
            if(currentEmailLine[0]=='business_name'){
              if(!data.merchant_info){
                data.merchant_info={};   
              }
              data.merchant_info.business_name=currentEmailLine[1];
            } 
          }
          if(currentEmailLine=='item'){
            if(!data.items){
              data.items=[];
            }
            var temp={};
            i = i+1;
            currentEmailLine=emailmessage[i].split(':=');
            if(currentEmailLine.length==2){
              if(currentEmailLine[0]=='name'){
                
                temp.name=currentEmailLine[1];
                temp.name=temp.name.trim();
              }
            }
            i = i+1;
            currentEmailLine=emailmessage[i].split(':=');
            if(currentEmailLine.length==2){
              if(currentEmailLine[0]=='quantity'){
                temp.quantity=currentEmailLine[1];
                temp.quantity=temp.quantity.trim();
              }
            }   
            i = i+1;
            currentEmailLine=emailmessage[i].split(':=');
            if(currentEmailLine.length==2){
              if(currentEmailLine[0]=='value'){
                temp.unit_price={
                  "currency": "USD",
                  "value": currentEmailLine[1].trim()
                }
              }
            }
            data.items.push(temp);
          }
          i=i+1;
        }
        flagInvoiceEmail=1;
       return data;
}}