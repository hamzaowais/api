function sendInvoice() {
  Logger.log('Starting the Script to check Email email for invoices and sending them');
  //Iterate through all emails 
  var threads = GmailApp.getInboxThreads();
  var flagInvoiceEmail=0;
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
  
 
    var merchant_info= {
  "email": "hamza.agreed@gmail.com",
  "first_name": "David",
  "last_name": "Larusso",
  "business_name": "Mitchell & Murray",
  "phone": {
    "country_code": "001",
    "national_number": "4085551234"
  }
    }
      
    var items= [
  {
    name: "Zoom System wireless headphones",
    quantity: 2,
    unit_price: {
    currency: "USD",
    value: 120
    },
    tax: {
    name: "Tax",
    percent: 0
    }
  }];
    
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
  for (var j = 0; j < threads.length; j++) {
    if(threads[j].isUnread()){
      var subject = threads[j].getFirstMessageSubject();
      subject = subject.trim();
      subject=subject.toLowerCase()
      if(subject=='send paypal invoice now!'){
        var emailmessages= threads[j].getMessages();
        var emailmessage=emailmessages[0].getPlainBody();
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
        Logger.log(JSON.stringify(data)); 
      } 
    }
 }
}