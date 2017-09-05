

function sendInvoice() {
  Logger.log('Starting the Script to check Email email for invoices and sending them');
  //Iterate through all emails 
  
//  var client_id='';
//  var secret_id='';
  
  var client_id='';
  var secret_id='';
  
  var threads = GmailApp.getInboxThreads();
  var invoiceThreads=[];
  var unpaidInvoiceThreads =[];
  var paidInvoiceThreads=[];
  var flagInvoiceEmail=0;
 
  
  for (var j = 0; j < threads.length; j++) {
    var tempMes=threads[j].getMessages()
    if(threads[j].isUnread() && tempMes.length==1){
      var subject = threads[j].getFirstMessageSubject();
      subject = subject.trim();
      subject=subject.toLowerCase()
      if(subject=='send paypal invoice'){
        invoiceThreads.push(threads[j]);
        flagInvoiceEmail=1;
      } 
      if(subject=='unpaid invoice'){
        unpaidInvoiceThreads.push(threads[j]);
        flagInvoiceEmail=1;
      } 
      if(subject=='paid invoice'){
        paidInvoiceThreads.push(threads[j]);
        flagInvoiceEmail=1;
      } 
    }
 }
  
  
  var authorizationToken;
  if(flagInvoiceEmail==1){
    var authorizationObj = getAuthorizationToken(client_id,secret_id);
    if(authorizationObj.error==true){
      //Send Error to admin 
      Logger.log(authorizationObj.message);
      MailApp.sendEmail("hamza.agreed@gmail.com","Contact Admin- Automated invoicing Api not working",authorizationObj.message);
      return;
    }
    authorizationToken=authorizationObj.access_token;
  }
  
  
  
  
  if(invoiceThreads.length>0){
    // get authorizationToken
    
    invoiceThreads.forEach(function(invoiceThread){
      var emailmessages= invoiceThread.getMessages();
      var emailmessage=emailmessages[0].getPlainBody();
      var invoiceData=getInvoiceDetailsFromMail(emailmessage);
      Logger.log(JSON.stringify(invoiceData));
      //check and validate invoiceData:= check valid email and the invoice amount
      var draftInvoiceObject=createInvoiceDraft(invoiceData,authorizationToken);
      if(draftInvoiceObject.error){
        Logger.log(draftInvoiceObject.message);
        MailApp.sendEmail("hamza.agreed@gmail.com","Contact Admin- Automated invoicing Api not working",authorizationObj.message);
        invoiceThread.reply('Failure<br>'+ draftInvoiceObject.message);
        Logger.log(draftInvoiceObject.message);
        return;
      }
      var invoiceConfirmation = sendDraftInvoice(draftInvoiceObject.id,authorizationToken);
      if(invoiceConfirmation.error){
        Logger.log(invoiceConfirmation.message);
        MailApp.sendEmail("hamza.agreed@gmail.com","Contact Admin- Automated invoicing Api not working",invoiceConfirmation.message);
        invoiceThread.reply('Failure<br>'+ invoiceConfirmation.message);
        Logger.log(invoiceConfirmation.message);
        return;
      }
      invoiceThread.reply('Invoice Sent<br>');      
    });
  }

  if(unpaidInvoiceThreads.length>0){
    //get list of all unpaid invoice 
    unpaidInvoiceObj=getUnpaidInvoice(authorizationToken);
    if(unpaidInvoiceObj.error){
      Logger.log(unpaidInvoiceObj.message);
      MailApp.sendEmail("hamza.agreed@gmail.com"," Contact Admin",unpaidInvoiceObj.message);

      Logger.log(unpaidInvoiceObj.message);
      return;
    }
    
    var htmlBody = 'Congrats No Unpaid Invoice exist'
    Logger.log(JSON.stringify(unpaidInvoiceObj));
    if(unpaidInvoiceObj.invoices.length>0){
      htmlBody = htmlifyObject(unpaidInvoiceObj.invoices);
    }
    unpaidInvoiceThreads.forEach(function(unpaidInvoiceThread){
      unpaidInvoiceThread.reply("Unpaid Invoices", {
        htmlBody: htmlBody  
      });
    });
  }

  if(paidInvoiceThreads.length>0){
    //get list of all paid invoice 
    paidInvoiceObj=getPaidInvoice(authorizationToken);
    if(paidInvoiceObj.error){
      Logger.log(paidInvoiceObj.message);
      MailApp.sendEmail("hamza.agreed@gmail.com"," Contact Admin",paidInvoiceObj.message);

      Logger.log(paidInvoiceObj.message);
      return;
    }
    
    var htmlBody = 'Crap!!  No paid Invoice exist'
    Logger.log(JSON.stringify(paidInvoiceObj));
    if(paidInvoiceObj.invoices.length>0){
      htmlBody = htmlifyObject(paidInvoiceObj.invoices);
    }
    paidInvoiceThreads.forEach(function(paidInvoiceThread){
      paidInvoiceThread.reply("paid Invoices", {
        htmlBody: htmlBody  
      });
    });
  }


}


function htmlifyObject(invoices){
  
  var htmlTable='<table><tr><th>Invoice No</th><th>Merchant Email</th><th>Invoice Date</th><th>Total Amount</th><th>Status</th></tr>';
  invoices.forEach(function(invoice){
    // add invoice number
    htmlTable=htmlTable+'<tr>';
    if(invoice.id){
      htmlTable=htmlTable+'<td>'+invoice.number+'</td>';
    }else{
      htmlTable=htmlTable+'<td>N/A</td>';
    }
    //add merchant email
    if(invoice.billing_info&&invoice.billing_info.email){
      htmlTable=htmlTable+'<td>'+invoice.billing_info.email+'</td>';
    }else{
      htmlTable=htmlTable+'<td>N/A</td>';
    }
    //add invoice DAte
    if(invoice.invoice_date){
      htmlTable=htmlTable+'<td>'+invoice.invoice_date+'</td>';
    }else{
      htmlTable=htmlTable+'<td>N/A</td>';
    }
    //add total amount
    if(invoice.total_amount&&invoice.total_amount.value){
      htmlTable=htmlTable+'<td>'+invoice.total_amount.value+'</td>';
    }else{
      htmlTable=htmlTable+'<td>N/A</td>';
    }
    // add status
    if(invoice.status){
      htmlTable=htmlTable+'<td>'+invoice.status+'</td>';
    }else{
      htmlTable=htmlTable+'<td>N/A</td>';
    }
    
    htmlTable=htmlTable+'</tr>';
  });
  htmlTable=htmlTable+'</table>'
  
  return htmlTable;
}

function getPaidInvoice(authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
   params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true, 
     payload : JSON.stringify({
       status: ["PAID"],
       page:0,
       page_size:50
     })
   }
  tokenEndpoint='https://api.paypal.com/v1/invoicing/search';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
   var paidInvoiceResponse={};
  if (responseCode === 200) {
    var responseJson = JSON.parse(responseBody);
    paidInvoiceResponse.error=false;
    paidInvoiceResponse.invoices=responseJson.invoices;
    return paidInvoiceResponse;

    } else {
      paidInvoiceResponse.error=true;
      paidInvoiceResponse.message=Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody);
      return paidInvoiceResponse;
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
function sendDraftInvoice(invoiceId,authorizationToken){
  head = {
    'Authorization':"Bearer "+ authorizationToken,
    'Content-Type': 'application/json'
  }
  
   params = {
    headers:  head,
    method : "post",
    muteHttpExceptions: true
  }
  tokenEndpoint='https://api.paypal.com/v1/invoicing/invoices/'+invoiceId+'/send';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
   var invoiceResponse={};
  if (responseCode === 202) {
    
    invoiceResponse.error=false;
    return invoiceResponse;

    } else {
      invoiceResponse.error=true;
      invoiceResponse.message=Utilities.formatString("Request failed. Expected 202, got %d: %s", responseCode, responseBody);
      return invoiceResponse;
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
  tokenEndpoint='https://api.paypal.com/v1/invoicing/invoices';
  request = UrlFetchApp.getRequest(tokenEndpoint, params); 
  response = UrlFetchApp.fetch(tokenEndpoint, params); 
  
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
    Logger.log(Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody))
  var invoiceResponse={};
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
              data.merchant_info.email=data.merchant_info.email.trim();
              
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
}