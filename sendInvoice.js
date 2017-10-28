function sendPaypalInvoice(e) {

 var  rowStart=e.range.rowStart;
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];
 var range = sheet.getRange(rowStart, 1, 1, 16);
 var values = range.getValues();
 var invoicesent=0;
  var errorMessage='No Error Found';
  var client_id='';
  var secret_id='';

  try{
    var authorizationToken;
    var authorizationObj = getAuthorizationToken(client_id,secret_id);
    if(authorizationObj.error==true){
      errorMessage=authorizationObj.message;
      throw(new Error(errorMessage));
    }
    authorizationToken=authorizationObj.access_token;

  
  
  var namedValues=e.namedValues;
  var invoiceData=getInvoiceDetailsFromMail(namedValues);
      if(invoiceData.error){
        throw(new Error(invoiceData.message));
      }
      //check and validate invoiceData:= check valid email and the invoice amount
    
      var draftInvoiceObject=createInvoiceDraft(invoiceData,authorizationToken);
      if(draftInvoiceObject.error){
        
        throw(new Error(JSON.stringify(draftInvoiceObject)));
        
      }
      var invoiceConfirmation = sendDraftInvoice(draftInvoiceObject.id,authorizationToken);
      if(invoiceConfirmation.error){
   
        throw(new Error(JSON.stringify(invoiceConfirmation)));

      }
    invoicesent=1;
    values[0][15]="Invoice Sent";
  }catch(err){
    values[0][15]=err.message;
    
      Logger.log(err.message);
      MailApp.sendEmail("","Contact Admin- Automated invoicing Api not working",err.message);
    
    }
  
  range.setValues(values);
  
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

function getInvoiceDetailsFromMail(e){
 

  var i =0;
  var data={};
  data.logo_url="https://pics.paypal.com/00/s/OTdYMjA5WFBORw/p/MDBjZTNkOTUtYTU1MC00ZGY2LWE4YmQtYjUzNzcwZDE0OTQw/image_109.PNG";
  var itemsFound=0;

  
  data.cc_info=[{email:'hamza.agreed@gmail.com'},
                {email:'hamza.agreed@gmail.com'},
                {email:'hamza.agreed@gmail.com'}
             ];
  data.merchant_info={};
  data.merchant_info.email="agreedtech.og@gmail.com";
  data.merchant_info.address= {
    "line1": "2515, westbankcroft street",
    "city": "New Delhi",
    "state": "NCT-Delhi",
    "postal_code": "110025",
    "country_code": "IN"
  }
  data.merchant_info.business_name="Agreed Technologies Private Limited";
  data.merchant_info.website="www.agreedtechnologies.com" ;
 
  var tempBillingInfo={};
  
  tempBillingInfo.email=e['Billing Email'][0].trim();
  if(e['First Name'][0].trim().length>0){
    tempBillingInfo.first_name=e['First Name'][0].trim();
  }
  if(e['Last Name'][0].trim().length>0){
    tempBillingInfo.first_name=e['Last Name'][0].trim();
  }
 data.items=[];
  for(i=1;i<=3;i++){
    var flag=1;
    var temp={};
    if(e['Item '+i+' Description'][0].trim().length>0){
      temp.name=e['Item '+i+' Description'][0].trim();
    }else{
      flag=0;
    }
    if(e['Item '+i+' Value'][0].trim().length>0){
      temp.quantity=e['Item '+i+' Value'][0].trim();
    }else{
      flag=0;
    }
    
    if(e['Item '+i+' Quantity'][0].trim().length>0){
      temp.unit_price={
        "currency": "USD",
        "value": e['Item '+i+' Quantity'][0].trim()
      }
      itemsFound=itemsFound+1;
    }else{
      flag=0;
    }
    if(flag==1){
      data.items.push(temp);
    }
    
    
    
  }
  data['merchant_memo']=e['Type of Sale'][0].trim()+'-'+e['Names'][0].trim();
    

  var error='';
  var errorFlag=0;
  if(tempBillingInfo&&tempBillingInfo.email){
    data.billing_info=[];
    data.billing_info.push(tempBillingInfo);
    
  }else{
    errorFlag=1;
    var tempData={};
    tempData.error=true;
    tempData.message='Billing Email missing';
    return tempData;
  }
  
  if(itemsFound==0){
    var tempData={};
    tempData.error=true;
    tempData.message='Items not found';
    return tempData;
  }
 
  return data;
}
