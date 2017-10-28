# API(Automated Paypal Invoicing) using Google App script  
  
## Getting Started:

The code in this repository will give you an idea on how we can use google app script to automate task. It will also give on how to use paypal resp api. How to make outh2.0 authentication in google app script. And on how to make paypal rest api call using google app script.  
You need get the client id and secret using the following [link](https://developer.paypal.com/developer/applications). You may also use the [sandbox](https://developer.paypal.com/developer/accounts/) account for testing purpose.  

**sendInvoice.gs:** This is triggered on form submit event. for more details on this [event](https://developers.google.com/apps-script/guides/triggers/events#form-submit), click the link. From the information that it gets from the google form, it uses the paypal rest api  to get the [access token)(https://developer.paypal.com/docs/api/overview/#authentication-and-authorization), [create invoice draft](https://developer.paypal.com/docs/api/invoicing/#invoices_create), and then [send the invoice](https://developer.paypal.com/docs/api/invoicing/#invoices_send).  
**generatemonthlyRevenueSpreadsheet.gs:** This is triggered on open event. for more details on this [event](https://developers.google.com/apps-script/guides/triggers/events#open), click the link. It uses the paypal rest api  to get the [access token)(https://developer.paypal.com/docs/api/overview/#authentication-and-authorization),  and get all the paid invoice from [paid invoice api](https://developer.paypal.com/docs/api/invoicing/#invoices_search)(This information is gathered using several api calls instead just one), and then using this information, we calculate the monthwise revenue.
 We just append this data to a sheet using [
Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet/).  
**generatePaidInvoiceSpreadsheet.gs:** This is triggered on open event. for more details on this [event](https://developers.google.com/apps-script/guides/triggers/events#open), click the link. It uses the paypal rest api  to get the [access token)(https://developer.paypal.com/docs/api/overview/#authentication-and-authorization),  and get last 50 paid invoice from [search invoice api](https://developer.paypal.com/docs/api/invoicing/#invoices_search) and then we just append this data to a sheet using [
Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet/).  
**generateUnPaidInvoiceSpreadsheet.gs:** This is triggered on open event. for more details on this [event](https://developers.google.com/apps-script/guides/triggers/events#open), click the link. It uses the paypal rest api  to get the [access token](https://developer.paypal.com/docs/api/overview/#authentication-and-authorization),  and get last 50 unpaid invoice from [search invoice api](https://developer.paypal.com/docs/api/invoicing/#invoices_search) and then we just append this data to a sheet using [
Spreadsheet Service](https://developers.google.com/apps-script/reference/spreadsheet/).


## Description: 
Using Google Apps Script we will automate some of the task like creating Paypal invoices, Email a list of unpaid invoices, and the list of paid Invoices. Not only does this automates few of the tasks but also able to make a third party paypal application without much code. We will do basic data analytics, like the monthly revenue. 
* Substitute of third party(android, iOS, web) Paypal application
* Ability to make reports/spreadsheets. Can apply google's Access Control. eg. grant permission only to selected people.
* Basic Data Analytics eg. Monthly Sales Report From Papal
* Report of Unpaid Paypal Invoices
* Report of Paid Paypal Invoices
* Using Google Form to Create Paypal invoice.
* Send notification email timely triggered. 
* Much more, Could be tailored according to User's specification.  

### Monthly Sales Report  
![](https://github.com/hamzaowais/api/blob/master/img/monthlySales.png?raw=true)

### Report of Unpaid Paypal Invoices  
![](https://github.com/hamzaowais/api/blob/master/img/unpaidInvoice.png?raw=true)

### Report of Paid Paypal Invoices  
![](https://github.com/hamzaowais/api/blob/master/img/paidInvoice.png?raw=true)

### Google Form to Create Paypal invoice  
![](https://github.com/hamzaowais/api/blob/master/img/googleForm1.png?raw=true)
![](https://github.com/hamzaowais/api/blob/master/img/googleFormBillingSalesInfo.png?raw=true)
![](https://github.com/hamzaowais/api/blob/master/img/googleInfoItem.png?raw=true)
 
### Example of an invoice sent to hamza.agreed@gmail.com worth of 2 dollars.
![](https://github.com/hamzaowais/api/blob/master/img/form1.png?raw=true)
![](https://github.com/hamzaowais/api/blob/master/img/form2.png?raw=true)
![](https://github.com/hamzaowais/api/blob/master/img/form2.png?raw=true)
### Confirmation of the Email Sent
![](https://github.com/hamzaowais/api/blob/master/img/sentInvoice.png?raw=true)


