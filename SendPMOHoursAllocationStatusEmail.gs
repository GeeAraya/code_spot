function SendPMOHoursAllocationStausEmail() {
//access data for report email generation
//code can be adopted to any list that has data and an assocoated email address field in google sheets
//this code is specifically written to work with the PMO Dashboard designed for RGO by Rob Pitts and maintained by Celena Scott
//code developed by Ghermay Araya 12/4/2020
//Code is copywrited and for private use of NLT.  Do not distribute.
  
var ss = SpreadsheetApp.getActiveSpreadsheet()  //access current active google sheet
var sheet1=ss.getSheetByName('RGO-Invoice_Data'); //assign a variable to the sheet tab of interest

// Setup iteration through list of rows in a column to send emails to
for (i = 17; i <= 29;i++){  //keeping it simple here by starting the iteration at the row number we want to start and ending at the count - 1 to get to the last row
    //prepare variables
    var count = i;
    var TeamMemberName = sheet1.getRange(i,4).getValue(); //access the value at (row,column) and assign to variable
    //var emailAddress = sheet1.getRange(i,5).getValue(); //access the value at (row,column) and assign to variable
    var TotalHours = sheet1.getRange(i,6).getValue(); //access the value at (row,column) and assign to variable
    var HoursRemaining = sheet1.getRange(i,20).getValue(); // access and assign value at (row,colunn) to a variable
    var HoursUsed = sheet1.getRange(i,19).getValue(); // access and assign value at (row,colunn) to a variable
    var emailAddress = sheet1.getRange(i,5).getValue(); //access the value at (row,column) and assign to variable, CANNOT have empty email address in field
    
    // setup email message variables
    var subject = 'PMO REPORT- FEMA RGO Time Allocation Update: ' + TeamMemberName;    
    var message = 'PMO REPORT: FEMA RGO Contract Utilization As of Last Payroll:\n -The Total Hours Allocated: ' + TotalHours +'\n - Hours Used: ' + HoursUsed +'\n - Hours Remaining: ' + HoursRemaining +'\n Please Contact the project manager or task manager for this project if you are running at two weeks or below on hours or pmo@nltgis.com for any inquries: ';
    MailApp.sendEmail(emailAddress, subject, message);
    //end the loop end here  
}
}
