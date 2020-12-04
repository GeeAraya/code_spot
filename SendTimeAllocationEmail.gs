function SendRGORemainingTimeAEmail() {
//access the specific tab of the active google sheet
var ss = SpreadsheetApp.getActiveSpreadsheet()  //access current active google sheet
var sheet1=ss.getSheetByName('RGO-Invoice_Data'); //assign a variable to the sheet tab of interest

// Setup iteration through list of rows in a column to send emails to
for (i = 17; i <= 27;i++){
    //prepare variables
    var TeamMemberName = sheet1.getRange(i,4).getValue(); //access the value at (row,column) and assign to variable
    //var emailAddress = sheet1.getRange(i,5).getValue(); //access the value at (row,column) and assign to variable
    var TotalHours = sheet1.getRange(i,6).getValue(); //access the value at (row,column) and assign to variable
    var HoursRemaining = sheet1.getRange(i,20).getValue(); // access and assign value at (row,colunn) to a variable
    var HoursUsed = sheet1.getRange(i,19).getValue(); // access and assign value at (row,colunn) to a variable
    var cell = sheet1.getCell(i, 10);
    if cell.isBlank() { 
      var emailAddress = sheet1.getRange(i,5).getValue(); //access the value at (row,column) and assign to variable
      else var emailAddress = ghermay.araya@gmail.com;
  }
    // setup email message variables
    var message = 'PMO REPORT: FEMA RGO Contract Utilization Report As of Last Payroll:\n -The Total Hours Allocated: ' + TotalHours +'\n - Hours Used: ' + HoursUsed +'\n - Hours Remaining: ' + HoursRemaining +'\n Please Contact the project manager or task manager for this project if you are running at two weeks or below on hours or pmo@nltgis.com for any inquries: ';
    var subject = 'LOOP TEST 2: PMO REPORT- FEMA RGO Time Allocation Update: ' + TeamMemberName;
    MailApp.sendEmail(emailAddress, subject, message);
    //end the loop end here  
}
}
