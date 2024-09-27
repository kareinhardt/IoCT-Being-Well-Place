function sendDailyReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const today = new Date();
  const currentDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy'); // Format today's date
  
  Logger.log("Today's Date: " + currentDate);  // Log the current date for comparison

  let firstCheckInTime = null;
  let lastCheckInTime = null;
  let checkInCount = 0;
  let emailContent = "";

  // Loop through the sheet to gather the check-ins
  for (let i = 1; i < data.length; i++) {
    const rawDateTimeString = data[i][3]; // Get the "Time & Date" value
    const buttonID = data[i][0];
    
    Logger.log("Raw DateTime from Sheet: " + rawDateTimeString); // Log the raw date and time from the sheet

    // Split the date and time
    let [datePart, timePart] = rawDateTimeString.split(" at ");

    // Add a space before the AM/PM if missing
    timePart = timePart.replace(/([APM]{2})$/, ' $1');
    
    // Combine date and time in a way that Date() can parse
    const formattedDateTimeString = datePart + " " + timePart;
    
    Logger.log("Formatted DateTime String: " + formattedDateTimeString);

    // Parse the "Time & Date" string into a Date object
    const dateTime = new Date(formattedDateTimeString);
    
    // Get the date part and time part
    const parsedDatePart = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'MMMM d, yyyy');
    const parsedTimePart = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'hh:mm a');

    Logger.log("Parsed Date: " + parsedDatePart);  // Log the parsed date
    Logger.log("Parsed Time: " + parsedTimePart);  // Log the parsed time

    // Check if the date matches today's date
    if (parsedDatePart === currentDate) {
      if (!firstCheckInTime) firstCheckInTime = parsedTimePart;  // First check-in time
      lastCheckInTime = parsedTimePart;  // Last check-in time
      checkInCount++;  // Increment the check-in count
      
      // Add to the email content
      emailContent += `${checkInCount}. Checked-in at ${parsedTimePart}, via ${buttonID}\n`;
    }
  }

  // Log or send the email content based on check-ins
 if (checkInCount > 0) {
    const subject = `Check-in at BWP on ${currentDate} - ${checkInCount} people checked in`;
    const body = `On ${currentDate}, a total of ${checkInCount} people visited the Being Well Place.\n\nThe first check-in was at ${firstCheckInTime}, and the last check-in was at ${lastCheckInTime}.\n\nHere is the detailed list:\n\n${emailContent}`;
    
    Logger.log("Email Subject: " + subject);
    Logger.log("Email Body: " + body);
    MailApp.sendEmail("X@newcastle.ac.uk, Y@newcastle.ac.uk", subject, body);
    //MailApp.sendEmail("X@gmail.com", "Test Subject", "This is a test email to verify email sending.");

  } else {
    Logger.log("No check-ins found for today.");
  }
}
