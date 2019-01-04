function sendFeedback() {
  // select the range from the Summary sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Summary");
  var lastRow = sheet.getLastRow();
  
  var range = sheet.getRange(4,1,lastRow-3,15).getValues();
  
  var timestamp = new Date();
  
  // loop over range and send if "Yes"
  for (var i = 0; i < range.length; i++) {
    if (range[i][12] == "Yes") {
      
      // choose email, slack or both channels
      switch (range[i][13]) {
        case "Email":
          // send email to agent 
          sendEmail(range[i]);
          break;
        
        case "Slack":
          // post message to slack
          sendToSlack(range[i]);
          break;
          
        case "Both":
          // send email and post to Slack
          sendEmail(range[i]);
          sendToSlack(range[i]);
          break;
      }
      
      // add timestamp to final column 
      sheet.getRange(i+4,15,1,1).setValue(timestamp);
    };
  }
}

function sendEmail(agent) {
  var timestamp = new Date();
  MailApp.sendEmail({
     to: agent[1],
     subject: "CS Feedback",
     htmlBody: 
      "Hi " + agent[0] +",<br><br>" +
      "Here are your scores and feedack for the week. If you have any questions, feel free to reach out to XXXXX!<br><br>" +
      "<table  border='1'><tr><td>Service Score</td>" +
      "<td>Data Entry Score</td>" +
      "<td>Timeliness Score</td>" +
      "<td>Total Cases</td></tr>" +
      "<tr><td>" + agent[5] + "</td>" +
      "<td>" + agent[6] +"</td>" +
      "<td>" + agent[7] + "</td>" +
      "<td>" + agent[8] + "</td></tr></table>" +
      "<br><b>Positive Notes:</b><br>" +
      agent[9] + "<br>" +
      "<br><b>Areas for improvement:</b><br>" +
      agent[10] + "<br>" +
      "<br><b>Any other comments:</b><br>" +
      agent[11] +
      "<br><br>Marked by: " + agent[3] +
      "<br>Date: " + timestamp,
   });
}

function sendToSlack(agent) {
  var timestamp = new Date();
  
  var url = "https://hooks.slack.com/services/XXXXXXXXXXXXXXXXXXXXX";
  
  var payload = {
    "channel": "@"+agent[2],
    "username": "XXXXXXXXX",
    "text": "Hi " + agent[0] +
      "\n Here are your scores and feedack for the week. If you have any questions, feel free to reach out to XXXXX! \n" +
      "\n Service Score: " + agent[5] +
      "\n Data Entry Score: " + agent[6] +
      "\n Timeliness Score: " + agent[7] +
      "\n Total Cases: " + agent[8] +
      "\n *Positive notes:* " + agent[9] +
      "\n *Areas for improvement:* " + agent[10] +
      "\n *Any other comments:* " + agent[11] +
      "\n \n Reviewed by: " + agent[3] +
      "\n Date: " + timestamp,
    "icon_emoji": ":inbox_tray:"
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  return UrlFetchApp.fetch(url,options);
}
