//Name: <?= sn ?> <?= fn ?> và <?= ln ?>


function Mail_Merge_App(){
  let mail_title = "TEST MAIL";
  let first = 0;
  let last = 1;
  let sex = 2;
  let email = 3;
  let emailTemp = HtmlService.createTemplateFromFile("index");
  let ws = SpreadsheetApp.openById("1WgBk2TiWY945u_Va-eLHZn7c3CfInp_Tl4C4JULcS8c") // <-- ID của Sheet
  let data = ws.getRange("A2:D" + ws.getLastRow()).getValues(); 
  data.forEach(function(row){
      emailTemp.fn = row[first];
      emailTemp.ln = row[last];
      if (row[sex]=="Nam") {emailTemp.sn = "anh";} else emailTemp.sn = "chị";
      let htmlMessage = emailTemp.evaluate().getContent();
      GmailApp.sendEmail(
        row[email], 
        mail_title , //+ ', ' + row[first] + '!',
        "Your email doesn't support HTML.",
        {name: "ngominhhung208@gmail.com", htmlBody: htmlMessage} 
        );        
      }
  );   
}
