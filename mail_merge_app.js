//Name: <?= fn ?> và <?= ln ?>


function Mail_Merge_App(){
  let mail_title = "Tiêu đề của email";
  let first = 0;
  let last = 1;
  let email = 2;
  let emailTemp = HtmlService.createTemplateFromFile("index");
  let ws = SpreadsheetApp.openById("1WgBk2TiWY945u_Va-eLHZn7c3CfInp_Tl4C4JULcS8c") // <-- ID của Sheet
  let data = ws.getRange("A2:C" + ws.getLastRow()).getValues(); 
  data.forEach(function(row){
      emailTemp.fn = row[first];
      emailTemp.ln = row[last];
      let htmlMessage = emailTemp.evaluate().getContent();
      GmailApp.sendEmail(
        row[email], 
        mail_title , //+ ', ' + row[first] + '!',
        "Your email doesn't support HTML.",
        {name: "Email gửi", htmlBody: htmlMessage} 
        );        
      }
  );   
}
