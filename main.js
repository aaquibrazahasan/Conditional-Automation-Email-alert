// This function is for filter only Overdue 

function PullingData(){
SpreadsheetApp.flush();
var sss = SpreadsheetApp.openById('1h-hkSAMOtZ-JMYh8RXskjV8R5ob0qQLr38UpAEl8Shg');
var ss = sss.getSheetByName('Insurance Policies');
var workingFile = sss.getSheetByName('Working file');
var Email = workingFile.getRange("B1").getValue();
var CC = workingFile.getRange("B2").getValue();
var BCC = workingFile.getRange("B3").getValue();
var lrw = ss.getLastRow();
var data1 = ss.getRange(2,1,lrw,16).getValues();
data = data1.filter(isColAAAWS);
 //  Logger.log(data)

  for (i = 0; i < data.length ;i++) { 
        if (data[i][1] != "" ){
        
     
              var SNo = data[i][0]
              var Type = data[i][1]
              var Description = data[i][2]
              var RenewalDueDate = data[i][3]
              var FromTo = data[i][4]
              var InsuranceCompany = data[i][5]
              var PolicyNo = data[i][6]
              var SumInsured = data[i][7]
              var PremiumPaid = data[i][8]
              var GST = data[i][9]
              var NCB = data[i][10]
              var ValueinP = data[i][11]
              var Remarks = data[i][12]
              var PendingDays = data[i][13]
              var Status = data[i][14]


              var TempleteBody = HtmlService.createTemplateFromFile('index.html').evaluate().getContent()

                .replace("{SNo}",SNo)  
                .replace("{Type}",Type) 
                .replace("{Description}",Description) 
                .replace("{RenewalDueDate}",RenewalDueDate) 
                .replace("{FromTo}",FromTo) 
                .replace("{InsuranceCompany}",InsuranceCompany) 
                .replace("{PolicyNo}",PolicyNo) 
                .replace("{SumInsured}",SumInsured) 
                .replace("{PremiumPaid}",PremiumPaid)
                .replace("{GST}",GST) 
                .replace("{NCB}",NCB)  
                .replace("{ValueinP}",ValueinP) 
                .replace("{Remarks}",Remarks) 
                .replace("{PendingDays}",PendingDays) 
                .replace("{Status}",Status)
                 

              Logger.log(TempleteBody)

              var Subject = "[Due Insurance Policy] - " + PolicyNo
            GmailApp.sendEmail(Email,Subject,TempleteBody,{htmlBody:TempleteBody,name:"Insurance Alerts",cc : CC,bcc : BCC })
            //   GmailApp.sendEmail("jasbir.kaur@twigafiber.com",Subject,TempleteBody,{htmlBody:TempleteBody,name:"Insurance Alerts"})
     
       }

  }

};



// Filter Applied
function isColAAAWS(arr) {
  // Logger.log(arr[4])
  return (arr[14] == "Overdue" &&  arr[15] != true) 


};
