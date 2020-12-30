function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [{
      name : "Invia Mail",
      functionName : "sendMail"
    }]; 
    sheet.addMenu("Richiesta Mail", entries);
};

  function sendMail(){
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ticket SUL');
    var active = s.getCurrentCell().getRow();

    var mail =  s.getRange(active, 3).getValue();
    var name =  s.getRange(active, 4).getValue();
    var date = s.getRange(active, 5).getValue();
    var dateok = new Date(date.getTime() + 24 * 60 * 60 * 1000);
    var date1 = Utilities.formatDate(dateok, "GMT", "dd/MM/yyyy");  
    var settore = s.getRange(active, 6).getValue();
    var platform = s.getRange(active, 7).getValue();
    var device = s.getRange(active, 8).getValue();
    var nameDevice = s.getRange(active, 9).getValue();
    var problem = s.getRange(active, 10).getValue();
    var app =  s.getRange(active, 11).getValue();
    var problemS =  s.getRange(active, 12).getValue();
    var app2 =   s.getRange(active, 13).getValue();
    var nomeCognome = s.getRange(active, 14).getValue();
    var birth = s.getRange(active, 15).getValue();
    var user = s.getRange(active, 16).getValue();
    var other = s.getRange(active, 17).getValue();
    var noFly = s.getRange(active, 18).getValue();
    
    

    var mailSettore;
    

    switch(settore){

      case "Rosso":
        var mailSettore = "rossobrandizzo@decathlon.net";
        break;
      
      case "Blu":
        var mailSettore = "blubrandizzo@decathlon.net";
        break;

      case "Arancio":
        var mailSettore = "aranciobrandizzo@decathlon.net";
        break;

      case "87":
        var mailSettore = "87rpal_brandizzo@decathlon.net";
        break;

      case "Voluminosi":
        var mailSettore = "volubrandizzo@decathlon.net";
        break;

      case "Ecommerce":
        var mailSettore = "brandizzo_logdeccom@decathlon.net";
        break;

      case "Qualità":
        var mailSettore = "pssc.brandizzo@decathlon.com";
        break;

      case "Trieuse":
        var mailSettore = "trieusebrandizzo@decathlon.net";
        break;

      case "Accueil":
        var mailSettore = "accueilbrandizzo@decathlon.net";
        break;

      case "Verde":
        var mailSettore = "verdebrandizzo@decathlon.net";
        break;

      case "Laboratorio":
        var mailSettore = "lrbrandizzo@decathlon.com";
        break;

      case "Responsabili":
        var mailSettore = "italygbrandizzorul@decathlon.net";
        break;

      case "Servizio clienti":
        var mailSettore = "servizioclientitorino@decathlon.com";
        break;
      }

      var corpo_mail = "";
      var mail_sul = "sulbrandizzo@decathlon.net";

      if(platform == 'Hardware'){
        var subject = '[SUL] Riparazione Hardware '+device;
        var templH = HtmlService.createTemplateFromFile("emailHardware");
          templH.name = name;
          templH.date1 = date1;
          templH.settore = settore;
          templH.platform = platform;
          templH.device = device;
          templH.nameDevice = nameDevice;
          templH.problem = problem;       

      var messageH = templH.evaluate().getContent();
      var optionsH = {
          htmlBody:messageH
      }
      MailApp.sendEmail(mail + "," + mail_sul + "," + mailSettore, subject, corpo_mail, optionsH); 
    }
    else if(platform == 'Software'){
      var subject = '[SUL] Assistenza software '+app;
      var templS = HtmlService.createTemplateFromFile("emailSoftware");
          templS.name = name;
          templS.date1 = date1;
          templS.settore = settore;
          templS.platform = platform;
          templS.app = app;
          templS.problemS = problemS;

          var messageS = templS.evaluate().getContent();
          var optionsS = {
              htmlBody:messageS
          }
          MailApp.sendEmail(mail + "," + mail_sul + "," + mailSettore, subject, corpo_mail, optionsS); 
    }
    else if(platform == 'Creazione profilo'){ 
      var subject = '[SUL] Creazione profilo';
      var templP = HtmlService.createTemplateFromFile("emailProfilo");
          templP.name = name;
          templP.date1 = date1;
          templP.settore = settore;
          templP.platform = platform;
          templP.app2 = app2;
          templP.nomeCognome = nomeCognome;
          templP.birth = birth;
          templP.user = user;
          templP.other = other;

          var messageP = templP.evaluate().getContent();
          var optionsP = {
              htmlBody:messageP
          }
          MailApp.sendEmail(mail + "," + mail_sul + "," + mailSettore, subject, corpo_mail, optionsP);
    }
    else if(platform == 'Sito'){ 
      var subject = '[SUL] Modifiche sito';
      var templP = HtmlService.createTemplateFromFile("emailSito");
          templIT.name = name;
          templIT.date1 = date1;
          tempIT.settore = settore;
          templIT.platform = platform;
          templIT.noFly = noFly;

          var messageIT = templIT.evaluate().getContent();
          var optionsIT = {
              htmlBody:messageIT
          }
          MailApp.sendEmail(mail + "," + mail_sul + "," + mailSettore, subject, corpo_mail, optionsIT);
    }
  }
