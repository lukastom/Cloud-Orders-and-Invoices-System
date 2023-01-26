/* Verze kódu: v 1.05

Changelog:    1.00 první verze
              1.01 rozdělení sloupců s předf. a f. v Zahraničie - platby + JCD
              1.02 zjednodušení zahr. mimo EU...mimo EU, zahraničie EU...EU
              1.03 bug fix
              1.04 přidáno "BETA"
              1.05 přejmenování - přidání [name]

Bug tracker: 
- !! když jsou dvě faktury za sebou (mezi nimi je čárka) a ta druhá končí nulou (nulami), tyto nuly jsou truncated, protože automatický formát si myslí, že to celé je desetinné číslo
- když se zadá fakultná nemocnica skalica a vyhledává se podle "skal" tak to hodí error (stejný, jako když nic není v listě)
- fakt. RD 57 (2014) - nevyplnila se adresa

Když předěláváme na nový rok, tak co nahradit (příklad 2013 -> 2014)

1) [sheet id] -> [sheet id]
2) 2013 -> 2014
3) ručně některé 13 -> 14
*/

//nastavení cache (jednoduchá náhrada global variables)
var cache = CacheService.getPrivateCache(); 
//-------------------------- OTEVŘÍT SOUBOR NA ZÁKLADĚ JMÉNA ---------------------------------
//výsledkem je objekt file
function getFileByName_(folder, filename) {
  var files = folder.find("title: \"" + filename + "\"");     
  for( var i in files){
    if ( files[i].getName() == filename ) 
    {
      //Logger.log("found " + files[i].getName());
      return files[i];
    }
  } //for i in files    
  //Logger.log("did not find " + filename);
} //getFileByName_

//-------------------------- VYHLEDÁVACÍ FUNKCE ---------------------------------
//extends Object 
//.findCells(String) or .findCells(Number) 
//returns array[[number(row),number(column)]] 
//vadí tomu message boxy, pak to nefunguje!!!
Object.prototype.findCells = function(key) {   
  var searchMatch = [];
  for (i = 0; i < this.length; i++){
    for (j = 0; j < this[i].length; j++){
      if (this[i][j].toString().toLowerCase().search(key.toString().toLowerCase()) != -1){
        searchMatch.push([i+1,j+1]); 
      }
    }
  }  
  return searchMatch;
}  

//-------------------------- EXPORTY XLSX, PDF ---------------------------------

function vyexportuj_pdf(cesta, idspreadsheetu, vystupni_jmeno){  
  //SpreadsheetApp.getActiveSpreadsheet().getId()
  var pdf = SpreadsheetApp.openById(idspreadsheetu).getAs("application/pdf");
  var file = DocsList.getFolder(cesta).createFile(pdf);
  file.rename(vystupni_jmeno + '.pdf');

  return file;
}

function vyexportuj_xlsx(cesta,id,vystupni_jmeno){
  var url = 'https://docs.google.com/feeds/';
  var doc = UrlFetchApp.fetch(url+'download/spreadsheets/Export?key='+id+'&exportFormat=xls',
  googleOAuth_('docs',url)).getBlob()
  DocsList.getFolder(cesta).createFile(doc).rename(vystupni_jmeno + '.xlsx');
}

//autorizační proces OAuth (pokud to nefunguje, tak debugovat, přitom vyběhne hláška, tu potvrdit, a je to)
function googleOAuth_(name,scope) {
  var oAuthConfig = UrlFetchApp.addOAuthService(name);
  oAuthConfig.setRequestTokenUrl("https://www.google.com/accounts/OAuthGetRequestToken?scope="+scope);
  oAuthConfig.setAuthorizationUrl("https://www.google.com/accounts/OAuthAuthorizeToken");
  oAuthConfig.setAccessTokenUrl("https://www.google.com/accounts/OAuthGetAccessToken");
  oAuthConfig.setConsumerKey('anonymous');
  oAuthConfig.setConsumerSecret('anonymous');
  return {oAuthServiceName:name, oAuthUseToken:"always"};
}

//-------------------------- PŘIDÁNÍ MENU DO souboru Fakturácia ---------------------------------

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{name : "Zaregistruj do RD", functionName : "Zaregistruj"},
                 {name : "Pridaj predfaktúru", functionName : "pridajPredfakturu"},
                 {name : "Pridaj + vystav predfaktúru SK (BETA)", functionName : "pridajvytvorPredfakturu"},
                 {name : "Pridaj faktúru", functionName : "pridajFakturu"}];
  sheet.addMenu("Fakturácia", entries);
};

//-------------------------------- VYSTAVENÍ PŘEDFAKTURY - NEPOUŽÍVÁ SE ------------------------------------

/*function vystavPredfakturu() {

  var id_sablony_predfaktury = "[sheet id]"
  
  //otevřít Šablona proforma faktura SK
  var spreadSheet = SpreadsheetApp.openById(id_sablony_predfaktury);
  var sheet = spreadSheet.getSheetByName("Proforma faktúra SK");  
  
  //vytvořit z kopie novou předfakturu
  var file_sablona = DocsList.getFileById(id_sablony_predfaktury);
  var folder = DocsList.getFolder("Work/2014/Proforma faktúry");
  var rootfolder = DocsList.getRootFolder();
  file_sablona.makeCopy('123_Land_Abbr').addToFolder(folder);
  var file_nova_predfaktura = getFileByName_(folder, '123_Land_Abbr');
  file_nova_predfaktura.removeFromFolder(rootfolder);
   
  //udělat tuto předfakturu aktivním spreadsheetem
  var novy_spreadSheet = SpreadsheetApp.openById(file_nova_predfaktura.getId());
  SpreadsheetApp.setActiveSpreadsheet(novy_spreadSheet); 
  
  //!!!tady je samotné vyplnění předfaktury!!!
  //+flush?  
  
  //export do XLSX 
  vyexportuj_xlsx("Work/2014/Proforma faktúry", id_sablony_predfaktury, "123_Land_Abbr");   
  
  //vykopírování do pomocného souboru, zaktivnění
  var folder_pomocny = DocsList.getFolder("Work/2014/Proforma faktúry/Pomocné súbory");  
  file_sablona.makeCopy('123_Land_Abbr' + '_pom').addToFolder(folder_pomocny); 
  var file_pomocny = getFileByName_(folder_pomocny, '123_Land_Abbr' + '_pom');
  file_pomocny.removeFromFolder(rootfolder);  
  var pomocny_spreadSheet = SpreadsheetApp.openById(file_pomocny.getId());
  SpreadsheetApp.setActiveSpreadsheet(pomocny_spreadSheet);  
  
  //smazání dvou sheetů
  var pomocny_sheet = pomocny_spreadSheet.getSheetByName("Export"); 
  SpreadsheetApp.setActiveSheet(pomocny_sheet);
  SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
  var pomocny_sheet = pomocny_spreadSheet.getSheetByName("Dodací list SK"); 
  SpreadsheetApp.setActiveSheet(pomocny_sheet);  
  SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();  
   
  //export do PDF
  vyexportuj_pdf("Work/2014/Proforma faktúry", file_pomocny.getId(), "123_Land_Abbr");
    
};*/
//------------------------------------------------ PŘIDÁNÍ + VYSTAVENÍ PREDFAKTÚRY ---------------------------------------------------------------

function pridajvytvorPredfakturu() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Číslo predfaktúry");   
  formular_pridani_do_predfaktury_vytvoreni(); 
};
  
function formular_pridani_do_predfaktury_vytvoreni(e) {
  var doc = SpreadsheetApp.openById("[sheet id]");
  var app = UiApp.createApplication().setWidth(500).setHeight(500).setTitle('Pridanie + vystavenie predfaktúry');
  var grid = app.createGrid(12, 2);
    
  var vlozene_cislo_z_rd = app.createTextBox().setName('cislo_z_rd');
  
  grid.setWidget(0, 0, app.createHTML('<b>Vložte číslo z RD</b>'));
  grid.setWidget(0, 1, vlozene_cislo_z_rd);
  
  //zobrazování odběratele z RD - nemazat
  //var label_odberatel1 = app.createLabel('Odberateľ (z RD):').setVisible(false).setId('label_odberatel1');
  //var label_odberatel2 = app.createLabel('').setVisible(false).setId('label_odberatel2');
  
  //grid.setWidget(1, 0, label_odberatel1);
  //grid.setWidget(1, 1, label_odberatel2);   
  
  var textBox = app.createTextBox().setName('textBox').setWidth('300px').setStyleAttribute("background", "#e8e8e8").setId('textBox');
  var tBoxHandler = app.createServerKeyHandler('search_podle_okenka');
  tBoxHandler.addCallbackElement(textBox);
  textBox.addKeyUpHandler(tBoxHandler);

  grid.setWidget(1, 0, app.createLabel('Vyhľadávanie vo fakturačných adresách'));
  grid.setWidget(1, 1, textBox);   

  var adresa_odberatela = app.createListBox().setWidth('300px').setName('adresa_odberatela_list').setId('adresa_odberatela_list');
  
  var sheet2 = doc.getSheetByName("Fakturačné adresy"); 

  var numItemList1 = sheet2.getLastRow()-1;//-1 is to exclude header row
  //get the item array
  var list1ItemArray = sheet2.getRange(2,1,numItemList1,1).getValues();
  //Add the items in ListBox
  for(var i=0; i<list1ItemArray.length; i++){
    adresa_odberatela.addItem(list1ItemArray[i][0])
  }  
  
  grid.setWidget(2, 0, app.createLabel('Vyberte zo zoznamu adresu odberateľa'));
  grid.setWidget(2, 1, adresa_odberatela);   
  
  var adresa_odberatela_na_fakture_label = app.createHTML('<b>Adresa odberateľa</b><br /><i>(tak, ako bude na faktúre)</i>').setVisible(false).setId('adresa_odberatela_na_fakture_label');    
  var adresa_odberatela_na_fakture = app.createHTML('').setVisible(false).setId('adresa_odberatela_na_fakture');  

  grid.setWidget(3, 0, adresa_odberatela_na_fakture_label);
  grid.setWidget(3, 1, adresa_odberatela_na_fakture); 
  
  var su_rovnake = app.createListBox().setName('su_rovnake').setId('su_rovnake');
  su_rovnake.addItem('Áno');
  su_rovnake.addItem('Nie');
  
  grid.setWidget(4, 0, app.createHTML('Sú adresy odberateľa<br />a dodací <b>rovnaké?</b>').setId('su_rovnake_label'));
  grid.setWidget(4, 1, su_rovnake);   
  
  //-----druhá adresa----
  
  cache.put('su_rovnake', 'Áno', 3600); //defaultní volba je Nie, pak se to mění handlery
  
  var textBoxB = app.createTextBox().setName('textBoxB').setWidth('300px').setStyleAttribute("background", "#e8e8e8").setVisible(false).setId('textBoxB');
  var tBoxHandlerB = app.createServerKeyHandler('search_podle_okenkaB');
  tBoxHandlerB.addCallbackElement(textBoxB);
  textBoxB.addKeyUpHandler(tBoxHandlerB);  
  
  grid.setWidget(5, 0, app.createLabel('Vyhľadávanie vo fakturačných adresách').setVisible(false).setId('vyhladavane_vo_fakturacnychB'));
  grid.setWidget(5, 1, textBoxB);   

  var adresa_odberatelaB = app.createListBox().setWidth('300px').setName('adresa_odberatela_listB').setVisible(false).setId('adresa_odberatela_listB'); 

  var numItemList1B = sheet2.getLastRow()-1;//-1 is to exclude header row
  //get the item array
  var list1ItemArrayB = sheet2.getRange(2,1,numItemList1B,1).getValues();
  //Add the items in ListBox
  for(var i=0; i<list1ItemArrayB.length; i++){
    adresa_odberatelaB.addItem(list1ItemArrayB[i][0])
  }  
  
  grid.setWidget(6, 0, app.createLabel('Vyberte zo zoznamu dodací adresu').setVisible(false).setId('vyberte_odberatelaB'));
  grid.setWidget(6, 1, adresa_odberatelaB);   
  
  var adresa_odberatela_na_fakture_labelB = app.createHTML('<b>Dodací adresa</b><br /><i>(tak, ako bude na faktúre)</i>').setVisible(false).setId('adresa_odberatela_na_fakture_labelB');    
  var adresa_odberatela_na_faktureB = app.createHTML('').setVisible(false).setId('adresa_odberatela_na_faktureB');  

  grid.setWidget(7, 0, adresa_odberatela_na_fakture_labelB);
  grid.setWidget(7, 1, adresa_odberatela_na_faktureB); 

  //-----konec druhé adresy------

  // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
  // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear();   

  //datum + 14
  var den = new Date();   
  den.setDate(den.getDate()+14);
  var datum_plus_ctrnact = den.getDate() + "." + (den.getMonth() + 1) + "." + den.getYear();  
  
  grid.setWidget(8, 0, app.createLabel('Dátum splatnosti:'));
  grid.setWidget(8, 1, app.createTextBox().setName('datum_splatnosti').setId('datum_splatnosti').setText(datum_plus_ctrnact));
  
  grid.setWidget(9, 0, app.createLabel('Dátum dodania tovaru:'));
  grid.setWidget(9, 1, app.createTextBox().setName('datum_dodania_tovaru').setId('datum_dodania_tovaru').setText(datum_string));    
  
  grid.setWidget(10, 0, app.createLabel('Kolik variant predfaktúr urobit?').setVisible(false));
  grid.setWidget(10, 1, app.createTextBox().setName('kolik_variant').setText('1').setVisible(false));

  if (Session.getActiveUser() == "[e-mail address]") {
   var vypln_jmeno = "l";
  } else {
   var vypln_jmeno = "s";
  }  
  
  grid.setWidget(11, 0, app.createLabel('Vaše jméno (1 malé písmeno)'));
  grid.setWidget(11, 1, app.createTextBox().setName('vase_jmeno').setText(vypln_jmeno));    

  // Create a vertical panel..
  var panel = app.createVerticalPanel();

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a callback element and the handler as a click handler
  // Identify the function b as the server click handler

  var button = app.createButton('Pridaj predfaktúru');
  var handler = app.createServerHandler('d');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);
  
  //Handlery na zobrazení odběratele - nepoužívá se, nemazat
  /*var handler1 = app.createServerKeyHandler('show_odberatel1'); //co se má spustit
  handler1.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler1); //čím se to spouští */
  var handler2 = app.createServerKeyHandler('show_odberatel2');
  handler2.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler2); 
  //Handler na vyhledání adresy podle čísla z RD
  var handler3 = app.createServerKeyHandler('search_podle_cisla'); //co se má spustit
  handler3.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler3); //čím se to spouští 
  //Handler na doplnění datumů podle čísla z RD
  /*var handler6 = app.createServerKeyHandler('show_datumy'); //co se má spustit
  handler6.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler6); //čím se to spouští   */
  
  //handler při změně výběru v seznamu adres 1
  var handler4 = app.createServerChangeHandler('uprava_adresy_pri_listovani'); //co se má spustit
  handler4.addCallbackElement(panel);
  adresa_odberatela.addChangeHandler(handler4); //čím se to spouští   

  //-----handlery na druhou adresu
  var handler5 = app.createServerChangeHandler('ukaz_druhou_adresu');
  handler5.addCallbackElement(panel);
  su_rovnake.addChangeHandler(handler5);
  
  //handler při změně výběru v seznamu adres B
  var handler4B = app.createServerChangeHandler('uprava_adresy_pri_listovaniB'); //co se má spustit
  handler4B.addCallbackElement(panel);
  adresa_odberatelaB.addChangeHandler(handler4B); //čím se to spouští    
  
  //podle čísla z RD vyplněního vyhledávacího okna adresy B
  var handler2B = app.createServerKeyHandler('show_odberatel2B');
  handler2B.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler2B); 
  
  //Handler na vyhledání adresy B podle čísla z RD
  var handler3B = app.createServerKeyHandler('search_podle_cislaB'); //co se má spustit
  handler3B.addCallbackElement(panel);
  vlozene_cislo_z_rd.addKeyUpHandler(handler3B); //čím se to spouští   
  
  panel.add(button);
  app.add(panel);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);

}

//--------------------------------- PŘIDÁNÍ + VYSTAVENÍ PREDFAKTÚRY - handlery ---------------------------------------
/*//doplnění do vyhledávacího okna adresy B
function show_datumy(e){
  var app = UiApp.getActiveApplication();
  
  // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
  // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear();   

  //datum + 14
  var den = new Date();   
  den.setDate(den.getDate()+14);
  var datum_plus_ctrnact = den.getDate() + "." + (den.getMonth() + 1) + "." + den.getYear();  

  app.getElementById('datum_splatnosti').setText(datum_plus_ctrnact);
  app.getElementById('datum_dodania_tovaru').setText(datum_string);  
  return app;  
}*/
//funkce hledání v rozbalovacím seznamu adres B na základě zadání čísla z RD
function search_podle_cislaB(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  app.getElementById('adresa_odberatela_listB').clear(); 

  var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RD");
  var radek_v_rd = Number(e.parameter.cislo_z_rd) + 1;
  var co_hledat = sheet4.getRange("D" + radek_v_rd).getValue(); 
  
  var searchKey = new RegExp(co_hledat,"gi");
    
  if (searchKey == "") app.getElementById('textBoxB').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            app.getElementById('adresa_odberatela_listB').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_labelB');
             app.getElementById('adresa_odberatela_na_faktureB').setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresyB', radek_fakturacni_adresy, 3600);             
            } 
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }

  return app;
}
//doplnění do vyhledávacího okna adresy B
function show_odberatel2B(e){
  var app = UiApp.getActiveApplication();

  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = spreadSheet.getSheetByName("RD");

  var cislo_z_rd = e.parameter.cislo_z_rd;  
  
  var radek_v_rd = Number(cislo_z_rd) + 1;
  var odberatel = sheet.getRange("D" + radek_v_rd).getValue();

  /*app.getElementById('label_odberatel2').setVisible(true).setText(odberatel);*/
  app.getElementById('textBoxB').setText(odberatel);
  return app;
}
//funkce při změně výběru v seznamu adres B
function uprava_adresy_pri_listovaniB(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  //app.getElementById('adresa_odberatela_list').clear(); 

  var searchKey = new RegExp(e.parameter.adresa_odberatela_listB,"gi");
  
  Logger.log(e.parameter.adresa_odberatela_listB); 
  
  //if (searchKey == "") app.getElementById('textBox').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            //app.getElementById('adresa_odberatela_list').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_labelB').setVisible(true);
             app.getElementById('adresa_odberatela_na_faktureB').setVisible(true).setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresyB', radek_fakturacni_adresy, 3600);              
            }        
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }  
  
  // set the top listbox item as the default
  //if (listBoxCount > 0) app.getElementById('adresa_odberatela_list').setItemSelected(0, true);
  
  //if (e.parameter.textBox.length < 1) app.getElementById('adresa_odberatela_list').clear();
  return app;
}
//funkce hledání v rozbalovacím seznamu B
function search_podle_okenkaB(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  app.getElementById('adresa_odberatela_listB').clear(); 

  var searchKey = new RegExp(e.parameter.textBoxB,"gi");
  
  if (searchKey == "") app.getElementById('textBoxB').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            app.getElementById('adresa_odberatela_listB').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_labelB').setVisible(true);
             app.getElementById('adresa_odberatela_na_faktureB').setVisible(true).setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresyB', radek_fakturacni_adresy, 3600);              
            }        
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }  
  
  // set the top listbox item as the default
  if (listBoxCount > 0) app.getElementById('adresa_odberatela_listB').setItemSelected(0, true);
  
  if (e.parameter.textBoxB.length < 1) app.getElementById('adresa_odberatela_listB').clear();
  return app;
}
//su adresy rovnake?
function ukaz_druhou_adresu(e){
  var app = UiApp.getActiveApplication();
  if (e.parameter.su_rovnake == "Nie"){
   app.getElementById('textBoxB').setVisible(true); 
   app.getElementById('vyhladavane_vo_fakturacnychB').setVisible(true); 
   app.getElementById('adresa_odberatela_listB').setVisible(true);     
   app.getElementById('vyberte_odberatelaB').setVisible(true);     
   app.getElementById('adresa_odberatela_na_fakture_labelB').setVisible(true); 
   app.getElementById('adresa_odberatela_na_faktureB').setVisible(true);
   cache.put('su_rovnake', 'Nie', 3600);    
  }
  if (e.parameter.su_rovnake == "Áno"){
   app.getElementById('textBoxB').setVisible(false); 
   app.getElementById('vyhladavane_vo_fakturacnychB').setVisible(false); 
   app.getElementById('adresa_odberatela_listB').setVisible(false);     
   app.getElementById('vyberte_odberatelaB').setVisible(false);     
   app.getElementById('adresa_odberatela_na_fakture_labelB').setVisible(false); 
   app.getElementById('adresa_odberatela_na_faktureB').setVisible(false); 
   cache.put('su_rovnake', 'Áno', 3600);    
  }  
  return app;
}
//funkce při změně výběru v seznamu adres 1
function uprava_adresy_pri_listovani(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  //app.getElementById('adresa_odberatela_list').clear(); 

  var searchKey = new RegExp(e.parameter.adresa_odberatela_list,"gi");
  
  Logger.log(e.parameter.adresa_odberatela_list); 
  
  //if (searchKey == "") app.getElementById('textBox').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            //app.getElementById('adresa_odberatela_list').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_label').setVisible(true);
             app.getElementById('adresa_odberatela_na_fakture').setVisible(true).setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresy', radek_fakturacni_adresy, 3600);              
            }        
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }  
  
  // set the top listbox item as the default
  //if (listBoxCount > 0) app.getElementById('adresa_odberatela_list').setItemSelected(0, true);
  
  //if (e.parameter.textBox.length < 1) app.getElementById('adresa_odberatela_list').clear();
  return app;
}
//funkce hledání v rozbalovacím seznamu na základě zadání čísla z RD
function search_podle_cisla(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  app.getElementById('adresa_odberatela_list').clear(); 

  var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RD");
  var radek_v_rd = Number(e.parameter.cislo_z_rd) + 1;
  var co_hledat = sheet4.getRange("D" + radek_v_rd).getValue(); 
  
  var searchKey = new RegExp(co_hledat,"gi");
    
  if (searchKey == "") app.getElementById('textBox').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            app.getElementById('adresa_odberatela_list').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_label').setVisible(true);
             app.getElementById('adresa_odberatela_na_fakture').setVisible(true).setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresy', radek_fakturacni_adresy, 3600);             
            } 
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }

  return app;
}
//funkce hledání v rozbalovacím seznamu
function search_podle_okenka(e){
  var ss = SpreadsheetApp.openById("[sheet id]");
  var sheet3 = ss.getSheetByName("Fakturačné adresy");   
  var numItemList1 = sheet3.getLastRow()-1;//-1 is to exclude header row
  
  var app = UiApp.getActiveApplication();
   
  app.getElementById('adresa_odberatela_list').clear(); 

  var searchKey = new RegExp(e.parameter.textBox,"gi");
  
  if (searchKey == "") app.getElementById('textBox').setValue('');
  var range = sheet3.getRange(2,1,numItemList1,1).getValues();
     
  var listBoxCount = 0;
  
  for (var i in range){
    if (i<numItemList1){ 
     if (range[i][0].search(searchKey) != -1){
            app.getElementById('adresa_odberatela_list').addItem(range[i][0].toString()); 
            //doplnit vybranou adresu níže
            if (listBoxCount == 0){
             var radek_fakturacni_adresy = Number(i) + 2;
             app.getElementById('adresa_odberatela_na_fakture_label').setVisible(true);
             app.getElementById('adresa_odberatela_na_fakture').setVisible(true).setHTML(range[i][0].toString() + '<br />' + sheet3.getRange("B" + radek_fakturacni_adresy).getValue() + '<br />' + sheet3.getRange("C" + radek_fakturacni_adresy).getValue());           
             cache.put('radek_fakturacni_adresy', radek_fakturacni_adresy, 3600);              
            }        
          var listBoxCount = listBoxCount + 1;
     }  
    }
  }  
  
  // set the top listbox item as the default
  if (listBoxCount > 0) app.getElementById('adresa_odberatela_list').setItemSelected(0, true);
  
  if (e.parameter.textBox.length < 1) app.getElementById('adresa_odberatela_list').clear();
  return app;
}

//Handlery na zobrazení odběratele - nepoužívá se, nemazat
/*function show_odberatel1(e){
  var app = UiApp.getActiveApplication();
  app.getElementById('label_odberatel1').setVisible(true);
  return app;
}*/

function show_odberatel2(e){
  var app = UiApp.getActiveApplication();

  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = spreadSheet.getSheetByName("RD");

  var cislo_z_rd = e.parameter.cislo_z_rd;  
  
  var radek_v_rd = Number(cislo_z_rd) + 1;
  var odberatel = sheet.getRange("D" + radek_v_rd).getValue();

  /*app.getElementById('label_odberatel2').setVisible(true).setText(odberatel);*/
  app.getElementById('textBox').setVisible(true).setText(odberatel);
  return app;
}

//--------------------------------- PŘIDÁNÍ + VYSTAVENÍ PREDFAKTÚRY - vyplnění hodnot z formuláře ---------------------------------------
// Vyplněné hodnoty z formuláře uloží do proměnných
function d(e) {
  
  //otevřít hárok, najít poslední řádek
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Číslo predfaktúry");   
  var lastRowIndex = sheet.getLastRow();  

  // You can access e.parameter.userName because you used setName('userName') above and
  // also added the grid containing those widgets as a callback element to the server
  // handler.
  var cislo_z_rd = e.parameter.cislo_z_rd;
  var kolik_variant = e.parameter.kolik_variant; 
  
 // zvýšit index o 1, takže ukazuje na nový řádek
  lastRowIndex++;
  
 // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
 // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear(); 
 
 // vyplnit datum
  sheet.getRange("A" + lastRowIndex).setValue(datum_string);   
  
  //vyplnění údaje do predfaktúry
  sheet.getRange("D" + lastRowIndex).setValue(cislo_z_rd); 

 //zjistit poslední číslo faktury
  var cislo_posledni_faktury = sheet.getRange("B" + (lastRowIndex - 1)).getValue(); 
  if (cislo_posledni_faktury == "Číslo predfaktúry") {
   cislo_posledni_faktury = "P00014";
  }  
 
 //zvětšení čísla poslední předfaktury o 1
  cislo_posledni_faktury = cislo_posledni_faktury.slice(1,4);
  cislo_posledni_faktury = Number(cislo_posledni_faktury) + 1;
  if (cislo_posledni_faktury < 10) {
   cislo_posledni_faktury = "P00" + cislo_posledni_faktury + "14";
  } else if (cislo_posledni_faktury > 99) {
   cislo_posledni_faktury = "P" + cislo_posledni_faktury + "14"; 
  } else {
   cislo_posledni_faktury = "P0" + cislo_posledni_faktury + "14";  
  }   

  //napodmínkování více variant předfaktury 
  if (kolik_variant == 2) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b";
  } else if (kolik_variant == 3) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b, " + cislo_posledni_faktury + "c";
  } else if (kolik_variant == 4) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b, " + cislo_posledni_faktury + "c, " + cislo_posledni_faktury + "d"; 
  }  
 
  // vyplnit číslo nové faktury
  sheet.getRange("B" + lastRowIndex).setValue(cislo_posledni_faktury);  
  
  //vyplnit uživatele
  if (e.parameter.vase_jmeno == "k") {
   var vase_jmeno="[name]";
   var vase_jmeno_prijmeni="[name surname]";        
  }
  else if (e.parameter.vase_jmeno == "m") {
   var vase_jmeno="[name]";
   var vase_jmeno_prijmeni="[name surname]";      
  } 
  else if (e.parameter.vase_jmeno == "s") {
   var vase_jmeno="[name]";
   var vase_jmeno_prijmeni="[name surname]";        
  } 
  else if (e.parameter.vase_jmeno == "v") {
   var vase_jmeno="[name]";
   var vase_jmeno_prijmeni="[name surname]";       
  }   
  else if (e.parameter.vase_jmeno == "l") {
   var vase_jmeno="[name]";
   var vase_jmeno_prijmeni="[name surname]";       
  }    
  sheet.getRange("C" + lastRowIndex).setValue(vase_jmeno);  
  
  //načtení vyplněné nové předfaktury
  cislo_posledni_faktury = sheet.getRange("B" + lastRowIndex).getValue();
  
  //vyplnit do RD číslo předfaktury
  sheet = spreadSheet.getSheetByName("RD");
  radek_v_rd = Number(cislo_z_rd) + 1;
  var dosavadni_predfaktury = sheet.getRange("H" + radek_v_rd).getValue();
  if (dosavadni_predfaktury == "") {
   sheet.getRange("H" + radek_v_rd).setValue(cislo_posledni_faktury);   
  } else {
   sheet.getRange("H" + radek_v_rd).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury);    
  }
  
  //>>do vystavené předfaktury: Odběratel, Předmět objednávky, ID objednávky
  var odberatel = sheet.getRange("D" + radek_v_rd).getValue();
  var predmet_objednavky = sheet.getRange("G" + radek_v_rd).getValue();
  var id_objednavky = sheet.getRange("F" + radek_v_rd).getValue();
  var objednavajuci = sheet.getRange("E" + radek_v_rd).getValue();   
  
  //opodmínkování - jestliže je to zahraničí - doplnit číslo předfaktury a obarvit
  var je_to_tuzex = sheet.getRange("I" + radek_v_rd).getValue();
  if ( (je_to_tuzex == "mimo EU") || (je_to_tuzex == "EU") ) {
  
   //vyplnit do Zahraničie - platby + JCD číslo predaftury
   sheet = spreadSheet.getSheetByName("Zahraničie - platby + JCD");
   lastRowIndex = sheet.getLastRow();
  
   //nalezení řádku, na kterém je cislo_z_rd 
   var bunka_s_cislem = sheet.getRange('P1:P' + lastRowIndex).getValues().findCells(e.parameter.cislo_z_rd);
   var radek_s_cislem = bunka_s_cislem[1,0][1,0];
   var dosavadni_predfaktury = sheet.getRange("C" + radek_s_cislem).getValue();
   if (dosavadni_predfaktury == "") {
    //sheet.getRange("C" + radek_s_cislem).setValue(cislo_posledni_faktury).setBackgroundColor("orange"); 
    sheet.getRange("C" + radek_s_cislem).setValue(cislo_posledni_faktury);     
   } else {
    //sheet.getRange("C" + radek_s_cislem).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury).setBackgroundColor("orange");    
    sheet.getRange("C" + radek_s_cislem).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury);      
   }  
    
  }   
    
  //----------------------------------------------------------------------------------
  //---------------------------PŘEDVYPLNĚNÍ PŘEDFAKTURY------------------------------
  //----------------------------------------------------------------------------------

  var id_sablony_predfaktury = "[sheet id]"
   
  //otevřít Šablona proforma faktura SK
  var spreadSheet = SpreadsheetApp.openById(id_sablony_predfaktury);
  var sheet = spreadSheet.getSheetByName("Proforma faktúra SK");  
  
  //vytvořit z kopie novou předfakturu
  var file_sablona = DocsList.getFileById(id_sablony_predfaktury);
  var folder = DocsList.getFolder("Work/2014/Proforma faktúry");
  var rootfolder = DocsList.getRootFolder();
  file_sablona.makeCopy(cislo_posledni_faktury + '_' + odberatel + '_' + predmet_objednavky).addToFolder(folder);
  var file_nova_predfaktura = getFileByName_(folder, cislo_posledni_faktury + '_' + odberatel + '_' + predmet_objednavky);
  file_nova_predfaktura.removeFromFolder(rootfolder);
   
  //udělat tuto předfakturu aktivním spreadsheetem
  var novy_spreadSheet = SpreadsheetApp.openById(file_nova_predfaktura.getId());
  SpreadsheetApp.setActiveSpreadsheet(novy_spreadSheet); 
  var novy_sheet = novy_spreadSheet.getSheetByName("Proforma faktúra SK"); 
  var novy_sheet_dl = novy_spreadSheet.getSheetByName("Dodací list SK");
  var novy_sheet_e = novy_spreadSheet.getSheetByName("Export")  
  
  //zjistit fakturační adresu 1 a ostatní věci na řádku
  var spreadSheet5 = SpreadsheetApp.openById("[sheet id]");  
  var sheet5 = spreadSheet5.getSheetByName("Fakturačné adresy");
  //z cache zjistit řádek
  var radek_fakturacni_adresy = cache.get('radek_fakturacni_adresy');
  cache.remove('radek_fakturacni_adresy');
  var radek_fakturacni_adresyB = cache.get('radek_fakturacni_adresyB');
  cache.remove('radek_fakturacni_adresyB');  
  var adresa1_odberatel = sheet5.getRange("A" + radek_fakturacni_adresy).getValue();  
  var adresa1_adresa = sheet5.getRange("B" + radek_fakturacni_adresy).getValue();
  var adresa1_pscmesto = sheet5.getRange("C" + radek_fakturacni_adresy).getValue();
  var adresaB_odberatel = sheet5.getRange("A" + radek_fakturacni_adresyB).getValue();  
  var adresaB_adresa = sheet5.getRange("B" + radek_fakturacni_adresyB).getValue();
  var adresaB_pscmesto = sheet5.getRange("C" + radek_fakturacni_adresyB).getValue();  
  var adresa1_ico = sheet5.getRange("E" + radek_fakturacni_adresy).getValue();
  var adresa1_icdph = sheet5.getRange("F" + radek_fakturacni_adresy).getValue();
  var adresa1_dic = sheet5.getRange("G" + radek_fakturacni_adresy).getValue(); 
  //z cache zjistit, jestli jsou adresy A i B rovnake
  var su_adresy_rovnake = cache.get('su_rovnake');
  cache.remove('su_rovnake');   
  
  //datum + 14
  /*var den = new Date();   
  den.setDate(den.getDate()+14);
  var datum_plus_ctrnact = den.getDate() + "." + (den.getMonth() + 1) + "." + den.getYear();*/ 

  //podpis
  //var url_podpisu_slavka = "http://docs.google.com/uc?id=0B42i0tjrFPTGMXVUZTdfWWZUM0E&hl=en";
  //var url_podpisu_slavka = "https://doc-08-20-docs.googleusercontent.com/docs/securesc/nt7203jrrufuce5bcbir9gkicajkj5qv/oaat4uqklav9k6u3ndptlaherfdr7ukc/1356609600000/02005143052652800140/02005143052652800140/0B42i0tjrFPTGMXVUZTdfWWZUM0E";
  var url_podpisu_slavka = "http://www.w-script.com/wp-content/frg6539jrT6un/j8eTvfrR4.png";
  var podpis = url_podpisu_slavka;  
      
  //---!!!tady je samotné vyplnění předfaktury!!!---(při změně v šabloně je potřebné změnit souřadnice, kam se vepisují údaje)   
  if (su_adresy_rovnake == 'Áno') {
   novy_sheet.getRange("D6").setValue(adresa1_odberatel); //adresa 1 - odberatel
   novy_sheet.getRange("D7").setValue(adresa1_adresa); //adresa 1 - adresa  
   novy_sheet.getRange("D8").setValue(adresa1_pscmesto); //adresa 1 - psc+mesto

   novy_sheet.getRange("D14").setValue(adresa1_odberatel); //adresa 1 - odberatel
   novy_sheet.getRange("D17").setValue(adresa1_adresa); //adresa 1 - adresa  
   novy_sheet.getRange("D18").setValue(adresa1_pscmesto); //adresa 1 - psc+mesto    
  } else if (su_adresy_rovnake == 'Nie') {
   novy_sheet.getRange("D6").setValue(adresa1_odberatel); //adresa 1 - odberatel
   novy_sheet.getRange("D7").setValue(adresa1_adresa); //adresa 1 - adresa  
   novy_sheet.getRange("D8").setValue(adresa1_pscmesto); //adresa 1 - psc+mesto

   novy_sheet.getRange("D14").setValue(adresaB_odberatel); //adresa 1 - odberatel
   novy_sheet.getRange("D17").setValue(adresaB_adresa); //adresa 1 - adresa  
   novy_sheet.getRange("D18").setValue(adresaB_pscmesto); //adresa 1 - psc+mesto        
  }  
  
  novy_sheet.getRange("B2").setValue(cislo_posledni_faktury); //číslo předfaktury
  novy_sheet.getRange("B3").setValue(id_objednavky); //objednávka    
  novy_sheet.getRange("I21").setValue(datum_string); //datum_dnes (vystavení)
  novy_sheet.getRange("A57").setValue("Fakturu vystavil: " + vase_jmeno_prijmeni); //kdo vystavil     
  novy_sheet.getRange("I22").setValue(e.parameter.datum_splatnosti); //datum+14 (splatnosti) 
  novy_sheet.getRange("I23").setValue(e.parameter.datum_dodania_tovaru); //datum dodania tovar (dnes - upravené)
  novy_sheet.getRange("E10").setValue(adresa1_ico); //adresa 1 - ico
  novy_sheet.getRange("E11").setValue(adresa1_icdph); //adresa 1 - ic dph  
  novy_sheet.getRange("E12").setValue(adresa1_dic); //adresa 1 - dic 
  novy_sheet.insertImage(podpis, 3, 43, 0, 0); //vložit podpis + razítko
  novy_sheet_dl.insertImage(podpis, 1, 35, 0, 0); //vložit podpis + razítko
  novy_sheet_e.getRange("F2").setValue(objednavajuci); 
  novy_sheet_e.getRange("E2").setValue(predmet_objednavky); 
  //novy_sheet.getRange("E5").setValue(odberatel); //odberatel (z RD) - není třeba  
  
  //export do XLSX
  SpreadsheetApp.flush(); //uložit změny
  vyexportuj_xlsx("Work/2014/Proforma faktúry", file_nova_predfaktura.getId(), cislo_posledni_faktury + '_' + odberatel + '_' + predmet_objednavky);     
    
  // Clean up - get the UiInstance object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  // The following line is REQUIRED for the widget to actually close.
  return app;
}

//-------------------------------------------------------- PŘIDÁNÍ FAKTÚRY --------------------------------------------------------

function pridajFakturu() {
  formular_pridani_faktury();  
};
  
function formular_pridani_faktury(e) {
  var doc = SpreadsheetApp.openById("[sheet id]");
  var app = UiApp.createApplication().setTitle('Pridanie faktúry');
  // Grid 5 řádků dotazů ve formuláři x 2 sloupce (Label a TextBox) 
  var grid = app.createGrid(2, 2);
   
  grid.setWidget(0, 0, app.createLabel('Vložte číslo z RD:'));
  grid.setWidget(0, 1, app.createTextBox().setName('cislo_z_rd'));
  
  if (Session.getActiveUser() == "[e-mail address]") {
   var vypln_jmeno = "l";
  } else {
   var vypln_jmeno = "s";
  }  

  grid.setWidget(1, 0, app.createLabel('Vaše jméno (1 malé písmeno):'));
  grid.setWidget(1, 1, app.createTextBox().setName('vase_jmeno').setText(vypln_jmeno));  

  // Create a vertical panel..
  var panel = app.createVerticalPanel();

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a callback element and the handler as a click handler
  // Identify the function b as the server click handler

  var button = app.createButton('Pridať faktúru');
  var handler = app.createServerHandler('a');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application, then display the application app
  panel.add(button);
  app.add(panel);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}

//----------------------------------- PŘIDÁNÍ FAKTÚRY - vyplnění hodnot z formuláře -------------------------------------------------------

function a(e) {
 
  var cislo_z_rd = e.parameter.cislo_z_rd;
  
  //  otevřít 2014 RD, hárok "Číslo faktúry"
  //var spreadSheet = SpreadsheetApp.openById("[sheet id]");
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Číslo faktúry"); 
  
  // najít index posledního řádku
  var lastRowIndex = sheet.getLastRow();
 
  // zvýšit index o 1, takže ukazuje na nový řádek
  lastRowIndex++;
  
  // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
  // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear(); 
 
  // vyplnit datum
  sheet.getRange("A" + lastRowIndex).setValue(datum_string);  
  
  //zjistit poslední číslo faktury
  var cislo_posledni_faktury = sheet.getRange("B" + (lastRowIndex - 1)).getValue(); 
  if (cislo_posledni_faktury == "Číslo faktúry") {
   cislo_posledni_faktury = 140000;
  }
  
  // vyplnit číslo nové faktury
  sheet.getRange("B" + lastRowIndex).setValue(cislo_posledni_faktury+1); 
  
  var cislo_uple_nove_faktury = Number(cislo_posledni_faktury) + 1;
  
  //vyplnit uživatele
  if (e.parameter.vase_jmeno == "k") {
   vase_jmeno="[name]";
  }
  else if (e.parameter.vase_jmeno == "m") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "s") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "v") {
   vase_jmeno="[name]";
  }  
  else if (e.parameter.vase_jmeno == "l") {
   vase_jmeno="[name]";
  }   
  sheet.getRange("C" + lastRowIndex).setValue(vase_jmeno);   
  
  //vyplnit číslo z RD
  sheet.getRange("D" + lastRowIndex).setValue(cislo_z_rd);   

  //vyplnit do RD číslo faktury
  sheet = spreadSheet.getSheetByName("RD");
  radek_v_rd = Number(cislo_z_rd) + 1;
  var dosavadni_faktury = sheet.getRange("H" + radek_v_rd).getValue();
  if (dosavadni_faktury == "") {
   sheet.getRange("H" + radek_v_rd).setValue(cislo_uple_nove_faktury);   
  } else {
   sheet.getRange("H" + radek_v_rd).setValue(dosavadni_faktury + ", " + cislo_uple_nove_faktury);    
  } 

  //opodmínkování - jestliže je to zahraničí
  var je_to_tuzex = sheet.getRange("I" + radek_v_rd).getValue();
  if ( (je_to_tuzex == "mimo EU") || (je_to_tuzex == "EU") ) {
  
   //vyplnit do Zahraničie - platby + JCD číslo faktury
   var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = spreadSheet.getSheetByName("Zahraničie - platby + JCD");
   var lastRowIndex = sheet.getLastRow();
  
   //nalezení řádku, na kterém je cislo_z_rd 
   var bunka_s_cislem = sheet.getRange('P1:P' + lastRowIndex).getValues().findCells(e.parameter.cislo_z_rd);
   var radek_s_cislem = bunka_s_cislem[1,0][1,0];
  
   
   //načtení toho, co uz v buňce je
   var uz_tam_je = sheet.getRange("D" + radek_s_cislem).getValue();  
   //přidat k tomu číslo faktury
   if (uz_tam_je == "") {
    //sheet.getRange("D" + radek_s_cislem).setValue(cislo_posledni_faktury).setBackgroundColor("red");
    sheet.getRange("D" + radek_s_cislem).setValue(cislo_uple_nove_faktury);     
   } else {
    //sheet.getRange("D" + radek_s_cislem).setValue(uz_tam_je + ", " + cislo_posledni_faktury).setBackgroundColor("red");   
    sheet.getRange("D" + radek_s_cislem).setValue(uz_tam_je + ", " + cislo_uple_nove_faktury);       
   } 
   
  
  }  
    
  // Clean up - get the UiInstance object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  // The following line is REQUIRED for the widget to actually close.
  return app;
}

//--------------------------------------------------- PŘIDÁNÍ PREDFAKTÚRY ---------------------------------------------------------------

function pridajPredfakturu() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Číslo predfaktúry"); 
  
 //uživatel zadává číslo z RD
  formular_pridani_do_predfaktury();
  
};

function formular_pridani_do_predfaktury(e) {
  var doc = SpreadsheetApp.openById("[sheet id]");
  var app = UiApp.createApplication().setTitle('Pridanie predfaktúry');
  // Create a grid with 4 text boxes and corresponding labels (druh číslo je vždy o 1 menší
  var grid = app.createGrid(3, 2);
  
  // Text entered in the text box "Age" is passed in to userName
  // The setName method will make those widgets available by
  // the given name to the server handlers later  
  // Text entered in the text box "City" is passed in to age
   // Text entered in the next text box is passed in to city.
  grid.setWidget(0, 0, app.createLabel('Vložte číslo z RD:'));
  grid.setWidget(0, 1, app.createTextBox().setName('cislo_z_rd'));
  
  grid.setWidget(1, 0, app.createLabel('Kolik variant predfaktúr urobit:'));
  grid.setWidget(1, 1, app.createTextBox().setName('kolik_variant').setText('1'));

  if (Session.getActiveUser() == "[e-mail address]") {
   var vypln_jmeno = "l";
  } else {
   var vypln_jmeno = "s";
  }  
  
  grid.setWidget(2, 0, app.createLabel('Vaše jméno (1 malé písmeno):'));
  grid.setWidget(2, 1, app.createTextBox().setName('vase_jmeno').setText(vypln_jmeno));   

  // Create a vertical panel..
  var panel = app.createVerticalPanel();

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a callback element and the handler as a click handler
  // Identify the function b as the server click handler

  var button = app.createButton('Pridaj predfaktúru');
  var handler = app.createServerHandler('c');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application, then display the application app
  panel.add(button);
  app.add(panel);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);

}

//--------------------------------------- PŘIDÁNÍ PREDFAKTÚRY - vyplnění hodnot z formuláře ---------------------------------------------------

function c(e) {
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("Číslo predfaktúry");   
  var lastRowIndex = sheet.getLastRow();
  lastRowIndex++;
  
  // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
  // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear(); 
 
  // vyplnit datum
  sheet.getRange("A" + lastRowIndex).setValue(datum_string);   
  
  // získání proměnných z formuláře
  cislo_z_rd = e.parameter.cislo_z_rd;
  kolik_variant = e.parameter.kolik_variant;
  
  //vyplnění údaje do predfaktúry
  sheet.getRange("D" + lastRowIndex).setValue(cislo_z_rd); 

  //zjistit poslední číslo faktury
  var cislo_posledni_faktury = sheet.getRange("B" + (lastRowIndex - 1)).getValue(); 
  if (cislo_posledni_faktury == "Číslo predfaktúry") {
   cislo_posledni_faktury = "P00014";
  }  
 
  //zvětšení čísla poslední předfaktury o 1
  cislo_posledni_faktury = cislo_posledni_faktury.slice(1,4);
  cislo_posledni_faktury = Number(cislo_posledni_faktury) + 1;
  if (cislo_posledni_faktury < 10) {
   cislo_posledni_faktury = "P00" + cislo_posledni_faktury + "14";
  } else if (cislo_posledni_faktury > 99) {
   cislo_posledni_faktury = "P" + cislo_posledni_faktury + "14"; 
  } else {
   cislo_posledni_faktury = "P0" + cislo_posledni_faktury + "14";  
  }   

  //napodmínkování více variant předfaktury 
  if (kolik_variant == 2) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b";
  } else if (kolik_variant == 3) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b, " + cislo_posledni_faktury + "c";
  } else if (kolik_variant == 4) {
   cislo_posledni_faktury = cislo_posledni_faktury + "a, " + cislo_posledni_faktury + "b, " + cislo_posledni_faktury + "c, " + cislo_posledni_faktury + "d"; 
  }  
 
  // vyplnit číslo nové faktury
  sheet.getRange("B" + lastRowIndex).setValue(cislo_posledni_faktury);  
  
  //vyplnit uživatele
  if (e.parameter.vase_jmeno == "k") {
   vase_jmeno="[name]";
  }
  else if (e.parameter.vase_jmeno == "m") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "s") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "v") {
   vase_jmeno="[name]";
  }    
  else if (e.parameter.vase_jmeno == "l") {
   vase_jmeno="[name]";
  }   
  sheet.getRange("C" + lastRowIndex).setValue(vase_jmeno);    
  
  //načtení vyplněné nové předfaktury
  cislo_posledni_faktury = sheet.getRange("B" + lastRowIndex).getValue();
  
  //vyplnit do RD číslo předfaktury
  sheet = spreadSheet.getSheetByName("RD");
  radek_v_rd = Number(cislo_z_rd) + 1;
  var dosavadni_predfaktury = sheet.getRange("H" + radek_v_rd).getValue();
  if (dosavadni_predfaktury == "") {
   sheet.getRange("H" + radek_v_rd).setValue(cislo_posledni_faktury);   
  } else {
   sheet.getRange("H" + radek_v_rd).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury);    
  }
  
  //opodmínkování - jestliže je to zahraničí
  var je_to_tuzex = sheet.getRange("I" + radek_v_rd).getValue();
  if ( (je_to_tuzex == "mimo EU") || (je_to_tuzex == "EU") ) {
  
   //vyplnit do Zahraničie - platby + JCD číslo predaftury
   sheet = spreadSheet.getSheetByName("Zahraničie - platby + JCD");
   lastRowIndex = sheet.getLastRow();  
   //nalezení řádku, na kterém je cislo_z_rd 
   var bunka_s_cislem = sheet.getRange('P1:P' + lastRowIndex).getValues().findCells(e.parameter.cislo_z_rd);
   var radek_s_cislem = bunka_s_cislem[1,0][1,0];
   var dosavadni_predfaktury = sheet.getRange("C" + radek_s_cislem).getValue();
   if (dosavadni_predfaktury == "") {
    //sheet.getRange("C" + radek_s_cislem).setValue(cislo_posledni_faktury).setBackgroundColor("orange");  
    sheet.getRange("C" + radek_s_cislem).setValue(cislo_posledni_faktury);      
   } else {
    //sheet.getRange("C" + radek_s_cislem).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury).setBackgroundColor("orange");    
    sheet.getRange("C" + radek_s_cislem).setValue(dosavadni_predfaktury + ", " + cislo_posledni_faktury);       
   }    
    
  }  
    
  // Clean up - get the UiInstance object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  // The following line is REQUIRED for the widget to actually close.
  return app;
}

//----------------------------------------------------- PŘIDÁNÍ ZÁZNAMU DO RD ----------------------------------------------------------------

function Zaregistruj() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("RD");
  formular_pridani_do_rd();
    
};
  
function formular_pridani_do_rd(e) {
  var doc = SpreadsheetApp.openById("[sheet id]");
  var app = UiApp.createApplication().setTitle('Zaregistrovanie do RD');
  // Grid 6 řádků dotazů ve formuláři x 2 sloupce (Label a TextBox) 
  var grid = app.createGrid(6, 3);
  
  // Text entered in the text box "Age" is passed in to userName
  // The setName method will make those widgets available by
  // the given name to the server handlers later  
  // Text entered in the text box "City" is passed in to age
   // Text entered in the next text box is passed in to city.
  grid.setWidget(0, 0, app.createLabel('Odberatel:'));
  grid.setWidget(0, 1, app.createTextBox().setName('odberatel'));
  
  grid.setWidget(1, 0, app.createLabel('Objednávajúci:'));
  grid.setWidget(1, 1, app.createTextBox().setName('objednavajuci'));

  grid.setWidget(2, 0, app.createLabel('ID objednávky:'));
  grid.setWidget(2, 1, app.createTextBox().setName('id_objednavky'));
  
  grid.setWidget(3, 0, app.createLabel('Predmet objednávky:'));
  grid.setWidget(3, 1, app.createTextBox().setName('predmet_objednavky'));
  
  var radio1 = app.createRadioButton('group1', 'EU').setName('zahranicny_klient_eu').setId('zahranicny_klient_eu');
  var radio2 = app.createRadioButton('group1', 'Mimo EU').setName('zahranicny_klient_mimo_eu').setId('zahranicny_klient_mimo_eu');
  var radio3 = app.createRadioButton('group1', 'Slovenský zákazník').setName('slovensky_klient').setId('slovensky_klient');
  
  grid.setWidget(4, 0, radio1);
  grid.setWidget(4, 1, radio2);  
  grid.setWidget(4, 2, radio3); 
 
  if (Session.getActiveUser() == "[e-mail address]") {
   var vypln_jmeno = "l";
  } else {
   var vypln_jmeno = "s";
  }  
  
  grid.setWidget(5, 0, app.createLabel('Vaše jméno (1 malé písmeno):'));
  grid.setWidget(5, 1, app.createTextBox().setName('vase_jmeno').setText(vypln_jmeno)); 
  
  // Create a vertical panel..
  var panel = app.createVerticalPanel();

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a callback element and the handler as a click handler
  // Identify the function b as the server click handler

  var button = app.createButton('Zaregistrovat do RD');
  var handler = app.createServerHandler('b');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);
  
  //handler, který vylučuje označení dvou radiobuttonů najednou
  //radio 1
  var handler1 = app.createServerValueChangeHandler('showstatus1');
  handler1.addCallbackElement(grid);
  radio1.addValueChangeHandler(handler1);
  //radio2
  var handler2 = app.createServerValueChangeHandler('showstatus2');
  handler2.addCallbackElement(grid);
  radio2.addValueChangeHandler(handler2);  
  //radio3
  var handler3 = app.createServerValueChangeHandler('showstatus3');
  handler3.addCallbackElement(grid);
  radio3.addValueChangeHandler(handler3);   

  // Add the button to the panel and the panel to the application, then display the application app
  panel.add(button);
  app.add(panel);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}

//funkce, které vylučují označení dvou radiobuttonů najednou
function showstatus1(e){
  var app = UiApp.getActiveApplication();
  var radioValue = e.parameter.zahranicny_klient_eu;
  app.getElementById('zahranicny_klient_mimo_eu').setValue(false);
  app.getElementById('slovensky_klient').setValue(false);  
  return app;
}
function showstatus2(e){
  var app = UiApp.getActiveApplication();
  var radioValue = e.parameter.zahranicny_klient_mimo_eu;
  app.getElementById('zahranicny_klient_eu').setValue(false);
  app.getElementById('slovensky_klient').setValue(false);  
  return app;
}
function showstatus3(e){
  var app = UiApp.getActiveApplication();
  var radioValue = e.parameter.slovensky_klient;
  app.getElementById('zahranicny_klient_eu').setValue(false);
  app.getElementById('zahranicny_klient_mimo_eu').setValue(false);  
  return app;
}

//----------------------------------- zpracování formuláře - PŘIDÁNÍ ZÁZNAMU DO RD -------------------------------------------------------

// Vyplněné hodnoty z formuláře uloží do proměnných
function b(e) {
  
  //otevřít hárok RD, najít poslední řádek
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName("RD");   
  var lastRowIndex = sheet.getLastRow();
  lastRowIndex++;
  var cislo_z_rd = lastRowIndex - 1;

  // zjistit dnešní datum
  var today = new Date();
  today = new Date(today.getYear(), today.getMonth(), today.getDate(), 0); 
  
  // konvertovat datum na string
  var datum_string = today.getDate() + "." + (today.getMonth() + 1) + "." + today.getYear(); 
 
  // vyplnit datum
  sheet.getRange("C" + lastRowIndex).setValue(datum_string);  
  
  //zjistit poslední číslo RD
  var cislo_posledni_rd = sheet.getRange("A" + (lastRowIndex - 1)).getValue(); 
  if (cislo_posledni_rd == "Číslo RD") {
   cislo_posledni_rd = 0;
  }
  
  // vyplnit číslo nové faktury
  sheet.getRange("A" + lastRowIndex).setValue(cislo_posledni_rd+1); 
  
  //vyplnění podle toho odkud je
  if (e.parameter.zahranicny_klient_eu == 'true') {
   sheet.getRange("I" + lastRowIndex).setValue("EU"); 
   var je_ze_zahranici = true;
  } else  if (e.parameter.zahranicny_klient_mimo_eu == 'true') {
   sheet.getRange("I" + lastRowIndex).setValue("mimo EU"); 
   var je_ze_zahranici = true; 
  } else  if (e.parameter.slovensky_klient == 'true') {  
   sheet.getRange("I" + lastRowIndex).setValue("Slovensko"); 
   var je_ze_zahranici = false;  
  }  
  
  //vyplnit uživatele
  if (e.parameter.vase_jmeno == "k") {
   vase_jmeno="[name]";
  }
  else if (e.parameter.vase_jmeno == "m") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "s") {
   vase_jmeno="[name]";
  } 
  else if (e.parameter.vase_jmeno == "v") {
   vase_jmeno="[name]";
  }    
  else if (e.parameter.vase_jmeno == "l") {
   vase_jmeno="[name]";
  }    
  sheet.getRange("B" + lastRowIndex).setValue(vase_jmeno);   
  
  //vyplnění údajú do RD
  sheet.getRange("D" + lastRowIndex).setValue(e.parameter.odberatel); 
  sheet.getRange("E" + lastRowIndex).setValue(e.parameter.objednavajuci); 
  sheet.getRange("F" + lastRowIndex).setValue(e.parameter.id_objednavky);   
  sheet.getRange("G" + lastRowIndex).setValue(e.parameter.predmet_objednavky);  
  
  //pokud je zahraniční, zapsat do hárku Zahraničie - platby + JCD
  if (je_ze_zahranici == true) {
  
   //otevřít hárok Zahraničie - platby + JCD, najít poslední řádek
   sheet = spreadSheet.getSheetByName("Zahraničie - platby + JCD");   
   lastRowIndex = sheet.getLastRow();  
   lastRowIndex++;
  
   //vyplnění údajů do Zahraničie - platby + JCD
   sheet.getRange("A" + lastRowIndex).setValue(e.parameter.odberatel);  
   sheet.getRange("B" + lastRowIndex).setValue(e.parameter.predmet_objednavky);   
   if (e.parameter.zahranicny_klient_eu == 'true') {
    sheet.getRange("K" + lastRowIndex).setValue("EU"); 
   } else  if (e.parameter.zahranicny_klient_mimo_eu == 'true') {
    sheet.getRange("K" + lastRowIndex).setValue("mimo EU"); 
   } 
   sheet.getRange("P" + lastRowIndex).setValue(cislo_z_rd); 
    
  }    
  
  // Clean up - get the UiInstance object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  // The following line is REQUIRED for the widget to actually close.
  return app;
}

//---------------------------------------------- NÁPOVĚDA ------------------------------------------------
/*

HANDLER, KTERÝ FUNGUJE NA ZÁKLADĚ ODENTROVÁNÍ (ODEJITÍ)

var handler1 = app.createServerChangeHandler('show_odberatel1'); //co se má spustit
handler1.addCallbackElement(panel);
vlozene_cislo_z_rd.addChangeHandler(handler1); //čím se to spouští

VÝPIS DATUMU Z OBJEKTU - nefunguje mi to, nevím proč (špatně píše měsíce)

var vysledek = Utilities.formatDate(den, "PST", "dd.mm.yyyy");

JAK NAJÍT POSLEDNÍ PLNOU BUŇKU VE SLOUPCI

var lastRow = doc.getLastRow(); // Determine the last row in the Spreadsheet that contains any values
var cell = doc.getRange('a1').offset(lastRow, 0); // determine the next free cell in column A

JAK LOGOVAT OBSAH PROMĚNNÝCH

Logger.log(promenna); 

FUNKCE, KTERÁ DO OKÉNKA (WIDGETU) PŘIDÁ HYPERLINK, TAKY TAM JE VELIKOST OKÉNKA

function test(){
showURL("http://www.google.com")
}
//
function showURL(href){
  var app = UiApp.createApplication().setHeight(50).setWidth(200);
  app.setTitle("Show URL");
  //var link = app.createAnchor('open ', href).setId("link").setTarget("_self");
  app.add(link);  
  var doc = SpreadsheetApp.getActive();
  doc.show(app);
  }

VYČERPANÉ JEDNOPÍSMENNÉ NÁZVY FUNKCÍ

a,b,c,d

OMEZENÍ

Session.getActiveUser() - funguje jen při Google Apps for Business (musí se platit)
.findCells() - vadí tomu message boxy, s nimi to nefunguje
nefunguje export do xls - var xls = SpreadsheetApp.openById(idspreadsheetu).getAs("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); //PRO XLSX: ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), PRO XLS: ("application/vnd.ms-excel")

//---------------------------------------------- NEOPRAVITELNÉ CHYBY ------------------------------------------------
1) Google Docs neumí tlusté čáry, při jejich náhradě vybarvenými buňkami se zase rozbije export do XLSX
2) podpis v šabloně se při exportu do XLSX nastaví jako unchecked "Print object" (nutno klik pravým v Excelu 2003 na obrázek, Formátovať obrázok>Vlastnosti>Tlač objektu 

//---------------------------------------------- DROBNÁ VYLEPŠENÍ PO DELŠÍM TESTOVÁNÍ ------------------------------------------------
-1) když je to EU tak do JCD dát čárku
0) odbarvení při přidání datumu zaplacení
1) jméno l/s/m z rozbalovacího seznamu
2) na výběr konst. symbol: produkty (defaultně) 0008, služby 0308
3) kurierská spol. - čím to išlo - do faktury (KAM DÁT?)
4) vkládání obrázků zvlášť do Google dokumentu a zvlášť do XLSX exportu
5) víc variant předfaktur (nakopírovat, přejmenovat na a, b, c, … - primitivní, variable je uvnitř funkce) (u pridaj + vystav predfakturu)
6) různé podpisy pro různé uživatele

*/



