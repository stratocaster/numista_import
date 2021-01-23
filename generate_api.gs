//@NotOnlyCurrentDoc

function myAPI2 () {

  Logger.clear();
  
  var NL = String.fromCharCode(13);
  var QM = String.fromCharCode(34);
  
  var myApp = SpreadsheetApp;
  
  var mySpread = myApp.getActiveSpreadsheet();
  
  // SET SHEETS
  //var mySheet = mySpread.getActiveSheet();
  var myImportSheet = mySpread.getSheetByName("query");
  var myAPIsheet = mySpread.getSheetByName("API_code");
  var myLinksSheet = mySpread.getSheetByName("Atom_Links");
  var myTwinImgSheet = mySpread.getSheetByName("Twin_Images");
  
  //***
  //CATALOG ID this is once per RUN!! &&& ISSUER
var myIssuer = "rome";
    var myCatID = "";
    var myTemp = "RIC IV.1";   //SET ME HERE********
    switch (myTemp)  {  
        case "RIC I":    //-31-69
            myCatID = 448;
            break;
        case "RIC II.1":   //69–96
            myCatID = 449;
            break;
        case "RIC II.2":   //96-117
            myCatID = 1418;
            break;
        case "RIC II.3":   //117–138 
            myCatID = 986;
            break;
        case "RIC II":   //96–138 
            myCatID = 305;
            break;
        case "RIC III":   //96–138 
            myCatID = 450;
            break;            
        case "RIC IV.1":   //193–217
            myCatID = 451;
            break;
        case "RIC IV.2":   //217–238
            myCatID = 452;
            break;
        case "RIC IV.3":   //238–253
            myCatID = 453;
            break;
        case "RIC V.1":   //253–276
            myCatID = 454;
            break;
        case "RIC V.2":   //276–310
            myCatID = 457;
            break;
        case "RIC VI":   //294-313
            myCatID = 458;
            break;
        case "RIC VII":   //313–337
            myCatID = 459;
            break;
        case "RIC VIII":   //337–364
            myCatID = 460;
            break;
        case "RIC IX":   //364–395
            myCatID = 461;
            break;
        case "RIC X":   //395–491
            myCatID = 462;
            break;
    }
  
  //________________________________________
  // SET RANGE OF ROWS TO RUN PARSER HERE:
  //________________________________________
 
  for(var i=40; i<=41; i++){   //SET ME HERE********
    
  //COMPUTE Start and end year
    var myStartYear = "";
    var myEndYear = "";
    var myYearsArray = [{}];
    myYearsArray = myImportSheet.getRange(i,26).getValue().toString().split("|");
    myStartYear = myYearsArray[0];     
      if (myYearsArray.length >= 2){
        myEndYear = myYearsArray[1];
      }
        else{
          myEndYear = myYearsArray[0];
        }
    
    //COMPUTE Image links
    var myOVlink = myTwinImgSheet.getRange(i,9).getValue();
    var myRVlink = myTwinImgSheet.getRange(i,10).getValue();
    var myImgCredits = myTwinImgSheet.getRange(i,6).getValue();
    var myCreditType = myTwinImgSheet.getRange(i,7).getValue();
    var myImgLink = myTwinImgSheet.getRange(i,3).getValue();
      
    //COMPUTE Mass Diameter
    var myTmass = myImportSheet.getRange(i,36).getValue();
    var myTdiam = myImportSheet.getRange(i,37).getValue();
    var myMassComment = myImportSheet.getRange(i,38).getValue();
    var myDiamComment = myImportSheet.getRange(i,39).getValue();  
    
    //IMPORT Deity, Ruler, Legends, descriptions
    var myDeity = myImportSheet.getRange(i,6).getValue();
    myDeity = myDeity.toString();
    var myDeityFr = myImportSheet.getRange(i,10).getValue().toString();
    //replace | EN
    if ((myDeity.length - myDeity.replace(/\|/g,'').length)>1){
    while ((myDeity.length - myDeity.replace(/\|/g,'').length)>1){
      myDeity = myDeity.replace('|',", ");
    }
    myDeity = myDeity.replace('|',", and ");
    }
    else {
      myDeity = myDeity.replace('|'," and ");
    }
    //replace | FR
    if ((myDeityFr.length - myDeityFr.replace(/\|/g,'').length)>1){
    while ((myDeityFr.length - myDeityFr.replace(/\|/g,'').length)>1){
      myDeityFr = myDeityFr.replace('|',", ");
    }
    myDeityFr = myDeityFr.replace('|'," et ");
    }
    else {
      myDeityFr = myDeityFr.replace('|'," et ");
    }
    //Descriptions
    var myOvDescription = myImportSheet.getRange(i,17).getValue().toString();
    if (myOvDescription[myOvDescription.length-1] !=="."){
      myOvDescription = myOvDescription + "."
    }
    var myRvDescription = myImportSheet.getRange(i,23).getValue();
    if (myRvDescription[myRvDescription.length-1] !=="."){
      myRvDescription = myRvDescription + "."
    }
    //Portrait
    //EN
    var myPortraitEn = myImportSheet.getRange(i,19).getValue().toString();
    var myPortraitFr = myImportSheet.getRange(i,20).getValue().toString();
    //replace | En
    if ((myPortraitEn.length - myPortraitEn.replace(/\|/g,'').length)>1){
    while ((myPortraitEn.length - myPortraitEn.replace(/\|/g,'').length)>1){
      myPortraitEn = myPortraitEn.replace('|',", ");
    }
    myPortraitEn = myPortraitEn.replace('|',", and ");
    }
    else {
      myPortraitEn = myPortraitEn.replace('|'," and ");
    }
   
    //replace | Fr
    if ((myPortraitFr.length - myPortraitFr.replace(/\|/g,'').length)>1){
    while ((myPortraitFr.length - myPortraitFr.replace(/\|/g,'').length)>1){
      myPortraitFr = myPortraitFr.replace('|',", ");
    }
    myPortraitFr = myPortraitFr.replace('|'," et ");
    }
    else {
      myPortraitFr = myPortraitFr.replace('|'," et ");
    }
    //AXIS
    var myAxis = myImportSheet.getRange(i,40).getValue();
    //MINT
     var myMint = myImportSheet.getRange(i,32).getValue();
    //RIC NO
    var myRICno = myImportSheet.getRange(i,33).getValue();
    //Legends
    var myOvTransEn = myImportSheet.getRange(i,28).getValue().toString();
    var myRvTransEn = myImportSheet.getRange(i,30).getValue().toString();
    var myOvTransFr = myImportSheet.getRange(i,29).getValue().toString();
    var myRvTransFr = myImportSheet.getRange(i,31).getValue().toString();
    var myRvLegend = myImportSheet.getRange(i,22).getValue().toString();
    var myOvLegend = myImportSheet.getRange(i,16).getValue().toString(); 
    //replace empty strings
    if (myOvTransEn.length <= 15) {
      myOvTransEn ="";
    }
    else if ((myOvTransEn.indexOf(' + chr(10) + ') == 1)||(myOvTransEn.indexOf(' + chr(10) + ') == myOvTransEn.length-14)){
      myOvTransEn = myOvTransEn.replace(/" \+ chr\(10\) \+ "/g,'');
    }
    if (myRvTransEn.length <= 15) {
      myRvTransEn ="";
    }
    else if ((myRvTransEn.indexOf(' + chr(10) + ') == 1)||(myRvTransEn.indexOf(' + chr(10) + ') == myRvTransEn.length-14)){
      myRvTransEn = myRvTransEn.replace(/" \+ chr\(10\) \+ "/g,'');
    }
    if (myOvTransFr.length <= 15) {
      myOvTransFr ="";
    }
    else if ((myOvTransFr.indexOf(' + chr(10) + ') == 1)||(myOvTransFr.indexOf(' + chr(10) + ') == myOvTransFr.length-14)){
      myOvTransFr = myOvTransFr.replace(/" \+ chr\(10\) \+ "/g,'');
    }
    if (myRvTransFr.length <= 15) {
      myRvTransEn ="";
    }
    else if ((myRvTransFr.indexOf(' + chr(10) + ') == 1)||(myRvTransFr.indexOf(' + chr(10) + ') == myRvTransFr.length-14)){
      myRvTransFr = myRvTransFr.replace(/" \+ chr\(10\) \+ "/g,'');
    }    
      
    //COMPUTE Metal ID
    /////////////////////////////////
    var myMetalID = "";
    var myTemp = myImportSheet.getRange(i,14).getValue();
    switch (myTemp)  {  
        case "Silver":
            myMetalID = 1;
            break;
        case "Bronze":
            myMetalID = 5;
            break;
        case "Gold":
            myMetalID = 6;
            break;
        case "Copper":
            myMetalID = 3;
            break;
        case "Billon":
            myMetalID = 7;
            break;
        case "Electrum":
            myMetalID = 17;
            break;
        case "Brass":
            myMetalID = 4;
            break;
        case "Potin":
            myMetalID = 33;
            break;        
        case "Aluminium":
            myMetalID = 45;
            break;
        case "Iron":
            myMetalID = 13;
            break;
        case "Lead":
            myMetalID = 21;
            break;
        case "Nickel":
            myMetalID = 8;
            break;
        case "Orichalcum":
            myMetalID = 52;
            break;
        case "Pewter":
            myMetalID = 25;
            break;
        case "Tin":
            myMetalID = 19;
            break;
        case "Tombac":
            myMetalID = 23;
            break;
        case "Zinc":
            myMetalID = 11;
            break;
    }
    
    //COMPUTE Ruler ID
            ////////////////////////////******************
    var myRulerID = "";
    var myTemp = myImportSheet.getRange(i,4).getValue().toString().split("|")[0];
    switch (myTemp)  {
//        case "Anonymous":
//            myRulerID = 4304;
//            break;
//        case "Clodius Macer":
//            myRulerID = 2791;
//            break;
//        case "Elagabalus":
//            myRulerID = 86;
//            break;
//        case "Augustus":
//            myRulerID = 61;
//            break;
//        case "Tiberius":
//            myRulerID = 62;
//            break;
//        case "Gaius/Caligula":
//            myRulerID = 63;
//            break;
//        case "Claudius":
//            myRulerID = 64;
//            break;
//        case "Nero":
//            myRulerID = 65;
//            break;
//        case "Galba":
//            myRulerID = 66;
//            break;
//        case "Otho":
//            myRulerID = 67;
//            break;
//        case "Vitellius":
//            myRulerID = 68;
//            break;
//        case "Vespasian":
//            myRulerID = 69;
//            break;
//        case "Titus":
//            myRulerID = 70;
//            break;
//        case "Domitian":
//            myRulerID = 71;
//            break;
//        case "Nerva":
//            myRulerID = 72;
//            break;
//        case "Trajan":
//            myRulerID = 73;
//            break;
//        case "Hadrian":
//            myRulerID = 74;
//            break;
//        case "Anonymous":
//            myRulerID = 4462;
//            break;
//        case "Antoninus Pius":
//            myRulerID = 75;
//            break;
//        case "Lucius Verus":
//            myRulerID = 76;
//            break;
//        case "Marcus Aurelius":
//            myRulerID = 77;
//            break;
// case "Commodus":
//     myRulerID = 78;
//     break;
      //  case "Pertinax":
      //      myRulerID = 79;
      //      break;
       case "Didius Julianus":
           myRulerID = 80;
           break;
//        case "Septimius Severus":
//            myRulerID = 81;
//            break;
//        case "Caracalla":
//            myRulerID = 82;
//            break;
//        case "Geta":
//            myRulerID = 83;
//            break;
//        case "Macrinus":
//            myRulerID = 84;
//            break;
//        case "Diadumenian":
//            myRulerID = 85;
//            break;
//        case "Severus Alexander":
//            myRulerID = 87;
//            break;
//        case "Gordian I":
//            myRulerID = 89;
//            break;
//        case "Gordian II":
//            myRulerID = 90;
//            break;
//        case "Pupienus":
//            myRulerID = 91;
//            break;
//        case "Balbinus":
//            myRulerID = 92;
//            break;
//        case "Gordian III":
//            myRulerID = 93;
//            break;
//        case "Philip the Arab":
//            myRulerID = 94;
//            break;
//        case "Herennius Etruscus":
//            myRulerID = 97;
//            break;
//        case "Aemilian":
//            myRulerID = 98;
//            break;
//        case "Volusian":
//            myRulerID = 99;
//            break;
//        case "Trebonianus Gallus":
//            myRulerID = 100;
//            break;
//        case "Hostilian":
//            myRulerID = 101;
//            break;
//        case "Valerian":
//            myRulerID = 102;
//            break;
//        case "Gallienus":
//            myRulerID = 103;
//            break;
//        case "Saloninus":
//            myRulerID = 104;
//            break;
//        case "Claudius II Gothicus":
//            myRulerID = 105;
//            break;
//        case "Quintillus":
//            myRulerID = 106;
//            break;
//        case "Tacitus":
//            myRulerID = 108;
//            break;
//        case "Probus":
//            myRulerID = 110;
//            break;
//        case "Numerian":
//            myRulerID = 112;
//            break;
//        case "Carinus":
//            myRulerID = 113;
//            break;
//        case "Diocletian":
//            myRulerID = 114;
//            break;
//        case "Maximian":
//            myRulerID = 115;
//            break;
//        case "Galerius":
//            myRulerID = 117;
//            break;
//        case "Constantine I":
//            myRulerID = 119;
//            break;
//        case "Licinius":
//            myRulerID = 121;
//            break;
//        case "Constantine II":
//            myRulerID = 122;
//            break;
//        case "Constans":
//            myRulerID = 123;
//            break;
//        case "Constantius II":
//            myRulerID = 124;
//            break;
//        case "Julian the Apostate":
//            myRulerID = 125;
//            break;
//        case "Jovianus":
//            myRulerID = 126;
//            break;
//        case "Valentinian I":
//            myRulerID = 127;
//            break;
//        case "Valens":
//            myRulerID = 128;
//            break;
//        case "Gratian":
//            myRulerID = 129;
//            break;
//        case "Valentinian II":
//            myRulerID = 130;
//            break;
//        case "Theodosius I":
//            myRulerID = 131;
//            break;
//        case "Arcadius":
//            myRulerID = 133;
//            break;
//        case "Honorius":
//            myRulerID = 134;
//            break;
//        case "Theodosius II":
//            myRulerID = 135;
//            break;
//        case "Constantius III":
//            myRulerID = 136;
//            break;
//        case "Valentinian III":
//            myRulerID = 137;
//            break;
//        case "Marcian":
//            myRulerID = 138;
//            break;
//        case "Postumus":
//            myRulerID = 1425;
//            break;
//        case "Tetricus I":
//            myRulerID = 1446;
//            break;
//        case "Victorinus":
//            myRulerID = 1447;
//            break;
//        case "Allectus":
//            myRulerID = 1448;
//            break;
//        case "Carausius":
//            myRulerID = 1449;
//            break;
//        case "Maxentius":
//            myRulerID = 1451;
//            break;
//        case "Vetranio":
//            myRulerID = 1495;
//            break;
//        case "Zeno":
//            myRulerID = 1529;
//            break;
//        case "Leo I":
//            myRulerID = 1530;
//            break;
//        case "Julius Nepos":
//            myRulerID = 1531;
//            break;
//        case "Anthemius":
//            myRulerID = 1533;
//            break;
//        case "Constantine III":
//            myRulerID = 1534;
//            break;
//        case "Petronius Maximus":
//            myRulerID = 1535;
//            break;
//        case "Avitus":
//            myRulerID = 1536;
//            break;
//        case "Crispus":
//            myRulerID = 1542;
//            break;
//        case "Constantius Gallus":
//            myRulerID = 1543;
//            break;
//        case "Justinian I":
//            myRulerID = 2139;
//            break;
//        case "Constans II":
//            myRulerID = 2222;
//            break;
//        case "Macrianus Minor":
//            myRulerID = 2795;
//            break;
//        case "Valerius Valens":
//            myRulerID = 2799;
//            break;
//        case "Pescennius Niger":
//            myRulerID = 2809;
//            break;
    }
  
    
    //COMPUTE CURRENCY
    //1662 - Denarius, Reform of Augustus (27 BC - AD 215)
    //1615 - Antoninianus, Reform of Caracalla (AD 215 - 301)
    //1618 - Argenteus, Reform of Diocletian (AD 293/301 - 310/324)
    //1619 - Solidus, Reform of Constantine (AD 310/324 - 395)

    
    var myCurrencyID = 1662;
    if (myStartYear>=310){
        myCurrencyID = 1619;
    }
    else if (myStartYear>=293){
        myCurrencyID = 1618;
    }
    else if (myStartYear>=215){
        myCurrencyID = 1615;
    }
    else if (myStartYear>=-27){
        myCurrencyID = 1662;
    }
//    myCurrencyID = 9372; //carthage usurpations
      
    //COMPUTE Values / DENOMINATIONS
    //1662 - Denarius, Reform of Augustus (27 BC - AD 215) 
      //// 1 Aureus = 2 Gold Quinarii = 25 Denarii 
      //// 1 Denarius = 2 Silver Quinarii = 4 Sestertii = 8 Dupondii = 16 Asses 
      //// 1 As = 2 Semisses = 4 Quadrantes
      var myImportValue = myImportSheet.getRange(i,7).getValue();
      var myImportValueFr = myImportValue;   
      var myFractionalFlag = true;
      var myNumerator = 1;
      var myIntegerValue = 1;
      var myDenominator = 1;
      var myValueTextEn = "";
      var myValueTextFr = "";
      if (myCurrencyID == 1662){
        switch (myImportValue) {
        case "As":
            myNumerator = 1;
            myDenominator = 16;
            myValueTextEn = "As = 1⁄16 Denarius";
            myValueTextFr = "as = 1⁄16 denier";
            break;    
        case "Aureus":
            myFractionalFlag = false;
            myIntegerValue = 25;
            myValueTextEn = "Aureus = 25 Denarii";
            myValueTextFr = "aureus = 25 deniers";
            break;
        case "Denarius":
            myFractionalFlag = false;
            myIntegerValue = 1;
            myValueTextEn = "Denarius";
            myValueTextFr = "denier";
            myImportValueFr = "Denier";
            break;
        case "Dupondius":
            myNumerator = 1;
            myDenominator = 8;
            myValueTextEn = "Dupondius = 1⁄8 Denarius"
            myValueTextFr = "dupondius = 1⁄8 denier"
            break;
        case "Quinarius":
            myNumerator = 1;
            myDenominator = 2;
            myValueTextEn = "Silver Quinarius = 1⁄2 Denarius"
            myValueTextFr = "quinaire d'argent = 1⁄2 denier"
            myImportValueFr = "Quinaire";
            break;
        case "Sestertius":
            myNumerator = 1;
            myDenominator = 4;
            myValueTextEn = "Sestertius = 1⁄4 Denarius"
            myValueTextFr = "sesterce = 1⁄4 denier"
            myImportValueFr = "Sesterce";
            break;
        case "Semis":
            myNumerator = 1;
            myDenominator = 32;
            myValueTextEn = "Semis = 1⁄2 As = 1⁄32 Denarius"
            myValueTextFr = "sesterce = 1⁄2 as = 1⁄32 denier"
            break;
        case "Quadrans":
            myNumerator = 1;
            myDenominator = 64;
            myValueTextEn = "Quadrans = 1⁄4 As = 1⁄64 Denarius"
            myValueTextFr = "quadrans = 1⁄4 as = 1⁄64 denier" 
            break;            
        case "Drachma":
            myFractionalFlag = false;
            myIntegerValue = 1;
            myValueTextEn = "Drachm = 1 Denarius";
            myValueTextFr = "drachme = 1 denier";
            myImportValue = "Drachm";
            myImportValueFr = "Drachme";
            break;            
        case "hemidrachm":
            myNumerator = 1;
            myDenominator = 2;
            myValueTextEn = "Hemidrachm = 1⁄2 Drachm = 1⁄2 Denarius";
            myValueTextFr = "hémidrachme = 1⁄2 drachme = 1⁄2 denier";
            myImportValue = "Hemidrachm";
            myImportValueFr = "Hémidrachme";
            break;
        case "Didrachm":
            myFractionalFlag = false;
            myIntegerValue = 2;
            myValueTextEn = "Didrachm = 2 Drachms = 2 Denarii"
            myValueTextFr = "didrachme = 2 drachmes = 2 deniers";
            myImportValueFr = "Didrachme";
            break;        
        case "12 As":
            myNumerator = 3;
            myDenominator = 4;
            myValueTextEn = "12 Assēs = 3⁄4 Denarius"
            myValueTextFr = "12 as = 3⁄4 denier"
            myImportValue = "12 Assēs";
            myImportValueFr = "12 Assēs";
            break;
        case "24 As":
            myNumerator = 3;
            myDenominator = 2;
            myValueTextEn = "24 Assēs = 1​1⁄2 Denarii"
            myValueTextFr = "24 as = 1​1⁄2 deniers"
            myImportValue = "24 Assēs";
            myImportValueFr = "24 Assēs";
            break;
        case "Quinarius aureus":
            myNumerator = 25;
            myDenominator = 2;
            myValueTextEn = "Gold Quinarius = 12​1⁄2 Denarii"
            myValueTextFr = "quinaire d'or = 12​1⁄2 deniers"
            myImportValue = "Quinarius Aureus";
            myImportValueFr = "Quinaire d'or";
            break;
        case "Dupondius|As":
            myNumerator = 1;
            myDenominator = 8;
            myValueTextEn = "Uncertain (Dupondius = 1⁄8 Denarius or As = 1⁄16 Denarius)"
            myValueTextFr = "Incertain (dupondius = 1⁄8 denier ou as = 1⁄16 denier)"            
            myImportValue = "Dupondius or As";
            myImportValueFr = "Dupondius ou as";            
            break;
        case "Cistophorus":
            myFractionalFlag = false;
            myIntegerValue = 3;
            myValueTextEn = "Cistophorus = 3 Drachms = 3 Denarii"
            myValueTextFr = "Cistophorus = 3 drachmes = 3 deniers";
            myImportValueFr = "Cistophore";            
            break;
        case "AE2":
            myValueTextEn = "Uncertain bronze";
            myValueTextFr = "Bronze incertain";
            myImportValueFr = "Æ2";            
            myImportValue = "Æ2"; 
            break;
        case "4 Aureus":
            myFractionalFlag = false;
            myIntegerValue = 100;            
            myValueTextEn = "4 Aurei = 100 Denarii";
            myValueTextFr = "4 aurei = 100 deniers";
            myImportValueFr = "4 Aurei";            
            myImportValue = "4 Aurei";
            break;
        case "5 Aureus":
            myFractionalFlag = false;
            myIntegerValue = 120;            
            myValueTextEn = "5 Aurei = 125 Denarii";
            myValueTextFr = "5 aurei = 125 deniers";
            myImportValueFr = "5 Aurei";            
            myImportValue = "5 Aurei";
            break;
        case "4 Denarius":
            myFractionalFlag = false;
            myIntegerValue = 4;            
            myValueTextEn = "4 Denarii";
            myValueTextFr = "4 deniers";
            myImportValueFr = "4 Denarii";            
            myImportValue = "4 Deniers";            
            break;
        case "5 Denarius":
            myFractionalFlag = false;
            myIntegerValue = 5;            
            myValueTextEn = "5 Denarii";
            myValueTextFr = "5 deniers";
            myImportValueFr = "5 Denarii";            
            myImportValue = "5 Deniers";            
            break; 
        case "8 Denarius":
            myFractionalFlag = false;
            myIntegerValue = 8;            
            myValueTextEn = "8 Denarii";
            myValueTextFr = "8 deniers";
            myImportValueFr = "8 Denarii";            
            myImportValue = "8 Deniers";            
            break;             
        case "AE Half-unit":           
            myValueTextEn = "Uncertain bronze half-unit";
            myValueTextFr = "Demi-unité de bronze incertaine";
            myImportValueFr = "Æ Half-Unit";            
            myImportValue = "Æ Demi-unité";
        case "AE Unit":           
            myValueTextEn = "Uncertain bronze unit";
            myValueTextFr = "Unité de bronze incertaine";
            myImportValueFr = "Æ Unit";            
            myImportValue = "Æ Unité";
        case "AE Large":           
            myValueTextEn = "Uncertain large bronze";
            myValueTextFr = "Bronze grand incertain";
            myImportValueFr = "Æ Large";            
            myImportValue = "Æ Grande"; 
        case "AE Medium":           
            myValueTextEn = "Uncertain medium bronze";
            myValueTextFr = "Bronze moyen incertain";
            myImportValueFr = "Æ Medium";            
            myImportValue = "Æ Moyen"; 
        case "AE Small":           
            myValueTextEn = "Uncertain small bronze";
            myValueTextFr = "Bronze petit incertain";
            myImportValueFr = "Æ Small";            
            myImportValue = "Æ Petit"; 
        case "AE Small|AE Medium":           
            myValueTextEn = "Uncertain small or medium bronze";
            myValueTextFr = "Bronze moyen ou petit incertaine";
            myImportValueFr = "Æ";            
            myImportValue = "Æ";             
        }
      }
    //replace |
//    if ((myImportValue.length - myImportValue.replace(/\|/g,'').length)>1){
//    while ((myImportValue.length - myImportValue.replace(/\|/g,'').length)>1){
//      myImportValue = myImportValue.replace('|',", ");
//    }
//    myImportValue = myImportValue.replace('|',", or ");
//    }
//    else {
//      myImportValue = myImportValue.replace('|'," or ");
//    }
    //1615 - Antoninianus, Reform of Caracalla (AD 215 - 301)
    //1618 - Argenteus, Reform of Diocletian (AD 293/301 - 310/324)
    //1619 - Solidus, Reform of Constantine (AD 310/324 - 395)   
  
//    var myMass = 0;
//    var myCountMass = 0;
//    var myDim = 0;
//    var myCountDim = 0;
    
   var myTempString = "";
   
    //////////////////////////////////////////////////////
    //PARSER
    //////////////////////////////////////////////////////
    
    //Assign variable
    myTempString =""; 
    myTempString = myTempString + "### NEW COIN TYPE ### RIC:" + myRICno + NL;
    myTempString = myTempString + "mycointype={" + NL;
    //TITLE
    myTempString = myTempString + "  " + QM + "title" + QM + ": [" + NL;
    myTempString = myTempString + "    {" + NL;
      ////en
    myTempString = myTempString + "      " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "      " + QM + "label" + QM + ": " + QM + myImportValue;
    if (myPortraitEn !== ""){
    myTempString = myTempString + " - " +  myPortraitEn;
    }
    if ((myDeity !== "") || (myRvLegend !== "")){
    myTempString = myTempString + " ("
      if (myRvLegend !== "") {
    myTempString = myTempString + myRvLegend;
      }
    if (myDeity !== "") {
      if (myRvLegend !== "") {
    myTempString = myTempString + "; ";
      }
      myTempString = myTempString +  myDeity; 
    }
      myTempString = myTempString + ")" + QM + NL;
    }
    else{
      myTempString = myTempString + QM + NL;
    }
    myTempString = myTempString + "    }," + NL;
      ////fr
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "language" + QM + ": " + QM + "fr" + QM + "," + NL;
    myTempString = myTempString + "      " + QM + "label" + QM + ": " + QM + myImportValueFr;
    if (myPortraitFr !== ""){
    myTempString = myTempString + " - " +  myPortraitFr;
    }
    if ((myDeityFr !== "") || (myRvLegend !== "")){
    myTempString = myTempString + " ("
      if (myRvLegend !== "") {
    myTempString = myTempString + myRvLegend;
      }
    if (myDeityFr !== "") {
      if (myRvLegend !== "") {
    myTempString = myTempString + "; ";
      }
      myTempString = myTempString +  myDeityFr; 
    }
      myTempString = myTempString + ")" + QM + NL;
    }
    else{
      myTempString = myTempString + QM + NL;
    }
    myTempString = myTempString + "    }" + NL;
    myTempString = myTempString + "  ]," + NL;
       
    //TYPE
    myTempString = myTempString + "  " + QM + "issuer" + QM + ": {" + NL;
    myTempString = myTempString + "    " + QM + "code" + QM + ": " + QM + myIssuer + QM + NL;
    myTempString = myTempString + "  }," + NL
    
    //ISSUER
    myTempString = myTempString + "  " + QM + "type" + QM + ": " + QM + "common_coin" + QM + "," + NL;
    
    //VALUE
    myTempString = myTempString + "  " + QM + "value" + QM + ": {" + NL;  
    myTempString = myTempString + "    " + QM + "text" + QM + ": [" + NL;
    //value en
    myTempString = myTempString + "      {" + NL;
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myValueTextEn + QM + NL;
    myTempString = myTempString + "      }," + NL;
    //value fr
    myTempString = myTempString + "      {" + NL;
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "fr" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myValueTextFr + QM + NL;
    myTempString = myTempString + "      }" + NL;
    myTempString = myTempString + "    ]," + NL;
    //further value data
        if (myFractionalFlag == 1){
          if (myNumerator !== ""){
    myTempString = myTempString + "    " + QM + "numerator" + QM + ": " + myNumerator + "," + NL;
    myTempString = myTempString + "    " + QM + "denominator" + QM + ": " + myDenominator + "," + NL;
          }
        }
        else {
          if (myIntegerValue !== ""){
    myTempString = myTempString + "    " + QM + "numeric_value" + QM + ": " + myIntegerValue + "," + NL; 
          }
        }
      myTempString = myTempString + "    " + QM + "currency" + QM + ": {" + NL;
    myTempString = myTempString + "      " + QM + "id" + QM + ": " + myCurrencyID + NL;
    myTempString = myTempString + "    }" + NL
    myTempString = myTempString + "  }," + NL  
    
    //RULING AUTHORITY
    myTempString = myTempString + "  " + QM + "ruling_authority" + QM + ": [" + NL;
    myTempString = myTempString + "      {" + NL;
    myTempString = myTempString + "      " + QM + "id" + QM + ": " + myRulerID + NL;
    myTempString = myTempString + "      }" + NL;
    myTempString = myTempString + "  ]," + NL;
    
    //SHAPE
    myTempString = myTempString + "  " + QM + "shape" + QM + ": {" + NL;
    myTempString = myTempString + "    " + QM + "id" + QM + ": " + "2" + NL;
//    myTempString = myTempString + "    " + QM + "additional_details" + QM + ": [" + NL;
//    myTempString = myTempString + "      {" + NL;
//    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;   
//    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + "NULL" + QM + NL;
//    myTempString = myTempString + "      }" + NL;
//    myTempString = myTempString + "    ]" + NL;
    myTempString = myTempString + "  }," + NL;
 
    //COMPOSITION
    myTempString = myTempString + "  " + QM + "composition" + QM + ": {" + NL;
    myTempString = myTempString + "    " + QM + "composition_type" + QM + ": " + QM + "plain" + QM + "," + NL;
    myTempString = myTempString + "    " + QM + "core" + QM + ": {" + NL;
    myTempString = myTempString + "      " + QM + "material" + QM + ": {" + NL;
    myTempString = myTempString + "        " + QM + "id" + QM + ": " + myMetalID + NL;
    myTempString = myTempString + "      }," + NL;
//    myTempString = myTempString + "      " + QM + "fineness" + QM + ": " + "NULL" + NL;
//    myTempString = myTempString + "    }," + NL;
//    myTempString = myTempString + "    " + QM + "additional_details" + QM + ": [" + NL;
//    myTempString = myTempString + "      {" + NL;
//    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;   
//    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + "NULL" + QM + NL;
//    myTempString = myTempString + "      }" + NL;
//    myTempString = myTempString + "    ]" + NL;
    myTempString = myTempString + "  }," + NL;
    myTempString = myTempString + " }," + NL;    
    //WEIGHT
    if (myTmass !== ""){
    myTempString = myTempString + "  " + QM + "weight" + QM + ": " + myTmass + "," + NL;
    }
    
    //Diameter
    if (myTdiam !== ""){
    myTempString = myTempString + "  " + QM + "size" + QM + ": " + myTdiam + "," + NL;
    }
    
    //Thickness
//    myTempString = myTempString + "  " + QM + "thickness" + QM + ": ," + NL;
    
    //Orientation
    if (myAxis !==""){
    myTempString = myTempString + "  " + QM + "orientation" + QM + ": " + QM + myAxis + QM + "," + NL;
    }
    
    //Demonetization
    myTempString = myTempString + "  " + QM + "demonetization" + QM + ": {" + NL;
    myTempString = myTempString + "    " + QM + "is_demonetized" + QM + ": " + "True," + NL;
//    myTempString = myTempString + "    " + QM + "demonetization_date" + QM + ": " + QM + "0000-00-00" + QM + NL;
    myTempString = myTempString + "  }," + NL;
    
    //Calendar
    myTempString = myTempString + "  " + QM + "calendar" + QM + ": {" + NL;
    myTempString = myTempString + "    " + QM + "code" + QM + ": "+ QM + "gregorien" + QM + NL;
    myTempString = myTempString + "  }," + NL;
    
    //OBVERSE=======================
    myTempString = myTempString + "  " + QM + "obverse" + QM + ": {" + NL;
    ////engravers
    if (myImportSheet.getRange(i,9).getValue() !== "") {
    myTempString = myTempString + "    " + QM + "engravers" + QM + ": [" + NL;
    myTempString = myTempString + "      " + QM + myImportSheet.getRange(i,9).getValue() + QM + NL;      
    myTempString = myTempString + "    ]," + NL;  
    } 
    ////description   
    myTempString = myTempString + "    " + QM + "description" + QM + ": [" + NL;
    myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;   
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myOvDescription + QM + "," + NL;
    myTempString = myTempString + "      }" + NL;
    myTempString = myTempString + "      ]," + NL;
    //lettering
    if (myOvLegend !== ""){
    myTempString = myTempString + "    " + QM + "lettering" + QM + ": " + QM + myOvLegend + QM + "," + NL;
          if ((myRvTransEn !== "")||(myRvTransFr !== "")){
    myTempString = myTempString + "    " + QM + "lettering_translation" + QM + ": [" + NL;
                if (myOvTransEn !== ""){
    myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myOvTransEn + QM + "," + NL;
    myTempString = myTempString + "      }," + NL; 
    }  
    if (myOvTransFr !== ""){
    myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "fr" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myOvTransFr + QM + "," + NL;
    myTempString = myTempString + "      }" + NL; 
    }
    myTempString = myTempString + "    ]," + NL; 
    }
    }
    ////Pictures
    if(myOVlink !== ""){
    myTempString = myTempString + "    " + QM + "picture" + QM + ": " + QM + myOVlink + QM + "," + NL;
    myTempString = myTempString + "    " + QM + "picture_copyright" + QM + ": " + QM + myImgCredits + QM + "," + NL;
    myTempString = myTempString + "    " + QM + "picture_copyright_type" + QM + ": " + QM + myCreditType + QM + NL;
    }
    myTempString = myTempString + "  }," + NL;    

    //REVERSE======================
    myTempString = myTempString + "  " + QM + "reverse" + QM + ": {" + NL;
    ////engravers
    if (myImportSheet.getRange(i,9).getValue() !== "") {
    myTempString = myTempString + "    " + QM + "engravers" + QM + ": [" + NL;
    myTempString = myTempString + "      " + QM + myImportSheet.getRange(i,9).getValue() + QM + NL;      
    myTempString = myTempString + "    ]," + NL;  
    } 
    ////description   
    myTempString = myTempString + "    " + QM + "description" + QM + ": [" + NL;
    myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myRvDescription + QM + "," + NL;
    myTempString = myTempString + "      }" + NL;
    myTempString = myTempString + "      ]," + NL;
    ////lettering
    if (myRvLegend !== ""){
    myTempString = myTempString + "    " + QM + "lettering" + QM + ": " + QM + myRvLegend + QM + "," + NL;
    if ((myRvTransEn !== "")||(myRvTransFr !== "")){
      myTempString = myTempString + "    " + QM + "lettering_translation" + QM + ": [" + NL;
      if (myRvTransEn !== ""){
        myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myRvTransEn + QM + "," + NL;
    myTempString = myTempString + "      }," + NL;
    }
          if (myRvTransFr !== ""){
    myTempString = myTempString + "      {" + NL;    
    myTempString = myTempString + "        " + QM + "language" + QM + ": " + QM + "fr" + QM + "," + NL;
    myTempString = myTempString + "        " + QM + "label" + QM + ": " + QM + myRvTransFr + QM + "," + NL;
    myTempString = myTempString + "      }" + NL;
          }
    myTempString = myTempString + "    ]," + NL; 
    }
    }
    ////Pictures
    if(myRVlink !== ""){
    myTempString = myTempString + "    " + QM + "picture" + QM + ":" + QM + myRVlink + QM + "," + NL;
    myTempString = myTempString + "    " + QM + "picture_copyright" + QM + ": " + QM + myImgCredits + QM + "," + NL;
    myTempString = myTempString + "    " + QM + "picture_copyright_type" + QM + ": " + QM + myCreditType + QM + NL;
    }
    myTempString = myTempString + "  }," + NL;    
    
    //M I N T S
    if (myMint !== "#N/A") {
    myTempString = myTempString + "  " + QM + "mints" + QM + ": [" + NL;
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "id" + QM + ": " + myMint + NL;
    myTempString = myTempString + "    }" + NL;      
    myTempString = myTempString + "    ]," + NL;    
    }
    
    //C O M M E N T S:
    myTempString = myTempString + "  " + QM + "comments" + QM + ": [" + NL;
    ////en
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "language" + QM + ": " + QM + "en" + QM + "," + NL;
    myTempString = myTempString + "      " + QM + "label" + QM + ": "
        
        var myComBool = false;
        
    myCell = myImportSheet.getRange(i,38).getValue();
    if (myCell !== ""){
        myTempString = myTempString + QM + "Mass varies: " + myCell + QM + "+ chr(10) +"
        myComBool = true;
      }
    
    myCell = myImportSheet.getRange(i,39).getValue();
    if (myCell !== ""){
        myTempString = myTempString + QM + "Diameter varies: " + myCell + QM + "+ chr(10) +";
        myComBool = true;
      }
      
      if (myComBool){
        myTempString = myTempString + " chr(10) +";    
        }
    //Sample image
    if(myImgLink !== ""){
    myTempString = myTempString + QM + "Example of this type:" + QM +  " + chr(10) + " + QM + "[url=" + myImgLink + "]" + myImgCredits + "[/url]" + QM +  "+ chr(10) + chr(10) +";
    }
    
    myTempString = myTempString + QM + "Source:"  + QM + " + chr(10) + " + QM +  "[url=http://numismatics.org/ocre/]Online Coins of the Roman Empire (OCRE)[/url]" + QM + NL;

    myTempString = myTempString + "    }," + NL;
    ////fr
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "language" + QM + ": " + QM + "fr" + QM + "," + NL;
    myTempString = myTempString + "      " + QM + "label" + QM + ": "
        myCell = myImportSheet.getRange(i,38).getValue();
    if (myCell !== ""){
        myTempString = myTempString + QM + "Poids variable: " + myCell + QM + "+ chr(10) +"
      }
    myCell = myImportSheet.getRange(i,39).getValue();
    if (myCell !== ""){
        myTempString = myTempString + QM + "Diamètre variable: " + myCell + QM + "+ chr(10) +";
      }
    if (myComBool){
        myTempString = myTempString + " chr(10) +";    
        }  
    //Sample image
    if(myImgLink !== ""){
    myTempString = myTempString + QM + "Exemple de ce type:" + QM  + " + chr(10) + " + QM + "[url=" + myImgLink + "]" + myImgCredits + "[/url]" + QM + "+ chr(10) + chr(10) +";
    }
    
    myTempString = myTempString + QM + "Source:" + QM + " + chr(10) + " + QM + "[url=http://numismatics.org/ocre/]Online Coins of the Roman Empire (OCRE)[/url]" + QM + NL;

    myTempString = myTempString + "    }," + NL;
//
    myTempString = myTempString + "  ]," + NL;
    
    //TAGS:
    var myTagArray = [];
    var myTags = myImportSheet.getRange(i,34).getValue().toString();
    if (myTags !== ""){
      myTagArray = myTags.split(", ");
      myTempString = myTempString + "  " + QM + "tags" + QM + ": [" + NL;
      for (var myTagCt = 0; myTagCt<myTagArray.length-1; myTagCt++){
        myTempString = myTempString + "    {" + NL;
        myTempString = myTempString + "      " + QM + "id" + QM + ": "+ myTagArray[myTagCt] + NL;    
        myTempString = myTempString + "    }," + NL;     
      }
      if (myTagArray.length>=1){
        myTempString = myTempString + "    {" + NL;
        myTempString = myTempString + "      " + QM + "id" + QM + ": "+ myTagArray[myTagArray.length-1] + NL;    
        myTempString = myTempString + "    }" + NL;     
      }
      myTempString = myTempString + "  ]," + NL;
    }
      
    //REFERENCES:
    myTempString = myTempString + "  " + QM + "references" + QM + ": [" + NL;
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "catalogue" + QM + ": {" + NL;
    myTempString = myTempString + "        " + QM + "id" + QM + ": " + myCatID + NL;
    myTempString = myTempString + "      }," + NL;
    myTempString = myTempString + "      " + QM + "number" + QM + ": " + QM + myRICno + QM + NL;    
    myTempString = myTempString + "    }," + NL;
    myTempString = myTempString + "    {" + NL;
    myTempString = myTempString + "      " + QM + "catalogue" + QM + ": {" + NL;
    myTempString = myTempString + "        " + QM + "id" + QM + ": 1020" + NL; //catalog ID of OCRE
    myTempString = myTempString + "      }," + NL;
    myTempString = myTempString + "      " + QM + "number" + QM + ": " + QM + myImportSheet.getRange(i,3).getValue() + QM + NL;    
    myTempString = myTempString + "    }" + NL;
    myTempString = myTempString + "  ]" + NL;      
    myTempString = myTempString + "}" + NL;
   
      //RESPONSE::
    myTempString = myTempString + "######## RESPONSE #######" + NL;
    myTempString = myTempString + "response = requests.post(" + NL;
    myTempString = myTempString + "    endpoint + '/coins'," + NL;
    myTempString = myTempString + "    params={'lang':'en'}," + NL;    
    myTempString = myTempString + "    headers={'Numista-API-Key': api_key}," + NL;
    myTempString = myTempString + "    data=json.dumps(mycointype)" + NL; 
    myTempString = myTempString + "    )" + NL; 
    myTempString = myTempString + "coinadd_response = response.json()" + NL; 
    myTempString = myTempString + "print(response)" + NL; 
    myTempString = myTempString + "print(" + QM + "Coin added succesfully; RIC# = " + myRICno + "; ID = " + QM + ", coinadd_response['id'])" + NL;
    ////////////////////////////////////
    //ISSUE
    myTempString = myTempString + "######## COIN ISSUE #######" + NL;  
    myTempString = myTempString + "coin_type_id =  coinadd_response['id']" + NL;             
    myTempString = myTempString + "mydateline={" + NL;
    myTempString = myTempString + "  " + QM + "is_dated" + QM + ": False" + "," + NL;
    myTempString = myTempString + "  " + QM + "min_year" + QM + ": " + myStartYear + "," + NL;
    myTempString = myTempString + "  " + QM + "max_year" + QM + ": " + myEndYear + "," + NL;  
    myTempString = myTempString + "}" + NL;
    
    //ISSUE Response
    myTempString = myTempString + "######## ISSUE RESPONSE #######" + NL;  
    myTempString = myTempString + "response = requests.post(" + NL;             
    myTempString = myTempString + "  endpoint + '/coins/' + str(coin_type_id) + '/issues'," + NL;
    myTempString = myTempString + "  params={'lang':'en'}," + NL;
    myTempString = myTempString + "  headers={'Numista-API-Key': api_key}," + NL;
    myTempString = myTempString + "  data=json.dumps(mydateline)" + NL;  
    myTempString = myTempString + "  )" + NL; 
    myTempString = myTempString + "coinadd_response = response.json()" + NL; 
    myTempString = myTempString + "print(response)" + NL; 
    myTempString = myTempString + "print(" + QM + "Date Line added succesfully; RIC# = " + myRICno + "; Date = " + myStartYear + "-" + myEndYear + QM + ")" + NL;
    
    
   myCell = myAPIsheet.getRange(i,1);
   myCell.setValue(myTempString); 
      }
  
}
