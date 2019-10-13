/*********************************************************
************ support functions ***************************
*********************************************************/


// affiche une boite d'alerte avec le message
function alert(prompt){
   SpreadsheetApp.getUi().alert(prompt);
}

/**
* @param {String} message : the message to display
* @return {String} : the text input by the user, empty if "close" clicked
*/
function prompt(message){
  var ui = SpreadsheetApp.getUi();
  return ui.prompt(message, ui.ButtonSet.OK).getResponseText();
}


// affiche le message dans le coin en bas à droite
function toast(msg, time){
  time = time || 3;  
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "", time);   
}



/** return the active sheet of the active spreadsheet
* @return {Sheet}
*/
function activeSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}


/**
* remove the protection of existing range of the sheet
* @param {Sheet} sheet
* @return {unprotectedRanges :[Range], sheet:{Sheet}, editors: [User]} : the original unprotected ranges, empty array if wasn't  protected, the sheet and editors associated
*/
function unprotectSheet(sheet){
  var protections=sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){
    return {unprotectedRanges : [], sheet : sheet, editors: []};
  }
   var originalUnprotected = protections[0].getUnprotectedRanges();
   protections[0].setUnprotectedRanges([sheet.getDataRange()]);
  return {unprotectedRanges : originalUnprotected, sheet : sheet, editors: protections[0].getEditors() };
}


function unprotectWholeSheet(sheet){
  var prot = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  prot.remove();
}

function protectWholeSheet(sheet){
 var prot = sheet.protect();
 prot.removeEditors(['lea.legrand@ceva.com', 'nelly.lesceau@ceva.com', 'magali.bossiere@ceva.com', 'garfrom@gmail.com', 'alexandre.brechet@ceva.com']); 
}
/**
* @param {unprotectedRanges:[Range],sheet: {Sheet}} originalUnprotected 
* @return {Protection} for chaining
*/
function restoreProtection(originalUnprotected){
  var protections=originalUnprotected.sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){
    return protections;
  }
  return protections[0].setUnprotectedRanges(originalUnprotected.unprotectedRanges);
}
 
/**
* @param {unprotectedRanges:[Range],sheet: {Sheet}} originalUnprotected 
* @return nothing
*/
function setProtection(originalProtection, ranges) {
  function protectRange(range) {
    var protection = originalProtection.sheet.getRange(range).protect();
    protection.addEditors(originalProtection.editors);
  }
  ranges.forEach(protectRange);
}


/**
* Copy the sheet protection from one sheet to another one
* @param {Sheet} fromSheet
* @param {Sheet} toSheet
* @return {Protection} : the protection object of the new sheet, null if fromSheet was not protected
* NOTE : only the first sheet protection is copied, including unprotected ranges
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function copyProtectionFromSheetToSheet(fromSheet, toSheet){
  var protections = fromSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){return null;}
  return copyProtectiontoSheet(protections[0],
                        toSheet);
}


/**
* Copy the given protection to another sheet
* @param {Protection} protection
* @param {Sheet} targetSheet
* @return {Protection} : the protection object of the new sheet
* NOTE : the protection is copied, including unprotected ranges
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function copyProtectiontoSheet(protection, targetSheet){
  var ur = protection.getUnprotectedRanges();
  // convert range into same range in new sheet
  var targetUr=[];
  for(i=0;i<ur.length;i++){
    targetUr.push(targetSheet.getRange(ur[i].getA1Notation()));
  }
  // set description to sheet name and copy unprotected ranges
  var newProtection= targetSheet.protect()
    .setDescription(targetSheet.getSheetName())
    .setUnprotectedRanges(targetUr);
  // allowed editors are those from original protection    
   return  newProtection.removeEditors(newProtection.getEditors())
    .addEditors(protection.getEditors());
}



/**
* force a sheet to refresh (when using query, Index(), custom function
* @param {Spreadsheet} the spreadsheet to whic the sheet belongs
* @param {Sheet} the sheet to refresh
* @return 
* NOTE : max 10 attempt done if legitimate #N/A in cell
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function refreshSheet(spreadsheet, sheet) {
  var dataArrayRange = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var dataArray = dataArrayRange.getValues(); // necessary to refresh custom functions
  var nanFound = true;
  var cpt = 10;
  while(nanFound && cpt > 0) {
    for(var i = 0; i < dataArray.length; i++) {
      if(dataArray[i].indexOf('#N/A') >= 0) {
        nanFound = true;
        dataArray = dataArrayRange.getValues();
        cpt--; // to avoid looping when formula result in #N/A legitimely
        break;
      } // end if
      else if(i == dataArray.length - 1) nanFound = false;
    } // end for
  } // end while
}


/** return the sheet in this spreadsheet with given name (null if doesn't exist)
* @param {String} name
* @return {Sheet}
*/
function getSheet(name){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}


/**
* @param {Range} col mono-dimensional range (column or row)
* @return {Integer} the last non-empty position (stops when more than 50 empty position to allow
* for "hole" in the data series)
*/
function getLastRowForColumn(col) {
  var values = col.getValues();
  var lign = 0, lastNonEmpty = 0;
  while( lign < values.length ) { 
    if( values[lign] != "") {lastNonEmpty = lign}
    if(lign++ - lastNonEmpty > 50) {break}
  } // end while
  return lastNonEmpty + 1; //because value array start at 0, rows at 1
}


/**
* @param {Int} nbL, nbC : nb de ligne et de colonne du tableau à créer
* @return {[][]} : an array of nbL x nbC initialised with empty string
*/
function createArray(nbL, nbC) {
  var arr = Array(nbL);
  for(var i=0; i < nbL; i++) {
    arr[i] = Array(nbC);
    for(j=0; j < nbC; arr[i][j++] = '');
  }
  return arr;
}



// make the string for a formula that create an hyperlink to link
// displaying "display" in the cell
function getHyperlinkFormulaToWithDisplay(link, display){
  return  '=HYPERLINK("'+link+'"; "'+display+'")'; 
 }
 
 
/** add an hyperlink to link to the cell
* it will display the current cell content
* @param {Cell} cell
* @param {String} link
*/
function addHyperlinkToCell(cell, link){
  cell.setFormula( 
    getHyperlinkFormulaToWithDisplay(
      link, cell.getValue())
    ).setShowHyperlink(true);
}


/** return the URL of a given sheet in this spreadsheet (for direct access, without opening a new tab)
* @param {Sheet} sheet
* @return {String}
*/
function getLinkToSheet(sheet){
  return "#gid="+sheet.getSheetId();
}

