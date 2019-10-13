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
    );
}


/** return the URL of a given sheet in this spreadsheet (for direct access, without opening a new tab)
* @param {Sheet} sheet
* @return {String}
*/
function getLinkToSheet(sheet){
  return "#gid="+sheet.getSheetId();
}

