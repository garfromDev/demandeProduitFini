//==== function common to many applicative function ============


/**
* copy the template to a new sheet with the given name
* in case of failure, ensure no sheet is created
* @param {Sheet} the template
* @param {String} the nema for the new sheet to create
* @return {Sheet} the newly created sheet
* @throw : "existing newName","copy failed", "rename failed" ou autre erreur liée à la copie ou le renommage
*/
function copyTemplateTo(template, newName){
  var newSheet;
    s = SpreadsheetApp.getActiveSpreadsheet();
  // 0 check if name already exist
  if (s.getSheetByName(newName) != null) {
    throw new Error("existing name "+newName);
  }
  // 1 find template and make a copy 
  try{
    newSheet = template.copyTo(s);
  }catch(error){
   throw new Error("copy of "+template.getName()+" failed because "+error.message); 
  }
  // 2 rename it
  try {
    newSheet = newSheet.setName(newName);
  }catch(error){
    s.deleteSheet(newSheet);
    throw new Error("rename to "+newName+" failed because "+error.message);
  }
  // 3 set protections same as template
  try {
    copyProtectionFromSheetToSheet(template, newSheet);
  }catch(error){
    s.deleteSheet(newSheet);
    throw new Error("Setting protection of "+newName+" failed because "+error.message);
  }  
  return newSheet;
}


function CopyTemplateWithData(template, newName, datas, locations){
   // copier le template vers une nouvelle feuille
  try{
    var newSheet = copyTemplateTo(template, newName);
  }catch (error) {
    throw new Error("erreur lors de la copie de "+template.getName()+" vers "+newName+" <"+error.message+">");
  }
    
  // transférer les données vers la nouvelle feuille
  try{
    transferDataTo(datas, locations, newSheet);
  }catch (error) { // in case of failure, we display a message and delete the created copy
    newSheet.getParent().deleteSheet(newSheet);
    throw new Error("Erreur lors du transfert des données vers "+newName+" <"+error.message+">");
  }
  return newSheet;
}

/**
* write all the value in the target sheet at given locations
* @param {[rawValue]} datas : an array of value to write (must be primitive value acceptable for Range.setValue)
* @param {[String]} targetLocations : an array of A1 notation for location to be written
* @param {Sheet} targetSheet
* @throw : "Impossible de copier le type et la date vers la feuille ", autre erreur liée à la copie
*/
function transferDataTo( datas, targetLocations, targetSheet){

  if(datas.length > targetLocations.length){
    throw new Error("Pas assez d'emplacement pour les données"  +datas.length+ " > "+targetLocations.length ); 
  }
  var originalProtection=unprotectSheet(targetSheet);
  for(i=0;i<datas.length; i++){
    targetSheet.getRange(targetLocations[i]).setValue(datas[i]); 
  }
  restoreProtection(originalProtection);
}


/**
* ATTENTION : le format de la ligne ligne sera copiée entièrement à partir de la 1ere colonne
*/
function insertLineInTable(fromFormula, toFormula){
  var s = activeSheet();
  var destinationRow = s.getRange(toFormula).getLastRow();
  // insert a new line
  s.insertRowBefore(destinationRow);
  try{
    // copy the format
    var sourceRange = s.getRange("A"+fromFormula.substring(1)); //on étend la ligne jusqu'à la première colonne
    sourceRange.copyFormatToRange(s,1,getRowFromA1(fromFormula),destinationRow,destinationRow); 
    // copy the formula
    s.getRange(fromFormula).copyTo(s.getRange(toFormula));
  }catch(error){
    s.deleteRow(destinationRow);
    throw error;
  }
}



function archiveFromTable(formulaFromZone, lineToArchive, sheetName){  
  var sprd =  SpreadsheetApp.getActiveSpreadsheet();
  var as = activeSheet();
 // 0 pas d'archivage si il n'y a plus qu'une ligne dans le tableau
  if(as.getLastRow() < getRowFromA1(formulaFromZone)){
    alert("Il doit toujours rester au moins une ligne dans le tableau");
    return;
  }

  if(sheetName =="" ){
     alert("Veuillez sélectionner un item à archiver dans la liste");
     return;
  }
  // 2 demande confirmation à l'utilisateur
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Mycoplasme',
                           'Archivage du lot '+sheetName+" ?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {
    return; // on arrête tout
  }
  
  toast("Archivage du lot "+ sheetName + " en cours");
  
  // Archivage de la feuille du lot
  try{
    performArchiving(sheetName);
  }catch(error){
    alert(error.message);
    // à ce point, rien n'a été effacé dans Stock AG
    return;
  }
  
  // archivage réussi, on efface la feuille
  var toDelete =  sprd.getSheetByName(sheetName);
  try{
     sprd.deleteSheet(toDelete);
  }catch(error){
    alert("Impossible de supprimer la feuille "+toDelete+" : "+error.message);
    // logiquement toujours rien d'effacé
    return;
  }
  // on retire la ligne du lot du tableau
  //alert("suppression ligne ");
  as.deleteRow(lineToArchive);
  
  toast("Lot "+sheetName+" archivé avec succès");
  
}



/**
* Crée un classeur avec le même nom que la feuille, copie la feuille dedans
* et efface les autres feuilles, déplace le classeur crée dans le répertoire 
* AG_ARCHIVE_FOLDER_ID
*
* @param {String} the name of the sheet in the current Spreadshhet that need archiving
*/
function performArchiving(sheetName){

  var AGspreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fromSheet = AGspreadsheet.getSheetByName(sheetName);
  if(fromSheet==null){
    throw new Error("La feuille à archiver " + sheetName + " n'existe pas  ");
  }

  // création d'un classeur vierge
  var targetSpreadsheet = SpreadsheetApp.create(sheetName);
  if(targetSpreadsheet==null){
    throw new Error("La création du classeur " + sheetName + " a échoué ");
  }
  // Copie de la feuille dans le classeur avec le même nom
  var targetSheet = copyUniqueTo(fromSheet,targetSpreadsheet);
  var archiveId = targetSpreadsheet.getId()
  // deplacement dans le répertoire d'archivage
  var targetFolder = DriveApp.getFoldersByName(CST().AG_ARCHIVE_FOLDER_ID);
  if(!targetFolder.hasNext()){
    throw new Error("Le dossier " + CST().AG_ARCHIVE_FOLDER_ID + " n'existe pas dans le drive!");
  }
  targetFolder = targetFolder.next();
  var archiveFile = DriveApp.getFileById(archiveId); 
  targetFolder.addFile(archiveFile);
}


/**
* @param {Sheet} sheet 
* @return {boolean} : true if the sheet is a template, i.e. the name containes "template"
*/
function isTemplate(sheet){
  return /template/i.test(sheet.getName()); 
}
