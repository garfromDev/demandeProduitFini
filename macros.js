

//fonction appelée par le bouton
function EnvoiMailSaisie() {
  sendEmail("aline.fromont@ceva.com","Saisie d'une nouvelle demande de produit R&D",6);
}

/**
Envoi un email depuis le compte de l'utilisateur courant
to : l'adresse (ou les adresses séparés par des virgules) de destination
subject : le sujet du mail
fromCol : le no de la colonne dans laquelle on trouve le contenu du message
*/
function sendEmail(to, subject, fromCol)
{
  //1 on récupère le nO de ligne de la cellule sélectionnée
  l = 62;
  // 2 le contenu du mail est dans la cellule de la même ligne, en colonne fromCol
  contenu = SpreadsheetApp.getActiveSheet().getRange(l, fromCol).getValue();
  
// 3 Display a dialog box with a title, message, and "Yes" and "No" buttons. The
// user can also close the dialog by clicking the close button in its title bar.
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Confirmer envoi email ?', contenu, ui.ButtonSet.YES_NO);

// Process the user's response.
if (response == ui.Button.YES)
  MailApp.sendEmail(to, subject, contenu);
else if (response == ui.Button.NO)
  return;// on arrête tout
else
  Logger.log('The user clicked the close button in the dialog\'s title bar.');
return; // on arrête tout
}  


//fonction pour récupérer l'url d'une feuille
function urlFeuille() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();  
  urlFeuille = spreadsheet.getFormUrl();

} 


/** Custom function  
* utilisation : `= NOM_FEUILLE()`
* @return {String} Le nom de la feuille courante
*/
function NOM_FEUILLE() {
  return activeSheet().getName();
}