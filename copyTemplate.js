var FEUILLE_RECAP = 'Récapitulatif demandes';
var COL_NOM = 1;
var FEUILLE_TEMP = 'Template demande';

function createNewDemand() {
    // 1 - trouver le premier nom disponible
    var sh = getSheet(FEUILLE_RECAP);
    var ligne = getLastRowForColumn(sh.getRange("B:B")) + 1;
    var nom_cible = sh.getRange(ligne, COL_NOM).getValue();
    // 2 copier le template
    var newSheet = copyTemplateTo(getSheet(FEUILLE_TEMP), nom_cible);
    // 3 insérer l'hyperlien
    addHyperlinkToCell(sh.getRange(ligne, COL_NOM),getLinkToSheet(newSheet));
}

function removeOtherProtection() {
    var shts = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(i=0; i < shts.length; i++) {
      var prot = shts[i].getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      if(prot) {
        prot.removeEditors(prot.getEditors());
      }
    }
}