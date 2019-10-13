var FEUILLE_RECAP = 'Récapitulatif demandes';
var COL_NOM = 1;
var FEUILLE_TEMP = 'Template demande';

function createNewDemand() {
    // 1 - trouver le premier nom disponible
    var sh = getSheet(FEUILLE_RECAP);
    var ligne = getLastRowForColumn(sh.getRange("B:B"));
    var nom_cible = sh.getRange(ligne, COL_NOM).getValue();
    // 2 copier le template
    var new_sheet = copyTemplateTo(FEUILLE_TEMP, nom_cible);
    // 3 insérer l'hyperlien
    addHyperlinkToCell(sh.getRange(ligne, COL_NOM),getLinkToSheet(newSheet));
}