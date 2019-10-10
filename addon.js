/* What should the add-on do after it is installed */
function onInstall() {
  onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu() // Add a new option in the Google Docs Add-ons Menu
  .addItem("Import", "ImportToSouchier")
  .addItem("Mise à jour emplacements", "initEmplacements")
  .addToUi();  // Run the showSidebar function when someone clicks the menu
}


