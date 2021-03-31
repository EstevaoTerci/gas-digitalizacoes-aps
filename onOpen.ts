function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('Digitalizações APS BSF')
  .addItem('Enviar Arquivo', 'showPicker')
  .addToUi();
}
