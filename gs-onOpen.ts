function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('Arquivo APS BSF')
  .addItem('Ver processos cadastrados', 'showAwTable')
  .addSeparator()
  .addItem('Enviar Arquivo', 'showPicker')
  .addItem('Gerar etiqueta para caixa', 'showSideBarEtiqueta')
  .addToUi();
}
