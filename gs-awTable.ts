function showAwTable() {  
  let html = HtmlService.createTemplateFromFile('AwTable.html')
  let htmlOutput = html.evaluate()
    .setTitle('Ver processos cadastrados')
    .setWidth(800)
    .setHeight(700)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ver processos cadastrados');
}