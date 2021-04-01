//FOI MONTADO A PARTIR DOS 2 LINKS: https://www.youtube.com/watch?v=rivZqJGHb0o     //   https://developers.google.com/apps-script/guides/dialogs


/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  Logger.log('Running showPicker')
  let html = HtmlService.createTemplateFromFile('Picker.html')
  let htmlOutput = html.evaluate().setTitle('Enviar PDFs').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * @description Filtra os NBs a serem exibidos na select
 */
const filtraNbsSemArquivos = (): RowData[] => {
  const sheetData = sheetDataObject()
  const headerFiltrar = 'Arquivo'
  const headerFiltrar2 = 'Numero'
  const filteredData = sheetData.filter(a => {    
    return (a[headerFiltrar] === '' || a[headerFiltrar] === null || a[headerFiltrar] === undefined) &&
            !(a[headerFiltrar2] === '' || a[headerFiltrar2] === null || a[headerFiltrar2] === undefined)
  })
  return filteredData
}

/**
 * 
 * @returns {Object} {rowNum: number, ...headers: rowData}
 */
const sheetDataObject = (): RowData[] => {
  const sheetData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Processos').getDataRange().getValues()
  const headers = sheetData.shift()
  const sheetEntries = <RowData[]>sheetData.map((arrLinha, i) => {
    let rowObject = {}
    rowObject['rowNum'] = i + 2
    for(let header of headers){
      let indexHeader = headers.indexOf(header)
      rowObject[header] = arrLinha[indexHeader]
    }
    return Object.assign({}, rowObject)
  })
  return <RowData[]>sheetEntries
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function onEnvioBemSucedido({id, rowNum}){
  //daqui para frene pegar oid ou url e prosseguir
  Logger.log(id);
  const rowData: RowData = sheetDataObject().filter(a => a['rowNum'] === +rowNum )[0]
  const fileName = `${rowData['Numero']} ${rowData['Nome']}`
  const fileUrl = DriveApp.getFileById(id).setName(fileName).getUrl()
  rowData['Arquivo'] = fileUrl
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Processos')
  const headers = sheet.getDataRange().getValues().shift()
  const headerCol = headers.indexOf('Arquivo') + 1
  sheet.getRange(rowNum, headerCol).setValue(fileUrl)
  Logger.log(fileName)
  return fileName
}

function highligthRow(rowNum: number){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Processos')
  const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn())
  row.setBackground('#76D7C4')
}

function unHighligthRow(rowNum: number){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Processos')
  const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn())
  row.setBackground('white')
}

function include(filename: string) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}


interface RowData {
  rowNum: number
  Numero: string
  Especie: string
  Tipo: string
  Caixa: string
  Nome: string
  CPF: string
  Origem: string
  Demanda: string
  Status_Digitalizacao: string
  Data_Digitalizacao: string
  Arquivo: string
}


