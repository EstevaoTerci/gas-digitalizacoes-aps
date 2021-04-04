const ONE_SECOND = 1000;
const ONE_MINUTE = ONE_SECOND * 60;
const START_TIME = Date.now();
const MAX_RETRIES = 5;

/**
 * @description Faz um retry relativamente aleatório se a função retornar erro.
 * @param func A função que será submetida aos retrys
 * @returns {any} O retorno da função
 */
const expBackoff = (func: Function): any => {
  for (let n = 0; n <= MAX_RETRIES; n += 1) {
    try {
      return func();
    } catch (e) {
      if (n === MAX_RETRIES) {
        throw e;
      }
      Utilities.sleep(2 ** n * ONE_SECOND + Math.round(Math.random() * ONE_SECOND));
    }
  }
  return null;
};

/**
 * @description Retorna todos os dados de uma determinada sheet no formato {rowNum: number, ...headers: rowData}
 * @param sheetName O nome da sheet buscada
 * @returns {Object} {rowNum: number, ...headers: rowData}
 */
 const sheetDataObjectGenerica = (sheetName: string) => {
  const sheetData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName).getDataRange().getValues()
  const headers = sheetData.shift()
  const sheetEntries = sheetData.map((arrLinha, i) => {
    let rowObject = {}
    rowObject['rowNum'] = i + 2
    for(let header of headers){
      let indexHeader = headers.indexOf(header)
      rowObject[header] = arrLinha[indexHeader]
    }
    return Object.assign({}, rowObject)
  })
  return sheetEntries
}

