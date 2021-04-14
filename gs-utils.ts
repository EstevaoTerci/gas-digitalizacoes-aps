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


/**
 * @description Recebe uma linha da planilha e devolve
 * a linha com os números normalizados
 */
function normalizaNumeroProcesso(rowData: RowData): RowData{

  if(!rowData.Tipo) return rowData;

  switch (rowData.Tipo){
    case 'Benefício':
      rowData.Numero = formataNb(rowData.Numero)
      rowData.Especie = formataEspecie(rowData.Especie)
      return rowData

    case 'CTC':
      rowData.Numero = formatCtc(rowData.Numero)
      return rowData

    case 'SIPPS':
      rowData.Numero = formatSipps(rowData.Numero)
      return rowData

    case 'Outros':
      return rowData
  }

}

function formataEspecie(num: string | number): string{
  
  const regex: RegExp = /[^\d]/g
  const cleanEsp = String((num + '' || '')).replace(regex, '')
  return cleanEsp.padStart(2, '0')

}

function formataNb(num: string | number): string{
  const regex: RegExp = /[^\d]/g
  const cleanNb = String((num + '' || '')).replace(regex, '')
  const paddedNb = cleanNb.padStart(10, '0')
  
  const subStr1 = paddedNb.substr(0, 3)
  const subStr2 = paddedNb.substr(3, 3)
  const subStr3 = paddedNb.substr(6, 3)
  const subStr4 = paddedNb.substr(9, 1)

  return `${subStr1}.${subStr2}.${subStr3}-${subStr4}`   
}

function formatCtc(num: string | number): string{
  let n = '07001140.1.00063/18-2'
  const regex: RegExp = /[^\d]/g
  const cleanNb = String((num + '' || '')).replace(regex, '')
  const paddedNb = cleanNb.padStart(17, '0')

  const subStr1 = paddedNb.substr(0, 8)
  const subStr2 = paddedNb.substr(8, 1)
  const subStr3 = paddedNb.substr(9, 5)
  const subStr4 = paddedNb.substr(14, 2)
  const subStr5 = paddedNb.substr(16, 1)

  return `${subStr1}.${subStr2}.${subStr3}/${subStr4}-${subStr5}`
}

function formatSipps(num: string | number): string{
  const regex: RegExp = /[^\d]/g
  const cleanNb = String((num + '' || '')).replace(regex, '')
  const paddedNb = cleanNb.padStart(17, '0')

  const subStr1 = paddedNb.substr(0, 5)
  const subStr2 = paddedNb.substr(5, 6)
  const subStr3 = paddedNb.substr(11, 4)
  const subStr4 = paddedNb.substr(15, 2)

  return `${subStr1}.${subStr2}/${subStr3}-${subStr4}`
}

function dataPadrao(): string {
  const mydate = new Date();
  
  let year = mydate.getFullYear();
  if (year < 2000){
    year += (year < 1900) ? 1900 : 0
  }

  let month: any = mydate.getMonth() + 1;
  if (month < 10){
    month = "0" + month;
  }

  let daym: any = mydate.getDate();
  if (daym < 10){
      daym = "0" + daym;
  }
  return daym + "/" + month + "/" + year;
}

function DataPorExtenso(): string{
  const mydate = new Date();

  let year = mydate.getFullYear();
  if (year < 2000){
    year += (year < 1900) ? 1900 : 0
  }
  var day = mydate.getDay();
  var month = mydate.getMonth();
  var daym: any = mydate.getDate();
  if (daym < 10){
      daym = "0" + daym;
  }
  var dayarray = new Array("Domingo","Segunda-feira","Terça-feira","Quarta-feira","Quinta-feira","Sexta-feira","Sábado");
  var montharray = new Array(" de Janeiro de "," de Fevereiro de "," de Março de ","de Abril de ","de Maio de ","de Junho de","de Julho de ","de Agosto de ","de Setembro de "," de Outubro de "," de Novembro de "," de Dezembro de ");
  var dataExtenso = ("   "+dayarray[day]+", "+daym+" "+montharray[month]+year+" ");
 
  return dataExtenso;

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
  Anotacoes: string
}
