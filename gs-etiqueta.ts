function showSideBarEtiqueta(){
    Logger.log('Running showSideBarEtiqueta...')
    let html = HtmlService.createTemplateFromFile('Etiquetas.html')
    let htmlOutput = html.evaluate().setTitle('Gerar etiqueta para caixa arquivo').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getCaixas(){
    Logger.log('Running getCaixas...')
    const sheetData: RowData[] = sheetDataObject()
    const arrCaixa = sheetData.map(a => a.Caixa)
    const data =  arrCaixa.reduce((a, b)=> { //Filtra as caixas únicas
        return [...a].includes(b) ? a : [...a, b]
    }, [arrCaixa[0]])
    return JSON.stringify(data)
}

function getBeneficiosPorCaixas(caixa: any){
    Logger.log('Running getBeneficiosPorCaixas...')
    const dadosBeneficios = sheetDataObject().filter(a => +a.Caixa === +caixa)
    const sorted = dadosBeneficios.sort((a,b) => {
        let numA = String(a.Numero).replace(/[^\d]+/g,'')
        let numB = String(b.Numero).replace(/[^\d]+/g,'')
        return +numA - +numB
    })
    return JSON.stringify(sorted)
}

function getQrCode(caixa: string){
    Logger.log('Running getQrCode...')
    const urlAwTable = `https://app.awesome-table.com/-LKBEJWIbUJE551Ql_HR/view?filterD=${caixa}`
    const apiQrCall = `https://chart.googleapis.com/chart?chs=100x100&cht=qr&chl=${urlAwTable}`
    return UrlFetchApp.fetch(apiQrCall).getAs('image/png')
}

function geraEtiqueta(caixa: string){
    Logger.log('Running geraEtiqueta...')
    const qrCodeImage = expBackoff(() => getQrCode(caixa))
    const sheetEtiqueta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Etiqueta')
    
    //Cria as stings da listagem de benefícios da caixa e dá classifica em ordem crescente de NBs
    const dadosBeneficios = sheetDataObject().filter(a => +a.Caixa === +caixa)
    const sorted = dadosBeneficios.sort((a,b) => {
        let numA = String(a.Numero).replace(/[^\d]+/g,'')
        let numB = String(b.Numero).replace(/[^\d]+/g,'')
        return +numA - +numB
    })
    const arrEspecieNbNome = sorted.map(d => `${d?.Especie}/${d.Numero} - ${d.Nome}`)

    //Organiza os Nbs para serem apresentados em duas colunas na etiqueta
    const numColunas = 2 
    const maxNbsPorColuna = Math.ceil(arrEspecieNbNome.length / numColunas)
    const arrColunas = separar(arrEspecieNbNome, maxNbsPorColuna)
    //Flateia o array de dados das colunas
    const flatArrColunas = arrColunas.map(arr => {
        return arr.reduce((a,b) => a.concat('\n' + b))
    })
    const dadosColuna1 = flatArrColunas[0]
    const dadosColuna2 = flatArrColunas[1]

    //Busca os demais dados da etiqueta
    const dadosEtiqueta = sheetDataObjectGenerica('Caixas').filter(a => +a['Caixa'] === +caixa)[0]
    //Remove as imagens que já estejam na sheet e insere o novo qrCode
    sheetEtiqueta.getImages().forEach(e => e.remove())
    sheetEtiqueta.insertImage(qrCodeImage, 2, 6, 250, 0)
    //Insere as listagens de benefícios nas colunas
    sheetEtiqueta.getRange('B7').setValue(dadosColuna1)
    sheetEtiqueta.getRange('D7').setValue(dadosColuna2)
    //Insere os demais dados da etiqueta
    sheetEtiqueta.getRange('B3').setValue(`Caixa \n ${dadosEtiqueta['Caixa']}`)
    sheetEtiqueta.getRange('B4').setValue(`Assunto \n ${dadosEtiqueta['Assunto']}`)
    sheetEtiqueta.getRange('B5').setValue(`Estante \n ${dadosEtiqueta['Estante']}`)
    //Torna a sheet etiqueta visível e seleciona a área de impressão
    sheetEtiqueta.activate()
    sheetEtiqueta.setActiveSelection('B2:E12')
}


/**
 * @description Recebe um array de para dividí-lo em subarrays com um determinada quantidade de itens
 * @param base O array com os dados para serem separados em grupos menores (subarrays)
 * @param maximo A quantidade máxima de itens que um subarray pode ter
 * @returns {Array} Um array de subarrays [[...], [...]]
 */
function separar(base, maximo) {
    var resultado = [[]];
    var grupo = 0;
  
    for (var indice = 0; indice < base.length; indice++) {
      if (resultado[grupo] === undefined) {
        resultado[grupo] = [];
      }
  
      resultado[grupo].push(base[indice]);
  
      if ((indice + 1) % maximo === 0) {
        grupo = grupo + 1;
      }
    }
  
    return resultado;
  }


interface Etiqueta {
    caixaNum: string
    assunto: string
    estante: string
    data: string
}