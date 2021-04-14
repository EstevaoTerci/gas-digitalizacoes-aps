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
        .map(rowData => normalizaNumeroProcesso(rowData))
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
            .map(rowData => normalizaNumeroProcesso(rowData))
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

    // Monta o payload da etiqueta
    const paramsEtiqueta: Etiqueta = {
        assunto: dadosEtiqueta['Assunto'],
        caixaNum: dadosEtiqueta['Caixa'],
        data: dataPadrao(),
        estante: dadosEtiqueta['Estante'],
        coluna1: dadosColuna1,
        coluna2: dadosColuna2
    }

    // Cria o doc da etiqueta e retorna o url
    const urlNovaEtiqueta = novaEtiqueta(paramsEtiqueta)

    insereQrCode(urlNovaEtiqueta, qrCodeImage)

    return urlNovaEtiqueta

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

  /**
   * Insere o qrCode na respectiva célula da table
   * @param urlDoc Url do google doc da etiqueta
   * @param image Blob da imagem do qrCode
   */
  function insereQrCode(urlDoc: string, image: GoogleAppsScript.Base.Blob){
      Logger.log('Running insereQrCode...')
      
      const tables = DocumentApp.openByUrl(urlDoc).getBody().getTables()
      
      let style = {}
      style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      
      return tables[0].getRow(3).getCell(1).insertImage(1, image).getParent()
        .setAttributes(style)
  } 


  /**
   * 
   * @param paramsEtiqueta Objeto contendo os dados da etiqueta
   * @returns Url do google doc gerado
   */
  function novaEtiqueta(paramsEtiqueta: Etiqueta): string {
    const docModeloId = '10LOtstqbFz10sQR9HD8uF1yorrdppjbZUhXPeMXNHco'
    const folderEtiquetasId = '1ZOWW9tEEnulQnMFSo7HIgrB2UcIbYK_K'

    const folderEtiquetasRef = DriveApp.getFolderById(folderEtiquetasId)
    
    const novoDocUrl = DriveApp.getFileById(docModeloId)
        .makeCopy(`Caixa n. ${paramsEtiqueta.caixaNum}`, folderEtiquetasRef).getUrl();

    const body = DocumentApp.openByUrl(novoDocUrl).getBody();

    for (var key in paramsEtiqueta) {
        body.replaceText(`<<${key}>>`, paramsEtiqueta[key]);
    }

    return novoDocUrl;
}

interface Etiqueta {
    caixaNum: string;
    assunto: string;
    estante: string;
    data: string;
    coluna1: string;
    coluna2: string;
}