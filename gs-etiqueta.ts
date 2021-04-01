function showSideBarEtiqueta(){
    Logger.log(getCaixas())
    let html = HtmlService.createTemplateFromFile('Etiquetas.html')
    let htmlOutput = html.evaluate().setTitle('Gerar etiqueta para caixa arquivo').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getCaixas(){
    const sheetData: RowData[] = sheetDataObject()
    const arrCaixa = sheetData.map(a => a.Caixa)
    return arrCaixa.reduce((a, b)=> { //Filtra as caixas únicas
        return [...a].includes(b) ? a : [...a, b]
    }, [arrCaixa[0]])
}

function getBeneficiosPorCaixas(caixa: string){
    Logger.log(sheetDataObject().filter(a => +a.Caixa === +caixa))
    return sheetDataObject().filter(a => +a.Caixa === +caixa)
}

function getQrCode(caixa: string){
    const urlAwTable = `https://app.awesome-table.com/-LKBEJWIbUJE551Ql_HR/view?filterD=${caixa}`
    const apiQrCall = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${urlAwTable}`
    
    const image = UrlFetchApp.fetch(apiQrCall).getAs('image/png')

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Etiqueta')
    sheet.insertImage(image, 1, 1);
}