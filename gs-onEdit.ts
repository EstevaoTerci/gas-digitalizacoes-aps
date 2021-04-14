function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit){
    // Set a comment on the edited cell to indicate when it was changed.
    Logger.log(e.range.getColumn())
    
    const sheet = e.range.getSheet()
    const sheetName = sheet.getName()
    
    if(sheetName === 'Processos'){
        
        const col = e.range.getColumn()
        const row = e.range.getRow()
        let rowData: RowData
    
        switch(col){
            case 1: //Coluna Numero
                rowData = sheetDataObject().filter(a => a['rowNum'] === +row )[0]
                if(!!rowData.Tipo && !!rowData.Numero){
                    let newRowData = normalizaNumeroProcesso(rowData)
                    e.range.setValue(newRowData.Numero)
                }
            break;

            case 2: //Coluna especie
                rowData = sheetDataObject().filter(a => a['rowNum'] === +row )[0]
                if(!!rowData.Tipo && !!rowData.Especie){
                    let newRowData = normalizaNumeroProcesso(rowData)
                    sheet.getRange(row, 2, 1, 1).setValue(newRowData.Especie)
                }
            break;

            case 3: //Coluna Tipo
                rowData = sheetDataObject().filter(a => a['rowNum'] === +row )[0]
                if(!!rowData.Tipo && !!rowData.Numero){
                    let newRowData = normalizaNumeroProcesso(rowData)
                    sheet.getRange(row, 1, 1, 1).setValue(newRowData.Numero)
                }
            break;
        }
    }

  }

