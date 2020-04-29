const { ipcRenderer } = require('electron')
const { outByTemplate,w2pTransform } = require('./office')

ipcRenderer.on('outByTemplate', (event, data) => {
    var excelFilePath = data[0]
    var sheetIndex = data[1]
    var templateFilePath = data[2]
    var outPath = data[3]
    var firstNameField = data[4]
    var lastName = data[5]
    var outPDFPath = data[6]
    outByTemplate(excelFilePath,sheetIndex,templateFilePath,outPath,firstNameField,lastName,outPDFPath)
})

ipcRenderer.on('w2pTransform',(event,data) => {
    var w2pInputPath =  data[0]
    var w2pOutPath = data[1]
    w2pTransform(w2pInputPath,w2pOutPath)
})