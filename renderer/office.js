const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const pizzzip = require('pizzip');
const docxTemplater = require('docxtemplater');
const { ipcRenderer } = require('electron')
const exec = require('child_process').exec
const execSync = require('child_process').execSync;

exports.getExcelSheet = (filePath) => {
    var workbook = xlsx.readFile(filePath)
    const sheetNames = workbook.SheetNames;
    return sheetNames
}

exports.readExcelHeader = (filePath,sheetIndex) => {
    var workbook = xlsx.readFile(filePath)
    const sheetNames = workbook.SheetNames;
    const worksheet = workbook.Sheets[sheetNames[sheetIndex]];
    var data =xlsx.utils.sheet_to_json(worksheet)[0];
    var header = [];
    for(var key in data){
        header.push(key)
    }
    return header
}

exports.outByTemplate = (excelFilePath,sheetIndex,
    templateFilePath,outPath,firstNameField,lastName,outPDFPath) =>{
    var workbook = xlsx.readFile(excelFilePath)
    const sheetNames = workbook.SheetNames;
    const worksheet = workbook.Sheets[sheetNames[sheetIndex]];
    var data =xlsx.utils.sheet_to_json(worksheet);
    var content = fs.readFileSync(templateFilePath,'binary')
    var zip = new pizzzip(content)
    var doc = new docxTemplater();
    doc.loadZip(zip)
    var length = 0
    $.each(data,function(){
        length++
    })
    if(outPDFPath){
        length *= 2
    }
    var faileWordCount = 0
    var failePdfCount = 0
    var resultIndex = 0
    var faileStr = '';
    $.each(data,(index,val) => {
        doc.setData(val)
        try {
            doc.render()
        } catch (error) {
            //返回报错信息
        }
        var buf = doc.getZip().generate({type:'nodebuffer'})
        var wordName = val[firstNameField]+lastName+'.docx'
        var wordPath = outPath+path.sep+wordName
        fs.writeFileSync(wordPath,buf)
        resultIndex++
        ipcRenderer.send('e2wOutByTemplate','任务进度：生成文件'+wordName+'（'+resultIndex+'/'+length+'）成功', Math.floor(resultIndex/length*100)+'%',false)
        if(outPDFPath){ //如果需要转换
            var pdfName = val[firstNameField]+lastName+'.pdf'
            var pdfPath = outPDFPath+path.sep+pdfName
            var tranResult = execSync('officeConvert.exe -w2p '+wordPath+' '+pdfPath).toString('utf-8')
            resultIndex++
            if(tranResult == 'true'){//转换成功
                ipcRenderer.send('e2wOutByTemplate','任务进度：转换文件'+wordName+'（'+resultIndex+'/'+length+'）成功', Math.floor(resultIndex/length*100)+'%',false)
            }else{//转换失败
                failePdfCount++
                faileStr += (pdfName+' ')
                ipcRenderer.send('e2wOutByTemplate','任务进度：转换文件'+wordName+'（'+resultIndex+'/'+length+'）失败', Math.floor(resultIndex/length*100)+'%',false)
            }
        }
        if(resultIndex == length){//生成完成
            ipcRenderer.send('e2wOutByTemplate','任务进度：完成，成功:'+(length-failePdfCount-faileWordCount)+
            '个，失败:'+(faileWordCount+failePdfCount)+'个,'+'失败任务：'+(faileStr?faileStr:'无'),
            Math.floor(resultIndex/length*100)+'%',true)
        }

    })
}

exports.w2pTransform = (w2pInputPath,w2pOutPath) => {
    var ext = '.docx'
    filesArr = getFiles(w2pInputPath,ext,[])
    var count = 0
    var length = filesArr.length
    ipcRenderer.send('testMessage','length:'+length)
    var faileCount = 0 
    var faileStr = ''
    if(length == 0){
        ipcRenderer.send('eventForward','mainWindow','w2pTransformPlan',['任务进度：未找到待转换文件','0%','true'])
    }else{
        $.each(filesArr,(index,file) => {
            count++
            var fileBaseName = path.basename(file,ext)
            var filePdfName = w2pOutPath + path.sep + fileBaseName + '.pdf'
            ipcRenderer.send('testMessage','file:'+file)
            ipcRenderer.send('testMessage','filePdfName:'+filePdfName)
            var tranResult = execSync('officeConvert.exe -w2p '+file+' '+filePdfName).toString('utf-8')
            if(tranResult == 'true'){
                ipcRenderer.send('testMessage','转换成功:'+filePdfName)
                ipcRenderer.send('eventForward','mainWindow','w2pTransformPlan',[
                    '任务进度：转换文件'+fileBaseName+'（'+count+'/'+length+'）成功',
                    Math.floor(count/length*100)+'%',
                    'false'])
            }else{
                ipcRenderer.send('testMessage','转换失败:'+filePdfName)
                faileCount++
                faileStr += (fileBaseName + ext + ' ')
                ipcRenderer.send('eventForward','mainWindow','w2pTransformPlan',[
                    '任务进度：转换文件'+fileBaseName+'（'+count+'/'+length+'）失败',
                    Math.floor(count/length*100)+'%',
                    'false'])
            }
            if(count == length){
                ipcRenderer.send('testMessage','转换完成')
                ipcRenderer.send('eventForward','mainWindow','w2pTransformPlan',[
                '任务进度：完成，成功:'+(length-faileCount)+
                '个，失败:'+faileCount+'个,'+'失败任务：'+(faileStr?faileStr:'无'),
                Math.floor(count/length*100)+'%','true'])
            }
        })
    }

}

//获取文件夹下所有相关文件
function getFiles(url, ext,filesArr) {
    fs.readdirSync(url).forEach(file=>{
        fileAllName = path.join(url, file)
        if(path.extname(fileAllName) === ext){
            filesArr.push(fileAllName)
        }
    })
    return filesArr
}