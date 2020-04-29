const { ipcRenderer } = require('electron')
const path = require('path')
const {getExcelSheet,readExcelHeader,outByTemplate} = require('./office')

layui.use(['form','layer','layedit', 'laydate','element'], function(){
    var $ = layui.jquery
    ,element = layui.element //Tab的切换功能，切换事件监听等，需要依赖element模块
    ,form = layui.form
    ,layer = layui.layer
    ,layedit = layui.layedit
    ,laydate = layui.laydate;
    var loading = -1

  //excel文件选择
  layui.$('#excelSelectBtn').on('click',function(){
    ipcRenderer.send('open-file',['xlsx'],'e2wExcelSelect')
  })
  ipcRenderer.on('e2wExcelSelect', (event, path) => {
      var filePath = path[0]
      form.val('e2wForm',{"excelPath":filePath})
  })

  //读取excel,生成sheets，生成文件名称选项
  layui.$('#sheetReadBtn').on('click',function(){
    var filePath = form.val('e2wForm')['excelPath'];
    if(filePath){
      //生成sheets
      var sheets = getExcelSheet(filePath)
      var sheetsHtml = ``
      $.each(sheets,function(index,val){
        sheetsHtml += `<option value="${index}">${val}</option>`
      })
      $('#sheet').empty().append(sheetsHtml)
      form.render('select');
      sheetChange(filePath,0)
    }else{
      layer.msg("请先选择excel文件")
    }
  })

  //excel-word word模板选择
  layui.$('#modelSelectBtn').on('click',function(){
    ipcRenderer.send('open-file',['docx'],'e2wModelSelect')
  })
  ipcRenderer.on('e2wModelSelect', (event, path) => {
      form.val('e2wForm',{"modelPath":path[0]})
  })
   //excel-word 输出路径
   layui.$('#outSelectBtn').on('click',function(){
    ipcRenderer.send('open-dir','outPathSelect')
  })
  ipcRenderer.on('outPathSelect', (event, path) => {
      form.val('e2wForm',{"outPath":path[0]})
  })
  //excel-word PDF输出路径
  layui.$('#outPDFSelectBtn').on('click',function(){
    ipcRenderer.send('open-dir','outPDFPathSelect')
  })
  ipcRenderer.on('outPDFPathSelect', (event, path) => {
    form.val('e2wForm',{"outPDFPath":path[0]})
  })
  //excel-word 监听sheet选择
  form.on('select(sheet)', function(data){
    var filePath = form.val('e2wForm')['excelPath'];
    var sheet = data.value
    sheetChange(filePath,sheet)
  })

  //excel-word开始任务
  layui.$('#e2wSubmit').on('click',function(){
    var data = form.val('e2wForm');
    var excelFilePath = data['excelPath']
    var sheetIndex = data['sheet']
    var templateFilePath = data['modelPath']
    var outPath = data['outPath']
    var firstNameField = data['wordName']
    var lastName = data['lastName']
    var outPDFPath = data['outPDFPath']
    if(excelFilePath && sheetIndex && sheetIndex && templateFilePath && outPath && firstNameField){
      loading = layer.load(2)
      ipcRenderer.send('eventForward','ioWindow','outByTemplate',[excelFilePath,sheetIndex,templateFilePath,outPath,firstNameField,lastName,outPDFPath])
    }else{
      layer.msg('请检查输入数据')
    }
  })

  //excel-word任务进度
  ipcRenderer.on('e2wOutPlan', (event, planTxt,planPercent,isFinish) => {
    $('#planTxt')[0].innerHTML = planTxt
    element.progress('planPercent', planPercent)
    if(isFinish){
      layer.close(loading)
    }
  })
  
  //word-pdf 输入路径
  layui.$('#w2pInputSelectBtn').on('click',function(){
    ipcRenderer.send('open-dir','w2pInputSelect')
  })
  ipcRenderer.on('w2pInputSelect', (event, path) => {
      form.val('w2pForm',{"w2pInputPath":path[0]})
  })
  //word-pdf 输出路径
  layui.$('#w2pOutSelectBtn').on('click',function(){
    ipcRenderer.send('open-dir','w2pOutSelect')
  })
  ipcRenderer.on('w2pOutSelect', (event, path) => {
      form.val('w2pForm',{"w2pOutPath":path[0]})
  })
  //word-pdf开始任务
  layui.$('#w2pSubmit').on('click',function(){
    var data = form.val('w2pForm');
    var w2pInputPath = data['w2pInputPath']
    var w2pOutPath = data['w2pOutPath']
    if(w2pInputPath && w2pOutPath){
      loading = layer.load(2)
      ipcRenderer.send('eventForward','ioWindow','w2pTransform',[w2pInputPath,w2pOutPath])
    }else{
      layer.msg('请检查输入数据')
    }
  })
  //word-pdf任务进度
   ipcRenderer.on('w2pTransformPlan', (event,data) => {
    var planTxt = data[0]
    var planPercent = data[1]
    var isFinish = data[2]
    $('#w2pPlanTxt')[0].innerHTML = planTxt
    element.progress('w2pPlanPercent', planPercent)
    if(isFinish == 'true'){
      layer.close(loading)
    }
  })
  
  function sheetChange(filePath,sheet){
    //生成文件名称选项
    var header = readExcelHeader(filePath,sheet)
    var wordNameHtml = ``
    $.each(header,function(index,val){
      wordNameHtml += `<option value="${val}">${val}</option>`
    })
    $('#wordName').empty().append(wordNameHtml)
    form.render('select');
  }
});

