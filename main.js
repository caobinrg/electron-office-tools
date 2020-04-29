const { app, BrowserWindow, Menu, ipcMain, dialog } = require('electron')
const path = require('path')

class AppWindow extends BrowserWindow {
  constructor(config, fileLocation) {
    const basicConfig = {
      icon: './build/icon.ico',
      resizable: false,
      width: 800,
      height: 620,
      webPreferences: {
        nodeIntegration: true
      }
    }
    const finalConfig = { ...basicConfig, ...config }
    super(finalConfig)
    Menu.setApplicationMenu(null)
    this.loadFile(fileLocation)
    this.once('ready-to-show', () => {
      this.show()
    })
  }
}
app.on('ready', () => {
  const mainWindow = new AppWindow({
      maximizable : false,
      fullscreenable : false,
      webPreferences: {
        nodeIntegration: true,
      }
    }, './renderer/index.html')
    // mainWindow.webContents.openDevTools()

  //用renderer执行io操作
  const ioWindow = new AppWindow({
    transparent: true,
    frame: false,
    width: 0,
    height: 0,
    webPreferences: {
      nodeIntegration: true,
    },
    parent: mainWindow
  }, './renderer/io.html')


  //事件转发
  ipcMain.on('eventForward', (event,win,forwardFun,args) => {
    sendByWin[win](forwardFun,args)
  })

  //测试子页面是否收到消息
  ipcMain.on('testMessage',(event,data) => {
    console.log('testmessage:',data)
  })

  //选择文件(单选)
  ipcMain.on('open-file', (event,types,sendMethod) => {
    dialog.showOpenDialog({
      filters: [
        { name: 'File', extensions: types }
      ]
      , properties: ['openFile']
      
    }).then(result => {
      if(!result.canceled){
        event.sender.send(sendMethod, result.filePaths)
      }
    }).catch(err => {
      console.log(err)
    })
  })

    //选择文件夹
    ipcMain.on('open-dir', (event,sendMethod) => {
      dialog.showOpenDialog({
        properties: ['openDirectory','createDirectory','promptToCreate']
      }).then(result => {
        if(!result.canceled){
          event.sender.send(sendMethod, result.filePaths)
        }
      }).catch(err => {
        console.log(err)
      })
    })
    //生成word文件进度
    ipcMain.on('e2wOutByTemplate', (event,planTxt,planPercent,isFinish) => {
      mainWindow.send('e2wOutPlan', planTxt,planPercent,isFinish)
    })

    var sendByWin = {
      name:'sendByWin',
      ioWindow:function(forwardFun,args){
        ioWindow.send(forwardFun,args)
      },
      mainWindow:function(forwardFun,args){
        mainWindow.send(forwardFun,args)
      }
    }
})