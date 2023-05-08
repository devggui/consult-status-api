const { app, BrowserWindow } = require('electron');
const path = require('path')

const createWindow = () => {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 600,    
  });
  
  const index = path.join(__dirname + '/index.html')
  // and load the index.html of the app.
  mainWindow.loadFile(index);
  mainWindow.focus();
};

app.whenReady().then(createWindow)