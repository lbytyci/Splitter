// const { contextBridge, ipcRenderer } = require('electron');

// contextBridge.exposeInMainWorld('electronAPI', {
//     invoke: (channel, data) => ipcRenderer.invoke(channel, data),
//     // Shto komandat e kontrollit të dritares
//     minimize: () => ipcRenderer.send('window-minimize'),
//     maximize: () => ipcRenderer.send('window-maximize'),
//     close: () => ipcRenderer.send('window-close')
// });


const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    invoke: (channel, data) => ipcRenderer.invoke(channel, data),
    minimize: () => ipcRenderer.send('window-minimize'),
    maximize: () => ipcRenderer.send('window-maximize'),
    close: () => ipcRenderer.send('window-close')
});