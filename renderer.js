const { contextBridge, ipcRenderer, shell } = require('electron');

contextBridge.exposeInMainWorld('EQCM', {
  getSettings: () => ipcRenderer.invoke('settings:get'),
  setSettings: (payload) => ipcRenderer.invoke('settings:set', payload),
  browseFolder: (which) => ipcRenderer.invoke('settings:browseFolder', which),
  deriveSheetId: (url) => ipcRenderer.invoke('settings:deriveSheetId', url),
  openExternal: (url) => shell.openExternal(url),
  getCovLists: () => ipcRenderer.invoke('cov:getLists')
});

ipcRenderer.on('settings:discoveredChars', (_e, chars) => {
  window.discoveredChars = chars||[];
  const s = window._lastSettings || {};
  renderFavs(s.favorites||[], window.discoveredChars||[]);
});
