const { contextBridge, ipcRenderer, shell } = require('electron');

contextBridge.exposeInMainWorld('EQCM', {
  getSettings: () => ipcRenderer.invoke('settings:get'),
  setSettings: (payload) => ipcRenderer.invoke('settings:set', payload),
  browseFolder: (which) => ipcRenderer.invoke('settings:browseFolder', which),
  deriveSheetId: (url) => ipcRenderer.invoke('settings:deriveSheetId', url),
  testWebhook: () => ipcRenderer.invoke('webhook:test'),
  openExternal: (url) => shell.openExternal(url)
});

ipcRenderer.on('settings:discoveredChars', (_e, chars) => {
  window.discoveredChars = chars||[];
  const s = window._lastSettings || {};
  renderFavs(s.favorites||[], window.discoveredChars||[]);
});
