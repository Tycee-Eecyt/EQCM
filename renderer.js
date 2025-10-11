const { contextBridge, ipcRenderer, shell } = require('electron');

contextBridge.exposeInMainWorld('EQCM', {
  getSettings: () => ipcRenderer.invoke('settings:get'),
  setSettings: (payload) => ipcRenderer.invoke('settings:set', payload),
  browseFolder: (which) => ipcRenderer.invoke('settings:browseFolder', which),
  deriveSheetId: (url) => ipcRenderer.invoke('settings:deriveSheetId', url),
  openExternal: (url) => shell.openExternal(url),
  getCovLists: () => ipcRenderer.invoke('cov:getLists'),
  replaceAll: (opts) => ipcRenderer.invoke('advanced:replaceAll', opts||{}),
  replaceAllForce: () => ipcRenderer.invoke('advanced:replaceAll', { force: true }),
  getRaidKit: () => ipcRenderer.invoke('raidkit:get'),
  setRaidKit: (payload) => ipcRenderer.invoke('raidkit:set', payload),
  saveRaidKitAndPush: (payload) => ipcRenderer.invoke('raidkit:saveAndPush', payload),
  getRaidKitCounts: (character) => ipcRenderer.invoke('raidkit:counts', character)
  ,copyPlayersLatest: () => ipcRenderer.invoke('players:copyLatest')
  ,getSheetCharacters: () => ipcRenderer.invoke('favorites:listFromSheet')
});

ipcRenderer.on('settings:discoveredChars', (_e, chars) => {
  window.discoveredChars = chars||[];
  const s = window._lastSettings || {};
  renderFavs(s.favorites||[], window.discoveredChars||[]);
});
