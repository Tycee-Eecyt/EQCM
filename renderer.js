const { contextBridge, ipcRenderer, shell } = require('electron');

contextBridge.exposeInMainWorld('EQCM', {
  getSettings: () => ipcRenderer.invoke('settings:get'),
  setSettings: (payload) => ipcRenderer.invoke('settings:set', payload),
  browseFolder: (which) => ipcRenderer.invoke('settings:browseFolder', which),
  deriveSheetId: (url) => ipcRenderer.invoke('settings:deriveSheetId', url),
  openExternal: (url) => shell.openExternal(url),
  getCovLists: () => ipcRenderer.invoke('cov:getLists'),
  forceBackscan: () => ipcRenderer.invoke('advanced:forceBackscan'),
  replaceAll: (opts) => ipcRenderer.invoke('advanced:replaceAll', opts||{}),
  replaceAllForce: () => ipcRenderer.invoke('advanced:replaceAll', { force: true }),
  replaceFactionsCsvFromLocal: () => ipcRenderer.invoke('advanced:replaceFactionsCsv'),
  getRaidKit: () => ipcRenderer.invoke('raidkit:get'),
  setRaidKit: (payload) => ipcRenderer.invoke('raidkit:set', payload),
  saveRaidKitAndPush: (payload) => ipcRenderer.invoke('raidkit:saveAndPush', payload),
  getRaidKitCounts: (character) => ipcRenderer.invoke('raidkit:counts', character)
});

ipcRenderer.on('settings:discoveredChars', (_e, chars) => {
  window.discoveredChars = chars||[];
  const s = window._lastSettings || {};
  renderFavs(s.favorites||[], window.discoveredChars||[]);
});
