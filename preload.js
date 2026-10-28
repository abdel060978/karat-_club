
const { contextBridge, ipcRenderer } = require('electron');

const CANAUX_AUTORISES = new Set([
  'ath:list',
  'ath:save',
  'ath:del',
  'ath:groupes',
  'ath:groupe',

  'doc:list',
  'doc:open',
  'doc:add',

  'xl:model',
  'xl:export',
  'xl:import',

  'row:list',
  'row:save',
  'row:del',

  'saison:ouvrir',
  'ins:set',
  'ins:mode',
  'ins:delai',
  'ins:assur',
  'auth:etat',
  'auth:setup',
  'auth:login',
  'auth:logout',
  'auth:changer',
  'auth:recuperer',
  'usr:list',
  'usr:save',
  'usr:reset',
  'jrn:list',
  'lien:ouvrir',
  'dos:set',
  'sauv:export',
  'sauv:restore',
  'list:export',

  'grp:save',
  'grp:del',

  'pay:cancel',
  'pc:save',
  'ann:set',
  'pres:save',

  'param:set',
  'param:get',

  'pres:get',
  'pres:stats'
]);

contextBridge.exposeInMainWorld('api', {
  call: (canal, ...args) => {
    if (!CANAUX_AUTORISES.has(canal)) {
      return Promise.reject(
        new Error(`Canal IPC non autorisé : ${canal}`)
      );
    }

    return ipcRenderer.invoke(canal, ...args);
  }
});
