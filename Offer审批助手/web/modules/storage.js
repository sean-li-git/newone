/* ======================================================================
 * Offer审批助手 · 存储层封装
 * IndexedDB（结构化数据）+ localStorage（配置项）
 * ====================================================================== */

const DB_NAME = 'OfferAdvisorDB';
const DB_VERSION = 1;

const STORES = {
  offers: 'offers',           // 当前 Offer 档案
  historyOffers: 'historyOffers', // 历史已审批 Offer
  profiles: 'profiles',       // 习惯画像条目
  rules: 'rules',             // 规则库（统一 + 个人）
  configs: 'configs',         // 配置表（因子权重、清单等）
  activities: 'activities',   // 活动记录
};

let _db = null;

/**
 * 初始化 IndexedDB
 */
function initDB() {
  return new Promise((resolve, reject) => {
    if (_db) { resolve(_db); return; }
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      // 创建 Object Stores
      if (!db.objectStoreNames.contains(STORES.offers)) {
        db.createObjectStore(STORES.offers, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORES.historyOffers)) {
        const hs = db.createObjectStore(STORES.historyOffers, { keyPath: 'id' });
        hs.createIndex('sliceKey', ['country', 'level', 'channel', 'jobFamily'], { unique: false });
      }
      if (!db.objectStoreNames.contains(STORES.profiles)) {
        db.createObjectStore(STORES.profiles, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORES.rules)) {
        const rs = db.createObjectStore(STORES.rules, { keyPath: 'id' });
        rs.createIndex('scope', 'scope', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORES.configs)) {
        db.createObjectStore(STORES.configs, { keyPath: 'key' });
      }
      if (!db.objectStoreNames.contains(STORES.activities)) {
        const as = db.createObjectStore(STORES.activities, { keyPath: 'id' });
        as.createIndex('timestamp', 'timestamp', { unique: false });
      }
    };
    request.onsuccess = (e) => { _db = e.target.result; resolve(_db); };
    request.onerror = (e) => reject(e.target.error);
  });
}

/**
 * 通用 CRUD 操作
 */
async function dbPut(storeName, data) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).put(data);
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

async function dbGet(storeName, key) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function dbGetAll(storeName) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function dbDelete(storeName, key) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

async function dbClear(storeName) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readwrite');
    tx.objectStore(storeName).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

async function dbCount(storeName) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const req = tx.objectStore(storeName).count();
    req.onsuccess = () => resolve(req.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

/**
 * 按索引查询
 */
async function dbGetByIndex(storeName, indexName, keyValue) {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, 'readonly');
    const idx = tx.objectStore(storeName).index(indexName);
    const req = idx.getAll(keyValue);
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = (e) => reject(e.target.error);
  });
}

/**
 * localStorage 配置管理
 */
const LS_PREFIX = 'offer_advisor_';

function lsGet(key, defaultValue) {
  try {
    const val = localStorage.getItem(LS_PREFIX + key);
    return val !== null ? JSON.parse(val) : defaultValue;
  } catch (e) { return defaultValue; }
}

function lsSet(key, value) {
  try { localStorage.setItem(LS_PREFIX + key, JSON.stringify(value)); } catch (e) {}
}

function lsRemove(key) {
  try { localStorage.removeItem(LS_PREFIX + key); } catch (e) {}
}

/**
 * 活动记录
 */
async function logActivity(type, message) {
  await dbPut(STORES.activities, {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    type,
    message,
    timestamp: new Date().toISOString(),
  });
}

/**
 * 导出全部数据为 JSON（备份）
 */
async function exportAllData() {
  const data = {};
  for (const store of Object.values(STORES)) {
    data[store] = await dbGetAll(store);
  }
  // 加入 localStorage 配置
  data._localStorage = {};
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key.startsWith(LS_PREFIX)) {
      data._localStorage[key] = localStorage.getItem(key);
    }
  }
  data._exportTime = new Date().toISOString();
  data._version = DB_VERSION;
  return data;
}

/**
 * 导入全部数据（还原）
 */
async function importAllData(data) {
  for (const store of Object.values(STORES)) {
    if (data[store] && Array.isArray(data[store])) {
      await dbClear(store);
      for (const item of data[store]) {
        await dbPut(store, item);
      }
    }
  }
  if (data._localStorage) {
    for (const [key, val] of Object.entries(data._localStorage)) {
      localStorage.setItem(key, val);
    }
  }
}
