/**
 * Utility Functions
 */

// CSV Parsing
function parseCSV(content) {
    const lines = content.split(/\r?\n/);
    if (!lines.length) return [];
    const header = parseCSVLine(lines[0]);
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const values = parseCSVLine(lines[i]);
        const row = {};
        header.forEach((col, idx) => row[col] = values[idx] || '');
        data.push(row);
    }
    return data;
}

function parseCSVLine(line) {
    const result = [];
    let current = '', inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
            else inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else current += char;
    }
    result.push(current.trim());
    return result;
}

// Math Utilities
function calculateStdDev(arr) {
    if (arr.length < 2) return 0;
    const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
    const squaredDiffs = arr.map(x => Math.pow(x - mean, 2));
    const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / arr.length;
    return Math.sqrt(avgSquaredDiff).toFixed(1);
}

// Toast Notifications
function showToast(message, type = 'info') {
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'toast-container';
        document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const icons = {
        success: '✓',
        error: '✗',
        info: 'ℹ'
    };

    toast.innerHTML = `<span>${icons[type] || 'ℹ'}</span> ${message}`;
    container.appendChild(toast);

    // Animation
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(-100%)';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// IndexedDB Utilities (Configurable)
// Apps can override DB_CONFIG before calling initDB()
let DB_CONFIG = {
    NAME: 'SharedAppDB',
    VERSION: 1,
    STORE_NAME: 'appData'
};
let db = null;

// Configure database - call before initDB()
function configureDB(name, version = 1, storeName = 'appData') {
    DB_CONFIG = { NAME: name, VERSION: version, STORE_NAME: storeName };
    db = null; // Reset connection when config changes
}

async function initDB() {
    if (db) return db; // Reuse existing connection
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_CONFIG.NAME, DB_CONFIG.VERSION);
        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            db = request.result;
            resolve(db);
        };
        request.onupgradeneeded = (e) => {
            const database = e.target.result;
            if (!database.objectStoreNames.contains(DB_CONFIG.STORE_NAME)) {
                database.createObjectStore(DB_CONFIG.STORE_NAME, { keyPath: 'id' });
            }
        };
    });
}

async function saveToIndexedDB(data, id = 'savedState') {
    if (!db) await initDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_CONFIG.STORE_NAME, 'readwrite');
        const store = tx.objectStore(DB_CONFIG.STORE_NAME);
        const request = store.put({
            id: id,
            rawData: data,
            timestamp: new Date().toISOString()
        });
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
    });
}

async function loadFromIndexedDB(id = 'savedState') {
    if (!db) await initDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_CONFIG.STORE_NAME, 'readonly');
        const store = tx.objectStore(DB_CONFIG.STORE_NAME);
        const request = store.get(id);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function clearSavedData(id = 'savedState') {
    if (!db) await initDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_CONFIG.STORE_NAME, 'readwrite');
        const store = tx.objectStore(DB_CONFIG.STORE_NAME);
        const request = store.delete(id);
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
    });
}

async function clearAllData() {
    if (!db) await initDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(DB_CONFIG.STORE_NAME, 'readwrite');
        const store = tx.objectStore(DB_CONFIG.STORE_NAME);
        const request = store.clear();
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
    });
}

// Theme Utilities
function loadTheme() {
    const savedTheme = localStorage.getItem('appTheme');
    if (savedTheme === 'light') {
        document.body.classList.add('light-mode');
    }
    updateThemeIcon();
}

function toggleTheme() {
    document.body.classList.toggle('light-mode');
    const isLight = document.body.classList.contains('light-mode');
    localStorage.setItem('appTheme', isLight ? 'light' : 'dark');
    updateThemeIcon();
}

function updateThemeIcon() {
    const btn = document.getElementById('themeToggle');
    if (btn) {
        btn.textContent = document.body.classList.contains('light-mode') ? '☀️' : '🌙';
    }
}

