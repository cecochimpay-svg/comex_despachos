const fs = require('fs');
const path = require('path');

const dbFile = path.join(__dirname, 'comex_despachos_local.json');

let store = {
  local_vista_pallet_despachados: {},
  local_pallet_detalle: {},
  local_pallet_detalle_calibres: {},
  local_eliminados: []
};

// Cargar estado inicial
if (fs.existsSync(dbFile)) {
  try {
    const raw = fs.readFileSync(dbFile, 'utf-8');
    const data = JSON.parse(raw);
    
    ['local_vista_pallet_despachados', 'local_pallet_detalle', 'local_pallet_detalle_calibres'].forEach(tabla => {
      if (Array.isArray(data[tabla])) {
        store[tabla] = {};
        const key = tabla === 'local_vista_pallet_despachados' ? 'id_registro_unico'
                  : tabla === 'local_pallet_detalle' ? 'id_detalle_pallet' 
                  : 'id_detalle_pallet_calibres';
        data[tabla].forEach(row => {
          if (row && row[key] !== undefined) store[tabla][row[key]] = row;
        });
      } else if (data[tabla]) {
        store[tabla] = data[tabla];
      }
    });

    store.local_eliminados = data.local_eliminados || [];
  } catch (e) {
    console.warn('⚠️ [BASE]: Reiniciando base local limpia:', e.message);
  }
}

let timeoutPersistir = null;
function persistirConDebounce() {
  clearTimeout(timeoutPersistir);
  timeoutPersistir = setTimeout(() => {
    try {
      fs.writeFileSync(dbFile, JSON.stringify(store), 'utf-8');
    } catch (err) {
      console.error('❌ [BASE]: Error al persistir en disco:', err.message);
    }
  }, 300);
}

const db = {
  prepare: (sql) => {
    const query = sql.trim().toUpperCase();

    return {
      run: (...params) => {
        const row = params[0] || {};

        if (query.includes('INTO LOCAL_VISTA_PALLET_DESPACHADOS')) {
          const id = Number(row.id_registro_unico || row.id_reg_despacho);
          store.local_vista_pallet_despachados[id] = { ...row, id_registro_unico: id };
          persistirConDebounce();
          return { changes: 1 };
        }

        if (query.includes('INTO LOCAL_PALLET_DETALLE_CALIBRES')) {
          const id = Number(row.id_detalle_pallet_calibres);
          store.local_pallet_detalle_calibres[id] = row;
          persistirConDebounce();
          return { changes: 1 };
        }

        if (query.includes('INTO LOCAL_PALLET_DETALLE')) {
          const id = Number(row.id_detalle_pallet);
          store.local_pallet_detalle[id] = row;
          persistirConDebounce();
          return { changes: 1 };
        }

        return { changes: 0 };
      },

      all: (...params) => {
        if (query.includes('FROM LOCAL_VISTA_PALLET_DESPACHADOS')) {
          return Object.values(store.local_vista_pallet_despachados);
        }
        if (query.includes('FROM LOCAL_PALLET_DETALLE_CALIBRES')) {
          return Object.values(store.local_pallet_detalle_calibres);
        }
        if (query.includes('FROM LOCAL_PALLET_DETALLE')) {
          return Object.values(store.local_pallet_detalle);
        }
        return [];
      }
    };
  }
};

module.exports = { db };