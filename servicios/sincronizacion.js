const { db } = require('../base/base.js');
const { supabaseCampo } = require('./conexion.js');

async function descargarTablaPorLotes(nombreTabla, columnaOrden, onProgreso) {
  let desde = 0;
  const TAMANO_PAGINA = 1000;
  let totalDescargado = 0;
  let hayMasFilas = true;
  let numeroPagina = 1;

  console.log(`📥 [DESCARGA]: Iniciando tabla '${nombreTabla}'...`);

  while (hayMasFilas) {
    const hasta = desde + TAMANO_PAGINA - 1;

    const { data, error } = await supabaseCampo
      .from(nombreTabla)
      .select('*')
      .order(columnaOrden, { ascending: false })
      .range(desde, hasta);

    if (error) {
      console.error(`❌ [${nombreTabla}] Error en rango [${desde}-${hasta}]:`, error.message);
      throw error;
    }

    if (!data || data.length === 0) {
      hayMasFilas = false;
      break;
    }

    const nombreTablaLocal = `local_${nombreTabla}`;
    data.forEach(item => {
      db.prepare(`INSERT INTO ${nombreTablaLocal}`).run(item);
    });

    totalDescargado += data.length;

    if (typeof onProgreso === 'function') {
      onProgreso(`${nombreTabla}: ${totalDescargado.toLocaleString()}`);
    }

    if (data.length < TAMANO_PAGINA) {
      hayMasFilas = false;
    } else {
      desde += TAMANO_PAGINA;
      numeroPagina++;
    }
  }

  console.log(`✅ [${nombreTabla} COMPLETADA]: ${totalDescargado.toLocaleString()} filas.`);
  return totalDescargado;
}

async function sincronizarDespachosCompleto(onProgreso) {
  const inicioT = performance.now();
  const resultado = {
    despachos: 0,
    detalles: 0,
    calibres: 0
  };

  // 1. Tabla Maestra Unificada de Despachos
  resultado.despachos = await descargarTablaPorLotes(
    'vista_pallet_despachados',
    'id_registro_unico',
    onProgreso
  );

  // 2. Pallet Detalle
  resultado.detalles = await descargarTablaPorLotes(
    'pallet_detalle',
    'id_detalle_pallet',
    onProgreso
  );

  // 3. Calibres
  resultado.calibres = await descargarTablaPorLotes(
    'pallet_detalle_calibres',
    'id_detalle_pallet_calibres',
    onProgreso
  );

  const duracion = ((performance.now() - inicioT) / 1000).toFixed(1);
  console.log(`🏁 Sincronización completada en ${duracion}s:`, resultado);

  return resultado;
}

module.exports = {
  sincronizarDespachosCompleto
};