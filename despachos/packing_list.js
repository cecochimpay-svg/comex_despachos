/* =============================================================================
   PACKING LIST · CECO S.A.
   Genera el Packing List (Exportación o Mercado Interno) con el formato oficial
   de la empresa, en Excel (.xlsx real, sin librerías externas: funciona offline)
   y en PDF (hoja de impresión → "Guardar como PDF").

   Uso desde la página:
     const pl = PackingList.construir(registrosDeLaVista, opciones);
     PackingList.descargarExcel(pl);
     PackingList.abrirPdf(pl);
   ============================================================================= */
(function (root) {
  'use strict';

  const LOGO_JPG_B64 = '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';
  const LOGO_PX = { w: 130, h: 107 };

  const EMPRESA = {
    nombre: 'CECO S.A',
    direccion1: 'Ruta 22 - Km - Chimpay 8364',
    direccion2: 'Rio Negro - Argentina',
    cuit: '30-68958736-0',
    ggn: '4050373290160',
    pie: [
      'CECO S.A. - Establecimiento Ruta 22 Km 1050 CHIMPAY - Rio Negro - ARGENTINA',
      'Tel: +54 2946 410901',
      'info@cecoargentina.com.ar',
    ],
  };

  // Paleta tomada del logo (bordó de las cerezas)
  const C = {
    bordo: 'A62A28',
    bordoSuave: 'F7E8E7',
    rotulo: 'F2F0EF',
    banda: 'FBF8F7',
    borde: '7F7F7F',
    gris: '595959',
    negro: '1A1A1A',
  };

  // ---------------------------------------------------------------------------
  // UTILIDADES
  // ---------------------------------------------------------------------------
  const txt = v => (v === null || v === undefined) ? '' : String(v).trim();

  function normalizarCalibre(calibreRaw) {
    if (!calibreRaw && calibreRaw !== 0) return '';
    const raw = String(calibreRaw).trim().toUpperCase();
    if (raw.includes('+')) return raw.replace(/\s+/g, '');
    if (raw.includes('-')) return raw.replace(/\s+/g, '');
    const num = parseInt(raw, 10);
    if (!isNaN(num)) return `${num}-${num + 2}`;
    return raw;
  }

  // Orden natural de calibres: S/C primero, luego por milímetros, los "+" al final de su rango
  function ordenCalibre(a, b) {
    const k = c => {
      const m = String(c).match(/(\d+)/);
      if (!m) return [-1, 0, String(c)];
      return [parseInt(m[1], 10), String(c).includes('+') ? 1 : 0, String(c)];
    };
    const ka = k(a), kb = k(b);
    return ka[0] - kb[0] || ka[1] - kb[1] || ka[2].localeCompare(kb[2]);
  }

  // "Cja 2.5 k" → 2.5 · "5 KG" → 5 · "2,5" → 2.5
  function pesoCaja(envase) {
    const m = txt(envase).replace(',', '.').match(/(\d+(?:\.\d+)?)/);
    if (!m) return null;
    const n = parseFloat(m[1]);
    return n > 0 && n < 50 ? n : null;
  }

  function fechaIso(f) {
    if (!f) return '';
    if (f instanceof Date) {
      return `${f.getFullYear()}-${String(f.getMonth() + 1).padStart(2, '0')}-${String(f.getDate()).padStart(2, '0')}`;
    }
    const s = String(f);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    return s;
  }

  function fechaVista(f) {
    const iso = fechaIso(f);
    const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso;
  }

  function numeroSiEsEntero(v) {
    const s = txt(v);
    return /^\d{1,9}$/.test(s) ? Number(s) : s;
  }

  const fmtMiles = n => Math.round(Number(n) || 0).toLocaleString('es-AR');
  const fmtKg = n => (Number(n) || 0).toLocaleString('es-AR', { maximumFractionDigits: 1 });
  const sinTildes = s => txt(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  const nombreArchivoSeguro = s => sinTildes(s).replace(/[^A-Za-z0-9_-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');

  // ---------------------------------------------------------------------------
  // 1. ARMADO DE DATOS
  // ---------------------------------------------------------------------------
  /**
   * registros: filas de vista_pallet_despachados (una fila por pallet + calibre).
   * op: { tipo:'EXPO'|'MI', cliente, fecha, remitos[], vuelo, awb, aerolinea, etd, eta,
   *       destino, ggn, producto, categoria, pesoDefecto, chofer, patente, telefono,
   *       transporte, observaciones }
   */
  function construir(registros, op = {}) {
    const tipo = op.tipo === 'MI' ? 'MI' : 'EXPO';
    const productoDef = txt(op.producto) || (tipo === 'EXPO' ? 'FRESH CHERRIES' : 'CEREZAS FRESCAS');
    const categoriaDef = txt(op.categoria) || 'CAT.1';
    const pesoDef = Number(op.pesoDefecto) > 0 ? Number(op.pesoDefecto) : 2.5;

    const pallets = new Map();
    const calibres = new Set();
    const infoRemito = new Map();   // remito → { fecha, cliente, vuelo, destino }

    (registros || []).forEach(r => {
      const remito = txt(r.nro_remito_embarque) || 'S/R';
      if (!infoRemito.has(remito)) {
        infoRemito.set(remito, { remito, fecha: fechaIso(r.fecha_salida), cliente: txt(r.cliente), vuelo: txt(r.nro_manifiesto_carga), destino: txt(r.nro_guia_frutihorticola) });
      }
      const nroPallet = txt(r.pallet_integrado) || txt(r.numero_romaneo_interno) || 'S/P';
      const clavePallet = `${remito}||${nroPallet}`;
      let p = pallets.get(clavePallet);
      if (!p) {
        p = { nro: nroPallet, remito, termografos: new Set(), filas: new Map() };
        pallets.set(clavePallet, p);
      }
      const termo = txt(r.pallet_termografo);
      if (termo && termo !== '0') p.termografos.add(termo);

      const peso = pesoCaja(r.envase) || pesoDef;
      const fila = {
        producto: productoDef,
        variedad: txt(r.variedad) || '-',
        peso,
        color: txt(r.color),
        categoria: txt(r.categoria) || categoriaDef,
        marca: txt(r.marca),
      };
      const clave = [fila.variedad, fila.peso, fila.color, fila.categoria, fila.marca].join('|');
      let f = p.filas.get(clave);
      if (!f) { f = { ...fila, sizes: {}, total: 0 }; p.filas.set(clave, f); }

      const cal = normalizarCalibre(r.calibre) || 'S/C';
      const cant = Math.abs(Number(r.cantidad || 0));
      if (!cant) return;
      calibres.add(cal);
      f.sizes[cal] = (f.sizes[cal] || 0) + cant;
      f.total += cant;
    });

    const listaCalibres = Array.from(calibres).sort(ordenCalibre);
    if (!listaCalibres.length) listaCalibres.push('S/C');

    // Orden: remito (por fecha y número) y dentro de cada remito, número de pallet
    const ordenRemitos = Array.from(infoRemito.values())
      .sort((a, b) => (a.fecha || '').localeCompare(b.fecha || '') || a.remito.localeCompare(b.remito, 'es', { numeric: true }))
      .map(x => x.remito);
    const posRemito = new Map(ordenRemitos.map((r, i) => [r, i]));
    const ordenPallet = (a, b) => {
      const dr = posRemito.get(a.remito) - posRemito.get(b.remito);
      if (dr) return dr;
      const na = Number(a.nro), nb = Number(b.nro);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return String(a.nro).localeCompare(String(b.nro));
    };

    const listaPallets = Array.from(pallets.values())
      .map(p => ({
        nro: p.nro,
        remito: p.remito,
        termografos: Array.from(p.termografos),
        filas: Array.from(p.filas.values()).filter(f => f.total > 0)
          .sort((a, b) => a.variedad.localeCompare(b.variedad) || a.peso - b.peso),
      }))
      .filter(p => p.filas.length)
      .sort(ordenPallet)
      .map((p, i) => ({ ...p, item: i + 1 }));

    const totalesDe = lista => {
      const porCal = {};
      let cajas = 0, kilos = 0;
      lista.forEach(p => p.filas.forEach(f => {
        cajas += f.total;
        kilos += f.total * f.peso;
        listaCalibres.forEach(c => { if (f.sizes[c]) porCal[c] = (porCal[c] || 0) + f.sizes[c]; });
      }));
      return { porCal, cajas, kilos, pallets: lista.length, filas: lista.reduce((s, p) => s + p.filas.length, 0) };
    };

    // Resumen por categoría + caja
    const resumenMap = new Map();
    listaPallets.forEach(p => p.filas.forEach(f => {
      const kr = `${f.categoria}|${f.peso}`;
      let rr = resumenMap.get(kr);
      if (!rr) { rr = { categoria: f.categoria, peso: f.peso, sizes: {}, total: 0 }; resumenMap.set(kr, rr); }
      listaCalibres.forEach(c => { if (f.sizes[c]) rr.sizes[c] = (rr.sizes[c] || 0) + f.sizes[c]; });
      rr.total += f.total;
    }));

    // Grupos por remito (para el packing consolidado con subtotales)
    const grupos = ordenRemitos
      .map(r => {
        const lista = listaPallets.filter(p => p.remito === r);
        return { ...infoRemito.get(r), pallets: lista, totales: totalesDe(lista) };
      })
      .filter(g => g.pallets.length);

    const remitosTxt = (op.remitos && op.remitos.length ? op.remitos : grupos.map(g => g.remito).filter(r => r !== 'S/R')).join(' / ');

    return {
      tipo,
      agrupado: !!op.agruparPorRemito && grupos.length > 1,
      cab: {
        fecha: fechaIso(op.fecha) || grupos.map(g => g.fecha).filter(Boolean).sort().pop() || fechaIso(new Date()),
        cliente: txt(op.cliente),
        remitos: remitosTxt,
        vuelo: txt(op.vuelo),
        awb: txt(op.awb),
        aerolinea: txt(op.aerolinea),
        etd: txt(op.etd),
        eta: txt(op.eta),
        destino: txt(op.destino),
        ggn: txt(op.ggn) || EMPRESA.ggn,
        chofer: txt(op.chofer),
        patente: txt(op.patente),
        telefono: txt(op.telefono),
        transporte: txt(op.transporte),
        observaciones: txt(op.observaciones),
      },
      calibres: listaCalibres,
      pallets: listaPallets,
      grupos,
      resumen: Array.from(resumenMap.values()).sort((a, b) => a.categoria.localeCompare(b.categoria) || a.peso - b.peso),
      totales: totalesDe(listaPallets),
    };
  }

  function textos(pl) {
    const expo = pl.tipo === 'EXPO';
    return {
      expo,
      consignee: expo ? 'CONSIGNEE:' : 'CLIENTE',
      remito: 'REMITO',
      resumenTitulo: expo ? 'SHIPMENT SUMMARY' : 'RESUMEN DEL ENVÍO',
      resPallets: expo ? 'Pallets' : 'Pallets',
      resCajas: expo ? 'Boxes' : 'Cajas',
      resKilos: expo ? 'Net weight (kg)' : 'Kilos netos',
      resRemito: expo ? 'Remito' : 'Remito(s)',
      obsTitulo: expo ? 'REMARKS' : 'OBSERVACIONES',
      termoSi: expo ? 'Yes' : 'SI',
      termoNo: expo ? '' : 'NO',
      sufijo: expo ? 'Exportacion' : 'Mercado_Interno',
    };
  }

  function textoTermografo(p, t) {
    if (!p.termografos.length) return t.termoNo;
    return `${t.termoSi} · ${p.termografos.join(', ')}`;
  }

  function nombreArchivo(pl, ext) {
    const t = textos(pl);
    return `PackingList_${t.sufijo}_${nombreArchivoSeguro(pl.cab.cliente) || 'Cliente'}_${pl.cab.fecha}.${ext}`;
  }

  // ---------------------------------------------------------------------------
  // 2. ESCRITOR XLSX MÍNIMO (Office Open XML + ZIP sin compresión)
  // ---------------------------------------------------------------------------
  const letraCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const ref = (col, fila) => letraCol(col) + fila;
  const escXml = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function crearHoja() {
    const celdas = new Map(); // key "fila:col"
    const hoja = {
      celdas, merges: [], anchos: {}, altos: {},
      set(col, fila, valor, estilo) {
        celdas.set(`${fila}:${col}`, { col, fila, valor, estilo: estilo || null });
      },
      // Combina un rango y aplica el estilo a todas sus celdas (así se ven los bordes completos)
      rango(c1, f1, c2, f2, valor, estilo) {
        for (let f = f1; f <= f2; f++) for (let c = c1; c <= c2; c++) hoja.set(c, f, (c === c1 && f === f1) ? valor : null, estilo);
        if (c1 !== c2 || f1 !== f2) hoja.merges.push(`${ref(c1, f1)}:${ref(c2, f2)}`);
      },
    };
    return hoja;
  }

  function registroEstilos() {
    const fonts = ['<font><sz val="10"/><name val="Arial"/></font>'];
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    const numFmts = [];
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const idx = { font: new Map(), fill: new Map(), border: new Map(), numFmt: new Map(), xf: new Map() };
    const builtinFmt = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4 };

    const id = (mapa, lista, xml) => { if (!mapa.has(xml)) { mapa.set(xml, lista.length); lista.push(xml); } return mapa.get(xml); };

    function lado(nombre, est, color) {
      return est ? `<${nombre} style="${est}"><color rgb="FF${color || C.borde}"/></${nombre}>` : `<${nombre}/>`;
    }

    return {
      xf(e) {
        if (!e) return 0;
        const f = e.font || {};
        const fontXml = `<font>${f.b ? '<b/>' : ''}${f.i ? '<i/>' : ''}${f.u ? '<u/>' : ''}<sz val="${f.sz || 9}"/><color rgb="FF${f.color || C.negro}"/><name val="${escXml(f.name || 'Arial')}"/></font>`;
        const fontId = id(idx.font, fonts, fontXml);
        const fillId = e.fill ? id(idx.fill, fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${e.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
        const b = e.border || {};
        const borderId = (b.l || b.r || b.t || b.b)
          ? id(idx.border, borders, `<border>${lado('left', b.l, b.color)}${lado('right', b.r, b.color)}${lado('top', b.t, b.color)}${lado('bottom', b.b, b.color)}<diagonal/></border>`)
          : 0;
        let numFmtId = 0;
        if (e.numFmt) {
          if (builtinFmt[e.numFmt] !== undefined) numFmtId = builtinFmt[e.numFmt];
          else {
            if (!idx.numFmt.has(e.numFmt)) { idx.numFmt.set(e.numFmt, 164 + numFmts.length); numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${escXml(e.numFmt)}"/>`); }
            numFmtId = idx.numFmt.get(e.numFmt);
          }
        }
        const a = e.align || {};
        const alignXml = (a.h || a.v || a.wrap) ? `<alignment${a.h ? ` horizontal="${a.h}"` : ''}${a.v ? ` vertical="${a.v}"` : ''}${a.wrap ? ' wrapText="1"' : ''}/>` : '';
        const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${alignXml}</xf>`;
        return id(idx.xf, xfs, xml);
      },
      xml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
      },
    };
  }

  function hojaXml(hoja, estilos, cfg) {
    const filas = new Map();
    hoja.celdas.forEach(c => { if (!filas.has(c.fila)) filas.set(c.fila, []); filas.get(c.fila).push(c); });
    Object.keys(hoja.altos).forEach(f => { if (!filas.has(Number(f))) filas.set(Number(f), []); });
    const nums = Array.from(filas.keys()).sort((a, b) => a - b);

    const sheetData = nums.map(f => {
      const alto = hoja.altos[f];
      const celdas = filas.get(f).sort((a, b) => a.col - b.col).map(c => {
        const r = ref(c.col, c.fila);
        const s = estilos.xf(c.estilo);
        const v = c.valor;
        const sAttr = s ? ` s="${s}"` : '';
        if (v === null || v === undefined || v === '') return `<c r="${r}"${sAttr}/>`;
        if (typeof v === 'number' && isFinite(v)) return `<c r="${r}"${sAttr}><v>${v}</v></c>`;
        if (typeof v === 'object' && v.f) {
          return `<c r="${r}"${sAttr}><f>${escXml(v.f)}</f>${v.v !== undefined ? `<v>${v.v}</v>` : ''}</c>`;
        }
        return `<c r="${r}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escXml(v)}</t></is></c>`;
      }).join('');
      return `<row r="${f}"${alto ? ` ht="${alto}" customHeight="1"` : ''}>${celdas}</row>`;
    }).join('');

    const cols = Object.keys(hoja.anchos).map(Number).sort((a, b) => a - b)
      .map(c => `<col min="${c}" max="${c}" width="${hoja.anchos[c]}" customWidth="1"/>`).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${ref(cfg.ultimaCol, cfg.ultimaFila)}"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${sheetData}</sheetData>${hoja.merges.length ? `<mergeCells count="${hoja.merges.length}">${hoja.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}<printOptions horizontalCentered="1"/><pageMargins left="0.35" right="0.35" top="0.4" bottom="0.5" header="0.2" footer="0.25"/><pageSetup paperSize="9" orientation="${cfg.horizontal ? 'landscape' : 'portrait'}" fitToWidth="1" fitToHeight="0"/><headerFooter><oddFooter>&amp;L&amp;8CECO S.A. · Packing List&amp;R&amp;8Página &amp;P de &amp;N</oddFooter></headerFooter>${cfg.conImagen ? '<drawing r:id="rId1"/>' : ''}</worksheet>`;
  }

  // ZIP "store" (sin compresión): suficiente y 100% compatible con Excel / LibreOffice / Google Sheets
  const TABLA_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = TABLA_CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function crearZip(archivos) {
    const enc = new TextEncoder();
    const partes = [];
    const central = [];
    let offset = 0;
    archivos.forEach(a => {
      const nombre = enc.encode(a.nombre);
      const datos = typeof a.datos === 'string' ? enc.encode(a.datos) : a.datos;
      const crc = crc32(datos);
      const local = new Uint8Array(30 + nombre.length);
      const dv = new DataView(local.buffer);
      dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0x0800, true); dv.setUint16(8, 0, true);
      dv.setUint16(10, 0, true); dv.setUint16(12, 0x0021, true);
      dv.setUint32(14, crc, true); dv.setUint32(18, datos.length, true); dv.setUint32(22, datos.length, true);
      dv.setUint16(26, nombre.length, true); dv.setUint16(28, 0, true);
      local.set(nombre, 30);
      partes.push(local, datos);

      const cen = new Uint8Array(46 + nombre.length);
      const dc = new DataView(cen.buffer);
      dc.setUint32(0, 0x02014b50, true); dc.setUint16(4, 20, true); dc.setUint16(6, 20, true); dc.setUint16(8, 0x0800, true);
      dc.setUint16(10, 0, true); dc.setUint16(12, 0, true); dc.setUint16(14, 0x0021, true);
      dc.setUint32(16, crc, true); dc.setUint32(20, datos.length, true); dc.setUint32(24, datos.length, true);
      dc.setUint16(28, nombre.length, true); dc.setUint32(42, offset, true);
      cen.set(nombre, 46);
      central.push(cen);
      offset += local.length + datos.length;
    });
    const tamCentral = central.reduce((s, c) => s + c.length, 0);
    const fin = new Uint8Array(22);
    const df = new DataView(fin.buffer);
    df.setUint32(0, 0x06054b50, true); df.setUint16(8, archivos.length, true); df.setUint16(10, archivos.length, true);
    df.setUint32(12, tamCentral, true); df.setUint32(16, offset, true);
    const total = offset + tamCentral + 22;
    const out = new Uint8Array(total);
    let p = 0;
    [...partes, ...central, fin].forEach(b => { out.set(b, p); p += b.length; });
    return out;
  }

  function base64ABytes(b64) {
    if (typeof atob === 'function') {
      const bin = atob(b64); const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    return Uint8Array.from(Buffer.from(b64, 'base64'));
  }

  // hojas: [{ hoja, cfg: { nombreHoja, ultimaCol, ultimaFila, horizontal, conImagen } }]
  function libroXlsx(hojas, meta = {}) {
    const estilos = registroEstilos();
    const emu = px => Math.round(px * 9525);
    const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // Nombres de hoja válidos y únicos (máx. 31, sin []:*?/\)
    const usados = new Set();
    hojas.forEach((h, i) => {
      let base = String(h.cfg.nombreHoja || `Hoja ${i + 1}`).replace(/[\[\]:*?\/\\]/g, '-').slice(0, 31).trim() || `Hoja ${i + 1}`;
      let nombre = base, k = 2;
      while (usados.has(nombre.toLowerCase())) { nombre = `${base.slice(0, 27)} (${k++})`; }
      usados.add(nombre.toLowerCase());
      h.nombreFinal = nombre;
    });

    const sheetsXml = hojas.map(h => hojaXml(h.hoja, estilos, h.cfg));   // registra estilos de todas las hojas
    const conImagen = hojas.some(h => h.cfg.conImagen);

    const archivos = [
      { nombre: '[Content_Types].xml', datos: XML + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${hojas.map((h, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>${h.cfg.conImagen ? `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : ''}`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>` },
      { nombre: '_rels/.rels', datos: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>` },
      { nombre: 'docProps/core.xml', datos: XML + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escXml(meta.titulo || 'Packing List')}</dc:title><dc:creator>CECO S.A.</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>` },
      { nombre: 'xl/workbook.xml', datos: XML + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${hojas.map((h, i) => `<sheet name="${escXml(h.nombreFinal)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets><definedNames>${hojas.map((h, i) => `<definedName name="_xlnm.Print_Area" localSheetId="${i}">'${escXml(h.nombreFinal.replace(/'/g, "''"))}'!$A$1:$${letraCol(h.cfg.ultimaCol)}$${h.cfg.ultimaFila}</definedName>`).join('')}</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
      { nombre: 'xl/_rels/workbook.xml.rels', datos: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${hojas.length + 1}" Type="${REL}/styles" Target="styles.xml"/></Relationships>` },
      ...sheetsXml.map((x, i) => ({ nombre: `xl/worksheets/sheet${i + 1}.xml`, datos: x })),
      { nombre: 'xl/styles.xml', datos: estilos.xml() },
    ];
    hojas.forEach((h, i) => {
      if (!h.cfg.conImagen) return;
      const n = i + 1;
      archivos.push(
        { nombre: `xl/worksheets/_rels/sheet${n}.xml.rels`, datos: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>` },
        { nombre: `xl/drawings/drawing${n}.xml`, datos: XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"><xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>${emu(6)}</xdr:rowOff></xdr:from><xdr:ext cx="${emu(LOGO_PX.w)}" cy="${emu(LOGO_PX.h)}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo CECO"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(LOGO_PX.w)}" cy="${emu(LOGO_PX.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>` },
        { nombre: `xl/drawings/_rels/drawing${n}.xml.rels`, datos: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/image" Target="../media/image1.jpeg"/></Relationships>` },
      );
    });
    if (conImagen) archivos.push({ nombre: 'xl/media/image1.jpeg', datos: base64ABytes(LOGO_JPG_B64) });
    return crearZip(archivos);
  }

  // ---------------------------------------------------------------------------
  // 3. PACKING LIST EN EXCEL (mismo formato que la planilla modelo)
  // ---------------------------------------------------------------------------
  const serialExcel = iso => {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000;
  };

  function llenarHoja(pl) {
    const t = textos(pl);
    const h = crearHoja();
    const n = pl.calibres.length;
    const COL_SIZE = 11;                    // K
    const COL_TOTAL = COL_SIZE + n;         // columna TOTAL BOXES
    const ULT = COL_TOTAL;
    const B = 2, D = 4, F = 6, I = 9, J = 10;

    // --- Estilos
    const fino = { l: 'thin', r: 'thin', t: 'thin', b: 'thin' };
    const sRotulo = { font: { name: 'Arial', sz: 9, b: true, color: C.negro }, fill: C.rotulo, border: fino, align: { h: 'left', v: 'center' } };
    const sValor = { font: { name: 'Arial', sz: 9, b: true }, border: fino, align: { h: 'left', v: 'center' } };
    const sValorNorm = { font: { name: 'Arial', sz: 9 }, border: fino, align: { h: 'left', v: 'center', wrap: true } };
    const sBanda = { font: { name: 'Arial', sz: 9, b: true, color: 'FFFFFF' }, fill: C.bordo, border: fino, align: { h: 'left', v: 'center' } };
    const sBandaCentro = { ...sBanda, align: { h: 'center', v: 'center' } };
    const sFecha = { ...sValor, numFmt: 'dd/mm/yyyy', align: { h: 'center', v: 'center' } };
    const fT = { name: 'Bookman Old Style', sz: 9 };
    const sHead = { font: { ...fT, b: true, color: 'FFFFFF' }, fill: C.bordo, border: { ...fino, color: '5A1716' }, align: { h: 'center', v: 'center', wrap: true } };
    const sCelda = (fill, extra = {}) => ({ font: { ...fT, ...(extra.font || {}) }, fill, border: fino, align: { h: 'center', v: 'center', ...(extra.align || {}) }, numFmt: extra.numFmt });
    const sTotal = { font: { ...fT, b: true }, fill: C.bordoSuave, border: { l: 'thin', r: 'thin', t: 'medium', b: 'medium' }, align: { h: 'center', v: 'center' }, numFmt: '#,##0' };

    // --- Anchos (idénticos a la planilla modelo; se ensancha lo justo para que entren los textos)
    Object.assign(h.anchos, { 1: 3.3, 2: 6, 3: 12.1, 4: 18.9, 5: 15.7, 6: 13, 7: 9, 8: 13.6, 9: 12.3, 10: 13.6 });
    for (let c = COL_SIZE; c < COL_TOTAL; c++) h.anchos[c] = 7.6;
    h.anchos[COL_TOTAL] = 9.4;
    for (let f = 1; f <= 9; f++) h.altos[f] = 15;

    // --- Encabezado: logo (imagen) + título
    h.rango(B, 4, ULT, 5, 'Packing List', { font: { name: 'Calibri', sz: 28, b: true, color: C.negro }, align: { h: 'center', v: 'center' } });
    h.rango(B, 6, ULT, 6, t.expo ? 'Fresh fruit from Patagonia · Argentina' : 'Fruta fresca de la Patagonia · Mercado interno',
      { font: { name: 'Calibri', sz: 10, i: true, color: C.gris }, align: { h: 'center', v: 'center' } });
    for (let c = B; c <= ULT; c++) h.set(c, 9, null, { border: { b: 'medium', color: C.bordo } });

    // --- Bloque de datos (B:F)
    const filaDato = (f, rotulo, valor, estiloValor = sValor) => {
      h.rango(B, f, 3, f, rotulo, sRotulo);
      h.rango(D, f, F, f, valor, estiloValor);
    };
    for (let f = 11; f <= 26; f++) h.altos[f] = 14.5;
    h.rango(B, 11, 3, 11, 'PACKING LIST', sBanda);
    h.set(D, 11, 'DATE', { ...sRotulo, align: { h: 'center', v: 'center' } });
    const serial = serialExcel(pl.cab.fecha);
    h.rango(5, 11, F, 11, serial !== null ? serial : fechaVista(pl.cab.fecha), sFecha);
    filaDato(12, 'SHIPPER:', EMPRESA.nombre);
    h.rango(B, 13, 3, 14, 'Address:', { ...sRotulo, align: { h: 'left', v: 'top' } });
    h.rango(D, 13, F, 13, EMPRESA.direccion1, sValorNorm);
    h.rango(D, 14, F, 14, EMPRESA.direccion2, sValorNorm);
    filaDato(15, 'CUIT:', EMPRESA.cuit, sValorNorm);
    filaDato(16, t.remito, pl.cab.remitos);
    filaDato(17, t.consignee, pl.cab.cliente);
    filaDato(18, 'AWB#', pl.cab.awb);
    filaDato(19, 'Airline:', pl.cab.aerolinea);
    filaDato(20, 'ETD:', pl.cab.etd, sValorNorm);
    filaDato(21, 'ETA:', pl.cab.eta, sValorNorm);
    filaDato(22, 'Destination:', pl.cab.destino);
    filaDato(23, 'Flight N°:', pl.cab.vuelo);
    filaDato(24, 'GGN:', numeroSiEsEntero(pl.cab.ggn), { ...sValor, numFmt: '0' });

    // --- Caja derecha: resumen del envío (I:ULT)
    const filaCaja = (f, rotulo, valor, est = sValor) => {
      h.set(I, f, rotulo, sRotulo);
      h.rango(J, f, ULT, f, valor, est);
    };
    h.rango(I, 11, ULT, 11, t.resumenTitulo, sBandaCentro);
    filaCaja(12, t.resPallets, pl.totales.pallets, { ...sValor, numFmt: '#,##0' });
    filaCaja(13, t.resCajas, pl.totales.cajas, { ...sValor, numFmt: '#,##0' });
    filaCaja(14, t.resKilos, Math.round(pl.totales.kilos * 10) / 10, { ...sValor, numFmt: '#,##0.0' });
    filaCaja(15, t.resRemito, pl.cab.remitos);

    let filaObs = 17;
    if (!t.expo) {
      h.rango(I, 17, ULT, 17, 'DATOS DEL CAMIÓN', sBandaCentro);
      filaCaja(18, 'CHOFER', pl.cab.chofer);
      filaCaja(19, 'PATENTE', pl.cab.patente);
      filaCaja(20, 'TELÉFONO', pl.cab.telefono);
      filaCaja(21, 'TRANSPORTE', pl.cab.transporte);
      filaObs = 22;
    }
    h.rango(I, filaObs, ULT, filaObs, t.obsTitulo, sBandaCentro);
    h.rango(I, filaObs + 1, ULT, 24, pl.cab.observaciones, { ...sValorNorm, align: { h: 'left', v: 'top', wrap: true } });

    // --- Tabla principal
    const F1 = 28, F2 = 29;
    h.altos[F1] = 22; h.altos[F2] = 18;
    const cab = [['ITEM', B], ['PALLET N°', 3], ['PRODUCT', 4], ['VARIETY', 5], ['NET WEIGHT / BOX', 6], ['Color', 7], ['Temperature recorder', 8], ['CATEGORY', 9], ['BRAND', 10]];
    cab.forEach(([titulo, col]) => h.rango(col, F1, col, F2, titulo, sHead));
    h.rango(COL_SIZE, F1, COL_TOTAL - 1, F1, 'SIZES MM', sHead);
    pl.calibres.forEach((c, i) => h.set(COL_SIZE + i, F2, c, sHead));
    h.set(COL_TOTAL, F1, 'TOTAL', sHead);
    h.set(COL_TOTAL, F2, 'BOXES', sHead);

    let f = 30;
    const sGrupo = { font: { ...fT, b: true, color: C.bordo }, fill: C.rotulo, border: { l: 'thin', r: 'thin', t: 'medium', b: 'thin' }, align: { h: 'left', v: 'center' } };
    const sSub = { font: { ...fT, b: true }, fill: 'FCF4F3', border: fino, align: { h: 'center', v: 'center' }, numFmt: '#,##0;-#,##0;""' };
    const grupos = pl.agrupado ? pl.grupos : [{ remito: null, pallets: pl.pallets }];
    const subtotales = [];
    let rangoDatos = null;
    let ip = 0;

    grupos.forEach(g => {
      if (pl.agrupado) {
        h.altos[f] = 17;
        h.rango(B, f, ULT, f, `REMITO ${g.remito}${g.fecha ? '  ·  ' + fechaVista(g.fecha) : ''}${g.vuelo ? '  ·  ' + g.vuelo : ''}${g.destino ? ' ' + g.destino : ''}  ·  ${g.totales.pallets} pallet(s)`, sGrupo);
        f++;
      }
      const gDesde = f;
      g.pallets.forEach(p => {
        const fill = ip++ % 2 ? C.banda : null;
        const desde = f;
        p.filas.forEach(fila => {
          h.altos[f] = 15;
          h.set(4, f, fila.producto, sCelda(fill));
          h.set(5, f, fila.variedad, sCelda(fill));
          h.set(6, f, fila.peso, sCelda(fill));
          h.set(7, f, fila.color, sCelda(fill));
          h.set(9, f, fila.categoria, sCelda(fill));
          h.set(10, f, numeroSiEsEntero(fila.marca), sCelda(fill));
          pl.calibres.forEach((c, i) => {
            const v = fila.sizes[c] || 0;
            h.set(COL_SIZE + i, f, v || null, sCelda(fill, { numFmt: '#,##0' }));
          });
          h.set(COL_TOTAL, f, { f: `SUM(${ref(COL_SIZE, f)}:${ref(COL_TOTAL - 1, f)})`, v: fila.total },
            sCelda(fill, { numFmt: '#,##0', font: { b: true } }));
          f++;
        });
        const hasta = f - 1;
        // Ítem, pallet y termógrafo combinados a lo alto del pallet
        h.rango(B, desde, B, hasta, p.item, sCelda(fill, { font: { b: true } }));
        h.rango(3, desde, 3, hasta, numeroSiEsEntero(p.nro), sCelda(fill, { font: { b: true, color: C.bordo } }));
        h.rango(8, desde, 8, hasta, textoTermografo(p, t), sCelda(fill, { align: { wrap: true } }));
      });
      const gHasta = f - 1;
      if (!rangoDatos) rangoDatos = [gDesde, gHasta]; else rangoDatos[1] = gHasta;

      if (pl.agrupado) {
        h.altos[f] = 15;
        h.rango(B, f, 10, f, `Subtotal remito ${g.remito}`, { ...sSub, align: { h: 'right', v: 'center' } });
        pl.calibres.forEach((c, i) => {
          const col = COL_SIZE + i;
          h.set(col, f, { f: `SUM(${ref(col, gDesde)}:${ref(col, gHasta)})`, v: g.totales.porCal[c] || 0 }, sSub);
        });
        h.set(COL_TOTAL, f, { f: `SUM(${ref(COL_SIZE, f)}:${ref(COL_TOTAL - 1, f)})`, v: g.totales.cajas }, sSub);
        subtotales.push(f);
        f++;
      }
    });
    if (!pl.pallets.length) { for (let c = B; c <= ULT; c++) h.set(c, f, null, sCelda(null)); rangoDatos = [f, f]; f++; }

    // Total general
    const fTot = f;
    h.altos[fTot] = 17;
    h.rango(B, fTot, 10, fTot, pl.agrupado ? `TOTAL GENERAL · ${pl.grupos.length} remitos` : 'TOTAL', { ...sTotal, align: { h: 'right', v: 'center' } });
    pl.calibres.forEach((c, i) => {
      const col = COL_SIZE + i;
      const formula = pl.agrupado ? subtotales.map(r => ref(col, r)).join('+') : `SUM(${ref(col, rangoDatos[0])}:${ref(col, rangoDatos[1])})`;
      h.set(col, fTot, { f: formula, v: pl.totales.porCal[c] || 0 }, sTotal);
    });
    h.set(COL_TOTAL, fTot, { f: `SUM(${ref(COL_SIZE, fTot)}:${ref(COL_TOTAL - 1, fTot)})`, v: pl.totales.cajas }, { ...sTotal, font: { ...fT, b: true, color: C.bordo } });

    // --- Resumen de calibres por categoría y caja
    let r = fTot + 3;
    h.altos[r] = 16; h.altos[r + 1] = 16;
    h.rango(I, r, I, r + 1, 'CATEGORY', sHead);
    h.rango(J, r, J, r + 1, 'BOX', sHead);
    h.rango(COL_SIZE, r, COL_TOTAL - 1, r, 'SIZES SUMMARY', sHead);
    pl.calibres.forEach((c, i) => h.set(COL_SIZE + i, r + 1, c, sHead));
    h.set(COL_TOTAL, r, 'TOTAL', sHead);
    h.set(COL_TOTAL, r + 1, 'BOXES', sHead);
    r += 2;
    const fRes0 = r;
    pl.resumen.forEach(rr => {
      h.set(I, r, rr.categoria, sCelda(null, { font: { b: true } }));
      h.set(J, r, `${String(rr.peso).replace('.', ',')} kg`, sCelda(null));
      pl.calibres.forEach((c, i) => h.set(COL_SIZE + i, r, rr.sizes[c] || null, sCelda(null, { numFmt: '#,##0' })));
      h.set(COL_TOTAL, r, { f: `SUM(${ref(COL_SIZE, r)}:${ref(COL_TOTAL - 1, r)})`, v: rr.total }, sCelda(null, { numFmt: '#,##0', font: { b: true } }));
      r++;
    });
    if (pl.resumen.length > 1) {
      h.rango(I, r, J, r, 'TOTAL', { ...sTotal, align: { h: 'right', v: 'center' } });
      pl.calibres.forEach((c, i) => {
        const col = COL_SIZE + i;
        h.set(col, r, { f: `SUM(${ref(col, fRes0)}:${ref(col, r - 1)})`, v: pl.totales.porCal[c] || 0 }, sTotal);
      });
      h.set(COL_TOTAL, r, { f: `SUM(${ref(COL_SIZE, r)}:${ref(COL_TOTAL - 1, r)})`, v: pl.totales.cajas }, sTotal);
      r++;
    }

    // --- Firmas (mercado interno: entrega en camión)
    if (!t.expo) {
      r += 3;
      h.altos[r] = 15;
      const linea = { border: { t: 'thin', color: C.negro }, font: { name: 'Arial', sz: 8, b: true, color: C.gris }, align: { h: 'center', v: 'top' } };
      h.rango(3, r, 5, r, 'Despachó (firma y aclaración)', linea);
      h.rango(8, r, 10, r, 'Recibí conforme · Chofer', linea);
      r += 1;
    }

    // --- Pie de empresa (como la planilla modelo)
    r += 3;
    const sPie = { font: { name: 'Arial', sz: 8, color: C.gris } };
    EMPRESA.pie.forEach((linea, i) => h.set(B, r + i, linea, i === 2 ? { font: { name: 'Arial', sz: 8, color: '0563C1', u: true } } : sPie));
    const ultimaFila = r + EMPRESA.pie.length - 1;

    return {
      hoja: h,
      cfg: { ultimaCol: ULT, ultimaFila, horizontal: n > 9, conImagen: true, nombreHoja: pl.nombreHoja || 'CECO' },
    };
  }

  function generarExcel(pl) {
    return libroXlsx([llenarHoja(pl)], { titulo: `Packing List ${pl.cab.cliente} ${fechaVista(pl.cab.fecha)}` });
  }

  // Un libro con una hoja por packing (p. ej. una por remito)
  function generarExcelMulti(pls) {
    return libroXlsx(pls.map(llenarHoja), { titulo: `Packing List · ${pls.length} remitos` });
  }

  function descargarBytes(bytes, nombre, mime) {
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nombre; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  function descargarExcel(pl) {
    const nombre = nombreArchivo(pl, 'xlsx');
    descargarBytes(generarExcel(pl), nombre, MIME_XLSX);
    return nombre;
  }

  function nombreArchivoMulti(pls, ext) {
    if (pls.length === 1) return nombreArchivo(pls[0], ext);
    const t = textos(pls[0]);
    const clientes = Array.from(new Set(pls.map(p => p.cab.cliente).filter(Boolean)));
    const fechas = pls.map(p => p.cab.fecha).filter(Boolean).sort();
    const rango = fechas.length ? (fechas[0] === fechas[fechas.length - 1] ? fechas[0] : `${fechas[0]}_a_${fechas[fechas.length - 1]}`) : '';
    return `PackingList_${t.sufijo}_${nombreArchivoSeguro(clientes.length === 1 ? clientes[0] : 'Varios')}_${pls.length}_remitos_${rango}.${ext}`;
  }

  function descargarExcelMulti(pls) {
    if (pls.length === 1) return descargarExcel(pls[0]);
    const nombre = nombreArchivoMulti(pls, 'xlsx');
    descargarBytes(generarExcelMulti(pls), nombre, MIME_XLSX);
    return nombre;
  }

  // ---------------------------------------------------------------------------
  // 4. PACKING LIST EN PDF (hoja de impresión con el mismo diseño)
  // ---------------------------------------------------------------------------
  const escHtml = s => txt(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function filasPallet(pl, t, p, ip) {
    return p.filas.map((f, i) => `
      <tr class="${ip % 2 ? 'banda' : ''} ${i === 0 ? 'inicio' : ''}">
        ${i === 0 ? `<td rowspan="${p.filas.length}" class="b">${p.item}</td>
        <td rowspan="${p.filas.length}" class="b pallet">${escHtml(p.nro)}</td>` : ''}
        <td>${escHtml(f.producto)}</td>
        <td>${escHtml(f.variedad)}</td>
        <td>${String(f.peso).replace('.', ',')}</td>
        <td>${escHtml(f.color)}</td>
        ${i === 0 ? `<td rowspan="${p.filas.length}" class="termo">${escHtml(textoTermografo(p, t))}</td>` : ''}
        <td>${escHtml(f.categoria)}</td>
        <td>${escHtml(f.marca)}</td>
        ${pl.calibres.map(c => `<td class="num">${f.sizes[c] ? fmtMiles(f.sizes[c]) : ''}</td>`).join('')}
        <td class="num b">${fmtMiles(f.total)}</td>
      </tr>`).join('');
  }

  function cuerpoTabla(pl, t) {
    const n = pl.calibres.length;
    if (!pl.pallets.length) return `<tr><td colspan="${10 + n}">Sin pallets</td></tr>`;
    if (!pl.agrupado) return pl.pallets.map((p, ip) => filasPallet(pl, t, p, ip)).join('');
    let ip = 0;
    return pl.grupos.map(g => `
      <tr class="grupo"><td colspan="${10 + n}">REMITO ${escHtml(g.remito)}${g.fecha ? ' · ' + escHtml(fechaVista(g.fecha)) : ''}${g.vuelo ? ' · ' + escHtml(g.vuelo) : ''}${g.destino ? ' ' + escHtml(g.destino) : ''} · ${g.totales.pallets} pallet(s)</td></tr>
      ${g.pallets.map(p => filasPallet(pl, t, p, ip++)).join('')}
      <tr class="sub"><td colspan="9" class="rot">Subtotal remito ${escHtml(g.remito)}</td>${pl.calibres.map(c => `<td class="num">${fmtMiles(g.totales.porCal[c] || 0)}</td>`).join('')}<td class="num">${fmtMiles(g.totales.cajas)}</td></tr>`).join('');
  }

  function hojaHtml(pl) {
    const t = textos(pl);
    const n = pl.calibres.length;
    const fila = (r, v, cls = '') => `<tr><th>${r}</th><td class="${cls}">${escHtml(v) || '&nbsp;'}</td></tr>`;

    const resumen = pl.resumen.map(rr => `
      <tr><td class="b">${escHtml(rr.categoria)}</td><td>${String(rr.peso).replace('.', ',')} kg</td>
      ${pl.calibres.map(c => `<td class="num">${rr.sizes[c] ? fmtMiles(rr.sizes[c]) : ''}</td>`).join('')}
      <td class="num b">${fmtMiles(rr.total)}</td></tr>`).join('');

    const camion = t.expo ? '' : `
      <table class="caja"><tr><th colspan="2" class="banda-t">DATOS DEL CAMIÓN</th></tr>
        ${fila('CHOFER', pl.cab.chofer)}${fila('PATENTE', pl.cab.patente)}${fila('TELÉFONO', pl.cab.telefono)}${fila('TRANSPORTE', pl.cab.transporte)}
      </table>`;

    return `
<div class="hoja">
  <div class="enc">
    <img src="data:image/jpeg;base64,${LOGO_JPG_B64}" alt="CECO S.A.">
    <div><h1>Packing List</h1><div class="sub">${t.expo ? 'Fresh fruit from Patagonia · Argentina' : 'Fruta fresca de la Patagonia · Mercado interno'}</div></div>
    <div class="enc-ref">${pl.grupos.length === 1 ? 'REMITO<br><b>' + escHtml(pl.grupos[0].remito) + '</b>' : ''}</div>
  </div>

  <div class="datos">
    <table class="info">
      <tr><th class="banda-t" style="text-align:left !important;">PACKING LIST</th><td style="text-align:center;">DATE: ${escHtml(fechaVista(pl.cab.fecha))}</td></tr>
      ${fila('SHIPPER:', EMPRESA.nombre)}
      <tr><th>Address:</th><td class="n">${escHtml(EMPRESA.direccion1)}<br>${escHtml(EMPRESA.direccion2)}</td></tr>
      ${fila('CUIT:', EMPRESA.cuit, 'n')}
      ${fila(t.remito, pl.cab.remitos)}
      ${fila(t.consignee, pl.cab.cliente)}
      ${fila('AWB#', pl.cab.awb)}
      ${fila('Airline:', pl.cab.aerolinea)}
      ${fila('ETD:', pl.cab.etd, 'n')}
      ${fila('ETA:', pl.cab.eta, 'n')}
      ${fila('Destination:', pl.cab.destino)}
      ${fila('Flight N°:', pl.cab.vuelo)}
      ${fila('GGN:', pl.cab.ggn)}
    </table>
    <div class="lado">
      <table class="caja">
        <tr><th colspan="2" class="banda-t">${t.resumenTitulo}</th></tr>
        ${fila(t.resPallets, fmtMiles(pl.totales.pallets))}
        ${fila(t.resCajas, fmtMiles(pl.totales.cajas))}
        ${fila(t.resKilos, fmtKg(pl.totales.kilos))}
        ${fila(t.resRemito, pl.cab.remitos)}
      </table>
      ${camion}
      <table class="caja obs"><tr><th class="banda-t">${t.obsTitulo}</th></tr><tr><td>${escHtml(pl.cab.observaciones) || '&nbsp;'}</td></tr></table>
    </div>
  </div>

  <table class="tabla">
    <thead>
      <tr>
        <th rowspan="2" style="width:26px;">ITEM</th><th rowspan="2">PALLET N°</th><th rowspan="2">PRODUCT</th><th rowspan="2">VARIETY</th>
        <th rowspan="2">NET WEIGHT / BOX</th><th rowspan="2">Color</th><th rowspan="2">Temperature recorder</th><th rowspan="2">CATEGORY</th><th rowspan="2">BRAND</th>
        <th colspan="${n}">SIZES MM</th><th>TOTAL</th>
      </tr>
      <tr>${pl.calibres.map(c => `<th>${escHtml(c)}</th>`).join('')}<th>BOXES</th></tr>
    </thead>
    <tbody>${cuerpoTabla(pl, t)}</tbody>
    <tfoot><tr><td colspan="9" class="rot">${pl.agrupado ? `TOTAL GENERAL · ${pl.grupos.length} remitos` : 'TOTAL'}</td>${pl.calibres.map(c => `<td class="num">${fmtMiles(pl.totales.porCal[c] || 0)}</td>`).join('')}<td class="num" style="color:#A62A28;">${fmtMiles(pl.totales.cajas)}</td></tr></tfoot>
  </table>

  <table class="tabla resumen">
    <thead>
      <tr><th rowspan="2">CATEGORY</th><th rowspan="2">BOX</th><th colspan="${n}">SIZES SUMMARY</th><th>TOTAL</th></tr>
      <tr>${pl.calibres.map(c => `<th>${escHtml(c)}</th>`).join('')}<th>BOXES</th></tr>
    </thead>
    <tbody>${resumen}
      ${pl.resumen.length > 1 ? `<tr class="tot"><td colspan="2" class="rot">TOTAL</td>${pl.calibres.map(c => `<td class="num">${fmtMiles(pl.totales.porCal[c] || 0)}</td>`).join('')}<td class="num">${fmtMiles(pl.totales.cajas)}</td></tr>` : ''}
    </tbody>
  </table>

  ${t.expo ? '' : '<div class="firmas"><div>Despachó (firma y aclaración)</div><div>Recibí conforme · Chofer</div></div>'}

  <div class="pie">${EMPRESA.pie[0]}<br>${EMPRESA.pie[1]}<br><a>${EMPRESA.pie[2]}</a></div>
</div>`;
  }

  // Documento de impresión: una hoja (página) por cada packing recibido
  function documentoHtml(pls, { autoImprimir = true } = {}) {
    const horizontal = pls.some(pl => pl.calibres.length > 9);
    const titulo = nombreArchivoMulti(pls, 'pdf').replace(/\.pdf$/, '');
    return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>${escHtml(titulo)}</title>
<style>
  @page { size: A4 ${horizontal ? 'landscape' : 'portrait'}; margin: 10mm 9mm 12mm; }
  * { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { margin: 0; font-family: Arial, Helvetica, sans-serif; color: #1A1A1A; background: #E5E5E5; }
  .hoja { background: #fff; width: ${horizontal ? '297mm' : '210mm'}; min-height: ${horizontal ? '210mm' : '297mm'}; margin: 12px auto; padding: 10mm 9mm 12mm; box-shadow: 0 4px 18px rgba(0,0,0,.15); display: flex; flex-direction: column; }
  @media print { body { background: #fff; } .hoja { width: auto; min-height: 0; margin: 0; padding: 0; box-shadow: none; } .hoja + .hoja { break-before: page; page-break-before: always; } .no-print { display: none !important; } }
  .barra { position: sticky; top: 0; z-index: 5; background: #1A1A1A; color: #fff; padding: 8px 14px; display: flex; gap: 10px; align-items: center; justify-content: space-between; font-size: 12px; }
  .barra button { background: #A62A28; color: #fff; border: 0; border-radius: 6px; padding: 7px 14px; font-weight: 700; cursor: pointer; }
  .enc { display: grid; grid-template-columns: 120px 1fr 120px; align-items: center; border-bottom: 2.5px solid #A62A28; padding-bottom: 6px; }
  .enc img { width: 104px; }
  .enc h1 { margin: 0; text-align: center; font-family: Calibri, Carlito, Arial, sans-serif; font-size: 30px; letter-spacing: .3px; }
  .enc .sub { text-align: center; font-size: 10px; color: #595959; font-style: italic; margin-top: 2px; }
  .enc-ref { text-align: right; font-size: 8.5px; color: #595959; letter-spacing: .4px; }
  .enc-ref b { font-size: 12px; color: #A62A28; letter-spacing: 0; }
  .datos { display: grid; grid-template-columns: 1.15fr 1fr; gap: 14px; margin-top: 10px; align-items: start; }
  table { border-collapse: collapse; width: 100%; }
  .info th, .caja th { background: #F2F0EF; text-align: left; font-size: 8.5px; padding: 3px 6px; border: .6px solid #7F7F7F; width: 34%; white-space: nowrap; }
  .info td, .caja td { font-size: 9px; font-weight: 700; padding: 3px 6px; border: .6px solid #7F7F7F; }
  .info td.n { font-weight: 400; }
  .banda-t { background: #A62A28 !important; color: #fff; text-align: center !important; letter-spacing: .4px; }
  .lado { display: flex; flex-direction: column; gap: 8px; }
  .obs td { height: 34px; vertical-align: top; font-weight: 400 !important; white-space: pre-wrap; }
  .tabla { margin-top: 12px; font-family: "Bookman Old Style", "URW Bookman", Georgia, serif; }
  .tabla th { background: #A62A28; color: #fff; font-size: 7.6px; padding: 3px 3px; border: .6px solid #5A1716; text-align: center; line-height: 1.15; }
  .tabla td { font-size: 8px; padding: 2.6px 3px; border: .6px solid #7F7F7F; text-align: center; }
  .tabla tr.banda td { background: #FBF8F7; }
  .tabla tr.inicio td { border-top: 1px solid #333; }
  .tabla tr.grupo td { background: #F2F0EF; color: #A62A28; font-weight: 700; text-align: left; padding: 4px 6px; border-top: 1.4px solid #1A1A1A; font-size: 8.4px; }
  .tabla tr.sub td { background: #FCF4F3; font-weight: 700; }
  .tabla td.rot { text-align: right; padding-right: 8px; }
  .tabla .num { text-align: center; }
  .tabla .b { font-weight: 700; }
  .tabla .pallet { color: #A62A28; font-size: 8.6px; }
  .tabla .termo { font-size: 7.4px; }
  .tabla tfoot td, .tot td { background: #F7E8E7; font-weight: 700; border-top: 1.4px solid #1A1A1A; border-bottom: 1.4px solid #1A1A1A; }
  .tabla thead { display: table-header-group; }
  .tabla tr { page-break-inside: avoid; }
  .resumen { width: auto; margin: 14px 0 0 auto; }
  .firmas { display: flex; justify-content: space-around; margin-top: 46px; }
  .firmas div { width: 34%; border-top: 1px solid #1A1A1A; text-align: center; font-size: 8px; font-weight: 700; color: #595959; padding-top: 3px; }
  .pie { margin-top: auto; padding-top: 18px; font-size: 7.6px; color: #595959; line-height: 1.5; }
  .pie a { color: #0563C1; }
</style></head><body>
<div class="barra no-print"><span>Vista previa · ${escHtml(titulo)}.pdf${pls.length > 1 ? ` · ${pls.length} páginas (una por remito)` : ''}</span><button onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button></div>
${pls.map(hojaHtml).join('\n')}
${autoImprimir ? '<script>window.onload = function () { setTimeout(function () { window.print(); }, 350); };<\/script>' : ''}
</body></html>`;
  }

  const generarHtml = (pl, op) => documentoHtml([pl], op);
  const generarHtmlMulti = (pls, op) => documentoHtml(pls, op);

  function abrirPdfMulti(pls) {
    const win = window.open('', '_blank');
    if (!win) throw new Error('El navegador bloqueó la ventana de impresión. Permití ventanas emergentes para este sitio.');
    win.document.open();
    win.document.write(documentoHtml(pls));
    win.document.close();
    return nombreArchivoMulti(pls, 'pdf');
  }
  const abrirPdf = pl => abrirPdfMulti([pl]);

  const api = {
    construir, generarExcel, generarExcelMulti, descargarExcel, descargarExcelMulti,
    generarHtml, generarHtmlMulti, abrirPdf, abrirPdfMulti, nombreArchivo, nombreArchivoMulti,
    normalizarCalibre, pesoCaja, fechaVista, fechaIso, EMPRESA,
  };
  root.PackingList = api;
  if (typeof module === 'object' && module && module.exports && typeof window === 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
