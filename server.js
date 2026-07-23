'use strict';

const express        = require('express');
const multer         = require('multer');
const pdfParse       = require('pdf-parse/lib/pdf-parse.js');
const XLSX           = require('xlsx');
const PDFDoc         = require('pdfkit');
const path           = require('path');
const { MongoClient} = require('mongodb');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── MongoDB + cache en memoria ────────────────────────────────────────────────
let db;
let _data   = { proveedores: [], productos: [], lotes: [], ventas: [] };
let _config = { empresa: 'Base de Alimentos Navarro S.A.', password: 'lotemania', logo: null };

async function initDB() {
  const uri = process.env.MONGODB_URI;
  if (!uri) { console.warn('⚠ MONGODB_URI no definida — datos solo en memoria'); return; }
  const client = new MongoClient(uri, { tls: true, tlsAllowInvalidCertificates: false, serverApi: { version: '1' } });
  await client.connect();
  db = client.db('lotetrack');
  console.log('✓ MongoDB conectado');
  const datos  = await db.collection('datos').findOne({ _id: 'main' });
  const config = await db.collection('config').findOne({ _id: 'main' });
  if (datos)  { const { _id, ...r } = datos;  _data   = { proveedores:[], productos:[], lotes:[], ventas:[], ...r }; }
  if (config) { const { _id, ...r } = config; _config = { empresa:'Base de Alimentos Navarro S.A.', password:'lotemania', logo:null, ...r }; }
  if (_config.logo) _config.logo = Buffer.from(_config.logo, 'base64');
}

function leer() { return _data; }
function guardar(d) {
  _data = d;
  if (db) db.collection('datos').replaceOne({ _id:'main' }, { _id:'main', ...d }, { upsert:true }).catch(console.error);
}
function leerConfig() { return _config; }
function guardarConfig(c) {
  _config = c;
  const toSave = { ...c, logo: c.logo ? c.logo.toString('base64') : null };
  if (db) db.collection('config').replaceOne({ _id:'main' }, { _id:'main', ...toSave }, { upsert:true }).catch(console.error);
}

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ storage: multer.memoryStorage() });

// Sirve el logo desde la config en memoria / MongoDB
app.get('/logo-empresa.png', (req, res) => {
  const logo = leerConfig().logo;
  if (logo) res.type('png').send(logo);
  else res.status(404).end();
});

// ── Persistencia helpers ──────────────────────────────────────────────────────
function nuevoId(p, lista) { return `${p}-${String(lista.length + 1).padStart(3, '0')}`; }

// ── Conversiones ──────────────────────────────────────────────────────────────
const LITROS_PROD = { '900ml': 0.9, '5L': 5, '10L': 10 };

// ── Parser de facturas PDF ────────────────────────────────────────────────────
function parsearFactura(texto) {
  const t = texto.replace(/\r/g, '').replace(/\t/g, ' ');

  const esVenta  = /CLIENTE\s*:/i.test(t);
  const esCompra = /SE[ÑN]OR(?:ES|\/ES)?\s*:/i.test(t);

  // ── Fecha: "FECHA: DD/MM/YYYY" evitando "Fecha de Vencimiento"
  let fecha = null;
  // Busca todas las fechas en el texto y toma la que está después de "FECHA:"
  const todasFechas = [...t.matchAll(/(\d{2}\/\d{2}\/\d{4})/g)];
  const idxFecha = t.search(/\bFECHA\s*[:\-]/i);
  if (idxFecha >= 0) {
    const despues = todasFechas.find(m => m.index > idxFecha && m.index < idxFecha + 30);
    if (despues) fecha = despues[1].split('/').reverse().join('-');
  }
  // Fallback: primera fecha que NO sea la de vencimiento
  if (!fecha && todasFechas.length) {
    const idxVto = t.search(/vencimiento/i);
    const candidata = todasFechas.find(m => idxVto < 0 || Math.abs(m.index - idxVto) > 40);
    if (candidata) fecha = candidata[1].split('/').reverse().join('-');
  }

  // ── Factura: "Nº 0008-00001215" o "0008 - 00001215"
  const mFac = t.match(/N[º°]\s*(\d{4}[-\s]\d{8})/i)
            || t.match(/(\d{4})\s*-\s*(\d{8})/);
  const factura = mFac
    ? (mFac[1]?.match(/\d{4}-\d{8}/) ? mFac[1] : mFac[1] && mFac[2] ? `${mFac[1]}-${mFac[2]}` : null)
    : null;

  // ════════════════════════════════ VENTA ════════════════════════════════════
  if (esVenta) {
    const mCliente = t.match(/CLIENTE\s*:\s*([^\n]+)/i);
    let cliente = mCliente ? mCliente[1].trim() : '';
    // Limpiar texto extra tipo "DOMICILIO: ..."
    cliente = cliente.replace(/\s*DOMICILIO\s*:.*/i, '').trim();

    const mRemito = t.match(/REMITO\s*:\s*([\d\s\-]+)/i);
    const remito  = mRemito ? mRemito[1].trim().replace(/\s/g, '') : null;

    // Items: "26520.000 u ... ACEITE GIRASOL 900 ml"
    const items = [];
    const reItem = /(\d[\d.]*)\s+u[^\n]*?ACEITE[^\n]*(900\s*m[lL]|5\s*[Ll]|10\s*[Ll])/gi;
    let m;
    while ((m = reItem.exec(t)) !== null) {
      const cant = Math.round(parseFloat(m[1]));
      const raw  = m[2].trim().replace(/\s/g, '').toLowerCase();
      const prod = raw.includes('900') ? '900ml' : raw.includes('10') ? '10L' : '5L';
      if (cant > 0) items.push({ producto: prod, cantidad: cant });
    }

    return { tipo: 'venta', fecha, factura, cliente, remito, items };
  }

  // ════════════════════════════════ COMPRA ══════════════════════════════════
  if (esCompra) {
    // ── Proveedor: en facturas tipo JM Falavigna, "Razón social:" precede a "Inicio de Actividades:"
    let proveedor = '';

    // 1. Línea antes de "Inicio de Actividades" (la más fiable en facturas AFIP)
    const mAntes = t.match(/([^\n\r]+)\r?\n[^\n\r]*[Ii]nicio\s+de\s+[Aa]ctividades/);
    if (mAntes) {
      proveedor = mAntes[1].trim()
        .replace(/[Rr]az[oó0]n\s*[Ss]ocial\s*:\s*/i, '')
        .replace(/^.*:\s*/, '') // eliminar cualquier "Label:" al inicio
        .trim();
    }

    // 2. "Razón social: XXX" o "Razon social: XXX" (cualquier encoding de ó)
    if (!proveedor || proveedor.length < 3) {
      const mRS = t.match(/[Rr][Aa][Zz].{0,3}[Nn]\s+[Ss][Oo][Cc][Ii][Aa][Ll]\s*:\s*([^\n\r]+)/);
      if (mRS) proveedor = mRS[1].trim();
    }

    // 3. Primera línea que contenga "S.A." o "S.R.L." y no sea BAN
    if (!proveedor || proveedor.length < 3) {
      const lineas = t.split('\n').map(l => l.trim());
      proveedor = lineas.find(l =>
        /\bS\.?\s?A\.?\b|\bS\.?\s?R\.?\s?L\.?\b/i.test(l) &&
        !/BASE DE ALIMENTOS|NAVARRO/i.test(l) &&
        l.length > 4 && l.length < 60
      ) || '';
    }

    // ── Remito
    const mRem = t.match(/S\/REMITO\s+([\d\-]+)/i)
              || t.match(/REMITO\s*:\s*([\d\s\-]+)/i);
    const remito = mRem ? mRem[1].trim().replace(/\s/g, '') : null;

    // ── Toneladas y Producto
    let producto = '', toneladas = 0;

    // Buscar la línea del detalle: "ACEITE ... S/REMITO XXXX-XXXXXXXX QTY PRECIO TOTAL"
    const mLinea = t.match(/ACEITE[^\n]+/i);
    if (mLinea) {
      const linea = mLinea[0];

      // Nombre: todo antes de S/REMITO o del primer número grande
      const parteNombre = linea.split(/S\/REMITO/i)[0];
      producto = parteNombre
        .replace(/\s{2,}/g, ' ')
        .trim()
        .replace(/\s+\d.*$/, '') // cortar si queda algún número al final
        .trim();
      if (!producto) producto = 'Aceite';

      // Toneladas: quitar el número de remito y buscar el primer decimal pequeño
      const lineaSinRemito = linea.replace(/S\/REMITO\s+[\d\-]+/gi, '');

      // Patrón: número de 1-3 dígitos con exactamente 2 decimales, seguido de un número más grande (precio)
      const mTon = lineaSinRemito.match(/\b(\d{1,3}[.,]\d{2})\s+[\d.,]{5,}/);
      if (mTon) {
        toneladas = parseFloat(mTon[1].replace(',', '.'));
      } else {
        // Fallback: todos los números < 500 con decimales en la línea
        const nums = [...lineaSinRemito.matchAll(/\b(\d{1,3}[.,]\d{1,2})\b/g)]
          .map(m => parseFloat(m[1].replace(',', '.')))
          .filter(n => n > 0.5 && n < 500);
        if (nums.length) toneladas = nums[0];
      }
    }

    return { tipo: 'compra', fecha, factura, proveedor, remito, producto, toneladas };
  }

  return { tipo: 'desconocido', fecha, factura };
}

// ══════════════════════════════════════════════════════════════════════════════
// RUTAS API
// ══════════════════════════════════════════════════════════════════════════════

// ── CONFIG / LOGIN / LOGO ─────────────────────────────────────────────────────
app.get('/api/config', (req, res) => {
  const c = leerConfig();
  res.json({ empresa: c.empresa, tienePassword: !!c.password, tienelogo: !!c.logo });
});

app.post('/api/config', (req, res) => {
  try {
    const c = leerConfig();
    if (req.body.empresa)          c.empresa  = req.body.empresa;
    if (req.body.passwordNueva)    c.password = req.body.passwordNueva;
    guardarConfig(c);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/login', (req, res) => {
  const c = leerConfig();
  if (req.body.password === c.password) res.json({ ok: true, empresa: c.empresa });
  else res.status(401).json({ ok: false, error: 'Contraseña incorrecta' });
});

app.post('/api/config/logo', upload.single('logo'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Sin archivo' });
    const c = leerConfig();
    c.logo = req.file.buffer;
    guardarConfig(c);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/config/logo', (req, res) => {
  const c = leerConfig();
  c.logo = null;
  guardarConfig(c);
  res.json({ ok: true });
});

// ── POST /api/parse-pdf ───────────────────────────────────────────────────────
app.post('/api/parse-pdf', upload.single('factura'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No se recibió archivo' });
    const data   = await pdfParse(req.file.buffer);
    const parsed = parsearFactura(data.text);
    res.json({ ok: true, parsed });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── PROVEEDORES CRUD ──────────────────────────────────────────────────────────
app.get('/api/proveedores', (req, res) => res.json(leer().proveedores));

app.post('/api/proveedores', (req, res) => {
  try {
    const { nombre, cuit, contacto } = req.body;
    if (!nombre) return res.status(400).json({ error: 'El nombre es requerido' });
    const d = leer();
    const p = { id: nuevoId('PRV', d.proveedores), nombre, cuit: cuit || '', contacto: contacto || '' };
    d.proveedores.push(p);
    guardar(d);
    res.json({ ok: true, proveedor: p });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/proveedores/:id', (req, res) => {
  try {
    const d = leer();
    const p = d.proveedores.find(x => x.id === req.params.id);
    if (!p) return res.status(404).json({ error: 'No encontrado' });
    Object.assign(p, { nombre: req.body.nombre || p.nombre, cuit: req.body.cuit ?? p.cuit, contacto: req.body.contacto ?? p.contacto });
    guardar(d);
    res.json({ ok: true, proveedor: p });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/proveedores/:id', (req, res) => {
  try {
    const d = leer();
    const idx = d.proveedores.findIndex(x => x.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'No encontrado' });
    d.proveedores.splice(idx, 1);
    guardar(d);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── PRODUCTOS CRUD ───────────────────────────────────────────────────────────
app.get('/api/productos', (req, res) => res.json(leer().productos));

app.post('/api/productos', (req, res) => {
  try {
    const { nombre } = req.body;
    if (!nombre) return res.status(400).json({ error: 'El nombre es requerido' });
    const d = leer();
    if (d.productos.some(p => p.nombre.toLowerCase() === nombre.toLowerCase()))
      return res.status(400).json({ error: 'Ya existe un producto con ese nombre' });
    const p = { id: nuevoId('PRD', d.productos), nombre: nombre.trim() };
    d.productos.push(p);
    guardar(d);
    res.json({ ok: true, producto: p });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/productos/:id', (req, res) => {
  try {
    const d = leer();
    const p = d.productos.find(x => x.id === req.params.id);
    if (!p) return res.status(404).json({ error: 'No encontrado' });
    p.nombre = req.body.nombre || p.nombre;
    guardar(d);
    res.json({ ok: true, producto: p });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/productos/:id', (req, res) => {
  try {
    const d = leer();
    const idx = d.productos.findIndex(x => x.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'No encontrado' });
    d.productos.splice(idx, 1);
    guardar(d);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── LOTES CRUD ────────────────────────────────────────────────────────────────
app.get('/api/lotes', (req, res) => res.json(leer().lotes));

app.get('/api/stock', (req, res) => {
  const d = leer();
  const total = d.lotes.reduce((s, l) => s + l.litros_disponibles, 0);
  // Agrupar por tipo de aceite
  const grupos = {};
  d.lotes.forEach(l => {
    if (!grupos[l.producto]) grupos[l.producto] = 0;
    grupos[l.producto] += l.litros_disponibles;
  });
  res.json({ lotes: d.lotes, total_disponible: parseFloat(total.toFixed(2)), grupos });
});

app.post('/api/lotes', (req, res) => {
  try {
    const { fecha, proveedor, factura, remito, producto, toneladas } = req.body;
    if (!fecha || !proveedor || !factura || !producto || !toneladas)
      return res.status(400).json({ error: 'Faltan campos obligatorios' });
    const d      = leer();
    const litros = parseFloat((parseFloat(toneladas) * 1080).toFixed(2));
    const lote   = {
      id: nuevoId('LOT', d.lotes), fecha, proveedor, factura,
      remito: remito || null, producto, toneladas: parseFloat(toneladas),
      litros_total: litros, litros_disponibles: litros
    };
    d.lotes.push(lote);
    guardar(d);
    res.json({ ok: true, lote });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/lotes/:id', (req, res) => {
  try {
    const d    = leer();
    const lote = d.lotes.find(l => l.id === req.params.id);
    if (!lote) return res.status(404).json({ error: 'Lote no encontrado' });

    // Si cambian toneladas, solo permitir si el lote no tiene ventas
    if (req.body.toneladas && parseFloat(req.body.toneladas) !== lote.toneladas) {
      const usado = d.ventas.some(v => v.consumo_lotes.some(c => c.lote_id === lote.id));
      if (usado) return res.status(400).json({ error: 'No se pueden cambiar las toneladas: el lote tiene ventas asociadas.' });
      const nuevasLitros = parseFloat((parseFloat(req.body.toneladas) * 1080).toFixed(2));
      lote.toneladas         = parseFloat(req.body.toneladas);
      lote.litros_total      = nuevasLitros;
      lote.litros_disponibles = nuevasLitros;
    }

    if (req.body.fecha)      lote.fecha      = req.body.fecha;
    if (req.body.proveedor)  lote.proveedor  = req.body.proveedor;
    if (req.body.factura)    lote.factura    = req.body.factura;
    if ('remito' in req.body) lote.remito    = req.body.remito || null;
    if (req.body.producto)   lote.producto   = req.body.producto;

    guardar(d);
    res.json({ ok: true, lote });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/lotes/:id', (req, res) => {
  try {
    const d   = leer();
    const idx = d.lotes.findIndex(l => l.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Lote no encontrado' });
    const usado = d.ventas.some(v => v.consumo_lotes.some(c => c.lote_id === req.params.id));
    if (usado) return res.status(400).json({ error: 'No se puede eliminar: el lote tiene ventas asociadas.' });
    d.lotes.splice(idx, 1);
    guardar(d);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── VENTAS CRUD ───────────────────────────────────────────────────────────────
app.get('/api/ventas', (req, res) => {
  const { proveedor } = req.query;
  const d = leer();
  let ventas = d.ventas;
  if (proveedor) ventas = ventas.filter(v => v.consumo_lotes.some(c => c.proveedor === proveedor));
  res.json(ventas);
});

function fifoDescontar(d, litrosPedidos, tipoProd) {
  const consumo_lotes = [];
  let pend = litrosPedidos;
  const candidatos = d.lotes
    .filter(l => l.litros_disponibles > 0 && (!tipoProd || l.producto === tipoProd))
    .sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
  for (const ref of candidatos) {
    if (pend <= 0) break;
    const usar = parseFloat(Math.min(ref.litros_disponibles, pend).toFixed(2));
    const lote = d.lotes.find(l => l.id === ref.id);
    lote.litros_disponibles = parseFloat((lote.litros_disponibles - usar).toFixed(2));
    consumo_lotes.push({ lote_id: lote.id, proveedor: lote.proveedor, factura_compra: lote.factura, litros_usados: usar });
    pend = parseFloat((pend - usar).toFixed(2));
  }
  return consumo_lotes;
}

function calcularItems(items) {
  let total = 0;
  const its = items.map(i => {
    const f = LITROS_PROD[i.producto];
    if (!f) throw new Error(`Producto inválido: ${i.producto}`);
    const l = parseFloat((i.cantidad * f).toFixed(2));
    total += l;
    return { producto: i.producto, cantidad: i.cantidad, litros_unitarios: f, litros_total: l };
  });
  return { its, total: parseFloat(total.toFixed(2)) };
}

app.post('/api/ventas', (req, res) => {
  try {
    const { fecha, cliente, factura, remito, items, tipo_aceite } = req.body;
    if (!fecha || !cliente || !factura || !items?.length)
      return res.status(400).json({ error: 'Faltan campos obligatorios' });
    if (!tipo_aceite)
      return res.status(400).json({ error: 'Falta especificar el tipo de aceite' });

    const d = leer();
    const { its, total } = calcularItems(items);
    const stock = d.lotes.filter(l => l.producto === tipo_aceite).reduce((s, l) => s + l.litros_disponibles, 0);
    if (stock < total) return res.status(400).json({ error: `Stock insuficiente de "${tipo_aceite}". Disponible: ${stock.toFixed(2)} L` });

    const consumo_lotes = fifoDescontar(d, total, tipo_aceite);
    const v = { id: nuevoId('VTA', d.ventas), fecha, cliente, factura, remito: remito || null, tipo_aceite, items: its, litros_total: total, consumo_lotes };
    d.ventas.push(v);
    guardar(d);
    res.json({ ok: true, venta: v });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/ventas/:id', (req, res) => {
  try {
    const d   = leer();
    const idx = d.ventas.findIndex(v => v.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Venta no encontrada' });

    const venta = d.ventas[idx];

    // Restaurar litros al stock
    venta.consumo_lotes.forEach(c => {
      const lote = d.lotes.find(l => l.id === c.lote_id);
      if (lote) lote.litros_disponibles = parseFloat((lote.litros_disponibles + c.litros_usados).toFixed(2));
    });

    // Recalcular con nuevos items
    const tipo_aceite = req.body.tipo_aceite || venta.tipo_aceite;
    const items  = req.body.items || venta.items.map(i => ({ producto: i.producto, cantidad: i.cantidad }));
    const { its, total } = calcularItems(items);
    const stock = d.lotes.filter(l => !tipo_aceite || l.producto === tipo_aceite).reduce((s, l) => s + l.litros_disponibles, 0);
    if (stock < total) {
      // Revertir restauración
      venta.consumo_lotes.forEach(c => {
        const lote = d.lotes.find(l => l.id === c.lote_id);
        if (lote) lote.litros_disponibles = parseFloat((lote.litros_disponibles - c.litros_usados).toFixed(2));
      });
      return res.status(400).json({ error: `Stock insuficiente. Disponible: ${stock.toFixed(2)} L` });
    }

    const consumo_lotes = fifoDescontar(d, total, tipo_aceite);

    // Actualizar campos
    venta.fecha         = req.body.fecha       || venta.fecha;
    venta.cliente       = req.body.cliente     || venta.cliente;
    venta.factura       = req.body.factura     || venta.factura;
    venta.remito        = 'remito' in req.body ? req.body.remito : venta.remito;
    venta.tipo_aceite   = tipo_aceite;
    venta.items         = its;
    venta.litros_total  = total;
    venta.consumo_lotes = consumo_lotes;

    guardar(d);
    res.json({ ok: true, venta });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/ventas/:id', (req, res) => {
  try {
    const d   = leer();
    const idx = d.ventas.findIndex(v => v.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Venta no encontrada' });
    d.ventas[idx].consumo_lotes.forEach(c => {
      const lote = d.lotes.find(l => l.id === c.lote_id);
      if (lote) lote.litros_disponibles = parseFloat((lote.litros_disponibles + c.litros_usados).toFixed(2));
    });
    d.ventas.splice(idx, 1);
    guardar(d);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── TRAZABILIDAD ──────────────────────────────────────────────────────────────
app.get('/api/trazabilidad/lote/:id', (req, res) => {
  const { lotes, ventas } = leer();
  const lote = lotes.find(l => l.id === req.params.id);
  if (!lote) return res.status(404).json({ error: 'Lote no encontrado' });
  res.json({ lote, ventas: ventas.filter(v => v.consumo_lotes.some(c => c.lote_id === lote.id)) });
});

app.get('/api/trazabilidad/venta/:id', (req, res) => {
  const { ventas } = leer();
  const venta = ventas.find(v => v.id === req.params.id);
  if (!venta) return res.status(404).json({ error: 'Venta no encontrada' });
  res.json(venta);
});

// ── EXPORTAR TRAZABILIDAD INDIVIDUAL ─────────────────────────────────────────
app.get('/api/export/excel/lote/:id', (req, res) => {
  const { lotes, ventas } = leer();
  const lote = lotes.find(l => l.id === req.params.id);
  if (!lote) return res.status(404).json({ error: 'No encontrado' });
  const ventasLote = ventas.filter(v => v.consumo_lotes.some(c => c.lote_id === lote.id));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{
    ID: lote.id, Fecha: lote.fecha, Proveedor: lote.proveedor, Factura: lote.factura,
    Remito: lote.remito || '', Producto: lote.producto, Toneladas: lote.toneladas,
    'Litros Total': lote.litros_total, 'Litros Disponibles': lote.litros_disponibles,
    'Litros Usados': parseFloat((lote.litros_total - lote.litros_disponibles).toFixed(2))
  }]), 'Lote');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
    ventasLote.map(v => {
      const c = v.consumo_lotes.find(x => x.lote_id === lote.id);
      return { 'ID Venta': v.id, 'Fecha': v.fecha, 'Cliente': v.cliente,
        'Factura Venta': v.factura, 'Remito': v.remito || '',
        'Litros de este lote': c?.litros_usados || 0, 'Total litros venta': v.litros_total };
    })
  ), 'Destino (Ventas)');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename="traz-${lote.id}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.get('/api/export/excel/venta/:id', (req, res) => {
  const { lotes, ventas } = leer();
  const venta = ventas.find(v => v.id === req.params.id);
  if (!venta) return res.status(404).json({ error: 'No encontrado' });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{
    ID: venta.id, Fecha: venta.fecha, Cliente: venta.cliente,
    Factura: venta.factura, Remito: venta.remito || '', 'Litros Total': venta.litros_total
  }]), 'Venta');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
    venta.items.map(i => ({ Producto: i.producto, Cantidad: i.cantidad,
      'L/unidad': i.litros_unitarios, 'Litros Total': i.litros_total }))
  ), 'Items');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
    venta.consumo_lotes.map(c => {
      const lote = lotes.find(l => l.id === c.lote_id);
      return { Lote: c.lote_id, Proveedor: c.proveedor, 'Factura Compra': c.factura_compra,
        'Fecha Compra': lote?.fecha || '', 'Litros Usados': c.litros_usados };
    })
  ), 'Origen (Lotes)');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename="traz-${venta.id}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.get('/api/export/pdf/lote/:id', (req, res) => {
  const d   = leer();
  const cfg = leerConfig();
  const lote = d.lotes.find(l => l.id === req.params.id);
  if (!lote) return res.status(404).json({ error: 'No encontrado' });
  const ventasLote = d.ventas.filter(v => v.consumo_lotes.some(c => c.lote_id === lote.id));
  const doc = new PDFDoc({ margin: 40, size: 'A4' });
  res.setHeader('Content-Disposition', `attachment; filename="traz-${lote.id}.pdf"`);
  res.setHeader('Content-Type', 'application/pdf');
  doc.pipe(res);

  pdfHeader(doc, `Trazabilidad de Lote — ${lote.id}`, cfg.empresa, null);

  // Ficha del lote
  pdfTabla(doc, [
    { x: 40, w: 100, label: 'Campo' },
    { x: 145, w: 410, label: 'Valor' },
  ], [
    ['ID', lote.id], ['Fecha', lote.fecha], ['Proveedor', lote.proveedor],
    ['Factura', lote.factura], ['Remito', lote.remito || '-'],
    ['Producto', lote.producto], ['Toneladas', lote.toneladas],
    ['Litros Total', lote.litros_total.toLocaleString('es-AR') + ' L'],
    ['Litros Disponibles', lote.litros_disponibles.toLocaleString('es-AR') + ' L'],
    ['Litros Usados', parseFloat((lote.litros_total - lote.litros_disponibles).toFixed(2)).toLocaleString('es-AR') + ' L'],
    ['% Consumido', ((lote.litros_total - lote.litros_disponibles) / lote.litros_total * 100).toFixed(1) + '%'],
  ], 'DATOS DEL LOTE');

  doc.moveDown(1);

  // Ventas destino
  const colsDest = [
    { x: 40,  w: 58,  label: 'ID Venta' },
    { x: 101, w: 60,  label: 'Fecha' },
    { x: 164, w: 140, label: 'Cliente' },
    { x: 307, w: 110, label: 'Factura' },
    { x: 420, w: 70,  label: 'Litros usados', align: 'right' },
    { x: 493, w: 62,  label: 'Total venta',   align: 'right' },
  ];
  pdfTabla(doc, colsDest,
    ventasLote.length ? ventasLote.map(v => {
      const c = v.consumo_lotes.find(x => x.lote_id === lote.id);
      return [v.id, v.fecha, v.cliente, v.factura,
        (c?.litros_usados || 0).toLocaleString('es-AR') + ' L',
        v.litros_total.toLocaleString('es-AR') + ' L'];
    }) : [['—', '—', 'Sin ventas asociadas', '', '', '']],
    'DESTINO — VENTAS QUE CONSUMIERON ESTE LOTE'
  );
  doc.end();
});

app.get('/api/export/pdf/venta/:id', (req, res) => {
  const d   = leer();
  const cfg = leerConfig();
  const venta = d.ventas.find(v => v.id === req.params.id);
  if (!venta) return res.status(404).json({ error: 'No encontrado' });
  const doc = new PDFDoc({ margin: 40, size: 'A4' });
  res.setHeader('Content-Disposition', `attachment; filename="traz-${venta.id}.pdf"`);
  res.setHeader('Content-Type', 'application/pdf');
  doc.pipe(res);

  pdfHeader(doc, `Trazabilidad de Venta — ${venta.id}`, cfg.empresa, null);

  // Ficha de venta
  pdfTabla(doc, [
    { x: 40, w: 100, label: 'Campo' },
    { x: 145, w: 410, label: 'Valor' },
  ], [
    ['ID', venta.id], ['Fecha', venta.fecha], ['Cliente', venta.cliente],
    ['Factura', venta.factura], ['Remito', venta.remito || '-'],
    ['Tipo aceite', venta.tipo_aceite || '-'],
    ['Total Litros', venta.litros_total.toLocaleString('es-AR') + ' L'],
  ], 'DATOS DE LA VENTA');

  doc.moveDown(1);

  // Items
  pdfTabla(doc, [
    { x: 40,  w: 120, label: 'Producto' },
    { x: 163, w: 90,  label: 'Cantidad (u)', align: 'right' },
    { x: 256, w: 80,  label: 'L/unidad',     align: 'right' },
    { x: 339, w: 90,  label: 'Litros Total',  align: 'right' },
  ], venta.items.map(i => [
    i.producto,
    i.cantidad.toLocaleString('es-AR'),
    i.litros_unitarios + ' L',
    i.litros_total.toLocaleString('es-AR') + ' L',
  ]), 'PRODUCTOS VENDIDOS');

  doc.moveDown(1);

  // Origen FIFO
  const colsOrig = [
    { x: 40,  w: 58,  label: 'Lote' },
    { x: 101, w: 140, label: 'Proveedor' },
    { x: 244, w: 110, label: 'Factura compra' },
    { x: 357, w: 60,  label: 'Fecha',        align: 'center' },
    { x: 420, w: 135, label: 'Litros usados', align: 'right' },
  ];
  pdfTabla(doc, colsOrig, venta.consumo_lotes.map(c => {
    const lote = d.lotes.find(l => l.id === c.lote_id);
    return [c.lote_id, c.proveedor, c.factura_compra,
      lote?.fecha || '-', c.litros_usados.toLocaleString('es-AR') + ' L'];
  }), 'ORIGEN (FIFO) — LOTES CONSUMIDOS');

  doc.end();
});

// ── HELPERS EXPORTAR ─────────────────────────────────────────────────────────
function aplicarFiltros(d, q) {
  let ls = [...d.lotes], vs = [...d.ventas];
  if (q.desde)     { ls = ls.filter(l => l.fecha >= q.desde); vs = vs.filter(v => v.fecha >= q.desde); }
  if (q.hasta)     { ls = ls.filter(l => l.fecha <= q.hasta); vs = vs.filter(v => v.fecha <= q.hasta); }
  if (q.proveedor) { ls = ls.filter(l => l.proveedor === q.proveedor); vs = vs.filter(v => v.consumo_lotes.some(c => c.proveedor === q.proveedor)); }
  return { ls, vs };
}

// Dibuja cabecera de tabla PDF con fondo azul y filas con alineación exacta
function pdfTabla(doc, cols, filas, titulo) {
  if (doc.y > 680) doc.addPage();
  if (titulo) {
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#0d1b4b').text(titulo, 40);
    doc.fillColor('black').moveDown(0.4);
  }
  const yH = doc.y;
  const rowH = 14;
  // Cabecera
  doc.rect(40, yH, 515, rowH).fill('#0d1b4b');
  cols.forEach(c => doc.fontSize(7.5).font('Helvetica-Bold').fillColor('white')
    .text(c.label, c.x, yH + 3, { width: c.w, align: c.align || 'left', lineBreak: false }));
  doc.fillColor('black');
  let y = yH + rowH + 1;
  filas.forEach((fila, ri) => {
    if (y > 750) { doc.addPage(); y = 40; }
    if (ri % 2 === 1) doc.rect(40, y - 1, 515, rowH).fill('#f4f6f8').fillColor('black');
    cols.forEach((c, ci) => {
      doc.fontSize(7.5).font('Helvetica').fillColor('#222')
         .text(String(fila[ci] ?? ''), c.x, y + 2, { width: c.w, align: c.align || 'left', lineBreak: false });
    });
    y += rowH;
  });
  doc.y = y + 4;
}

function pdfHeader(doc, titulo, empresa, filtros) {
  const logo = leerConfig().logo;
  let startX = 40;
  if (logo) {
    try { doc.image(logo, 40, 35, { height: 40 }); startX = 120; } catch(e) {}
  }
  doc.fontSize(17).font('Helvetica-Bold').fillColor('#0d1b4b')
     .text('Lote', startX, 38, { continued: true, align: 'left' });
  doc.fillColor('#2563eb').text(titulo.replace('Reporte de Trazabilidad','Track — Reporte').replace('Trazabilidad —','Track —'), { align: 'left' });
  doc.fontSize(9).font('Helvetica').fillColor('#555')
     .text(empresa + '  —  ' + new Date().toLocaleDateString('es-AR'), startX, 58, { align: 'center' });
  if (filtros) doc.fontSize(8).fillColor('#888').text(filtros, startX, 72, { align: 'center' });
  doc.fillColor('black').moveDown(filtros ? 2.5 : 2);
}

// ── EXPORTAR (con filtros) ────────────────────────────────────────────────────
app.get('/api/export/excel', (req, res) => {
  const d = leer();
  const { ls, vs } = aplicarFiltros(d, req.query);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ls.map(l => ({
    ID: l.id, Fecha: l.fecha, Proveedor: l.proveedor, Factura: l.factura,
    Remito: l.remito || '', Producto: l.producto, 'Tipo Aceite': l.producto,
    Toneladas: l.toneladas, 'Litros Total': l.litros_total,
    'Litros Disponibles': l.litros_disponibles,
    'Litros Usados': parseFloat((l.litros_total - l.litros_disponibles).toFixed(2)),
    '% Disponible': parseFloat(((l.litros_disponibles / l.litros_total) * 100).toFixed(1))
  }))), 'Lotes');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(vs.flatMap(v =>
    v.consumo_lotes.map(c => ({
      'ID Venta': v.id, 'Fecha Venta': v.fecha, Cliente: v.cliente,
      'Tipo Aceite': v.tipo_aceite || '', 'Factura Venta': v.factura,
      'Litros Venta': v.litros_total, 'Lote Origen': c.lote_id,
      'Proveedor Origen': c.proveedor, 'Factura Compra': c.factura_compra,
      'Litros del Lote': c.litros_usados
    }))
  )), 'Trazabilidad');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="lotemania-reporte.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.get('/api/export/pdf', (req, res) => {
  const d  = leer();
  const cfg = leerConfig();
  const { ls, vs } = aplicarFiltros(d, req.query);
  const { desde, hasta, proveedor } = req.query;
  const doc = new PDFDoc({ margin: 40, size: 'A4' });
  res.setHeader('Content-Disposition', 'attachment; filename="lotemania-reporte.pdf"');
  res.setHeader('Content-Type', 'application/pdf');
  doc.pipe(res);

  const filtroLabel = [
    desde ? `Desde: ${desde}` : '', hasta ? `Hasta: ${hasta}` : '',
    proveedor ? `Proveedor: ${proveedor}` : ''
  ].filter(Boolean).join('   |   ') || null;

  pdfHeader(doc, 'Reporte de Trazabilidad', cfg.empresa, filtroLabel);

  // Tabla de lotes
  const colsLote = [
    { x: 40,  w: 52,  label: 'ID' },
    { x: 95,  w: 60,  label: 'Fecha' },
    { x: 158, w: 115, label: 'Proveedor' },
    { x: 276, w: 105, label: 'Producto' },
    { x: 384, w: 57,  label: 'L Total',   align: 'right' },
    { x: 444, w: 57,  label: 'L Dispon.', align: 'right' },
    { x: 504, w: 51,  label: '%',         align: 'right' },
  ];
  pdfTabla(doc, colsLote, ls.map(l => {
    const pct = l.litros_total > 0 ? ((l.litros_disponibles / l.litros_total) * 100).toFixed(1) + '%' : '0%';
    return [l.id, l.fecha, l.proveedor, l.producto,
      l.litros_total.toLocaleString('es-AR'), l.litros_disponibles.toLocaleString('es-AR'), pct];
  }), 'STOCK POR LOTE');

  const totalDisp = ls.reduce((s, l) => s + l.litros_disponibles, 0);
  doc.fontSize(8).font('Helvetica-Bold').fillColor('#0d1b4b')
     .text(`TOTAL DISPONIBLE: ${totalDisp.toLocaleString('es-AR')} L`, 40);
  doc.fillColor('black').moveDown(1.5);

  // Tabla de ventas
  const colsVta = [
    { x: 40,  w: 52,  label: 'ID Venta' },
    { x: 95,  w: 60,  label: 'Fecha' },
    { x: 158, w: 115, label: 'Cliente' },
    { x: 276, w: 105, label: 'Tipo Aceite' },
    { x: 384, w: 57,  label: 'Litros',  align: 'right' },
    { x: 444, w: 111, label: 'Lotes origen' },
  ];
  pdfTabla(doc, colsVta, vs.map(v => [
    v.id, v.fecha, v.cliente, v.tipo_aceite || '',
    v.litros_total.toLocaleString('es-AR'),
    v.consumo_lotes.map(c => `${c.lote_id}(${c.litros_usados}L)`).join(', ')
  ]), 'TRAZABILIDAD DE VENTAS');

  doc.end();
});

// ── Inicio ────────────────────────────────────────────────────────────────────
initDB()
  .then(() => app.listen(PORT, () => console.log(`\n✓ LoteTrack en http://localhost:${PORT}\n`)))
  .catch(err => {
    console.error('Error conectando MongoDB:', err.message);
    app.listen(PORT, () => console.log(`\n✓ LoteTrack en http://localhost:${PORT} (sin DB)\n`));
  });

