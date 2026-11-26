// Cocina — web app personal sobre la hoja "Cocina · Inventario v1".
const SHEET_ID = '1BWGuyolvqc8fqS3zIcanCg3bPuqqCidPa-pWIpDQTWM';
const TZ = 'Asia/Jerusalem';
const PERE = ['Perecedero', 'Sobra'];

// ===== API para el front-end de GitHub Pages (reemplaza al doGet anterior) =====
const APP_URL = 'https://trypartylol-max.github.io/cocina/';
const SB_URL = 'https://llpyxktcfubhmthqujpa.supabase.co';
const SB_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxscHl4a3RjZnViaG10aHF1anBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk4MTQyMzEsImV4cCI6MjA5NTM5MDIzMX0.4dutcaJzjhuR9IQRlpBetGgLalaSZCiEDdl3oSI_f2A';
const OWNER = '40ad6da8-c25d-4e28-8667-08e7c96ce555'; // tu usuario de SHAAH

function doGet() {
  return HtmlService.createHtmlOutput('<base target="_top"><p style="font:16px system-ui;padding:24px">Cocina se mudó: <a href="' + APP_URL + '">abrir Cocina</a></p>').setTitle('Cocina');
}

function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    const API = { getData, salida, setCantidad, addCompra, addCompras, seAcabo, addInventario, comprado, descartarReceta, addInventarioMany, cocine, usarPronto, guardarMenu, elegirMenu, leerRecibo, reciboOk, avisoOn, avisoOff, avisoAhora, setFlash, retirar, salidaMany, setFlashMany, extender, mover, renombrar };
    if (!authOk_(b.jwt)) out = { ok: false, error: 'auth' };
    else if (!Object.prototype.hasOwnProperty.call(API, b.fn)) out = { ok: false, error: 'función no permitida' };
    else out = { ok: true, data: API[b.fn].apply(null, b.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function authOk_(jwt) {
  if (!jwt) return false;
  const cache = CacheService.getScriptCache();
  const k = 'jwt_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, jwt));
  if (cache.get(k)) return true;
  const r = UrlFetchApp.fetch(SB_URL + '/auth/v1/user', { headers: { apikey: SB_ANON, Authorization: 'Bearer ' + jwt }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return false;
  if (JSON.parse(r.getContentText()).id !== OWNER) return false;
  cache.put(k, '1', 300);
  return true;
}

// Ejecuta esta una vez desde el editor para dar el permiso nuevo (conexión a Supabase).
function autorizar() { UrlFetchApp.fetch(SB_URL + '/auth/v1/health', { headers: { apikey: SB_ANON }, muteHttpExceptions: true }); }
// Ejecuta esta desde el editor si el recibo dice que falta permiso de Drive: pide y prueba los permisos de Drive/Docs.
function autorizarDrive() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']);
  Drive.Files.list({ pageSize: 1 });
  Logger.log('Drive OK');
}
// ===== fin API =====


let _ss;
function ss_() { return _ss || (_ss = SpreadsheetApp.openById(SHEET_ID)); }
function sh_(n) { return ss_().getSheetByName(n); }
function hdr_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }
function rowOf_(sh, id) {
  const n = Math.max(sh.getLastRow() - 1, 1);
  const f = sh.getRange(2, 1, n, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) throw new Error('No encontrado: ' + id);
  return f.getRow();
}
function get_(name, id, col) { const sh = sh_(name), h = hdr_(sh), r = rowOf_(sh, id); return sh.getRange(r, h.indexOf(col) + 1).getValue(); }
function set_(name, id, obj) { const sh = sh_(name), h = hdr_(sh), r = rowOf_(sh, id); Object.keys(obj).forEach(k => sh.getRange(r, h.indexOf(k) + 1).setValue(obj[k])); }
function append_(name, obj) { const sh = sh_(name), h = hdr_(sh); sh.appendRow(h.map(k => obj[k] !== undefined ? obj[k] : '')); }
function rows_(name) {
  const sh = sh_(name);
  const v = sh.getDataRange().getValues();
  const h = v.shift();
  return { sh, h, rows: v.map((r, i) => { const o = { _r: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; }).filter(o => o[h[0]] !== '') };
}
function today_() { const s = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); return new Date(s + 'T00:00:00'); }
function ymd_(d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || ''); }
function asDate_(d) { if (!d) return null; if (d instanceof Date) return new Date(ymd_(d) + 'T00:00:00'); return new Date(String(d).slice(0, 10) + 'T00:00:00'); }
function vence_(o) {
  if (!PERE.includes(o.Tipo) || o.Ubicacion === 'Congelador' || o.Dias_vida === '' || o.Dias_vida == null) return null;
  const e = asDate_(o.Fecha_entrada); if (!e) return null;
  return new Date(e.getTime() + Number(o.Dias_vida) * 864e5);
}
function urg_(dias) { return dias == null ? 'Sin caducidad' : dias < 0 ? 'Vencido' : dias === 0 ? 'Crítico' : dias <= 2 ? 'Por vencer' : 'Fresco'; }
function split_(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }

// Resumen del mes: comidos vs botados (por fecha de salida) y lo que más se bota.
function mes_(invAll) {
  const mk = Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), M = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const rows = invAll.filter(o => (o.Salida === 'Consumido' || o.Salida === 'Desperdiciado') && o.Fecha_salida && ymd_(o.Fecha_salida).slice(0, 7) === mk);
  const bot = {}; rows.filter(o => o.Salida === 'Desperdiciado').forEach(o => { bot[o.Ingrediente] = (bot[o.Ingrediente] || 0) + 1; });
  return { nombre: M[Number(mk.slice(5)) - 1], comido: rows.filter(o => o.Salida === 'Consumido').length, botado: rows.length - rows.filter(o => o.Salida === 'Consumido').length, top: Object.keys(bot).map(k => [k, bot[k]]).sort((a, b) => b[1] - a[1]).slice(0, 3) };
}
function getData() {
  seedRecetas_();
  migrarCategorias_();
  moverAfuera_();
  const t = today_();
  const cat = rows_('Catalogo').rows.map(c => ({ n: c.Ingrediente, cat: c.Categoria, tipo: c.Tipo, dias: c.Dias_vida, ubi: c.Ubicacion_habitual, ret: isTrue_(c.Retirado) }));
  const invAll = rows_('Inventario').rows, hist = {};
  invAll.forEach(o => { if (o.Salida === 'Consumido') hist[o.Ingrediente] = (hist[o.Ingrediente] || 0) + 1; });
  const inv = invAll.filter(o => !o.Salida).map(o => {
    const x = o.Vence_nuevo ? asDate_(o.Vence_nuevo) : null, v0 = x || vence_(o), u = o.Usar_antes ? asDate_(o.Usar_antes) : null;
    const v = u && (!v0 || u < v0) ? u : v0, dias = v ? Math.round((v - t) / 864e5) : null;
    return { id: o.ID, n: o.Ingrediente, cat: o.Categoria, tipo: o.Tipo, ubi: o.Ubicacion, cant: o.Cantidad, nota: o.Nota, entrada: ymd_(o.Fecha_entrada), v0: v0 ? ymd_(v0) : '', usar: u ? ymd_(u) : '', vence: v ? ymd_(v) : '', dias, urg: urg_(dias) };
  });
  const have = new Set(inv.map(i => i.n)), urgent = new Set(inv.filter(i => ['Vencido', 'Crítico', 'Por vencer'].includes(i.urg)).map(i => i.n));
  const rec = rows_('Recetas').rows.filter(r => r.Estado !== 'Descartada').map(r => {
    const ing = split_(r.Ingredientes), tengo = ing.filter(x => have.has(x)), faltan = ing.filter(x => !have.has(x)), rescata = ing.filter(x => urgent.has(x));
    return { id: r.ID, n: r.Nombre, tipo: r.Tipo_comida, pasos: r.Pasos_notas, estado: r.Estado, flex: r.Flexibilidad, min: r.Tiempo_min, link: r.Link, ing, tengo, faltan, rescata, score: rescata.length * 10 + tengo.length - faltan.length * 3 };
  }).sort((a, b) => b.score - a.score);
  const compras = rows_('Lista_de_Compras').rows.filter(c => c.Comprado !== true && String(c.Comprado).toUpperCase() !== 'TRUE').map(c => ({ id: c.ID, n: c.Ingrediente, nota: c.Cantidad_nota, origen: c.Origen, flash: c.Flash === true || String(c.Flash).toUpperCase() === 'TRUE' }));
  return { hoy: ymd_(t), inv, rec, compras, cat, hist, menu: menuRows_(t), aviso: avisoOnState_(), avisoHora: avisoHora_(), mes: mes_(invAll) };
}

// Acciones rápidas: el cliente ya actualizó la pantalla; aquí solo se escribe en la hoja.
function salida(id, tipo, cid) {
  if (!['Consumido', 'Desperdiciado', 'Eliminado'].includes(tipo)) throw new Error('Tipo no válido');
  if (tipo !== 'Eliminado') reponer_(id, tipo === 'Consumido' ? 'Consumido' : 'Se perdió', cid);
  set_('Inventario', id, { Salida: tipo, Fecha_salida: today_() }); return 'ok';
}
// Lo que sale de casa vuelve a compras (menos las sobras cocinadas).
function reponer_(id, origen, cid) {
  if (get_('Inventario', id, 'Tipo') === 'Sobra') return;
  const ing = get_('Inventario', id, 'Ingrediente');
  if (retirado_(ing)) return;
  addCompra_(ing, origen, cid);
}
// Retirar: sale de compras y no vuelve sola. Añadirla a mano (o 'Se acabó') la reactiva.
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function retirado_(ing) { return rows_('Catalogo').rows.some(c => c.Ingrediente === ing && isTrue_(c.Retirado)); }
function unretire_(ing) {
  const d = rows_('Catalogo'), k = d.h.indexOf('Retirado'); if (k < 0) return;
  d.rows.filter(c => c.Ingrediente === ing && isTrue_(c.Retirado)).forEach(c => d.sh.getRange(c._r, k + 1).setValue(false));
}
function retirar(ing) {
  ensureCat_(ing); ensureCol_('Catalogo', 'Retirado');
  const d = rows_('Catalogo'), k = d.h.indexOf('Retirado');
  d.rows.filter(c => c.Ingrediente === ing).forEach(c => d.sh.getRange(c._r, k + 1).setValue(true));
  const L = rows_('Lista_de_Compras');
  L.rows.filter(c => c.Ingrediente === ing && !isTrue_(c.Comprado)).map(c => c._r).sort((a, b) => b - a).forEach(r => L.sh.deleteRow(r));
  return 'ok';
}
function ensureCol_(name, col) { const sh = sh_(name), h = hdr_(sh); if (h.indexOf(col) < 0) sh.getRange(1, h.length + 1).setValue(col); }
// Todavía está bueno: nueva fecha de vencimiento (reemplaza la calculada) y quita el "usar pronto".
function extender(id, ymd) {
  ensureCol_('Inventario', 'Vence_nuevo');
  const o = { Vence_nuevo: ymd ? new Date(ymd + 'T00:00:00') : '' };
  if (hdr_(sh_('Inventario')).includes('Usar_antes')) o.Usar_antes = '';
  set_('Inventario', id, o); return 'ok';
}
function salidaMany(list) { (list || []).forEach(a => salida(a[0], a[1], a[2])); return 'ok'; }
function setFlashMany(ids, on) { ensureCol_('Lista_de_Compras', 'Flash'); (ids || []).forEach(id => set_('Lista_de_Compras', id, { Flash: !!on })); return 'ok'; }
function setFlash(cid, on) { ensureCol_('Lista_de_Compras', 'Flash'); set_('Lista_de_Compras', cid, { Flash: !!on }); return 'ok'; }
// Cambiar el nombre de un producto: en casa, en compras y en el catálogo. El nombre viejo queda retirado (no vuelve a compras).
function renombrar(viejo, nuevo) {
  nuevo = String(nuevo || '').trim().replace(/\s+/g, ' ');
  if (!viejo || !nuevo || nuevo === viejo || nuevo.length > 60) throw new Error('Nombre no válido');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureCol_('Catalogo', 'Retirado');
    const C = rows_('Catalogo'), h = C.h, old = C.rows.find(c => c.Ingrediente === viejo), nw = C.rows.find(c => c.Ingrediente === nuevo);
    const cocinado = !old || old.Categoria === 'Cocinado' || old.Tipo === 'Sobra';
    if (!nw) {
      if (cocinado) ensureCat_(nuevo);
      else append_('Catalogo', { Ingrediente: nuevo, Categoria: old.Categoria, Tipo: old.Tipo, Dias_vida: old.Dias_vida, Ubicacion_habitual: old.Ubicacion_habitual, Retirado: false });
    } else unretire_(nuevo);
    if (old) C.sh.getRange(old._r, h.indexOf('Retirado') + 1).setValue(true);
    const catNuevo = rows_('Catalogo').rows.find(c => c.Ingrediente === nuevo).Categoria;
    rows_('Inventario').rows.filter(o => o.Ingrediente === viejo && !o.Salida).forEach(o => {
      const x = { Ingrediente: nuevo }; if (o.Tipo === 'Sobra') x.Tipo = 'Perecedero'; if (o.Categoria === 'Cocinado') x.Categoria = catNuevo;
      set_('Inventario', o.ID, x);
    });
    const L = rows_('Lista_de_Compras'), pend = c => !isTrue_(c.Comprado);
    const dup = L.rows.some(c => c.Ingrediente === nuevo && pend(c)), ni = L.h.indexOf('Ingrediente');
    L.rows.filter(c => c.Ingrediente === viejo && pend(c)).map(c => c._r).sort((a, b) => b - a).forEach(r => { if (dup) L.sh.deleteRow(r); else L.sh.getRange(r, ni + 1).setValue(nuevo); });
  } finally { lock.releaseLock(); }
  return 'ok';
}
function mover(id, ubi, cat) {
  if (!['Congelador', 'Nevera', 'Despensa', 'Encimera'].includes(ubi)) throw new Error('Lugar no válido');
  if (!cat || String(cat).length > 40) throw new Error('Categoría no válida');
  const o = rows_('Inventario').rows.find(x => x.ID === id); if (!o) throw new Error('No encontrado');
  set_('Inventario', id, { Ubicacion: ubi, Categoria: cat });
  ensureCat_(o.Ingrediente);
  const d = rows_('Catalogo'), c = d.rows.find(x => x.Ingrediente === o.Ingrediente);
  if (c) { const h = d.h; d.sh.getRange(c._r, h.indexOf('Ubicacion_habitual') + 1).setValue(ubi); d.sh.getRange(c._r, h.indexOf('Categoria') + 1).setValue(cat); }
  return 'ok';
}
function setCantidad(id, cant) {
  if (!['Lleno', 'Medio', 'Poco'].includes(cant)) throw new Error('Cantidad no válida');
  set_('Inventario', id, { Cantidad: cant }); return 'ok';
}
function addCompra_(ing, origen, cid) {
  if (origen === 'Manual' || origen === 'Se acabó') unretire_(ing);
  const d = rows_('Lista_de_Compras');
  if (d.rows.some(c => c.Ingrediente === ing && String(c.Comprado).toUpperCase() !== 'TRUE')) return;
  append_('Lista_de_Compras', { ID: cid || ('COM-' + Date.now().toString(36).toUpperCase()), Ingrediente: ing, Cantidad_nota: '', Origen: origen, Fecha_agregado: today_(), Comprado: false });
}
function addCompra(ing, cid) { ensureCat_(ing); addCompra_(ing, 'Manual', cid); return 'ok'; }
function addCompras(list, origen) { (list || []).forEach(p => addCompra_(p[0], origen || 'Receta', p[1])); return 'ok'; }
function seAcabo(id, cid) { addCompra_(get_('Inventario', id, 'Ingrediente'), 'Se acabó', cid); set_('Inventario', id, { Salida: 'Consumido', Fecha_salida: today_() }); return 'ok'; }
function ensureCat_(ing) {
  const d = rows_('Catalogo'); if (d.rows.some(c => c.Ingrediente === ing)) return;
  const k = classify(ing);
  append_('Catalogo', { Ingrediente: ing, Categoria: k.cat, Tipo: k.tipo, Dias_vida: k.dias, Ubicacion_habitual: k.ubi });
}
function addInventario(ing, cant, nota, sobra, iid) {
  ing = String(ing || '').trim(); if (!ing) throw new Error('Falta el ingrediente');
  let c;
  const sd = sobra ? (Number(sobra) > 0 && Number(sobra) <= 2 ? Number(sobra) : 2) : 0;
  if (sd) c = { Categoria: 'Cocinado', Tipo: 'Sobra', Dias_vida: sd, Ubicacion_habitual: 'Nevera' };
  else { ensureCat_(ing); c = rows_('Catalogo').rows.find(x => x.Ingrediente === ing); }
  const tipo = c.Tipo, dias = tipo === 'Sobra' ? (sd || 2) : tipo === 'Perecedero' ? c.Dias_vida : '';
  const o = { ID: iid || ('INV-' + Date.now().toString(36).toUpperCase()), Ingrediente: ing, Categoria: c.Categoria, Tipo: tipo, Ubicacion: c.Ubicacion_habitual, Cantidad: cant || 'Lleno', Fecha_entrada: today_(), Dias_vida: dias, Salida: '', Fecha_salida: '', Nota: nota || '' };
  o.Fecha_vence = vence_(o) || '';
  append_('Inventario', o); return 'ok';
}
function comprado(cid, iid) {
  set_('Lista_de_Compras', cid, { Comprado: true });
  return addInventario(get_('Lista_de_Compras', cid, 'Ingrediente'), 'Lleno', '', false, iid);
}

function descartarReceta(id) { set_('Recetas', id, { Estado: 'Descartada' }); return 'ok'; }
function addInventarioMany(list) { (list || []).forEach(a => addInventario(a[0], a[1], a[2], a[3], a[4])); return 'ok'; }
function cocine(ids, sobraName, dias, iid, cids) {
  (ids || []).forEach((id, k) => { reponer_(id, 'Cocinado', (cids || [])[k]); set_('Inventario', id, { Salida: 'Consumido', Fecha_salida: today_() }); });
  if (sobraName) addInventario(sobraName, 'Lleno', 'Sobra de receta', dias || 2, iid);
  return 'ok';
}
function usarPronto(id, ymd) {
  const sh = sh_('Inventario'), h = hdr_(sh);
  if (!h.includes('Usar_antes')) sh.getRange(1, h.length + 1).setValue('Usar_antes');
  set_('Inventario', id, { Usar_antes: ymd ? new Date(ymd + 'T00:00:00') : '' }); return 'ok';
}
// Recetas sugeridas con lo que había en casa el 20-nov-2026 (se cargan una sola vez).
function seedRecetas_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('seed_recetas_v1')) return;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try { if (p.getProperty('seed_recetas_v1')) return; seedRecetasRun_(p); } finally { lock.releaseLock(); }
}
function seedRecetasRun_(p) {
  const L = [
    ['REC-008', 'Shakshuka verde de zucchini', 'Desayuno/Cena', 'Huevos , Zucchini , Calabacín , Cebolla blanca , Ajo , Chile picante , Cilantro , Queso crema', 'Sofríe cebolla, ajo y chile. Agrega zucchini y calabacín en cubos hasta que se ablanden. Haz huecos, rompe los huevos, tapa 5-6 min. Termina con cucharadas de queso crema y cilantro.', 20],
    ['REC-009', 'Crema de calabacín y zucchini con coco', 'Almuerzo/Cena', 'Calabacín , Zucchini , Crema de coco , Cebolla blanca , Ajo , Perejil', 'Sofríe cebolla y ajo, agrega calabacín y zucchini, cubre con agua y cocina 15 min. Licúa con la crema de coco, sal y pimienta. Perejil arriba.', 30],
    ['REC-010', 'Ensalada de quinoa, pepino y atún', 'Almuerzo/Cena', 'Quinoa , Atún en lata , Pepino , Cebolla morada , Perejil , Cilantro', 'Cocina la quinoa y enfríala. Mezcla con atún, pepino y cebolla morada en cubos, hierbas picadas, limón, aceite y sal.', 20],
    ['REC-011', 'Arroz frito con huevo y edamame', 'Almuerzo/Cena', 'Arroz blanco , Huevos , Edamame , Ajo triturado , Cebolla morada , Chile picante', 'Mejor con arroz del día anterior. Saltea cebolla, ajo y chile, agrega edamame y arroz a fuego alto. Empuja a un lado, revuelve los huevos y mezcla todo. Salsa de soya si tienes.', 20],
    ['REC-012', 'Brisket con puré', 'Almuerzo/Cena', 'Brisket cocinado , Bolitas de puré , Cebolla morada', 'Descongela el brisket la noche anterior. Caliéntalo en sartén tapada con un chorrito de agua. Hornea las bolitas de puré. Cebolla morada encurtida rápida (limón + sal) arriba.', 15],
    ['REC-013', 'Tostadas de queso crema y pepino', 'Desayuno/Snack', 'Pan integral , Queso crema , Pepino , Nana', 'Tuesta el pan, unta queso crema, rodajas finas de pepino, nana picada, sal y pimienta.', 5],
    ['REC-014', 'Avena horneada de banana y manzana', 'Desayuno', 'Avena , Banana , Manzana , Huevos , Leche de macadamia , Chocolate amargo', 'Mezcla avena, banana aplastada, huevos y leche. Agrega manzana en cubos y chocolate picado. Hornea 25 min a 180°C. Sirve para 2-3 desayunos.', 30],
    ['REC-015', 'Tabule de quinoa', 'Almuerzo/Acompañante', 'Quinoa , Perejil , Nana , Pepino , Cebolla morada', 'Quinoa cocida y fría con mucho perejil y nana picados, pepino y cebolla en cubitos, limón, aceite de oliva y sal.', 20],
    ['REC-016', 'Pollo al coco con arroz', 'Almuerzo/Cena', 'Pechuga de pollo , Crema de coco , Cebolla blanca , Ajo , Chile picante , Cilantro , Arroz blanco', 'Dora el pollo en cubos, saca. Sofríe cebolla, ajo y chile, agrega la crema de coco y el pollo, cocina 10 min. Cilantro y arroz.', 35],
    ['REC-017', 'Topinambur asado con ghee', 'Acompañante', 'jerusalem artichoke , Ghee , Ajo , Perejil', 'Lava y corta en mitades. Mezcla con ghee derretido, ajo y sal. Horno 220°C por 30-35 min hasta dorar. Perejil arriba.', 35]
  ];
  const d = rows_('Recetas'), have = new Set(d.rows.map(r => r.ID));
  L.filter(r => !have.has(r[0])).forEach(r => append_('Recetas', { ID: r[0], Nombre: r[1], Tipo_comida: r[2], Ingredientes: r[3], Pasos_notas: r[4], Estado: 'Por probar', Flexibilidad: 'Flexible', Tiempo_min: r[5], Fuente: 'Claude' }));
  p.setProperty('seed_recetas_v1', '1');
}

// [regex, categoria, tipo, dias, lugar]  — el primero que coincide gana. Pensado para una cocina en Israel.
const CLS_RULES = [
  ['papel|aluminio|bolsa|jab[oó]n|detergente|servillet|esponja|film|toalla|נייר|אלומיניום|שקיות|סבון', 'Hogar', 'Despensa', '', 'Despensa'],
  ['helado|גלידה', 'Dulces y snacks', 'Congelado', '', 'Congelador'],
  ['en lata|enlatad|conserva|crema de coco|leche de coco|שימורים|קופסת', 'Conservas', 'Despensa', '', 'Despensa'],
  ['tahini|tehina|טחינה', 'Salsas y especias', 'Despensa', '', 'Despensa'],
  ['hummus|humus|חומוס מוכן|מטבוחה|matbuja|baba|chatzil|guacamole', 'Salsas y especias', 'Perecedero', 5, 'Nevera'],
  ['salsa|ketchup|mayonesa|mostaza|soya|soja|vinagre|aceite|\\bsal\\b|pimienta|comino|p[aá]prika|zaatar|za.atar|sumac|especia|caldo|miel|sil[aá]n|\\bamba\\b|schug|zhug|s.hug|harissa|ajo triturado|רוטב|קטשופ|מיונז|חרדל|סויה|חומץ|שמן|מלח|כמון|פפריקה|זעתר|סומק|דבש|סילאן|עמבה|סחוג', 'Salsas y especias', 'Despensa', '', 'Despensa'],
  ['ch[ií]a|nuez|nueces|almendra|anacardo|caj[uú]|pistach|semilla|s[eé]samo|linaza|man[ií]\\b|cacahuate|girasol|pecan|פיצוח|שקדים|אגוז|צ.יה|פשתן|שומשום|בוטנים|גרעינים', 'Semillas y frutos secos', 'Despensa', '', 'Despensa'],
  ['chocolate|galleta|dulce|caramelo|chips|snack|\\bbamba\\b|bisli|halv|במבה|ביסלי|שוקולד|עוגיות|חלבה|ופל', 'Dulces y snacks', 'Despensa', '', 'Despensa'],
  ['\\bagua\\b|jugo|zumo|soda|cerveza|vino|caf[eé]|\\bt[eé]\\b|kombucha|refresco|\\bcola\\b|coca|מים|מיץ|בירה|יין|קפה|\\bתה\\b|סודה', 'Bebidas', 'Despensa', '', 'Despensa'],
  ['sobra|cocinad|cocido|guiso|\\bsopa\\b', 'Cocinado', 'Sobra', 2, 'Nevera'],
  ['pizza|nuggets|schnitzel congel|papas prefrit|bolitas|burekas|bureka|schnitzel|shnitzel|malawach|jachnun|ziva|בורקס|מלאווח|ג.חנון|שניצל', 'Congelados listos', 'Congelado', '', 'Congelador'],
  ['pollo|\\bres\\b|carne|molida|steak|bife|salm[oó]n|at[uú]n|pescado|brisket|tu[eé]tano|kebab|shawarma|hamburguesa|pavo|cordero|merguez|ternera|camar[oó]n|עוף|בשר|דג|סלמון|טונה|הודו|כבש|שווארמה|קבב|המבורגר', 'Carnes y pescados', 'Perecedero', 2, 'Nevera'],
  ['queso|manchego|amsterdam|parmesan|mozzarella|gouda|brie|feta|ricotta|halloumi|emmental|leche|yog|labneh|leben|lebe|cottage|crema|mantequilla|ghee|huevo|gvina|tzfat|bulgar|גבינה|חלב|יוגורט|לבנה|לבן|קוטג|שמנת|חמאה|ביצ|מעדן|אשל|גיל', 'Lácteos y huevos', 'Perecedero', 7, 'Nevera'],
  ['cilantro|perejil|\\bnana\\b|menta|eneldo|albahaca|ceboll[ií]n|romero|tomillo|כוסברה|פטרוזיליה|נענע|שמיר|בזיליקום|בצל ירוק|רוזמרין|טימין', 'Hierbas frescas', 'Perecedero', 5, 'Nevera'],
  ['\\bpapas?\\b|cebolla|\\bajo\\b|batata|camote|calabaza|תפוח אדמה|תפו.א|בצל|שום|בטטה|דלעת', 'Verduras', 'Despensa', '', 'Despensa'],
  ['\\bpan\\b|pita|lafa|laffa|jal[aá]|challah|bagel|tortilla|wrap|baguette|croissant|לחם|פית|לאפה|חלה|בייגל|באגט|קרואסון', 'Panes', 'Perecedero', 3, 'Despensa'],
  ['pepino|zucchini|calabac|chile|topinambur|artichoke|מלפפון|קישוא|ארטישוק ירושלמי', 'Verduras', 'Perecedero', 6, 'Despensa'],
  ['manzana|תפוח(?! אדמה)', 'Frutas', 'Perecedero', 6, 'Despensa'],
  ['banana|pl[aá]tano|aguacate|palta|בננה|אבוקדו', 'Frutas', 'Perecedero', 5, 'Despensa'],
  ['manzana|naranja|mandarina|lim[oó]n|lima|uva|fresa|frutilla|mango|pi[nñ]a|ar[aá]ndano|pera|durazno|melocot|nectarin|ciruela|sand[ií]a|mel[oó]n|granada|kiwi|d[aá]til|higo|caqui|persimon|frambuesa|cereza|תפוח|תפוז|קלמנטינה|לימון|ענב|תות|מנגו|אננס|אוכמנ|אגס|אפרסק|נקטרינה|שזיף|אבטיח|מלון|רימון|קיווי|תמר|תאנה|אפרסמון|דובדבן|פטל', 'Frutas', 'Perecedero', 6, 'Nevera'],
  ['tomate|pepino|zucchini|calabac|berenjena|pimiento|piment|zanahoria|lechuga|espinaca|\\bcol\\b|repollo|kale|apio|champi|hongo|seta|chile|jalape|r[uú]cula|br[oó]coli|coliflor|remolacha|ma[ií]z|edamame|ejote|jud[ií]a|arveja|guisante|alcachofa|topinambur|artichoke|esp[aá]rrago|puerro|hinojo|עגבני|מלפפון|קישוא|חציל|פלפל|גזר|חסה|תרד|כרוב|קייל|סלרי|פטריות|סלק|תירס|שעועית ירוקה|אפונה|ארטישוק|אספרגוס|כרשה|שומר', 'Verduras', 'Perecedero', 6, 'Nevera'],
  ['lenteja|garbanzo|frijol|poroto|alubia|\\bhabas?\\b|עדשים|חומוס יבש|גרגרי חומוס|שעועית|פול', 'Legumbres', 'Despensa', '', 'Despensa'],
  ['arroz|quinoa|avena|pasta|fideo|espagueti|macarr|mac and cheese|cusc|cous|bulgur|harina|panko|ptitim|cereal|granola|polenta|אורז|קינואה|שיבולת|פסטה|ספגטי|קוסקוס|בורגול|קמח|פתיתים|דגני|גרנולה', 'Granos y pastas', 'Despensa', '', 'Despensa']
];
function classify(name) {
  const n = String(name || '').toLowerCase();
  let r = null;
  for (const x of CLS_RULES) if (new RegExp(x[0], 'i').test(n)) { r = x; break; }
  const out = r ? { cat: r[1], tipo: r[2], dias: r[3], ubi: r[4] } : { cat: 'Otro', tipo: 'Despensa', dias: '', ubi: 'Despensa' };
  if (/congel|frozen|קפוא/i.test(n)) { out.tipo = 'Congelado'; out.dias = ''; out.ubi = 'Congelador'; }
  return out;
}

// Categoría y lugar fijos para lo que ya está en el catálogo (cocina en Israel). Se aplica una sola vez.
const FIX_ = {
  'Steak de atún': ['Carnes y pescados', 'Congelador'], 'Pechuga de pollo': ['Carnes y pescados', 'Congelador'], 'Carne molida': ['Carnes y pescados', 'Congelador'],
  'Steak de res': ['Carnes y pescados', 'Nevera'], 'Brisket': ['Carnes y pescados', 'Nevera'], 'Salmón': ['Carnes y pescados', 'Nevera'], 'Salmón ahumado': ['Carnes y pescados', 'Nevera'], 'Tuétano': ['Carnes y pescados', 'Congelador'],
  'Brisket cocinado': ['Cocinado', 'Nevera'], 'Arroz cocido': ['Cocinado', 'Nevera'], 'Pollo cocinado': ['Cocinado', 'Nevera'], 'Sobras varias': ['Cocinado', 'Nevera'],
  'Papas prefritas': ['Congelados listos', 'Congelador'], 'Bolitas de puré': ['Congelados listos', 'Congelador'], 'Mini pizza': ['Congelados listos', 'Congelador'],
  'Edamame': ['Verduras', 'Congelador'], 'Mango congelado': ['Frutas', 'Congelador'], 'Piña congelada': ['Frutas', 'Congelador'], 'Arándanos congelados': ['Frutas', 'Congelador'], 'Fresa congelada': ['Frutas', 'Congelador'],
  'Pan integral': ['Panes', 'Congelador'], 'Pan pita': ['Panes', 'Despensa'], 'Ajo triturado': ['Salsas y especias', 'Congelador'],
  'Old Amsterdam': ['Lácteos y huevos', 'Nevera'], 'Manchego trufado': ['Lácteos y huevos', 'Nevera'], 'Yogur griego': ['Lácteos y huevos', 'Nevera'], 'Ghee': ['Lácteos y huevos', 'Nevera'], 'Mantequilla de trufa': ['Lácteos y huevos', 'Nevera'], 'Huevos': ['Lácteos y huevos', 'Nevera'], 'Queso crema': ['Lácteos y huevos', 'Nevera'], 'Crema de leche': ['Lácteos y huevos', 'Nevera'], 'Queso búlgaro 5%': ['Lácteos y huevos', 'Nevera'], 'Labneh': ['Lácteos y huevos', 'Nevera'], 'Leche de macadamia': ['Lácteos y huevos', 'Nevera'], 'Leche proteica': ['Lácteos y huevos', 'Nevera'],
  'Cilantro': ['Hierbas frescas', 'Nevera'], 'Perejil': ['Hierbas frescas', 'Nevera'], 'Nana': ['Hierbas frescas', 'Nevera'], 'Cebollín': ['Hierbas frescas', 'Nevera'],
  'Zucchini': ['Verduras', 'Nevera'], 'Calabacín': ['Verduras', 'Nevera'], 'Pepino': ['Verduras', 'Nevera'], 'Chile picante': ['Verduras', 'Nevera'], 'jerusalem artichoke': ['Verduras', 'Nevera'],
  'Papas': ['Verduras', 'Despensa'], 'Cebolla morada': ['Verduras', 'Despensa'], 'Cebolla blanca': ['Verduras', 'Despensa'], 'Ajo': ['Verduras', 'Despensa'],
  'Manzana': ['Frutas', 'Nevera'], 'Banana': ['Frutas', 'Despensa'],
  'Arroz blanco': ['Granos y pastas', 'Despensa'], 'Quinoa': ['Granos y pastas', 'Despensa'], 'Avena': ['Granos y pastas', 'Despensa'], 'Panko': ['Granos y pastas', 'Despensa'], 'Mac and cheese (paquete)': ['Granos y pastas', 'Despensa'],
  'Lentejas rojas': ['Legumbres', 'Despensa'], 'Lentejas beluga': ['Legumbres', 'Despensa'], 'Chía': ['Semillas y frutos secos', 'Despensa'],
  'Atún en lata': ['Conservas', 'Despensa'], 'Alcachofas en conserva': ['Conservas', 'Despensa'], 'Crema de coco': ['Conservas', 'Despensa'], 'Salsa de tomate': ['Salsas y especias', 'Despensa'],
  'Chocolate amargo': ['Dulces y snacks', 'Despensa'], 'Agua de coco': ['Bebidas', 'Despensa'], 'Agua con gas': ['Bebidas', 'Despensa'], 'Cerveza': ['Bebidas', 'Despensa'],
  'Papel aluminio': ['Hogar', 'Despensa'], 'Papel de flojo': ['Hogar', 'Despensa']
};
function migrarCategorias_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('cats_v1')) return;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    if (p.getProperty('cats_v1')) return;
    const look = n => FIX_[n] || (k => [k.cat, k.ubi])(classify(n));
    // Catálogo: categoría y lugar habitual
    let sh = sh_('Catalogo'), v = sh.getDataRange().getValues(), h = v[0];
    let ci = h.indexOf('Categoria'), ui = h.indexOf('Ubicacion_habitual'), ni = h.indexOf('Ingrediente');
    for (let r = 1; r < v.length; r++) { if (!v[r][ni]) continue; const k = look(v[r][ni]); v[r][ci] = k[0]; v[r][ui] = k[1]; }
    sh.getRange(1, 1, v.length, h.length).setValues(v);
    // Inventario: categoría de todo; lugar de lo que sigue en casa (lo que congelaste se queda en el congelador)
    sh = sh_('Inventario'); v = sh.getDataRange().getValues(); h = v[0];
    ci = h.indexOf('Categoria'); ui = h.indexOf('Ubicacion'); ni = h.indexOf('Ingrediente');
    const si = h.indexOf('Salida'), ti = h.indexOf('Tipo');
    const cats = v.map((row, r) => [r === 0 ? h[ci] : (row[ni] ? (row[ti] === 'Sobra' ? 'Cocinado' : look(row[ni])[0]) : row[ci])]);
    const ubis = v.map((row, r) => [r === 0 || !row[ni] || row[si] || row[ui] === 'Congelador' || row[ti] === 'Sobra' ? row[ui] : look(row[ni])[1]]);
    sh.getRange(1, ci + 1, v.length, 1).setValues(cats);
    sh.getRange(1, ui + 1, v.length, 1).setValues(ubis);
    p.setProperty('cats_v1', '1');
  } finally { lock.releaseLock(); }
}

// David guarda estas cosas fuera de la nevera (20-nov-2026). Se aplica una sola vez.
const AFUERA_ = ['Zucchini', 'Calabacín', 'Pepino', 'Manzana', 'Chile picante', 'jerusalem artichoke'];
function moverAfuera_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('afuera_v1')) return;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    if (p.getProperty('afuera_v1')) return;
    [['Catalogo', 'Ubicacion_habitual', false], ['Inventario', 'Ubicacion', true]].forEach(([name, col, onlyActive]) => {
      const sh = sh_(name), v = sh.getDataRange().getValues(), h = v[0], ui = h.indexOf(col), ni = h.indexOf('Ingrediente'), si = h.indexOf('Salida');
      const out = v.map((row, r) => [r > 0 && AFUERA_.includes(row[ni]) && (!onlyActive || !row[si]) && row[ui] !== 'Congelador' ? 'Despensa' : row[ui]]);
      sh.getRange(1, ui + 1, v.length, 1).setValues(out);
    });
    p.setProperty('afuera_v1', '1');
  } finally { lock.releaseLock(); }
}

// Menú de la semana (pestaña Menu_y_Opciones)
function menuRows_(t) {
  return rows_('Menu_y_Opciones').rows.filter(m => m.Fecha && asDate_(m.Fecha) >= t)
    .map(m => ({ id: m.ID, fecha: ymd_(m.Fecha), momento: m.Momento, ops: [m.Opcion_1, m.Opcion_2, m.Opcion_3].filter(Boolean).map(String), elegida: m.Elegida || '', nota: m.Nota || '' }));
}
function guardarMenu(list) {
  const sh = sh_('Menu_y_Opciones'), t = today_(), v = sh.getDataRange().getValues(), fi = v[0].indexOf('Fecha');
  for (let r = v.length - 1; r >= 1; r--) { const f = v[r][fi]; if (v[r][0] && f && asDate_(f) >= t) sh.deleteRow(r + 1); }
  (list || []).forEach(m => append_('Menu_y_Opciones', { ID: m.id, Fecha: new Date(m.fecha + 'T00:00:00'), Momento: m.momento || 'Comida', Opcion_1: m.ops[0] || '', Opcion_2: m.ops[1] || '', Opcion_3: m.ops[2] || '', Elegida: m.elegida || '', Nota: m.nota || '' }));
  return 'ok';
}
function elegirMenu(id, nombre) { set_('Menu_y_Opciones', id, { Elegida: nombre || '' }); return 'ok'; }
// Recibo: OCR con Google Drive (hebreo) y alta en inventario
function leerRecibo(b64, mime) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', 'recibo.jpg');
  const meta = { name: 'cocina-recibo-ocr', mimeType: 'application/vnd.google-apps.document' };
  let f;
  try { f = Drive.Files.create(meta, blob, { ocrLanguage: 'iw' }); } // Google usa 'iw' para hebreo
  catch (e) { f = Drive.Files.create(meta, blob); }
  try { return DocumentApp.openById(f.id).getBody().getText(); } finally { try { Drive.Files.remove(f.id); } catch (e) {} }
}
function reciboOk(list, cids) {
  (cids || []).forEach(id => { try { set_('Lista_de_Compras', id, { Comprado: true }); } catch (e) {} });
  return addInventarioMany(list);
}

// ===== Aviso diario en Google Calendar (12:00, hora de Israel) =====
// Cada día a las ~11:00 se crea un evento a las 12:00 con lo que vence; el celular avisa con la notificación del calendario.
function avisoOnState_() { try { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'avisoCalendario'); } catch (e) { return false; } }
// Hora del aviso (HH:MM, hora de Israel). El trigger corre ~1 h antes y crea el evento a esa hora.
function avisoHora_() { return PropertiesService.getScriptProperties().getProperty('AVISO_HORA') || '12:00'; }
function avisoMin_() { const p = avisoHora_().split(':'); return Number(p[0]) * 60 + Number(p[1] || 0); }
function avisoOn(hhmm) {
  if (hhmm != null) {
    if (!/^([01]?\d|2[0-3]):[0-5]\d$/.test(String(hhmm))) throw new Error('Hora no válida');
    const p = String(hhmm).split(':'); PropertiesService.getScriptProperties().setProperty('AVISO_HORA', ('0' + Number(p[0])).slice(-2) + ':' + p[1]);
  }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'avisoCalendario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('avisoCalendario').timeBased().everyDays(1).atHour(Math.floor(((avisoMin_() - 60 + 1440) % 1440) / 60)).nearMinute(0).inTimezone(TZ).create();
  avisoCalendario(true);
  return true;
}
function avisoOff() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'avisoCalendario').forEach(t => ScriptApp.deleteTrigger(t));
  borrarEvento_();
  return false;
}
function avisoAhora() { avisoCalendario(true); return avisoOnState_(); }
function borrarEvento_() {
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('AVISO_EVT');
  if (id) { try { const ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) {} P.deleteProperty('AVISO_EVT'); }
}
function avisoCalendario(manual) {
  const ahora = new Date(), mins = Number(Utilities.formatDate(ahora, TZ, 'H')) * 60 + Number(Utilities.formatDate(ahora, TZ, 'm'));
  const off = mins >= avisoMin_() ? 1 : 0; // pasada la hora del aviso, se arma el de mañana
  const D = getData();
  const dd = i => i.dias == null ? null : i.dias - off;
  const hoy = D.inv.filter(i => dd(i) != null && dd(i) <= 0), man = D.inv.filter(i => dd(i) === 1);
  const fecha = Utilities.formatDate(new Date(ahora.getTime() + off * 864e5), TZ, 'yyyy-MM-dd');
  borrarEvento_();
  if (!hoy.length && !man.length && manual !== true) return;
  const urg = new Set(hoy.concat(man).map(i => i.n));
  const rec = D.rec.map(r => ({ r, n: r.ing.filter(x => urg.has(x)).length })).filter(x => x.n > 0).sort((a, b) => b.n - a.n || a.r.faltan.length - b.r.faltan.length)[0];
  const el = (D.menu || []).find(m => m.fecha === fecha && m.elegida);
  const titulo = hoy.length ? '🍳 Usa hoy: ' + hoy.map(i => i.n).join(', ') : man.length ? '🍳 Vence mañana: ' + man.map(i => i.n).join(', ') : '🍳 Cocina: nada vence hoy';
  const det = [];
  if (hoy.length) det.push('Hoy: ' + hoy.map(i => i.n + (dd(i) < 0 ? ' (vencido)' : '')).join(', '));
  if (man.length) det.push('Mañana: ' + man.map(i => i.n).join(', '));
  if (el) det.push('', 'Hoy toca: ' + el.elegida);
  else if (rec) det.push('', 'Idea: ' + rec.r.n + ' (usa ' + rec.r.ing.filter(x => urg.has(x)).join(', ') + ')');
  det.push('', 'Abrir Cocina: ' + APP_URL);
  const start = Utilities.parseDate(fecha + ' ' + avisoHora_(), TZ, 'yyyy-MM-dd HH:mm');
  const ev = CalendarApp.getDefaultCalendar().createEvent(titulo.slice(0, 120), start, new Date(start.getTime() + 15 * 60000), { description: det.join('\n') });
  ev.removeAllReminders(); ev.addPopupReminder(0);
  PropertiesService.getScriptProperties().setProperty('AVISO_EVT', ev.getId());
}
