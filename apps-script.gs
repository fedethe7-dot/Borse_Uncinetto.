/**
 * Le borse della mamma - collegamento al foglio Google
 * Incolla tutto questo codice in Estensioni > Apps Script del foglio,
 * poi pubblica come App web (Esegui come: Me, Chi può accedere: Chiunque).
 */

const FOGLIO_DATI = 'Dati app';      // copia completa usata dall'app (non toccarla)
const PEZZO = 40000;                  // caratteri per cella

function doGet(e) {
  return risposta({ ok: true, dati: leggi() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const corpo = JSON.parse(e.postData.contents);
    if (corpo.azione !== 'scrivi' || !corpo.dati) return risposta({ ok: false, errore: 'richiesta non valida' });
    const attuali = leggi();
    // Se sul foglio c'è già una versione più recente, non la sovrascrivo
    if (attuali && Number(attuali.updatedAt) > Number(corpo.dati.updatedAt)) {
      return risposta({ ok: true, vecchio: true });
    }
    scrivi(corpo.dati);
    aggiornaTabelle(corpo.dati);
    return risposta({ ok: true });
  } catch (err) {
    return risposta({ ok: false, errore: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- copia completa ---------- */
function leggi() {
  const sh = SpreadsheetApp.getActive().getSheetByName(FOGLIO_DATI);
  if (!sh || sh.getLastRow() === 0) return null;
  const testo = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]).join('');
  if (!testo) return null;
  try { return JSON.parse(testo); } catch (e) { return null; }
}

function scrivi(dati) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FOGLIO_DATI) || ss.insertSheet(FOGLIO_DATI);
  const testo = JSON.stringify(dati);
  const pezzi = [];
  for (let i = 0; i < testo.length; i += PEZZO) pezzi.push([testo.slice(i, i + PEZZO)]);
  sh.clear();
  sh.getRange(1, 1, pezzi.length, 1).setNumberFormat('@').setValues(pezzi);
  sh.hideSheet();
}

/* ---------- tabelle leggibili ---------- */
function aggiornaTabelle(d) {
  const mat = {};
  (d.materiali || []).forEach(m => mat[m.id] = m);
  const costo = mod => (mod.materiali || []).reduce((t, r) => t + (mat[r.id] ? Number(mat[r.id].prezzo) * Number(r.qta) : 0), 0);
  const tondo = n => Math.round(Number(n || 0) * 100) / 100;

  tabella('Vendite',
    ['Borsa', 'Stato', 'Finita il', 'Venduta il', 'Prezzo vendita', 'Costo materiali', 'Guadagno', 'Ore', 'Cliente'],
    (d.lavori || []).slice().sort((a, b) => String(b.dataFine).localeCompare(String(a.dataFine))).map(l => [
      l.nome, l.stato === 'venduta' ? 'Venduta' : 'Da vendere', data(l.dataFine), data(l.dataVendita),
      l.stato === 'venduta' ? tondo(l.prezzo) : '', tondo(l.costo),
      l.stato === 'venduta' ? tondo(l.prezzo - l.costo) : '', l.ore || '', l.cliente || ''
    ]), [5, 6, 7]);

  tabella('Acquisti',
    ['Data', 'Materiale', 'Quantità', 'Unità', 'Totale'],
    (d.spese || []).slice().sort((a, b) => String(b.data).localeCompare(String(a.data))).map(s => [
      data(s.data), s.nome, s.quantita, s.unita, tondo(s.totale)
    ]), [5]);

  tabella('Materiali',
    ['Materiale', 'Tipo', 'Si compra a', 'Prezzo'],
    (d.materiali || []).map(m => [m.nome, m.tipo, m.unita, tondo(m.prezzo)]), [4]);

  tabella('Modelli',
    ['Borsa', 'Ore di lavoro', 'Costo materiali', 'Prezzo', 'Resta'],
    (d.modelli || []).map(m => [m.nome, m.ore, tondo(costo(m)), tondo(m.prezzo), tondo(m.prezzo - costo(m))]), [3, 4, 5]);
}

function tabella(nome, intestazioni, righe, colonneEuro) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
  sh.clear();
  sh.getRange(1, 1, 1, intestazioni.length).setValues([intestazioni]).setFontWeight('bold').setBackground('#F4E4EC');
  sh.setFrozenRows(1);
  if (righe.length) {
    sh.getRange(2, 1, righe.length, intestazioni.length).setValues(righe);
    (colonneEuro || []).forEach(c => sh.getRange(2, c, righe.length, 1).setNumberFormat('€ #,##0.00'));
  }
  sh.autoResizeColumns(1, intestazioni.length);
}

function data(s) {
  if (!s) return '';
  const p = String(s).split('-');
  return p.length === 3 ? new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])) : s;
}

function risposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
