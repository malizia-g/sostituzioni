/**
 * Funzioni di supporto: normalizzazione testi, date, lettura fogli.
 */

/** Normalizza un nome per il confronto: minuscolo, senza accenti, spazi singoli. */
function normalizza_(valore) {
  return String(valore === null || valore === undefined ? '' : valore)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Chiave di confronto per una data (ignora l'orario). */
function chiaveData_(valore) {
  const data = aData_(valore);
  if (!data) { return ''; }
  return Utilities.formatDate(data, fusoOrario_(), 'yyyy-MM-dd');
}

/** Data in formato italiano, per le intestazioni del documento. */
function dataItaliana_(valore) {
  const data = aData_(valore);
  return data ? Utilities.formatDate(data, fusoOrario_(), 'dd/MM/yyyy') : '';
}

/**
 * Converte in Date un valore che può essere Date, stringa gg/mm/aaaa o aaaa-mm-gg.
 *
 * Le date scritte come testo vengono fissate a mezzogiorno: lo script e il foglio
 * possono avere fusi orari diversi e la mezzanotte dell'uno può cadere nel giorno
 * precedente dell'altro. A mezzogiorno il giorno di calendario resta lo stesso.
 */
function aData_(valore) {
  if (valore instanceof Date && !isNaN(valore.getTime())) { return valore; }
  const testo = String(valore === null || valore === undefined ? '' : valore).trim();
  if (!testo) { return null; }

  let m = testo.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    const anno = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
    return new Date(anno, Number(m[2]) - 1, Number(m[1]), 12);
  }
  m = testo.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) { return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12); }
  return null;
}

/** Nome del giorno della settimana ("Lunedì", ...), calcolato nel fuso orario del foglio. */
function giornoSettimana_(data) {
  // "u" restituisce 1 = lunedì ... 7 = domenica; GIORNI parte dalla domenica.
  return GIORNI[Number(Utilities.formatDate(data, fusoOrario_(), 'u')) % 7];
}

function fusoOrario_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Europe/Rome';
}

/** Legge un foglio come array di oggetti {intestazione: valore, _riga: numeroRiga}. */
function leggiFoglio_(ss, nomeFoglio) {
  const foglio = ss.getSheetByName(nomeFoglio);
  if (!foglio) {
    throw new Error('Manca il foglio "' + nomeFoglio + '". Usa Sostituzioni ▸ Prepara il foglio.');
  }
  const ultimaRiga = foglio.getLastRow();
  const ultimaColonna = foglio.getLastColumn();
  if (ultimaRiga < 2) { return []; }

  const valori = foglio.getRange(1, 1, ultimaRiga, ultimaColonna).getValues();
  const intestazioni = valori[0].map(function (v) { return String(v).trim(); });

  return valori.slice(1).map(function (riga, indice) {
    const oggetto = { _riga: indice + 2 };
    intestazioni.forEach(function (nome, i) { if (nome) { oggetto[nome] = riga[i]; } });
    return oggetto;
  }).filter(function (oggetto) {
    return intestazioni.some(function (nome) { return nome && String(oggetto[nome]).trim() !== ''; });
  });
}

/** Numero d'ora come intero, tollerante a formati tipo "3ª" o "3". */
function aOra_(valore) {
  const n = parseInt(String(valore).replace(/[^0-9]/g, ''), 10);
  return isNaN(n) ? null : n;
}

/** Dialogo di errore/informazione, con fallback al log se l'interfaccia non è disponibile. */
function avvisa_(titolo, messaggio) {
  try {
    SpreadsheetApp.getUi().alert(titolo, messaggio, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(titolo + ': ' + messaggio);
  }
}
