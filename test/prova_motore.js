// Banco di prova: simula i fogli Google per verificare il motore di assegnazione.
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const DIR = __dirname;

// Come in Apps Script con il manifest del progetto: lo script gira in Europe/Rome,
// mentre il foglio può avere un fuso orario diverso (vedi scenario 4).
process.env.TZ = 'Europe/Rome';
let fusoFoglio = 'Europe/Rome';

// ---- foglio finto --------------------------------------------------------
function Foglio(valori) { this.v = valori; }
Foglio.prototype.getLastRow = function () { return this.v.length; };
Foglio.prototype.getLastColumn = function () { return this.v[0].length; };
Foglio.prototype.getRange = function (r, c, nr, nc) {
  const self = this;
  nr = nr === undefined ? 1 : nr;
  nc = nc === undefined ? 1 : nc;
  return {
    getValues: function () {
      const out = [];
      for (let i = 0; i < nr; i++) {
        const riga = [];
        for (let j = 0; j < nc; j++) {
          const src = self.v[r - 1 + i] || [];
          riga.push(src[c - 1 + j] === undefined ? '' : src[c - 1 + j]);
        }
        out.push(riga);
      }
      return out;
    },
    setValues: function (blocco) {
      blocco.forEach(function (riga, i) {
        while (self.v.length <= r - 1 + i) { self.v.push([]); }
        riga.forEach(function (val, j) { self.v[r - 1 + i][c - 1 + j] = val; });
      });
    }
  };
};

const fogli = {};
const SpreadsheetAppStub = {
  getActive: function () {
    return {
      getSheetByName: function (nome) { return fogli[nome] || null; },
      getSpreadsheetTimeZone: function () { return fusoFoglio; }
    };
  }
};
// Utilities.formatDate rispetta il fuso orario richiesto, come quello vero.
const UtilitiesStub = {
  formatDate: function (data, tz, formato) {
    const parti = {};
    new Intl.DateTimeFormat('en-GB', {
      timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', weekday: 'short'
    }).formatToParts(data).forEach(function (p) { parti[p.type] = p.value; });
    const u = { Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7 }[parti.weekday];
    return formato.replace(/^u$/, String(u))
      .replace('yyyy', parti.year).replace('MM', parti.month).replace('dd', parti.day);
  }
};

const codice = ['00_Configurazione.gs', '10_Utilita.gs', '30_Motore.gs']
  .map(function (f) { return fs.readFileSync(path.join(DIR, '..', 'src', f), 'utf8'); })
  .join('\n\n');

const api = new Function('SpreadsheetApp', 'Utilities', 'Logger',
  codice + '\nreturn { calcolaSostituzioni_, leggiDisponibilita_, leggiAssenzeDelGiorno_, aData_, giornoSettimana_, dataItaliana_ };'
)(SpreadsheetAppStub, UtilitiesStub, { log: console.log });

// ---- dati di prova -------------------------------------------------------
const LUN = new Date(2026, 8, 14); // lunedì 14 settembre 2026

fogli['Consigli di classe'] = new Foglio([
  ['Classe', 'Docente', 'Disciplina', 'Ruolo'],
  ['3A', 'Rossi Mario', 'Matematica', 'Docente'],
  ['3A', 'Bianchi Anna', 'Italiano', 'Coordinatore'],
  ['3A', 'Verdi Luca', 'Informatica', 'Docente'],
  ['3A', 'Gialli Paolo', 'Sostegno', 'Sostegno'],
  ['4B', 'Rossi Mario', 'Matematica', 'Docente'],
  ['4B', 'Neri Sara', 'Matematica', 'Docente'],
  ['4B', 'Blu Elena', 'Inglese', 'Docente']
]);

fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [LUN, 2, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 3, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 1, '4B', 'Rossi Mario', 'Matematica', 'Permesso', '', '', ''],
  [LUN, 4, '4B', 'Rossi Mario', 'Matematica', 'Permesso', '', 'Prof. Bianchi', ''] // deciso a mano
]);

fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Rossi Mario', 'Lunedì', 2, 'Disposizione', '', '', ''],       // assente in giornata → escluso
  ['Neri Sara', 'Lunedì', 2, 'Disposizione', '', '', ''],
  ['Gialli Paolo', 'Lunedì', 2, 'Disposizione', '', '', ''],
  ['Verdi Luca', 'Lunedì', 3, 'Recupero permesso', '', '', ''],
  ['Neri Sara', 'Lunedì', 3, 'Disposizione', '', '', ''],
  ['Neri Sara', 'Lunedì', 1, 'Disposizione', '', '', ''],
  ['Blu Elena', 'Lunedì', 1, 'Disposizione', '', '', ''],
  ['Verdi Luca', 'Martedì', 1, 'Disposizione', '', '', ''],        // altro giorno → ignorata
  ['Blu Elena', '20/09/2026', 2, 'Disposizione', '', '', ''],      // altra data → ignorata
  ['Gialli Paolo', 'Lunedì', 1, 'Disposizione', '01/10/2026', '', ''] // non ancora valida
]);

fogli['Storico'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Sostituto', 'Criterio', 'Generato il', 'Documento']
]);
fogli['Impostazioni'] = new Foglio([['Impostazione', 'Valore', 'A cosa serve']]);

// ---- esecuzione ----------------------------------------------------------
const esito = api.calcolaSostituzioni_(LUN);
const righe = fogli['Assenti'].v;
const risultato = {};
righe.slice(1).forEach(function (r) { risultato[r[1] + 'ª ' + r[2]] = { sostituto: r[7], criterio: r[8] }; });
console.log(JSON.stringify(risultato, null, 2));
console.log('esito:', JSON.stringify(esito));

// ---- verifiche -----------------------------------------------------------
assert.strictEqual(risultato['1ª 4B'].sostituto, 'Neri Sara', 'stessa classe + stessa disciplina batte solo-classe');
assert.ok(/stessa disciplina/.test(risultato['1ª 4B'].criterio));
assert.strictEqual(risultato['2ª 3A'].sostituto, 'Neri Sara', 'con Rossi assente e Gialli di sostegno resta Neri');
assert.strictEqual(risultato['3ª 3A'].sostituto, 'Verdi Luca', 'consiglio di classe batte docente esterno');
assert.strictEqual(risultato['4ª 4B'].sostituto, 'Prof. Bianchi', 'la riga manuale non viene toccata');
assert.strictEqual(risultato['4ª 4B'].criterio, '', 'la riga manuale resta senza criterio automatico');
assert.strictEqual(esito.manuali, 1);
assert.strictEqual(esito.assegnate, 3);
assert.strictEqual(esito.scoperte.length, 0);

// Un docente non può coprire due classi nella stessa ora.
const perOra = {};
righe.slice(1).forEach(function (r) {
  const chiave = r[1] + '|' + String(r[7]).toLowerCase();
  assert.ok(!perOra[chiave], 'doppia assegnazione nella stessa ora: ' + chiave);
  perOra[chiave] = true;
});

// ---- scenario 2: tetto giornaliero e ore scoperte ------------------------
fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [LUN, 1, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 2, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 3, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 5, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''] // nessuno disponibile alla 5ª
]);
fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Neri Sara', 'Lunedì', 1, 'Disposizione', '', '', ''],
  ['Neri Sara', 'Lunedì', 2, 'Disposizione', '', '', ''],
  ['Neri Sara', 'Lunedì', 3, 'Disposizione', '', '', '']  // la 3ª supera il tetto di 2
]);

const esito2 = api.calcolaSostituzioni_(LUN);
const righe2 = fogli['Assenti'].v.slice(1);
console.log('\nscenario 2:', JSON.stringify(esito2));

assert.strictEqual(righe2[0][7], 'Neri Sara');
assert.strictEqual(righe2[1][7], 'Neri Sara');
assert.strictEqual(righe2[2][7], '', 'oltre il tetto di 2 sostituzioni al giorno non viene proposta');
assert.strictEqual(righe2[3][7], '', 'nessuna disponibilità alla 5ª ora');
assert.strictEqual(esito2.scoperte.length, 2);
assert.ok(/nessun docente disponibile/.test(righe2[3][8]));

// ---- scenario 3: nella stessa ora vince l'affinità, non l'ordine delle righe --
fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [LUN, 6, '4B', 'Blu Elena', 'Inglese', 'Malattia', '', '', ''],     // riga che viene prima
  [LUN, 6, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', '']  // Verdi fa parte del consiglio della 3A
]);
fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Verdi Luca', 'Lunedì', 6, 'Disposizione', '', '', '']
]);

api.calcolaSostituzioni_(LUN);
const righe3 = fogli['Assenti'].v.slice(1);
console.log('\nscenario 3:', JSON.stringify(righe3.map(function (r) { return [r[2], r[7]]; })));
assert.strictEqual(righe3[1][7], 'Verdi Luca', 'il docente va nella classe del suo consiglio');
assert.strictEqual(righe3[0][7], '', 'l\'altra classe resta scoperta');

// ---- scenario 4: foglio con fuso orario diverso da quello dello script -------
// Un file caricato su Drive può avere il fuso GMT: le date nelle celle sono la
// mezzanotte GMT, mentre la data scritta nel popup viene costruita nel fuso dello script.
fusoFoglio = 'GMT';
const GIO_CELLA = new Date(Date.UTC(2026, 8, 17)); // come Sheets restituisce 17/09/2026 in un foglio GMT

fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [GIO_CELLA, 2, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', '']
]);
fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Verdi Luca', 'Giovedì', 2, 'Disposizione', GIO_CELLA, GIO_CELLA, ''] // valida solo quel giorno
]);

const dalPopup = api.aData_('17/09/2026');
assert.strictEqual(api.dataItaliana_(dalPopup), '17/09/2026', 'la data del popup non deve slittare al giorno prima');
assert.strictEqual(api.giornoSettimana_(dalPopup), 'Giovedì');

const esito4 = api.calcolaSostituzioni_(dalPopup);
console.log('\nscenario 4:', JSON.stringify(esito4));
assert.strictEqual(fogli['Assenti'].v[1][7], 'Verdi Luca', 'assenza trovata e disponibilità valida nell\'ultimo giorno');
fusoFoglio = 'Europe/Rome';

// ---- scenario 5: "A pagamento" è l'ultima risorsa ---------------------------
// Alla 2ª ora c'è un docente pagato del consiglio di classe e uno gratuito estraneo:
// deve vincere il gratuito. Alla 5ª ora resta solo il pagato, che viene usato.
fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [LUN, 2, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 5, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', '']
]);
fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Verdi Luca', 'Lunedì', 2, 'A pagamento', '', '', ''],   // consiglio della 3A, ma retribuito
  ['Blu Elena', 'Lunedì', 2, 'Disposizione', '', '', ''],   // estranea alla 3A, ma gratuita
  ['Verdi Luca', 'Lunedì', 5, 'A pagamento', '', '', '']    // unica possibilità
]);

const esito5 = api.calcolaSostituzioni_(LUN);
const righe5 = fogli['Assenti'].v.slice(1);
console.log('\nscenario 5:', JSON.stringify(righe5.map(function (r) { return [r[1], r[7], r[8]]; })));

assert.strictEqual(righe5[0][7], 'Blu Elena', 'il docente gratuito batte quello a pagamento del consiglio di classe');
assert.strictEqual(righe5[1][7], 'Verdi Luca', 'se non resta altro si usa l\'ora a pagamento');
assert.ok(/A pagamento/.test(righe5[1][8]), 'il criterio dichiara che è un\'ora retribuita');
assert.strictEqual(esito5.aPagamento, 1, 'il riepilogo conta le ore retribuite');
assert.strictEqual(esito5.scoperte.length, 0);

// ---- scenario 6: più classi nella stessa ora, una sola copribile a pagamento --
// I gratuiti prendono le classi del proprio consiglio; il pagato copre quella che
// resterebbe scoperta, non una qualsiasi.
fogli['Assenti'] = new Foglio([
  ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  [LUN, 6, '4B', 'Blu Elena', 'Inglese', 'Malattia', '', '', ''],
  [LUN, 6, '3A', 'Bianchi Anna', 'Italiano', 'Malattia', '', '', ''],
  [LUN, 6, '5C', 'Neri Sara', 'Matematica', 'Malattia', '', '', '']  // classe senza docenti liberi
]);
fogli['Disponibilità'] = new Foglio([
  ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  ['Verdi Luca', 'Lunedì', 6, 'Disposizione', '', '', ''],   // consiglio della 3A
  ['Rossi Mario', 'Lunedì', 6, 'Disposizione', '', '', ''],  // consigli di 3A e 4B
  ['Gialli Paolo', 'Lunedì', 6, 'A pagamento', '', '', '']   // sostegno, retribuito
]);

const esito6 = api.calcolaSostituzioni_(LUN);
const perClasse6 = {};
fogli['Assenti'].v.slice(1).forEach(function (r) { perClasse6[r[2]] = r[7]; });
console.log('\nscenario 6:', JSON.stringify(perClasse6), JSON.stringify(esito6));

assert.strictEqual(perClasse6['3A'], 'Verdi Luca');
assert.strictEqual(perClasse6['4B'], 'Rossi Mario');
assert.strictEqual(perClasse6['5C'], 'Gialli Paolo', 'il pagato copre la classe rimasta senza nessuno');
assert.strictEqual(esito6.aPagamento, 1);
assert.strictEqual(esito6.scoperte.length, 0);

console.log('\nTutte le verifiche superate.');
