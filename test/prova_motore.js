// Banco di prova: simula i fogli Google per verificare il motore di assegnazione.
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const DIR = __dirname;

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
      getSpreadsheetTimeZone: function () { return 'Europe/Rome'; }
    };
  }
};
const UtilitiesStub = {
  formatDate: function (data, tz, formato) {
    const p = function (n) { return String(n).padStart(2, '0'); };
    if (formato === 'yyyy-MM-dd') { return data.getFullYear() + '-' + p(data.getMonth() + 1) + '-' + p(data.getDate()); }
    if (formato === 'dd/MM/yyyy') { return p(data.getDate()) + '/' + p(data.getMonth() + 1) + '/' + data.getFullYear(); }
    return String(data);
  }
};

const codice = ['00_Configurazione.gs', '10_Utilita.gs', '30_Motore.gs']
  .map(function (f) { return fs.readFileSync(path.join(DIR, '..', 'src', f), 'utf8'); })
  .join('\n\n');

const api = new Function('SpreadsheetApp', 'Utilities', 'Logger',
  codice + '\nreturn { calcolaSostituzioni_, leggiDisponibilita_, leggiAssenzeDelGiorno_ };'
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

console.log('\nTutte le verifiche superate.');
