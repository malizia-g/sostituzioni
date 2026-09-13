/**
 * Creazione e manutenzione della struttura del foglio di lavoro.
 * È sicura da rieseguire: non cancella i dati già inseriti.
 */

/** Voce di menu "Prepara il foglio". */
function preparaFoglio() {
  const ss = SpreadsheetApp.getActive();

  preparaImpostazioni_(ss);
  const imp = leggiImpostazioni_(ss);
  const oreMax = Math.max(1, imp.numero('Numero di ore giornaliere') || 8);

  preparaConsigli_(ss);
  preparaElenchi_(ss); // prima dei fogli che ne usano le colonne per i menu a tendina
  preparaAssenti_(ss, oreMax);
  preparaDisponibilita_(ss, oreMax);
  preparaStorico_(ss);
  ordinaFogli_(ss);

  avvisa_('Foglio pronto',
    'Struttura creata o aggiornata.\n\n' +
    '1. Compila "Consigli di classe" (una riga per docente/classe/disciplina).\n' +
    '2. Compila "Disponibilità" con le ore a disposizione.\n' +
    '3. Registra le assenze del giorno in "Assenti".\n' +
    '4. Usa il menu Sostituzioni ▸ Calcola e crea il documento.');
}

function foglioOCrea_(ss, nome) {
  let foglio = ss.getSheetByName(nome);
  if (!foglio) { foglio = ss.insertSheet(nome); }
  return foglio;
}

/** Scrive le intestazioni e applica lo stile della riga di testata. */
function intestazioni_(foglio, intestazioni) {
  foglio.getRange(1, 1, 1, intestazioni.length).setValues([intestazioni]);
  foglio.getRange(1, 1, 1, intestazioni.length)
    .setFontWeight('bold')
    .setBackground('#1f3864')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  foglio.setFrozenRows(1);
  foglio.setRowHeight(1, 32);
  if (foglio.getMaxColumns() > intestazioni.length) {
    foglio.deleteColumns(intestazioni.length + 1, foglio.getMaxColumns() - intestazioni.length);
  }
}

function preparaAssenti_(ss, oreMax) {
  const foglio = foglioOCrea_(ss, FOGLI.ASSENTI);
  intestazioni_(foglio, INTESTAZIONI.ASSENTI);

  const righe = Math.max(foglio.getMaxRows() - 1, 1);
  foglio.getRange(2, COL_ASSENTI.DATA, righe, 1).setNumberFormat('dd/mm/yyyy');
  validazioneElenco_(foglio.getRange(2, COL_ASSENTI.ORA, righe, 1), serieOre_(oreMax));
  validazioneDaRange_(foglio, COL_ASSENTI.CLASSE, 'A');     // classi
  validazioneDaRange_(foglio, COL_ASSENTI.DOCENTE, 'B');    // docenti
  validazioneDaRange_(foglio, COL_ASSENTI.DISCIPLINA, 'C'); // discipline
  validazioneElenco_(foglio.getRange(2, COL_ASSENTI.TIPO, righe, 1), TIPI_ASSENZA);

  // Le due colonne di output sono compilate dallo script: sfondo grigio per distinguerle.
  foglio.getRange(2, COL_ASSENTI.SOSTITUTO, righe, 2).setBackground('#f3f3f3');

  // Evidenzia in rosso le ore rimaste scoperte.
  const regola = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($D2<>"",$H2="")')
    .setBackground('#fce8e6')
    .setRanges([foglio.getRange(2, 1, righe, INTESTAZIONI.ASSENTI.length)])
    .build();
  foglio.setConditionalFormatRules([regola]);

  larghezze_(foglio, [95, 55, 90, 200, 180, 130, 220, 200, 260]);
  return foglio;
}

function preparaDisponibilita_(ss, oreMax) {
  const foglio = foglioOCrea_(ss, FOGLI.DISPONIBILITA);
  intestazioni_(foglio, INTESTAZIONI.DISPONIBILITA);

  const righe = Math.max(foglio.getMaxRows() - 1, 1);
  validazioneDaRange_(foglio, 1, 'B'); // docenti
  // Il giorno accetta sia il nome del giorno (ricorrente) sia una data singola:
  // niente validazione rigida, solo una nota esplicativa.
  foglio.getRange(1, 2).setNote('Nome del giorno (es. "Lunedì") per una disponibilità ricorrente,\n' +
    'oppure una data gg/mm/aaaa per una disponibilità valida solo quel giorno.');
  validazioneElenco_(foglio.getRange(2, 3, righe, 1), serieOre_(oreMax));
  validazioneElenco_(foglio.getRange(2, 4, righe, 1), TIPI_DISPONIBILITA);
  foglio.getRange(2, 5, righe, 2).setNumberFormat('dd/mm/yyyy');
  foglio.getRange(1, 5).setNote('Facoltativi: limitano la validità della riga a un periodo (es. un quadrimestre).');

  larghezze_(foglio, [200, 120, 55, 170, 100, 100, 240]);
  return foglio;
}

function preparaConsigli_(ss) {
  const foglio = foglioOCrea_(ss, FOGLI.CONSIGLI);
  intestazioni_(foglio, INTESTAZIONI.CONSIGLI);

  const righe = Math.max(foglio.getMaxRows() - 1, 1);
  validazioneElenco_(foglio.getRange(2, 4, righe, 1), RUOLI);
  foglio.getRange(1, 1).setNote('Una riga per ogni abbinamento docente-classe-disciplina.\n' +
    'È la fonte degli elenchi a tendina degli altri fogli.');

  larghezze_(foglio, [90, 220, 220, 140]);
  return foglio;
}

function preparaStorico_(ss) {
  const foglio = foglioOCrea_(ss, FOGLI.STORICO);
  intestazioni_(foglio, INTESTAZIONI.STORICO);
  foglio.getRange(2, 1, Math.max(foglio.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy');
  foglio.getRange(1, 1).setNote('Compilato automaticamente a ogni documento generato.\n' +
    'Serve a distribuire equamente le sostituzioni: non modificarlo a mano.');
  larghezze_(foglio, [95, 55, 90, 200, 200, 260, 140, 260]);
  return foglio;
}

/**
 * Foglio nascosto con gli elenchi univoci ricavati dai consigli di classe:
 * A = classi, B = docenti, C = discipline. Alimenta i menu a tendina.
 */
function preparaElenchi_(ss) {
  const foglio = foglioOCrea_(ss, FOGLI.ELENCHI);
  foglio.clear();
  foglio.getRange(1, 1, 1, 3).setValues([['Classi', 'Docenti', 'Discipline']]).setFontWeight('bold');

  // IFERROR: finché i consigli di classe sono vuoti la formula non deve mostrare errori.
  const consigli = "'" + FOGLI.CONSIGLI + "'";
  ['A', 'B', 'C'].forEach(function (colonna, indice) {
    foglio.getRange(2, indice + 1).setFormula(
      '=IFERROR(SORT(UNIQUE(FILTER(' + consigli + '!' + colonna + '2:' + colonna + ', ' +
      consigli + '!' + colonna + '2:' + colonna + '<>""))), "")');
  });

  foglio.hideSheet();
  return foglio;
}

function preparaImpostazioni_(ss) {
  const foglio = foglioOCrea_(ss, FOGLI.IMPOSTAZIONI);
  const esistenti = {};
  if (foglio.getLastRow() > 1) {
    foglio.getRange(2, 1, foglio.getLastRow() - 1, 2).getValues().forEach(function (riga) {
      const chiave = String(riga[0]).trim();
      if (chiave) { esistenti[chiave] = riga[1]; }
    });
  }

  foglio.clear();
  intestazioni_(foglio, ['Impostazione', 'Valore', 'A cosa serve']);

  const righe = IMPOSTAZIONI_DEFAULT.map(function (riga) {
    const valore = esistenti[riga[0]] !== undefined && String(esistenti[riga[0]]).trim() !== ''
      ? esistenti[riga[0]] : riga[1];
    return [riga[0], valore, riga[2]];
  });
  foglio.getRange(2, 1, righe.length, 3).setValues(righe);
  foglio.getRange(2, 2, righe.length, 1).setBackground('#fff2cc').setHorizontalAlignment('left');
  foglio.getRange(2, 3, righe.length, 1).setFontColor('#666666').setWrap(true);
  larghezze_(foglio, [300, 260, 520]);
  return foglio;
}

/** Mette i fogli nell'ordine di lavoro abituale. */
function ordinaFogli_(ss) {
  [FOGLI.ASSENTI, FOGLI.DISPONIBILITA, FOGLI.CONSIGLI, FOGLI.STORICO, FOGLI.IMPOSTAZIONI]
    .forEach(function (nome, indice) {
      const foglio = ss.getSheetByName(nome);
      if (foglio) {
        ss.setActiveSheet(foglio);
        ss.moveActiveSheet(indice + 1);
      }
    });
  ss.setActiveSheet(ss.getSheetByName(FOGLI.ASSENTI));
}

function serieOre_(oreMax) {
  const ore = [];
  for (let i = 1; i <= oreMax; i++) { ore.push(String(i)); }
  return ore;
}

function validazioneElenco_(range, valori) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(valori, true).setAllowInvalid(true).build());
}

/** Menu a tendina alimentato da una colonna del foglio Elenchi. */
function validazioneDaRange_(foglio, colonna, colonnaElenchi) {
  const righe = Math.max(foglio.getMaxRows() - 1, 1);
  const origine = foglio.getParent().getSheetByName(FOGLI.ELENCHI);
  if (!origine) { return; }
  foglio.getRange(2, colonna, righe, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(origine.getRange(colonnaElenchi + '2:' + colonnaElenchi), true)
      .setAllowInvalid(true)
      .build());
}

function larghezze_(foglio, larghezze) {
  larghezze.forEach(function (larghezza, i) { foglio.setColumnWidth(i + 1, larghezza); });
}
