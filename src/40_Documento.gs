/**
 * Generazione del documento giornaliero delle sostituzioni a partire da un
 * Google Doc usato come modello, e creazione del modello stesso.
 *
 * Segnaposto riconosciuti nel modello (corpo, intestazione e piè di pagina):
 *   {{ISTITUTO}} {{DATA}} {{GIORNO}} {{GENERATO_IL}}
 *   {{TOTALE_ORE}} {{ELENCO_ASSENTI}} {{ORE_SCOPERTE}}
 * e, in una riga di tabella che viene ripetuta per ogni sostituzione:
 *   {{ORA}} {{CLASSE}} {{DOCENTE_ASSENTE}} {{DISCIPLINA}} {{SOSTITUTO}} {{CRITERIO}} {{NOTE}}
 */

/**
 * Crea il documento del giorno leggendo ciò che è scritto nel foglio Assenti
 * (comprese le correzioni manuali) e registra l'esito nello Storico.
 *
 * @param {!Date} data giorno di riferimento.
 * @return {{url: string, nome: string, righe: number}}
 */
function creaDocumento_(data) {
  const ss = SpreadsheetApp.getActive();
  const imp = leggiImpostazioni_(ss);
  const assenze = leggiAssenzeDelGiorno_(ss, data);

  if (!assenze.length) {
    throw new Error('Nessuna assenza registrata per il ' + dataItaliana_(data) + '.');
  }

  const idModello = imp.testo('ID documento template');
  if (!idModello) {
    throw new Error('Manca l\'ID del documento modello.\n\n' +
      'Usa Sostituzioni ▸ Crea documento modello, oppure incolla l\'ID del tuo modello ' +
      'nel foglio "' + FOGLI.IMPOSTAZIONI + '".');
  }

  const nome = (imp.testo('Prefisso nome documento') || 'Sostituzioni') + ' - ' + dataItaliana_(data);
  const copia = copiaModello_(idModello, nome, imp.testo('ID cartella di destinazione'));
  const doc = DocumentApp.openById(copia.getId());

  const scoperte = assenze.filter(function (a) { return !String(a.sostituto).trim(); });
  const generali = {
    ISTITUTO: imp.testo('Nome istituto'),
    DATA: dataItaliana_(data),
    GIORNO: giornoSettimana_(data),
    GENERATO_IL: Utilities.formatDate(new Date(), fusoOrario_(), 'dd/MM/yyyy HH:mm'),
    TOTALE_ORE: String(assenze.length),
    ELENCO_ASSENTI: elencoAssenti_(assenze),
    ORE_SCOPERTE: scoperte.length
      ? scoperte.map(function (a) { return a.ora + 'ª ' + a.classe; }).join(', ')
      : 'nessuna'
  };

  const righe = assenze.map(function (a) {
    return {
      ORA: String(a.ora),
      CLASSE: a.classe,
      DOCENTE_ASSENTE: a.docente,
      DISCIPLINA: a.disciplina,
      SOSTITUTO: String(a.sostituto).trim() || '— da assegnare —',
      CRITERIO: String(a.criterio).replace(MARCA_AUTO, ''),
      NOTE: a.note
    };
  });

  [doc.getBody(), doc.getHeader(), doc.getFooter()].forEach(function (sezione) {
    if (!sezione) { return; }
    Object.keys(generali).forEach(function (chiave) {
      sostituisci_(sezione, chiave, generali[chiave]);
    });
  });

  if (!riempiTabella_(doc.getBody(), righe)) {
    elencoDiRiserva_(doc.getBody(), righe);
  }

  doc.saveAndClose();
  registraStorico_(ss, data, assenze, copia.getUrl());

  return { url: copia.getUrl(), nome: nome, righe: righe.length };
}

/** Copia il modello nella cartella indicata (o nel Drive principale se non impostata). */
function copiaModello_(idModello, nome, idCartella) {
  const modello = DriveApp.getFileById(idModello);
  if (idCartella) {
    return modello.makeCopy(nome, DriveApp.getFolderById(idCartella));
  }
  return modello.makeCopy(nome);
}

/** Sostituisce un segnaposto, con le parentesi graffe protette dal motore delle espressioni regolari. */
function sostituisci_(sezione, chiave, valore) {
  sezione.replaceText('\\{\\{' + chiave + '\\}\\}', String(valore === null || valore === undefined ? '' : valore));
}

/**
 * Trova la riga di tabella che contiene {{ORA}} e la ripete per ogni sostituzione.
 * @return {boolean} true se la tabella modello è stata trovata.
 */
function riempiTabella_(corpo, righe) {
  const tabelle = corpo.getTables();

  for (let t = 0; t < tabelle.length; t++) {
    const tabella = tabelle[t];
    let indiceModello = -1;
    for (let r = 0; r < tabella.getNumRows(); r++) {
      if (tabella.getRow(r).getText().indexOf('{{ORA}}') !== -1) { indiceModello = r; break; }
    }
    if (indiceModello === -1) { continue; }

    const modello = tabella.getRow(indiceModello);
    righe.forEach(function (riga, i) {
      const nuova = modello.copy();
      Object.keys(riga).forEach(function (chiave) { sostituisci_(nuova, chiave, riga[chiave]); });
      tabella.insertTableRow(indiceModello + 1 + i, nuova);
    });
    modello.removeFromParent();
    return true;
  }
  return false;
}

/** Se il modello non contiene la tabella con i segnaposto, scrive comunque l'elenco in fondo. */
function elencoDiRiserva_(corpo, righe) {
  corpo.appendParagraph('Sostituzioni').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  righe.forEach(function (riga) {
    corpo.appendListItem(riga.ORA + 'ª ora · ' + riga.CLASSE + ' · ' + riga.DOCENTE_ASSENTE +
      ' → ' + riga.SOSTITUTO + (riga.NOTE ? ' (' + riga.NOTE + ')' : ''))
      .setGlyphType(DocumentApp.GlyphType.BULLET);
  });
}

/** Riepilogo testuale "Docente (ore 1, 3) - motivo", usato dal segnaposto {{ELENCO_ASSENTI}}. */
function elencoAssenti_(assenze) {
  const perDocente = {};
  const ordine = [];
  assenze.forEach(function (a) {
    if (!perDocente[a.docente]) { perDocente[a.docente] = { ore: [], motivo: a.tipoAssenza }; ordine.push(a.docente); }
    perDocente[a.docente].ore.push(a.ora);
  });
  return ordine.map(function (docente) {
    const voce = perDocente[docente];
    const ore = voce.ore.sort(function (x, y) { return x - y; }).join(', ');
    return docente + ' (ore ' + ore + (voce.motivo ? ' · ' + voce.motivo : '') + ')';
  }).join('; ');
}

/** Aggiunge allo Storico le sostituzioni effettivamente assegnate. */
function registraStorico_(ss, data, assenze, url) {
  const foglio = ss.getSheetByName(FOGLI.STORICO);
  if (!foglio) { return; }

  const chiaveOggi = chiaveData_(data);
  // Elimina eventuali registrazioni precedenti dello stesso giorno: il documento
  // può essere rigenerato più volte e lo storico non deve contare doppio.
  if (foglio.getLastRow() > 1) {
    const date = foglio.getRange(2, 1, foglio.getLastRow() - 1, 1).getValues();
    for (let i = date.length - 1; i >= 0; i--) {
      if (chiaveData_(date[i][0]) === chiaveOggi) { foglio.deleteRow(i + 2); }
    }
  }

  const generatoIl = new Date();
  const righe = assenze
    .filter(function (a) { return String(a.sostituto).trim() !== ''; })
    .map(function (a) {
      return [data, a.ora, a.classe, a.docente, a.sostituto,
        String(a.criterio).replace(MARCA_AUTO, ''), generatoIl, url];
    });

  if (righe.length) {
    foglio.getRange(foglio.getLastRow() + 1, 1, righe.length, INTESTAZIONI.STORICO.length).setValues(righe);
  }
}

/**
 * Crea un documento modello già impaginato e ne salva l'ID nelle impostazioni.
 * Voce di menu "Crea documento modello".
 */
function creaDocumentoModello() {
  const ss = SpreadsheetApp.getActive();
  const imp = leggiImpostazioni_(ss);

  const doc = DocumentApp.create('Modello sostituzioni - ' + imp.testo('Nome istituto'));
  const corpo = doc.getBody();
  corpo.clear();

  corpo.appendParagraph('{{ISTITUTO}}')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  corpo.appendParagraph('Sostituzioni dei docenti assenti')
    .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  corpo.appendParagraph('{{GIORNO}} {{DATA}}')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  corpo.appendParagraph('Docenti assenti: {{ELENCO_ASSENTI}}');
  corpo.appendParagraph('Ore da coprire: {{TOTALE_ORE}} — ore scoperte: {{ORE_SCOPERTE}}');
  corpo.appendParagraph('');

  const tabella = corpo.appendTable([
    ['Ora', 'Classe', 'Docente assente', 'Sostituto', 'Criterio', 'Note'],
    ['{{ORA}}', '{{CLASSE}}', '{{DOCENTE_ASSENTE}}', '{{SOSTITUTO}}', '{{CRITERIO}}', '{{NOTE}}']
  ]);
  const testata = tabella.getRow(0);
  for (let c = 0; c < testata.getNumCells(); c++) {
    testata.getCell(c).setBackgroundColor('#1f3864')
      .editAsText().setForegroundColor('#ffffff').setBold(true);
  }

  corpo.appendParagraph('');
  corpo.appendParagraph('Il Dirigente Scolastico')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  corpo.appendParagraph('_______________________')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const pie = doc.addFooter();
  pie.appendParagraph('Prospetto generato il {{GENERATO_IL}}')
    .setFontSize(8).setForegroundColor('#666666');

  doc.saveAndClose();
  scriviImpostazione_(ss, 'ID documento template', doc.getId());

  // DocumentApp.create salva sempre nella radice del Drive: se è impostata una
  // cartella di destinazione, il modello va messo lì insieme ai documenti generati.
  let posizione = 'nella cartella principale del tuo Drive';
  const idCartella = imp.testo('ID cartella di destinazione');
  if (idCartella) {
    const cartella = DriveApp.getFolderById(idCartella);
    DriveApp.getFileById(doc.getId()).moveTo(cartella);
    posizione = 'nella cartella "' + cartella.getName() + '"';
  }

  avvisa_('Modello creato',
    'Ho creato il documento "' + doc.getName() + '" ' + posizione +
    ' e ne ho salvato l\'ID nelle impostazioni.\n\n' +
    'Puoi aprirlo e personalizzarlo (logo, intestazione, firme): basta non cancellare ' +
    'i segnaposto fra doppie parentesi graffe.\n\n' + doc.getUrl());
}
