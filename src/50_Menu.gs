/**
 * Punti di ingresso: menu del foglio, bottone e dialoghi.
 *
 * Funzione da collegare al bottone disegnato nel foglio: calcolaECreaDocumento
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 Sostituzioni')
    .addItem('▶️  Calcola e crea il documento', 'calcolaECreaDocumento')
    .addSeparator()
    .addItem('🔎  Solo calcolo (proposte nel foglio)', 'calcolaProposte')
    .addItem('📄  Solo documento (da ciò che c\'è nel foglio)', 'generaSoloDocumento')
    .addItem('🧹  Azzera le proposte del giorno', 'azzeraProposte')
    .addSeparator()
    .addItem('🧩  Prepara il foglio', 'preparaFoglio')
    .addItem('🗎  Crea documento modello', 'creaDocumentoModello')
    .addItem('❓  Guida rapida', 'mostraGuida')
    .addToUi();
}

/** Azione del bottone: calcola le sostituzioni e genera subito il documento. */
function calcolaECreaDocumento() {
  const lucchetto = LockService.getDocumentLock();
  if (!lucchetto.tryLock(30000)) {
    avvisa_('Attendi', 'Un\'altra elaborazione è in corso. Riprova fra qualche secondo.');
    return;
  }
  try {
    const data = chiediData_();
    if (!data) { return; }

    const esito = calcolaSostituzioni_(data);
    const documento = creaDocumento_(data);
    mostraEsito_(data, esito, documento);
  } catch (errore) {
    avvisa_('Impossibile completare', errore.message);
  } finally {
    lucchetto.releaseLock();
  }
}

/** Calcola le proposte e le scrive nel foglio, senza generare il documento. */
function calcolaProposte() {
  try {
    const data = chiediData_();
    if (!data) { return; }
    const esito = calcolaSostituzioni_(data);
    mostraEsito_(data, esito, null);
  } catch (errore) {
    avvisa_('Impossibile completare', errore.message);
  }
}

/** Genera il documento da quanto già presente nel foglio (comprese le modifiche manuali). */
function generaSoloDocumento() {
  try {
    const data = chiediData_();
    if (!data) { return; }
    const documento = creaDocumento_(data);
    mostraEsito_(data, null, documento);
  } catch (errore) {
    avvisa_('Impossibile completare', errore.message);
  }
}

/** Svuota le colonne Sostituto e Criterio per la data indicata, comprese le righe manuali. */
function azzeraProposte() {
  try {
    const data = chiediData_();
    if (!data) { return; }
    const ss = SpreadsheetApp.getActive();
    const foglio = ss.getSheetByName(FOGLI.ASSENTI);
    const assenze = leggiAssenzeDelGiorno_(ss, data);
    assenze.forEach(function (assenza) {
      foglio.getRange(assenza._riga, COL_ASSENTI.SOSTITUTO, 1, 2).clearContent();
    });
    avvisa_('Fatto', 'Azzerate ' + assenze.length + ' righe del ' + dataItaliana_(data) + '.');
  } catch (errore) {
    avvisa_('Impossibile completare', errore.message);
  }
}

/** Chiede la data di riferimento; vuoto = oggi. Restituisce null se l'utente annulla. */
function chiediData_() {
  const ui = SpreadsheetApp.getUi();
  const oggi = new Date();
  const risposta = ui.prompt(
    'Giorno delle sostituzioni',
    'Data in formato gg/mm/aaaa.\nLascia vuoto per usare oggi (' + dataItaliana_(oggi) + ').',
    ui.ButtonSet.OK_CANCEL);

  if (risposta.getSelectedButton() !== ui.Button.OK) { return null; }

  const testo = risposta.getResponseText().trim();
  if (!testo) { return oggi; }

  const data = aData_(testo);
  if (!data) {
    avvisa_('Data non valida', '"' + testo + '" non è una data riconoscibile. Usa il formato gg/mm/aaaa.');
    return null;
  }
  return data;
}

/** Riepilogo finale con link cliccabile al documento. */
function mostraEsito_(data, esito, documento) {
  let html = '<div style="font-family:Google Sans,Roboto,Arial,sans-serif;font-size:13px;line-height:1.6">';
  html += '<p><b>' + giornoSettimana_(data) + ' ' + dataItaliana_(data) + '</b></p>';

  if (esito) {
    html += '<p>Ore da coprire: <b>' + esito.totale + '</b><br>' +
      'Assegnate automaticamente: <b>' + esito.assegnate + '</b><br>' +
      'Confermate a mano: <b>' + esito.manuali + '</b></p>';
    if (esito.scoperte.length) {
      html += '<p style="color:#b00020">Restano scoperte:<br>• ' +
        esito.scoperte.join('<br>• ') + '</p>';
    }
  }

  if (documento) {
    html += '<p>Documento creato: <a href="' + documento.url + '" target="_blank">' +
      documento.nome + '</a></p>';
  } else {
    html += '<p style="color:#5f6368">Controlla le proposte nel foglio, correggile se serve, ' +
      'poi usa <i>Solo documento</i>.</p>';
  }
  html += '</div>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(430).setHeight(300), 'Sostituzioni');
}

function mostraGuida() {
  const html =
    '<div style="font-family:Google Sans,Roboto,Arial,sans-serif;font-size:13px;line-height:1.6">' +
    '<p><b>Uso quotidiano</b></p>' +
    '<ol>' +
    '<li>Nel foglio <b>Assenti</b> inserisci una riga per ogni <i>ora</i> da coprire: ' +
    'data, ora, classe, docente assente, disciplina.</li>' +
    '<li>Premi il bottone <b>Genera sostituzioni</b> (o il menu ▶️).</li>' +
    '<li>Le colonne <b>Sostituto</b> e <b>Criterio</b> si compilano da sole e viene creato il documento.</li>' +
    '</ol>' +
    '<p><b>Correzioni manuali</b><br>Se scrivi un nome nella colonna Sostituto e cancelli il ' +
    'contenuto di Criterio, quella riga viene considerata decisa da te e il calcolo non la tocca più.</p>' +
    '<p><b>Disponibilità</b><br>Nella colonna <i>Giorno</i> scrivi il giorno della settimana ' +
    '(disponibilità ricorrente da orario) oppure una data precisa (disponibilità straordinaria).</p>' +
    '<p><b>Come sceglie il sostituto</b><br>Nell\'ordine: docente del consiglio di classe che insegna ' +
    'la stessa disciplina, docente del consiglio di classe, docente della stessa disciplina, ' +
    'altro docente a disposizione. A parità pesa il tipo di disponibilità e quante sostituzioni ' +
    'ha già fatto. Tutti i parametri sono nel foglio <b>Impostazioni</b>.</p>' +
    '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(420), 'Guida rapida');
}
