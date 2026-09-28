/**
 * Gestione sostituzioni docenti - Configurazione.
 *
 * Questo file contiene solo costanti e lettura delle impostazioni.
 * Non va modificato per l'uso quotidiano: i parametri si cambiano dal
 * foglio "Impostazioni".
 */

/** Nomi dei fogli usati dall'applicazione. */
const FOGLI = {
  ASSENTI: 'Assenti',
  DISPONIBILITA: 'Disponibilità',
  CONSIGLI: 'Consigli di classe',
  STORICO: 'Storico',
  IMPOSTAZIONI: 'Impostazioni',
  ELENCHI: 'Elenchi'
};

/** Intestazioni attese, in ordine, per ogni foglio dati. */
const INTESTAZIONI = {
  ASSENTI: ['Data', 'Ora', 'Classe', 'Docente assente', 'Disciplina', 'Tipo assenza', 'Note', 'Sostituto', 'Criterio'],
  DISPONIBILITA: ['Docente', 'Giorno', 'Ora', 'Tipo', 'Valida dal', 'Valida al', 'Note'],
  CONSIGLI: ['Classe', 'Docente', 'Disciplina', 'Ruolo'],
  STORICO: ['Data', 'Ora', 'Classe', 'Docente assente', 'Sostituto', 'Criterio', 'Generato il', 'Documento']
};

/** Indici di colonna (1-based) del foglio Assenti, usati per la riscrittura. */
const COL_ASSENTI = {
  DATA: 1, ORA: 2, CLASSE: 3, DOCENTE: 4, DISCIPLINA: 5,
  TIPO: 6, NOTE: 7, SOSTITUTO: 8, CRITERIO: 9
};

const GIORNI = ['Domenica', 'Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato'];
const TIPI_DISPONIBILITA = ['Disposizione', 'Recupero permesso', 'Completamento orario',
  'Potenziamento', 'Compresenza', 'A pagamento'];

/** Ore retribuite: ultima risorsa, da usare solo se non c'è nessun altro modo di coprire l'ora. */
const TIPO_A_PAGAMENTO = 'A pagamento';

const TIPI_ASSENZA = ['Malattia', 'Permesso', 'Ferie', 'Formazione', 'Uscita didattica', 'Sciopero', 'Altro'];
const RUOLI = ['Docente', 'Coordinatore', 'ITP', 'Sostegno'];

/** Prefisso che marca un'assegnazione generata dallo script (le altre sono manuali e non vengono toccate). */
const MARCA_AUTO = 'Auto · ';

/**
 * Impostazioni predefinite: [chiave, valore, descrizione].
 * Vengono scritte nel foglio "Impostazioni" da preparaFoglio() e lette a ogni esecuzione.
 */
const IMPOSTAZIONI_DEFAULT = [
  ['Nome istituto', 'Istituto Tecnico Superiore', 'Compare nell\'intestazione del documento generato.'],
  ['ID documento template', '', 'ID del Google Doc modello. Lascia vuoto e usa "Crea documento modello" dal menu.'],
  ['ID cartella di destinazione', '', 'ID della cartella Drive dove salvare i documenti. Vuoto = cartella principale del Drive.'],
  ['Prefisso nome documento', 'Sostituzioni', 'Il documento si chiamerà "Prefisso - gg/mm/aaaa".'],
  ['Numero di ore giornaliere', '8', 'Quante ore ha la giornata scolastica (serve per i menu a tendina).'],
  ['Max sostituzioni per docente al giorno', '2', 'Oltre questa soglia il docente non viene più proposto.'],
  ['Peso stessa classe', '100', 'Punti per un docente che insegna nella classe scoperta.'],
  ['Peso stessa disciplina', '60', 'Punti per un docente che insegna la stessa disciplina dell\'ora scoperta.'],
  ['Priorità tipi disponibilità', 'Recupero permesso, Disposizione, Completamento orario, Potenziamento, Compresenza, A pagamento',
    'Ordine di preferenza: il primo tipo vale di più. Modifica l\'ordine per cambiare la precedenza.'],
  ['Penalità docente di sostegno', '500', 'Sottratta ai docenti con ruolo "Sostegno". Con un valore alto ' +
    'vengono impiegati solo se non c\'è nessun altro; metti 0 per trattarli come gli altri.'],
  ['Penalità disponibilità a pagamento', '1000', 'Sottratta alle ore di tipo "A pagamento". Più alta di quella ' +
    'del sostegno: sono l\'ultima risorsa e vengono proposte solo se l\'ora resterebbe scoperta.'],
  ['Penalità per ogni sostituzione già assegnata oggi', '25', 'Distribuisce il carico nella giornata.'],
  ['Penalità per ogni sostituzione recente', '3', 'Distribuisce il carico nel periodo (vedi finestra sotto).'],
  ['Giorni di storico considerati', '30', 'Ampiezza della finestra per la penalità "sostituzione recente".'],
  ['Escludi i docenti assenti per tutta la giornata', 'SI', 'SI = chi risulta assente anche a una sola ora non è proponibile in nessun\'ora.']
];

/**
 * Legge il foglio Impostazioni e restituisce un oggetto con accessori tipizzati.
 * @return {{testo: function(string): string, numero: function(string): number,
 *           booleano: function(string): boolean, elenco: function(string): !Array<string>}}
 */
function leggiImpostazioni_(ss) {
  const foglio = ss.getSheetByName(FOGLI.IMPOSTAZIONI);
  const valori = {};
  IMPOSTAZIONI_DEFAULT.forEach(function (riga) { valori[riga[0]] = riga[1]; });

  if (foglio && foglio.getLastRow() > 1) {
    foglio.getRange(2, 1, foglio.getLastRow() - 1, 2).getValues().forEach(function (riga) {
      const chiave = String(riga[0]).trim();
      if (chiave) { valori[chiave] = riga[1]; }
    });
  }

  return {
    testo: function (chiave) { return String(valori[chiave] === undefined ? '' : valori[chiave]).trim(); },
    numero: function (chiave) {
      const n = Number(String(valori[chiave]).replace(',', '.'));
      return isNaN(n) ? 0 : n;
    },
    booleano: function (chiave) {
      const v = String(valori[chiave]).trim().toLowerCase();
      return v === 'si' || v === 'sì' || v === 'true' || v === 'vero' || v === 'x';
    },
    elenco: function (chiave) {
      return String(valori[chiave]).split(',').map(function (v) { return v.trim(); })
        .filter(function (v) { return v !== ''; });
    }
  };
}

/** Scrive una singola impostazione (usata da "Crea documento modello"). */
function scriviImpostazione_(ss, chiave, valore) {
  const foglio = ss.getSheetByName(FOGLI.IMPOSTAZIONI);
  if (!foglio) { return; }
  const chiavi = foglio.getRange(2, 1, Math.max(foglio.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < chiavi.length; i++) {
    if (String(chiavi[i][0]).trim() === chiave) {
      foglio.getRange(i + 2, 2).setValue(valore);
      return;
    }
  }
  foglio.appendRow([chiave, valore, '']);
}
