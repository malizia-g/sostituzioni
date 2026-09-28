/**
 * Motore di assegnazione delle sostituzioni.
 *
 * Criterio: fra i docenti che risultano disponibili in quell'ora si preferisce,
 * nell'ordine, chi fa parte del consiglio di quella classe e insegna la stessa
 * disciplina, chi fa parte del consiglio di classe, chi insegna la stessa
 * disciplina altrove, infine chiunque sia a disposizione. A parità di profilo
 * conta il tipo di disponibilità e il carico di sostituzioni già sostenuto.
 */

/**
 * Calcola le sostituzioni per una data e le scrive nel foglio Assenti.
 * Le assegnazioni inserite a mano (criterio diverso da "Auto · ...") non vengono toccate.
 *
 * @param {!Date} data giorno da coprire.
 * @return {{assegnate: number, manuali: number, aPagamento: number,
 *           scoperte: !Array<string>, totale: number}}
 */
function calcolaSostituzioni_(data) {
  const ss = SpreadsheetApp.getActive();
  const imp = leggiImpostazioni_(ss);
  const assenze = leggiAssenzeDelGiorno_(ss, data);

  if (!assenze.length) {
    throw new Error('Nessuna assenza registrata per il ' + dataItaliana_(data) +
      '.\nControlla la colonna Data del foglio "' + FOGLI.ASSENTI + '".');
  }

  const contesto = costruisciContesto_(ss, imp, data, assenze);
  const disponibilita = leggiDisponibilita_(ss, data);

  // Le assegnazioni manuali valgono come già fatte: occupano il docente e ne aumentano il carico.
  assenze.forEach(function (assenza) {
    if (assenza.bloccata) {
      registraImpegno_(contesto, normalizza_(assenza.sostituto), assenza.ora);
    }
  });

  const daCoprire = assenze.filter(function (a) { return !a.bloccata; });
  const scoperte = [];
  let aPagamento = 0;

  const perOra = {};
  daCoprire.forEach(function (assenza) {
    if (!assenza.ora || !assenza.classe) {
      assenza.sostituto = '';
      assenza.criterio = MARCA_AUTO + 'riga incompleta: manca l\'ora o la classe';
      scoperte.push('riga ' + assenza._riga + ' (' + assenza.docente + '): ora o classe mancante');
      return;
    }
    if (!perOra[assenza.ora]) { perOra[assenza.ora] = []; }
    perOra[assenza.ora].push(assenza);
  });

  // Si parte dalle ore con meno candidati per classe scoperta: il tetto giornaliero
  // lega le ore fra loro e i docenti scarsi vanno spesi dove servono di più.
  const ore = Object.keys(perOra).map(Number).sort(function (a, b) {
    const ra = (disponibilita[a] || []).length / perOra[a].length;
    const rb = (disponibilita[b] || []).length / perOra[b].length;
    return ra !== rb ? ra - rb : a - b;
  });

  ore.forEach(function (ora) {
    const aperte = perOra[ora].slice();

    // Nella stessa ora si assegna prima la coppia classe-docente con l'affinità più alta:
    // così un docente del consiglio di classe non finisce in una classe che non conosce
    // solo perché quella riga compariva prima nel foglio.
    while (aperte.length) {
      let migliore = null;
      let indice = -1;
      aperte.forEach(function (assenza, i) {
        const scelta = scegliSostituto_(assenza, disponibilita[ora] || [], contesto);
        if (scelta && (!migliore || scelta.punti > migliore.punti)) {
          migliore = scelta;
          indice = i;
        }
      });
      if (!migliore) { break; }

      const assenza = aperte.splice(indice, 1)[0];
      assenza.sostituto = migliore.nome;
      assenza.criterio = MARCA_AUTO + migliore.criterio + ' (' + migliore.tipo + ')';
      if (normalizza_(migliore.tipo) === normalizza_(TIPO_A_PAGAMENTO)) { aPagamento++; }
      registraImpegno_(contesto, migliore.chiave, ora);
    }

    aperte.forEach(function (assenza) {
      assenza.sostituto = '';
      assenza.criterio = MARCA_AUTO + 'nessun docente disponibile';
      scoperte.push(ora + 'ª ora, ' + assenza.classe + ' (' + assenza.docente + ')');
    });
  });

  scriviAssegnazioni_(ss, assenze);

  return {
    totale: assenze.length,
    manuali: assenze.filter(function (a) { return a.bloccata; }).length,
    assegnate: daCoprire.filter(function (a) { return a.sostituto; }).length,
    aPagamento: aPagamento,
    scoperte: scoperte
  };
}

/** Sceglie il miglior candidato per un'ora scoperta, o null se non ce ne sono. */
function scegliSostituto_(assenza, candidati, contesto) {
  const ammessi = [];

  candidati.forEach(function (candidato) {
    if (candidato.chiave === normalizza_(assenza.docente)) { return; }        // è lui l'assente
    if (contesto.nonDisponibili[candidato.chiave]) { return; }                // assente in giornata
    if (contesto.assentiPerOra[assenza.ora] &&
        contesto.assentiPerOra[assenza.ora][candidato.chiave]) { return; }    // assente in quest'ora
    if (contesto.occupati[assenza.ora] &&
        contesto.occupati[assenza.ora][candidato.chiave]) { return; }         // già impegnato in quest'ora
    if ((contesto.assegnateOggi[candidato.chiave] || 0) >= contesto.maxAlGiorno) { return; }

    const valutazione = valutaCandidato_(candidato, assenza, contesto);
    ammessi.push({
      nome: candidato.nome,
      chiave: candidato.chiave,
      tipo: candidato.tipo,
      punti: valutazione.punti,
      criterio: valutazione.criterio
    });
  });

  if (!ammessi.length) { return null; }

  // Punteggio decrescente; a parità, ordine alfabetico per avere un esito riproducibile.
  ammessi.sort(function (a, b) {
    return b.punti !== a.punti ? b.punti - a.punti : a.chiave.localeCompare(b.chiave);
  });
  return ammessi[0];
}

/** Assegna un punteggio al candidato per una specifica ora scoperta. */
function valutaCandidato_(candidato, assenza, contesto) {
  const classi = contesto.classiPerDocente[candidato.chiave] || {};
  const discipline = contesto.disciplinePerDocente[candidato.chiave] || {};

  const stessaClasse = !!classi[normalizza_(assenza.classe)];
  const stessaDisciplina = !!(assenza.disciplinaChiave && discipline[assenza.disciplinaChiave]);

  let punti = 0;
  if (stessaClasse) { punti += contesto.pesoClasse; }
  if (stessaDisciplina) { punti += contesto.pesoDisciplina; }
  punti += contesto.bonusTipo[normalizza_(candidato.tipo)] || 0;

  if (normalizza_(candidato.tipo) === normalizza_(TIPO_A_PAGAMENTO)) {
    punti -= contesto.penalitaPagamento;
  }
  if (normalizza_(contesto.ruoloPerDocente[candidato.chiave]) === 'sostegno') {
    punti -= contesto.penalitaSostegno;
  }
  punti -= contesto.penalitaGiorno * (contesto.assegnateOggi[candidato.chiave] || 0);
  punti -= contesto.penalitaStorico * (contesto.caricoStorico[candidato.chiave] || 0);

  let criterio;
  if (stessaClasse && stessaDisciplina) { criterio = 'consiglio di classe, stessa disciplina'; }
  else if (stessaClasse) { criterio = 'consiglio di classe'; }
  else if (stessaDisciplina) { criterio = 'stessa disciplina'; }
  else { criterio = 'docente a disposizione'; }

  return { punti: punti, criterio: criterio };
}

/** Segna il docente come impegnato in quell'ora e ne aggiorna il carico giornaliero. */
function registraImpegno_(contesto, chiaveDocente, ora) {
  if (!chiaveDocente) { return; }
  if (!contesto.occupati[ora]) { contesto.occupati[ora] = {}; }
  contesto.occupati[ora][chiaveDocente] = true;
  contesto.assegnateOggi[chiaveDocente] = (contesto.assegnateOggi[chiaveDocente] || 0) + 1;
}

/** Raccoglie in un solo oggetto indici, pesi e stato di avanzamento. */
function costruisciContesto_(ss, imp, data, assenze) {
  const consigli = indicizzaConsigli_(ss);

  const bonusTipo = {};
  const priorita = imp.elenco('Priorità tipi disponibilità');
  priorita.forEach(function (tipo, indice) {
    bonusTipo[normalizza_(tipo)] = (priorita.length - indice) * 5;
  });

  const nonDisponibili = {};
  const assentiPerOra = {};
  const escludiTuttoIlGiorno = imp.booleano('Escludi i docenti assenti per tutta la giornata');
  assenze.forEach(function (assenza) {
    const chiave = normalizza_(assenza.docente);
    if (!chiave) { return; }
    if (escludiTuttoIlGiorno) { nonDisponibili[chiave] = true; }
    if (!assentiPerOra[assenza.ora]) { assentiPerOra[assenza.ora] = {}; }
    assentiPerOra[assenza.ora][chiave] = true;
  });

  return {
    classiPerDocente: consigli.classiPerDocente,
    disciplinePerDocente: consigli.disciplinePerDocente,
    ruoloPerDocente: consigli.ruoloPerDocente,
    caricoStorico: caricoStorico_(ss, data, imp.numero('Giorni di storico considerati')),
    assegnateOggi: {},
    occupati: {},
    nonDisponibili: nonDisponibili,
    assentiPerOra: assentiPerOra,
    bonusTipo: bonusTipo,
    pesoClasse: imp.numero('Peso stessa classe'),
    pesoDisciplina: imp.numero('Peso stessa disciplina'),
    penalitaSostegno: imp.numero('Penalità docente di sostegno'),
    penalitaPagamento: imp.numero('Penalità disponibilità a pagamento'),
    penalitaGiorno: imp.numero('Penalità per ogni sostituzione già assegnata oggi'),
    penalitaStorico: imp.numero('Penalità per ogni sostituzione recente'),
    maxAlGiorno: Math.max(1, imp.numero('Max sostituzioni per docente al giorno') || 2)
  };
}

/** Indicizza i consigli di classe per docente. */
function indicizzaConsigli_(ss) {
  const righe = leggiFoglio_(ss, FOGLI.CONSIGLI);
  const classiPerDocente = {};
  const disciplinePerDocente = {};
  const ruoloPerDocente = {};

  righe.forEach(function (riga) {
    const chiave = normalizza_(riga['Docente']);
    if (!chiave) { return; }
    if (!classiPerDocente[chiave]) { classiPerDocente[chiave] = {}; }
    if (!disciplinePerDocente[chiave]) { disciplinePerDocente[chiave] = {}; }

    const classe = normalizza_(riga['Classe']);
    if (classe) { classiPerDocente[chiave][classe] = true; }
    const disciplina = normalizza_(riga['Disciplina']);
    if (disciplina) { disciplinePerDocente[chiave][disciplina] = true; }
    const ruolo = String(riga['Ruolo'] || '').trim();
    if (ruolo) { ruoloPerDocente[chiave] = ruolo; }
  });

  if (!righe.length) {
    throw new Error('Il foglio "' + FOGLI.CONSIGLI + '" è vuoto: senza consigli di classe ' +
      'non è possibile scegliere i sostituti per affinità.');
  }
  return {
    classiPerDocente: classiPerDocente,
    disciplinePerDocente: disciplinePerDocente,
    ruoloPerDocente: ruoloPerDocente
  };
}

/** Assenze registrate per la data indicata, in ordine di ora. */
function leggiAssenzeDelGiorno_(ss, data) {
  const chiave = chiaveData_(data);
  return leggiFoglio_(ss, FOGLI.ASSENTI).filter(function (riga) {
    return chiaveData_(riga['Data']) === chiave && String(riga['Docente assente']).trim() !== '';
  }).map(function (riga) {
    const sostituto = String(riga['Sostituto'] || '').trim();
    const criterio = String(riga['Criterio'] || '').trim();
    return {
      _riga: riga._riga,
      ora: aOra_(riga['Ora']) || 0,
      classe: String(riga['Classe'] || '').trim(),
      docente: String(riga['Docente assente']).trim(),
      disciplina: String(riga['Disciplina'] || '').trim(),
      disciplinaChiave: normalizza_(riga['Disciplina']),
      tipoAssenza: String(riga['Tipo assenza'] || '').trim(),
      note: String(riga['Note'] || '').trim(),
      sostituto: sostituto,
      criterio: criterio,
      // Una riga compilata a mano non viene sovrascritta dal calcolo automatico.
      bloccata: sostituto !== '' && criterio.indexOf(MARCA_AUTO) !== 0
    };
  }).sort(function (a, b) { return a.ora - b.ora; });
}

/** Disponibilità valide per la data, raggruppate per ora. */
function leggiDisponibilita_(ss, data) {
  const righe = leggiFoglio_(ss, FOGLI.DISPONIBILITA);
  const giorno = normalizza_(giornoSettimana_(data));
  const chiave = chiaveData_(data);
  const perOra = {};
  const gia = {};

  righe.forEach(function (riga) {
    const nome = String(riga['Docente'] || '').trim();
    const ora = aOra_(riga['Ora']);
    if (!nome || !ora) { return; }

    // "Giorno" accetta un giorno della settimana (ricorrente) o una data singola.
    const dataSpecifica = aData_(riga['Giorno']);
    const valeOggi = dataSpecifica
      ? chiaveData_(dataSpecifica) === chiave
      : normalizza_(riga['Giorno']) === giorno;
    if (!valeOggi) { return; }

    // Confronto sui giorni di calendario (aaaa-mm-gg), non sugli istanti: gli estremi sono inclusi.
    const dal = chiaveData_(riga['Valida dal']);
    if (dal && chiave < dal) { return; }
    const al = chiaveData_(riga['Valida al']);
    if (al && chiave > al) { return; }

    const chiaveDocente = normalizza_(nome);
    const chiaveRiga = ora + '|' + chiaveDocente;
    if (gia[chiaveRiga]) { return; }  // stesso docente già presente per quell'ora
    gia[chiaveRiga] = true;

    if (!perOra[ora]) { perOra[ora] = []; }
    perOra[ora].push({
      nome: nome,
      chiave: chiaveDocente,
      tipo: String(riga['Tipo'] || '').trim() || 'Disposizione',
      note: String(riga['Note'] || '').trim()
    });
  });

  return perOra;
}

/** Quante sostituzioni ha già fatto ogni docente nei giorni precedenti. */
function caricoStorico_(ss, data, giorni) {
  const carico = {};
  const foglio = ss.getSheetByName(FOGLI.STORICO);
  if (!foglio || foglio.getLastRow() < 2 || !giorni) { return carico; }

  const limite = new Date(data.getTime() - giorni * 24 * 60 * 60 * 1000);
  const chiaveOggi = chiaveData_(data);

  foglio.getRange(2, 1, foglio.getLastRow() - 1, INTESTAZIONI.STORICO.length).getValues()
    .forEach(function (riga) {
      const quando = aData_(riga[0]);
      const sostituto = normalizza_(riga[4]);
      if (!quando || !sostituto) { return; }
      if (quando < limite || chiaveData_(quando) === chiaveOggi) { return; }
      carico[sostituto] = (carico[sostituto] || 0) + 1;
    });

  return carico;
}

/** Riscrive in blocco le colonne Sostituto e Criterio del foglio Assenti. */
function scriviAssegnazioni_(ss, assenze) {
  const foglio = ss.getSheetByName(FOGLI.ASSENTI);
  const righe = assenze.map(function (a) { return a._riga; });
  const prima = Math.min.apply(null, righe);
  const ultima = Math.max.apply(null, righe);

  const blocco = foglio.getRange(prima, COL_ASSENTI.SOSTITUTO, ultima - prima + 1, 2).getValues();
  assenze.forEach(function (assenza) {
    if (assenza.bloccata) { return; }
    blocco[assenza._riga - prima] = [assenza.sostituto, assenza.criterio];
  });
  foglio.getRange(prima, COL_ASSENTI.SOSTITUTO, blocco.length, 2).setValues(blocco);
}
