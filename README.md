# Sostituzioni docenti — Google Spreadsheet + Apps Script

Foglio di lavoro per la gestione quotidiana delle sostituzioni in un istituto tecnico
superiore. Si registrano le assenze, si preme un bottone e viene generato un documento
Google con il prospetto delle sostituzioni, scegliendo i sostituti fra i docenti del
consiglio di quella classe o della stessa disciplina.

## Struttura del foglio di lavoro

| Foglio | A cosa serve | Chi lo compila |
|---|---|---|
| **Assenti** | Una riga per ogni *ora* da coprire | ogni giorno, chi gestisce le sostituzioni |
| **Disponibilità** | Ore a disposizione di ciascun docente | a inizio anno, aggiornato all'occorrenza |
| **Consigli di classe** | Docente / classe / disciplina | a inizio anno |
| **Storico** | Registro delle sostituzioni già assegnate | automatico |
| **Impostazioni** | Modello, cartella, pesi dell'algoritmo | una volta |
| *Elenchi* (nascosto) | Alimenta i menu a tendina | automatico |

### Assenti

| Data | Ora | Classe | Docente assente | Disciplina | Tipo assenza | Note | Sostituto | Criterio |
|---|---|---|---|---|---|---|---|---|
| 14/09/2026 | 2 | 3A | Bianchi Anna | Italiano | Malattia | | *(script)* | *(script)* |
| 14/09/2026 | 3 | 3A | Bianchi Anna | Italiano | Malattia | | *(script)* | *(script)* |

Una riga per ora, non per docente: è l'unità di lavoro reale ed è ciò che finisce nel
documento. Le ultime due colonne le compila lo script; le righe rimaste scoperte si
colorano di rosso.

### Disponibilità

| Docente | Giorno | Ora | Tipo | Valida dal | Valida al | Note |
|---|---|---|---|---|---|---|
| Rossi Mario | Lunedì | 2 | Disposizione | | | |
| Neri Sara | 14/09/2026 | 3 | Recupero permesso | | | solo questa settimana |

La colonna **Giorno** accetta due cose: il nome del giorno della settimana, per la
disponibilità ricorrente da orario, oppure una data precisa, per una disponibilità
straordinaria. *Valida dal / al* permette di limitare una riga a un periodo (per esempio
un quadrimestre) senza doverla cancellare.

### Consigli di classe

| Classe | Docente | Disciplina | Ruolo |
|---|---|---|---|
| 3A | Rossi Mario | Matematica | Docente |
| 3A | Bianchi Anna | Italiano | Coordinatore |
| 3A | Gialli Paolo | Sostegno | Sostegno |

Una riga per ogni abbinamento docente-classe-disciplina. È anche la fonte di tutti i
menu a tendina degli altri fogli: i nomi vanno scritti qui una volta sola e poi scelti
dall'elenco, così non nascono varianti di ortografia.

## Come sceglie il sostituto

Fra i soli docenti che risultano disponibili in quell'ora, e che non siano assenti né
già impegnati altrove nella stessa ora, la preferenza è, nell'ordine:

1. docente del **consiglio di classe** che insegna anche la **stessa disciplina**;
2. docente del **consiglio di classe**;
3. docente della **stessa disciplina** in altre classi;
4. qualunque altro docente a disposizione.

A parità di profilo pesano il tipo di disponibilità (di norma prima i recuperi di
permesso, poi le ore di disposizione) e il carico già sostenuto, sia nella giornata sia
nei 30 giorni precedenti: a parità di titolo viene proposto chi ha sostituito di meno.
I docenti di sostegno sono impiegati solo se non resta nessun altro.

Tutti i pesi sono numeri modificabili nel foglio **Impostazioni**, senza toccare il
codice. Le ore con meno candidati vengono assegnate per prime, così l'unico docente
possibile per un'ora difficile non viene consumato da un'ora che si sarebbe coperta
comunque.

## Installazione

1. Crea un nuovo foglio Google (`sheets.new`) e dagli un nome, per esempio *Sostituzioni
   2026/27*.
2. **Estensioni ▸ Apps Script**. Cancella il file `Codice.gs` che trovi già aperto.
3. Crea un file per ciascuno dei file in [src/](src/) (il pulsante **+ ▸ Script**),
   con lo stesso nome, e incollaci il contenuto. I prefissi numerici servono solo a
   tenere l'ordine leggibile.
4. In **Impostazioni progetto** spunta *Mostra il file manifest appsscript.json* e
   sostituisci il contenuto del manifest con quello di
   [src/appsscript.json](src/appsscript.json).
5. Salva e torna al foglio. Ricaricando la pagina compare il menu **📋 Sostituzioni**.
6. **📋 Sostituzioni ▸ Prepara il foglio**. Alla prima esecuzione Google chiede
   l'autorizzazione: è normale, lo script è tuo e gira con il tuo account.
7. **📋 Sostituzioni ▸ Crea documento modello**: genera un modello già impaginato e ne
   salva l'ID nelle impostazioni. Puoi aprirlo e personalizzarlo (logo, intestazione,
   firma del dirigente) purché lasci i segnaposto.
8. Compila **Consigli di classe** e **Disponibilità**.

### Il bottone

Nel foglio *Assenti*: **Inserisci ▸ Disegno**, disegna un rettangolo con la scritta
*Genera sostituzioni*, **Salva e chiudi**. Clicca il disegno, poi i tre puntini in alto
a destra ▸ **Assegna script** ▸ scrivi `calcolaECreaDocumento`.

## Uso quotidiano

1. Registri le assenze del giorno nel foglio **Assenti**, una riga per ora.
2. Premi il bottone. Lo script chiede la data (invio = oggi), riempie *Sostituto* e
   *Criterio* e crea il documento; una finestra riepiloga le ore coperte, quelle rimaste
   scoperte e il link al documento.
3. Se una proposta non va bene, scrivi tu il nome nella colonna *Sostituto* e **cancella
   il contenuto della colonna Criterio**: quella riga risulta decisa a mano e il calcolo
   non la tocca più. Poi usa *Solo documento* per rigenerare il prospetto.

Le altre voci di menu servono a separare i due passaggi: **Solo calcolo** scrive le
proposte nel foglio senza creare nulla su Drive, **Solo documento** genera il prospetto
da quello che c'è nel foglio, **Azzera le proposte del giorno** ricomincia da capo.

Rigenerare il documento per uno stesso giorno sostituisce le righe di quel giorno nello
Storico, quindi il conteggio dei carichi non viene falsato.

## Il documento modello

Un normale Google Doc con dei segnaposto fra doppie graffe:

| Segnaposto | Contenuto |
|---|---|
| `{{ISTITUTO}}` `{{DATA}}` `{{GIORNO}}` `{{GENERATO_IL}}` | intestazione |
| `{{ELENCO_ASSENTI}}` | «Bianchi Anna (ore 2, 3 · Malattia); …» |
| `{{TOTALE_ORE}}` `{{ORE_SCOPERTE}}` | riepilogo |

Più una riga di tabella che contiene `{{ORA}}`, `{{CLASSE}}`, `{{DOCENTE_ASSENTE}}`,
`{{DISCIPLINA}}`, `{{SOSTITUTO}}`, `{{CRITERIO}}`, `{{NOTE}}`: viene ripetuta una volta
per ogni sostituzione, mantenendo la formattazione che le hai dato. Funzionano anche
nell'intestazione e nel piè di pagina.

## Verifica del motore

[test/prova_motore.js](test/prova_motore.js) simula i fogli e controlla le regole di
scelta (affinità, docente già impegnato, tetto giornaliero, righe manuali, ore scoperte).
Si esegue con Node, senza collegarsi a Google:

```bash
node test/prova_motore.js
```
