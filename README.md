# Sociogramma

Analisi delle preferenze di collaborazione in un gruppo di lavoro secondo il
metodo di Moreno. A partire da un file Excel con le scelte e i rifiuti di ogni
persona, `sociogramma.py` produce i diagrammi, gli indici sociometrici e un
report commentato con indicazioni per comporre i team.

Il programma è un unico file Python: si può usare con un'interfaccia grafica
(doppio clic) oppure da riga di comando.

## Contenuto della cartella

| File | Descrizione |
|------|-------------|
| `sociogramma.py` | Il programma (interfaccia grafica + CLI) |
| `LEGGIMI - istruzioni.txt` | Istruzioni passo passo per chi usa il programma su Windows |
| `esempio_input.xlsx` | File Excel di esempio con dati fittizi |

## Requisiti

- Python 3.9 o successivo (con Tkinter per l'interfaccia grafica)
- Librerie: `pandas`, `numpy`, `networkx`, `matplotlib`, `openpyxl`,
  `python-docx` (facoltativa, serve per il report)
- LibreOffice, solo se si vuole il report anche in PDF (`--pdf`)

Al primo avvio dall'interfaccia grafica il programma propone di installare da
solo le librerie mancanti. In alternativa:

```bash
pip install pandas numpy networkx matplotlib openpyxl python-docx
```

## Uso

### Interfaccia grafica

```bash
python sociogramma.py
```

Avviato senza argomenti (o con doppio clic) apre una finestra in cui scegliere
il file Excel, la cartella dei risultati e le opzioni. Su Windows il pulsante
"Crea l'avvio rapido" mette sul Desktop un collegamento al programma.

### Riga di comando

```bash
python sociogramma.py dati.xlsx
python sociogramma.py dati.xlsx -o risultati --formato pdf --insieme --pdf
python sociogramma.py --crea-esempio esempio.xlsx
```

Opzioni principali:

| Opzione | Effetto |
|---------|---------|
| `-o, --output DIR` | Cartella di output (default `output_sociogramma`) |
| `-f, --foglio` | Nome o indice del foglio Excel (default: il primo) |
| `--etichetta COLONNA` | Colonna da usare come etichetta dei nodi al posto di NOME + COGNOME |
| `--etichette-dentro` | Scrive le etichette dentro il pallino |
| `--layout` | `bersaglio` (default), `spring`, `kamada`, `circolare` |
| `--insieme` | Genera anche il grafico con scelte e rifiuti sovrapposti |
| `--includi-esterni` | Mantiene come nodi gli ID citati ma assenti dalla colonna ID |
| `--formato` | Formato dei grafici: `png` (default), `pdf`, `svg` |
| `--dpi N` | Risoluzione dei grafici (default 200) |
| `--pdf` | Converte il report anche in PDF (richiede LibreOffice) |
| `--no-html` / `--no-report` | Salta la versione interattiva / il report |
| `--seed N` | Seed per i layout casuali (default 42) |
| `--titolo TESTO` | Titolo dei grafici |
| `--crea-esempio FILE.xlsx` | Genera un Excel di esempio e termina |

Elenco completo: `python sociogramma.py --help`.

## Formato del file Excel

Una riga per persona, con queste colonne nella prima riga:

| Colonna | Contenuto |
|---------|-----------|
| `ID` | Codice breve della persona (es. `P01`) |
| `COGNOME`, `NOME` | Facoltative: se mancano si usa l'ID |
| `SI-1`, `SI-2`, `SI-3` | ID delle persone con cui vorrebbe lavorare, in ordine di preferenza |
| `NO-1`, `NO-2` | ID delle persone con cui non vorrebbe lavorare |

Le caselle possono restare vuote. Maiuscole/minuscole e varianti come `SI 1` o
`si_1` sono accettate. Gli errori di compilazione (ID inesistenti,
autoindicazioni, duplicati) vengono scartati e segnalati nel foglio
"Anomalie" di `indici.xlsx`.

## Output

| File | Contenuto |
|------|-----------|
| `sociogramma_scelte.png` | Diagramma delle preferenze positive |
| `sociogramma_rifiuti.png` | Diagramma dei rifiuti, con le stesse posizioni |
| `sociogramma_completo.png` | Scelte e rifiuti sovrapposti (solo con `--insieme`) |
| `sociogramma.html` | Versione interattiva offline (zoom, nodi trascinabili, filtri) |
| `indici.xlsx` | Indici individuali e di gruppo, coppie reciproche, sottogruppi, matrice, anomalie |
| `report.docx` | Report commentato con metodo, tabelle, diagrammi e indicazioni per i team |
| `report.pdf` | Solo con `--pdf` |

## Riservatezza

Diagrammi e report contengono nomi e informazioni delicate su chi viene scelto
o rifiutato dai colleghi: vanno tenuti riservati a chi conduce l'analisi. Per
una diffusione più ampia conviene usare codici al posto dei nomi
(`--etichetta ID`).

Per lo stesso motivo il `.gitignore` esclude le cartelle di output: non
committare dati reali nel repository.
