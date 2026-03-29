# Cartelle Sanitarie — Generatore automatico

Applicazione web per la generazione automatica di cartelle sanitarie individuali
a partire da file TOTALI Word e anagrafiche Excel.

## Funzionamento

1. Carica i file TOTALI `.docx` (idoneità e cartelle sanitarie) e le anagrafiche `.xlsx`
2. Clicca **Genera cartelle**
3. Scarica lo ZIP con una cartella per ogni lavoratore

## Struttura ZIP generato

```
cartelle_sanitarie.zip
├── COGNOME_NOME_CF/
│   ├── idoneita_COGNOME_NOME_CF.pdf
│   └── cartella_COGNOME_NOME_CF.pdf
└── ...
```

## Privacy

I file vengono elaborati interamente in memoria (RAM).
Non vengono salvati su disco né trasmessi a terzi.
Al termine della sessione non rimane nulla.

## Deploy

Pubblicata su [Streamlit Community Cloud](https://streamlit.io/cloud).
