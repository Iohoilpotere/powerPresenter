# PowerPresenter

PowerPresenter è un visualizzatore a tutto schermo pensato per eventi in presenza. Permette di individuare rapidamente i file PowerPoint contenuti in una cartella, mostrarne le anteprime e avviare la presentazione con un singolo click.

## Requisiti

- Windows 10 o successivo
- .NET 6 SDK (per la compilazione)
- Microsoft PowerPoint installato (necessario per l'esportazione delle anteprime tramite COM e per l'avvio delle presentazioni)

## Struttura della soluzione

- `PowerPresenter.sln`: soluzione Visual Studio 2022
- `src/PowerPresenter.Core`: componenti riutilizzabili (servizi, modelli, strategie di anteprima, gestione preferenze)
- `src/PowerPresenter.App`: applicazione WPF full-screen con interfaccia moderna e comandi operatore
- `install/PowerPresenterContextMenu.reg`: script di registrazione per aggiungere "Apri con PowerPresenter" al menu contestuale di Windows

## Build

1. Aprire `PowerPresenter.sln` in Visual Studio 2022.
2. Ripristinare i pacchetti NuGet.
3. Compilare in modalità `Release` per generare `PowerPresenter.exe`.

## Esecuzione

- Avviare l'eseguibile passando come argomento opzionale la cartella da esplorare:
  ```powershell
  PowerPresenter.exe "C:\\Percorso\\Alla\\Cartella"
  ```
- In assenza di argomenti è possibile scegliere la cartella tramite il pulsante *Seleziona cartella*.

## Menu contestuale di Windows

Per aprire PowerPresenter dal menu contestuale di una cartella:

1. Copiare `install/PowerPresenterContextMenu.reg` sul PC di destinazione.
2. Modificare i percorsi `C:\\Program Files\\PowerPresenter\\PowerPresenter.exe` inserendo il percorso reale dell'eseguibile.
3. Importare il file nel registro di sistema (doppio click o `reg import`).
4. Ripetere l'operazione per rimuovere la voce eliminando le chiavi corrispondenti.

Il comando aggiunge sia l'opzione sul menu della cartella sia su quello dello sfondo interno.

## Preferenze utente

- **Sfondo personalizzato:** il pulsante *Imposta sfondo* permette di scegliere un'immagine locale. Il percorso viene salvato in `%LOCALAPPDATA%\PowerPresenter\user-preferences.json`.
- **Monitor di destinazione:** il pulsante *Monitor* abilita tre modalità:
  - *Automatico* → avvia la presentazione su monitor esteso quando presente, altrimenti su quello principale.
  - *Primario* → forza la riproduzione sul monitor base disattivando la presenter view.
  - *Esteso* → abilita la presenter view quando disponibili più schermi.

Le preferenze vengono persistite tramite `UserPreferencesStore` (Singleton thread-safe) e riapplicate all'avvio.

## Anteprime

Le anteprime vengono generate seguendo una pipeline di strategie:
1. `InteropPreviewGenerationStrategy` utilizza PowerPoint per esportare la prima slide.
2. `ThumbnailPreviewGenerationStrategy` recupera eventuali miniature incorporate e, in mancanza, crea un segnaposto con titolo.

Le immagini sono memorizzate in `%LOCALAPPDATA%\PowerPresenter\Previews` e riutilizzate ai caricamenti successivi.

## Avvio delle presentazioni

`PowerPointPresentationLauncher` si occupa di configurare lo slide-show tramite COM. Il decorator `MonitorAwarePresentationLauncherDecorator` normalizza la preferenza monitor e garantisce la scelta corretta in base agli schermi disponibili.

## Note

- Le librerie COM di PowerPoint richiedono l'esecuzione su un ambiente Windows con l'applicazione Office installata.
- Tutte le chiamate potenzialmente lunghe vengono eseguite in background per mantenere reattiva la UI.
