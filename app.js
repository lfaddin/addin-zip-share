// *** SOSTITUISCI QUESTO PLACEHOLDER CON L'URL HTTP POST DEL TUO FLUSSO POWER AUTOMATE ***
const POWER_AUTOMATE_FLOW_URL = "https://default76581f54c29d419abc4b7c30934a22.32.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a70a4ca4fe00405dacd64e46e5ea9921/triggers/manual/paths/invoke/?api-version=1&tenantId=tId&environmentName=Default-76581f54-c29d-419a-bc4b-7c30934a2232&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8m7DhWSpYA3txsxlcAIuh_LRX04qJ3pHAs8QJJr1QyA";

let globalFileList = []; // Variabile globale per memorizzare l'elenco dei file ZIP

Office.onReady(function(info ) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Add-in di Outlook pronto!");
    }
});

// Gestione della selezione del file per aggiornare il nome visualizzato e leggere il contenuto ZIP
document.getElementById('fileZip').addEventListener('change', async function() {
    const fileNameSpan = document.getElementById('selectedFileName');
    globalFileList = []; // Resetta l'elenco ogni volta che un nuovo file viene selezionato

    if (this.files.length > 0) {
        const file = this.files[0];
        fileNameSpan.textContent = file.name;

        if (file.name.toLowerCase().endsWith('.zip')) {
            try {
                const zip = new JSZip();
                const zipContent = await zip.loadAsync(file);
                
                zipContent.forEach((relativePath, zipEntry) => {
                    if (!zipEntry.dir) { // Aggiungi solo i file, non le directory
                        globalFileList.push(relativePath);
                    }
                });
                console.log("Files in ZIP:", globalFileList);
            } catch (e) {
                console.error("Errore nella lettura del file ZIP:", e);
                alert("Errore nella lettura del file ZIP. Assicurati che sia un file ZIP valido.");
                fileNameSpan.textContent = 'Nessun file selezionato (Errore ZIP)';
                this.value = ''; // Pulisci l'input del file per forzare una nuova selezione
            }
        } else {
            alert("Per favore, seleziona un file .zip.");
            fileNameSpan.textContent = 'Nessun file selezionato';
            this.value = ''; // Pulisci l'input del file
        }
    } else {
        fileNameSpan.textContent = 'Nessun file selezionato';
    }
});

document.getElementById('zipShareForm').addEventListener('submit', async function(event) {
    event.preventDefault();

    const email = document.getElementById('emailDestinatario').value;
    const file = document.getElementById('fileZip').files[0];
    const submitButton = document.getElementById('submitButton');

    if (!email || !file) {
        alert('Per favore, compila tutti i campi.');
        return;
    }

    // Verifica che il file ZIP sia stato letto correttamente
    if (file.name.toLowerCase().endsWith('.zip') && globalFileList.length === 0) {
        alert('Attendere il caricamento del contenuto del file ZIP o selezionare un file ZIP valido.');
        return;
    }

    // Disabilita il pulsante per prevenire invii multipli
    submitButton.disabled = true;
    submitButton.textContent = 'Invio in corso...';

    try {
        // 1. Imposta l'email del destinatario in Outlook
        if (Office.context.mailbox.item) {
            const recipients = [{
                displayName: email,
                emailAddress: email
            }];

            const setEmailResult = await new Promise((resolve, reject) => {
                Office.context.mailbox.item.to.setAsync(recipients, function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Email destinatario impostata con successo in Outlook.");
                        resolve(true);
                    } else {
                        console.error("Errore nell'impostazione dell'email destinatario:", asyncResult.error.message);
                        reject(new Error("Errore nell'impostazione dell'email destinatario: " + asyncResult.error.message));
                    }
                });
            });

            if (!setEmailResult) {
                throw new Error("Impossibile impostare l'email del destinatario.");
            }
        } else {
            console.warn("Office.context.mailbox.item non disponibile. Impossibile impostare il destinatario.");
        }

        // 2. Leggi il contenuto del file ZIP come Base64
        const fileContentBase64 = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result.split(',')[1]); // Prende solo la parte Base64
            reader.onerror = error => reject(error);
            reader.readAsDataURL(file);
        });

        // 3. Invia i dati a Power Automate (incluso l'elenco dei file)
        const response = await fetch(POWER_AUTOMATE_FLOW_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                emailDestinatario: email,
                fileName: file.name,
                fileContent: fileContentBase64,
                fileList: globalFileList // Invia l'elenco dei file generato lato client
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Errore dal flusso Power Automate: ${response.status} - ${errorText}`);
        }

        const data = await response.json();
        const shareLink = data.shareLink;
        const receivedFileList = data.fileList; // Ricevi l'elenco dei file dal flusso (per consistenza)

        if (!shareLink) {
            throw new Error("Il flusso Power Automate non ha restituito un link di condivisione valido.");
        }

        console.log("Link di condivisione ricevuto:", shareLink);

        // 4. Inserisci il link e l'elenco dei file nel corpo della mail di Outlook
        if (Office.context.mailbox.item) {
            let emailBodyContent = `
                <p>Ecco il link al tuo file ZIP: <a href="${shareLink}">${shareLink}</a></p>
            `;
            
            if (receivedFileList && receivedFileList.length > 0) {
                emailBodyContent += `<p>Contenuto del file ZIP:</p><ul>`;
                receivedFileList.forEach(fileName => {
                    emailBodyContent += `<li>${fileName}</li>`;
                });
                emailBodyContent += `</ul>`;
            } else {
                emailBodyContent += `<p>Impossibile recuperare l'elenco dei file nel ZIP.</p>`;
            }

            await new Promise((resolve, reject) => {
                Office.context.mailbox.item.body.setSelectedDataAsync(
                    emailBodyContent,
                    { coercionType: Office.CoercionType.Html },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Contenuto inserito con successo nel corpo della mail.");
                            resolve(true);
                        } else {
                            console.error("Errore nell'inserimento del contenuto:", asyncResult.error.message);
                            reject(new Error("Errore nell'inserimento del contenuto: " + asyncResult.error.message));
                        }
                    }
                );
            });
        } else {
            console.warn("Office.context.mailbox.item non disponibile. Impossibile inserire il contenuto.");
        }

        alert('Operazione completata con successo! Link e dettagli inseriti nella mail.');
        // Puoi chiudere il task pane qui se l'operazione è conclusa
        // Office.context.ui.closeContainer();

    } catch (error) {
        console.error("Si è verificato un errore:", error);
        alert("Si è verificato un errore: " + error.message);
    } finally {
        // Riabilita il pulsante alla fine dell'operazione
        submitButton.disabled = false;
        submitButton.textContent = 'Invia';
    }
});
