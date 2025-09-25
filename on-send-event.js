/*
 * Event-Handler für Smart Alerts in Outlook.
 *
 * Dieses Skript registriert einen Handler für den `OnMessageSend`-Event
 * und öffnet ein eigenes Dialogfenster mit drei Optionen: normal senden,
 * verschlüsselt senden oder Abbrechen. Je nach Auswahl wird die E-Mail
 * entweder normal versendet, der Betreff mit einem HIN-Marker versehen
 * oder das Senden abgebrochen.
 */

(() => {
  /**
   * Registriert die Funktionen, sobald Office bereit ist.
   */
  Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", async (event) => {
    // deine Logik ...
    event.completed({ allowEvent: true });
  });
  Office.actions.associate("encryptSendFunction", async (event) => {
    // deine Button-Logik ...
    event.completed();
  });
});


  /**
   * Handler für den OnMessageSend-Event.
   * Wird aufgerufen, wenn der Benutzer in Outlook auf „Senden“ klickt.
   * Öffnet ein Dialogfenster mit Optionen und gibt das Ergebnis an Outlook zurück.
   *
   * @param {Office.AddinCommands.Event} event Das Ereignisobjekt, über das der
   *    Sendvorgang zugelassen oder blockiert werden kann.
   */
  function onMessageSendHandler(event) {
    try {
      // Ermittelt die Basis-URL des Add-Ins. Entspricht der URL der Manifestdatei.
      const baseUrl =
        Office.context.extensionBaseUri ||
        (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.addIns
          ? Office.context.mailbox.item.addIns.extensionBaseUri
          : "");
      // Definiert die Dialog-URL relativ zur Basis-URL.
      // Verwenden Sie die ASPX-Datei des Dialogs im GitHub-Repository.
      const dialogUrl = `${baseUrl}/encrypt-dialog.aspx`;
      // Öffnet das Dialogfenster. Höhe und Breite sind Prozentsätze der Bildschirmgröße.
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 45, width: 30, requireHTTPS: true },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const dialog = asyncResult.value;
            // Handler für Nachrichten aus dem Dialogfenster.
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
              const choice = args.message;
              // Schließt das Dialogfenster.
              dialog.close();
              switch (choice) {
                case "cancel":
                  // Abbruch – Senden blockieren.
                  event.completed({ allowEvent: false });
                  break;
                case "normal":
                  // Normal senden – Senden zulassen.
                  event.completed({ allowEvent: true });
                  break;
                case "encrypt":
                  // Verschlüsselt senden – Betreff mit HIN-Marker versehen.
                  const item = Office.context.mailbox.item;
                  item.subject.getAsync((getResult) => {
                    let currentSubject = "";
                    if (getResult.status === Office.AsyncResultStatus.Succeeded) {
                      currentSubject = getResult.value || "";
                    }
                    const newSubject = `[HIN] ${currentSubject}`;
                    item.subject.setAsync(newSubject, () => {
                      // Unabhängig vom Ergebnis das Senden zulassen.
                      event.completed({ allowEvent: true });
                    });
                  });
                  break;
                default:
                  // Unbekannte Option – zur Sicherheit senden zulassen.
                  event.completed({ allowEvent: true });
                  break;
              }
            });
          } else {
            // Fehler beim Öffnen des Dialogs – E-Mail normal senden lassen.
            event.completed({ allowEvent: true });
          }
        }
      );
    } catch (error) {
      // Bei unerwarteten Fehlern das Senden zulassen.
      event.completed({ allowEvent: true });
    }
  }
})();
