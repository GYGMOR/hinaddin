/*
 * Event‑Handler für Smart Alerts in Outlook.
 *
 * Dieses Skript registriert einen Handler für den `OnMessageSend`‑Event
 * und öffnet ein eigenes Dialogfenster mit drei Optionen: normal senden,
 * verschlüsselt senden oder Abbrechen. Je nach Auswahl wird die E‑Mail
 * entweder normal versendet, das Betreff um einen HIN‑Marker erweitert
 * oder das Senden abgebrochen. 
 *
 * Hinweise:
 * – Die Datei muss über denselben Domänennamen ausgeliefert werden wie
 *   Ihre anderen Add‑In‑Ressourcen (manifest, HTML, Icons), z. B.
 *   https://addins.example.com/on-send-event.js.
 * – Outlook lädt für Windows nur die JavaScript‑Datei, daher
 *   implementiert dieses Skript den vollständigen Handler.
 */

(function () {
  /* Registriert die Funktionen, sobald Office bereit ist. */
  Office.onReady(function () {
    // Weist den Funktionsnamen in der Manifestdatei der Implementierung zu.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    // Optionaler Ribbon‑Button: ruft denselben Handler auf.
    Office.actions.associate("encryptSendFunction", onMessageSendHandler);
  });

  /**
   * Handler für den OnMessageSend‑Event. Dieser wird aufgerufen, wenn der
   * Benutzer in Outlook auf „Senden“ klickt. Er öffnet ein Dialogfenster,
   * das die Auswahl zwischen normalem Senden, HIN‑Verschlüsselung oder
   * Abbrechen ermöglicht.  Das Senden wird solange blockiert, bis der
   * Benutzer eine Auswahl trifft.
   *
   * @param {Office.AddinCommands.Event} event Das Ereignisobjekt, über das
   *        wir den Sendvorgang zulassen oder blockieren können.
   */
  function onMessageSendHandler(event) {
    try {
      // Ermittelt die Basis‑URL des Add‑Ins. Diese wird von Outlook zur
      // Laufzeit festgelegt und entspricht der URL, unter der die
      // Manifestdatei bereitgestellt wurde.
      var baseUrl = Office.context.extensionBaseUri || Office.context.mailbox?.item?.addIns?.extensionBaseUri || "";
      // URL des Dialogs. Der Pfad muss relativ zur Basis‑URL sein.
      var dialogUrl = baseUrl + "/encrypt-dialog.html";
      // Öffnet das Dialogfenster. Höhe und Breite sind in Prozenten der
      // Bildschirmgröße angegeben. Ein HTTPS‑Endpunkt ist Pflicht.
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 45, width: 30, requireHTTPS: true },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            var dialog = asyncResult.value;
            // Handler für vom Dialog gesendete Nachrichten.
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
              var choice = args.message;
              // Dialog schließen.
              dialog.close();
              // Je nach Auswahl fortfahren.
              if (choice === "cancel") {
                // Benutzer hat Abbrechen gewählt – Senden blockieren.
                event.completed({ allowEvent: false });
              } else if (choice === "normal") {
                // Normal senden – Senden zulassen.
                event.completed({ allowEvent: true });
              } else if (choice === "encrypt") {
                // Verschlüsselt senden – Betreff mit HIN‑Marker versehen.
                var item = Office.context.mailbox.item;
                // Aktuelles Betreff lesen.
                item.subject.getAsync(function (getResult) {
                  var currentSubject = "";
                  if (getResult.status === Office.AsyncResultStatus.Succeeded) {
                    currentSubject = getResult.value || "";
                  }
                    // Betreff mit HIN‑Marker versehen. Passen Sie den Marker
                    // nach Absprache mit dem HIN‑Gateway an (z. B. [HIN]).
                  var newSubject = "[HIN] " + currentSubject;
                  item.subject.setAsync(newSubject, function (setResult) {
                    // Unabhängig davon, ob das Setzen erfolgreich war,
                    // Senden zulassen (Fehler werden von Outlook gemeldet).
                    event.completed({ allowEvent: true });
                  });
                });
              } else {
                // Unbekannte Option – zur Sicherheit senden zulassen.
                event.completed({ allowEvent: true });
              }
            });
          } else {
            // Falls das Dialogfenster nicht geöffnet werden kann,
            // E‑Mail normal senden lassen (fail‑safe).
            event.completed({ allowEvent: true });
          }
        }
      );
    } catch (error) {
      // Bei unerwarteten Fehlern das Senden zulassen, um den Benutzer
      // nicht zu blockieren. Debug‑Informationen können in der Konsole
      // ausgegeben werden, wenn Outlook das erlaubt.
      event.completed({ allowEvent: true });
    }
  }
})();