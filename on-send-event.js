// Lädt office.js über die HTML-Runtime (encrypt-dialog.html) – hier nur Logik.
// WICHTIG: Handler müssen global (window) sein.

(function () {
  Office.onReady(() => {
    // Globale Zuordnung der Handler-Namen aus dem Manifest
    Office.actions.associate("onMessageSendHandler", window.onMessageSendHandler);
    Office.actions.associate("encryptSendFunction", window.encryptSendFunction);
  });

  // Button-Funktion (optional)
  window.encryptSendFunction = async (event) => {
    // Hier könntest du z. B. auch das Dialogfenster öffnen
    event.completed();
  };

  // OnMessageSend-Handler
  window.onMessageSendHandler = function (event) {
    try {
      // **FESTE absolute URL** zur Dialogseite auf dem CDN
      const dialogUrl = "https://cdn.jsdelivr.net/gh/GYGMOR/hinaddin@main/encrypt-dialog.html";

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 45, width: 30, requireHTTPS: true },
        (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            // Wenn der Dialog nicht öffnet, Senden erlauben
            event.completed({ allowEvent: true });
            return;
          }

          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            const choice = args.message;
            dialog.close();

            switch (choice) {
              case "cancel":
                event.completed({ allowEvent: false }); // Senden blockieren
                break;

              case "encrypt": {
                const item = Office.context.mailbox.item;
                item.subject.getAsync((res) => {
                  const current = (res.status === Office.AsyncResultStatus.Succeeded && res.value) ? res.value : "";
                  item.subject.setAsync(`[HIN] ${current}`, () => {
                    event.completed({ allowEvent: true });
                  });
                });
                break;
              }

              case "normal":
              default:
                event.completed({ allowEvent: true });
                break;
            }
          });
        }
      );
    } catch (e) {
      event.completed({ allowEvent: true });
    }
  };
})();
