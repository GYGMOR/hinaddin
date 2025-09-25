(function () {
  Office.onReady(() => {
    Office.actions.associate("onMessageSendHandler", window.onMessageSendHandler);
    Office.actions.associate("encryptSendFunction", window.encryptSendFunction);
  });

  window.encryptSendFunction = async (event) => {
    // Optional: zusätzliche Logik
    event.completed();
  };

  window.onMessageSendHandler = function (event) {
    try {
      const dialogUrl = "https://cdn.jsdelivr.net/gh/GYGMOR/hinaddin@main/encrypt-dialog.html";

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 45, width: 30, requireHTTPS: true },
        (r) => {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true }); // Dialog nicht geöffnet -> senden erlauben
            return;
          }
          const dialog = r.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            const choice = args.message;
            dialog.close();

            switch (choice) {
              case "cancel":
                event.completed({ allowEvent: false });
                break;
              case "encrypt": {
                const item = Office.context.mailbox.item;
                item.subject.getAsync((gr) => {
                  const current = (gr.status === Office.AsyncResultStatus.Succeeded && gr.value) ? gr.value : "";
                  item.subject.setAsync(`[HIN] ${current}`, () => event.completed({ allowEvent: true }));
                });
                break;
              }
              case "normal":
              default:
                event.completed({ allowEvent: true });
            }
          });
        }
      );
    } catch {
      event.completed({ allowEvent: true });
    }
  };
})();
