/* commands.js
   Hosted over HTTPS (e.g., GitHub Pages). No HTML needed. */

function replaceWithHelloWorld(event) {
  // Write or replace selection in compose body
  const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
  if (!item || !item.body || typeof item.body.setSelectedDataAsync !== "function") {
    // In case someone clicks in a non-compose surface (shouldn’t happen due to rules)
    notifyAndComplete(event, "This command works only in a compose window.");
    return;
  }

  item.body.setSelectedDataAsync(
    "hello world",
    { coercionType: Office.CoercionType.Text },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        notifyAndComplete(event); // success, no message necessary
      } else {
        const msg = (result.error && result.error.message) ? result.error.message : "Unknown error";
        notifyAndComplete(event, "Error: " + msg);
      }
    }
  );
}

// Helper to finish command (and optionally show a simple notification)
function notifyAndComplete(event, message) {
  try {
    if (message && Office.context.ui && Office.context.mailbox) {
      Office.context.mailbox.item.notificationMessages.replaceAsync("helloStatus", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "icon32",
        persistent: false
      });
    }
  } catch (e) {
    // best-effort; ignore notification errors
  } finally {
    event.completed(); // MUST call to return control to Outlook
  }
}

// Associate the function name defined in manifest
Office.actions.associate("replaceWithHelloWorld", replaceWithHelloWorld);
