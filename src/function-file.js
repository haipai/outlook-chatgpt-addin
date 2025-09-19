/* global Office */

const ENDPOINT = "https://chubby-ducks-sing.loca.lt/draft"; // we'll switch to HTTPS in the next task

Office.onReady(() => { /* ready */ });

async function onDraftClicked(event) {
  try {
    Office.context.mailbox.item.getSelectedDataAsync(
      Office.CoercionType.Text,
      async (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          notify("Could not read selection.");
          return event.completed();
        }
        const selection = (res.value || "").trim();
        if (!selection) {
          notify("Select a placeholder first (e.g., 'help draft: ...').");
          return event.completed();
        }

        const prompt =
          `Rewrite the text into a clear, friendly, concise email paragraph.\n---\n${selection}\n---\nReturn only the paragraph.`;

        let text = "";
        try {
          const r = await fetch(ENDPOINT, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ prompt }),
          });
          if (!r.ok) throw new Error(`Service error ${r.status}`);
          const data = await r.json();
          text = (data?.text || "").trim();
        } catch (e) {
          notify(`Request failed: ${e.message || e}`);
          return event.completed();
        }

        if (!text) {
          notify("No text returned.");
          return event.completed();
        }

        Office.context.mailbox.item.setSelectedDataAsync(
          text,
          { coercionType: Office.CoercionType.Text },
          (setRes) => {
            if (setRes.status !== Office.AsyncResultStatus.Succeeded) {
              notify("Failed to insert generated text.");
            }
            event.completed();
          }
        );
      }
    );
  } catch (err) {
    notify(`Unexpected error: ${err.message || err}`);
    event.completed();
  }
}

// REQUIRED for command buttons: register the function name globally
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onDraftClicked", onDraftClicked);
} else {
  // Fallback for older clients
  // eslint-disable-next-line no-undef
  window.onDraftClicked = onDraftClicked;
}

function notify(message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("draft-msg", {
      type: "informationalMessage",
      message,
      icon: "icon-16",
      persistent: false,
    });
  } catch (_) {}
}
