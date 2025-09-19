/* global Office */
Office.onReady(() => { /* ready */ });

// Point this to your backend once it’s deployed over HTTPS
const ENDPOINT = "http://localhost:8787/draft";

export async function onDraftClicked(event) {
  Office.context.mailbox.item.getSelectedDataAsync(
    Office.CoercionType.Text,
    async (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        return done(event, "Could not read selection.");
      }
      const selection = (res.value || "").trim();
      if (!selection) {
        return done(event, "Select a placeholder first, e.g., 'help draft: ...'.");
      }

      const prompt =
        `Rewrite the text into a clear, friendly, concise email paragraph.\n---\n${selection}\n---\nReturn only the paragraph.`;

      try {
        const r = await fetch(ENDPOINT, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ prompt }),
        });
        if (!r.ok) {
          return done(event, `Service error ${r.status}`);
        }
        const { text } = await r.json();
        const output = (text || "").trim() || "(no result)";

        Office.context.mailbox.item.setSelectedDataAsync(
          output,
          { coercionType: Office.CoercionType.Text },
          (setRes) => {
            if (setRes.status !== Office.AsyncResultStatus.Succeeded) {
              notify("Failed to insert generated text.");
            }
            event.completed();
          }
        );
      } catch (e) {
        done(event, String(e && e.message ? e.message : e));
      }
    }
  );
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
function done(event, message) { notify(message); try { event.completed(); } catch (_) {} }
