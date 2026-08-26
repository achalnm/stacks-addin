/*
 * Stacks Pharmacy - Recipient Confirmation
 * Smart Alerts add-in. Fires on Send, checks the recipient list,
 * and asks the user to confirm before sending to a watched address.
 *
 * TO CHANGE WHO IS WATCHED: edit WATCHED_RECIPIENTS below.
 */

// ---------------------------------------------------------------------------
// CONFIGURATION - this is the only part that normally needs editing
// ---------------------------------------------------------------------------

const WATCHED_RECIPIENTS = [
  {
    address: "darraglynnnursinghome@healthmail.ie",
    label: "Darraglynn Nursing Home",
    message:
      "You are sending to DARRAGLYNN NURSING HOME. " +
      "Please double check this is the correct nursing home before sending."
  },
  {
    // TEMPORARY - proves the add-in still fires after this edit.
    // Remove in a second commit once confirmed working.
    address: "achalnm02@gmail.com",
    label: "Achal personal (TEST)",
    message:
      "TEST WARNING. You are sending to achalnm02@gmail.com. " +
      "Please confirm this is the correct recipient before sending."
  }
];

// ---------------------------------------------------------------------------

// Register the handler as early as possible AND again once Office is ready.
// Belt and braces: on some clients Office.actions is not populated at parse time,
// and a single guarded call can silently no-op, leaving the send event unhandled.
function registerHandler() {
  try {
    if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
      Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
      console.log("Recipient Confirmation: handler associated.");
    }
  } catch (e) {
    console.error("Recipient Confirmation: associate failed", e);
  }
}

Office.onReady(function () {
  console.log("Recipient Confirmation: Office.onReady fired.");
  registerHandler();
});

/**
 * Entry point. Registered against OnMessageSend in the manifest.
 */
function onMessageSendHandler(event) {
  getAllRecipients()
    .then(function (recipients) {
      const hit = findWatchedRecipient(recipients);

      if (!hit) {
        // Nothing watched in the recipient list. Send normally.
        event.completed({ allowEvent: true });
        return;
      }

      // Watched recipient found. Stop the send and show the confirmation.
      // Because SendMode is PromptUser, the user gets "Send Anyway" and "Don't Send".
      event.completed({
        allowEvent: false,
        errorMessage: hit.message
      });
    })
    .catch(function (err) {
      // FAIL OPEN, DELIBERATELY.
      // If anything goes wrong we allow the send rather than trapping the user.
      // A broken add-in must never stop a pharmacy emailing a nursing home.
      console.error("Recipient confirmation add-in error:", err);
      event.completed({ allowEvent: true });
    });
}

/**
 * Collects To, Cc and Bcc into a single array of lowercased addresses.
 */
function getAllRecipients() {
  const item = Office.context.mailbox.item;

  return Promise.all([
    getRecipientField(item.to),
    getRecipientField(item.cc),
    getRecipientField(item.bcc)
  ]).then(function (fields) {
    return fields
      .reduce(function (all, field) {
        return all.concat(field);
      }, [])
      .map(function (r) {
        return (r.emailAddress || "").trim().toLowerCase();
      })
      .filter(Boolean);
  });
}

/**
 * Wraps a recipient field's getAsync in a promise.
 */
function getRecipientField(field) {
  return new Promise(function (resolve) {
    if (!field || typeof field.getAsync !== "function") {
      resolve([]);
      return;
    }
    field.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
      } else {
        resolve([]);
      }
    });
  });
}

/**
 * Returns the first watched entry present in the recipient list, or null.
 */
function findWatchedRecipient(recipients) {
  for (let i = 0; i < WATCHED_RECIPIENTS.length; i++) {
    const watched = WATCHED_RECIPIENTS[i];
    const target = watched.address.trim().toLowerCase();
    if (recipients.indexOf(target) !== -1) {
      return watched;
    }
  }
  return null;
}

// Also attempt registration immediately at parse time (covers clients that
// populate Office.actions synchronously and may dispatch before onReady).
registerHandler();
