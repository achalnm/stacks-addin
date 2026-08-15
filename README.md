# Recipient Confirmation — Outlook Smart Alerts add-in

**Purpose.** When anyone at Stacks sends an email to a nominated address, Outlook stops the send and shows a confirmation they must accept before it goes.

**Why this and not Purview DLP.** ⭐ **The DLP oversharing dialog requires Microsoft 365 E5. Stacks does not have it** — the tenant runs Business Premium, Business Standard, Exchange Online Plan 1 and 2, and Defender for Office P1. The DLP rule that was built is correct and enforces the block, but the override button and custom text are E5 features and will never appear. Verified against the tenant's own licence list on Fri 14 Aug 2026.

⭐ **Smart Alerts has no E5 requirement.** It works in new Outlook for Windows, classic Outlook, Outlook on the web and Mac.

---

## Files

| File | What it is |
|---|---|
| `manifest.xml` | The add-in definition. Uploaded to the admin centre. |
| `commands.js` | ⭐ **The logic. This is the only file you normally edit.** |
| `commands.html` | Invisible loader page. Never seen by users. |
| `icon-*.png` | Icons the manifest requires. |

---

## ⭐ Changing who gets the warning

Everything lives at the top of **`commands.js`**:

```javascript
const WATCHED_RECIPIENTS = [
  {
    address: "achalnm02@gmail.com",
    label: "Achal personal (TEST)",
    message: "TEST WARNING. You are sending to achalnm02@gmail.com. ..."
  }
];
```

To go live at Darraglynn, replace that entry (or uncomment the second one already in the file):

```javascript
{
  address: "darraglynnnursinghome@healthmail.ie",
  label: "Darraglynn Nursing Home",
  message: "You are sending to DARRAGLYNN NURSING HOME. Please check this is the correct nursing home for this resident before sending."
}
```

⭐ **Multiple addresses are supported** — add as many entries as you like, each with its own wording. That matters: covering every nursing home is better than covering one, because warning on a single address quietly teaches staff that no warning means safe.

⚠️ **Message text limit is roughly 500 characters.**

---

## Deployment

### Step 1 — Host the files on HTTPS

The files must be reachable over HTTPS. **GitHub Pages is free and adequate.**

1. Create a **public** GitHub repository, e.g. `stacks-recipient-confirmation`.
2. Upload `manifest.xml`, `commands.html`, `commands.js` and all `icon-*.png` files.
3. **Settings → Pages → Source: `main` branch, `/root`** → Save.
4. Wait a minute, then note the URL: `https://<username>.github.io/<repo>/`
5. Confirm `https://<username>.github.io/<repo>/commands.js` loads in a browser.

### Step 2 — Point the manifest at that URL

In `manifest.xml`, **find and replace every `https://REPLACE-ME` with your Pages URL** (no trailing slash).

There are 10 occurrences. Miss one and the add-in silently fails to load.

⚠️ The `<AppDomain>` entry needs the bare domain, e.g. `https://<username>.github.io`.

### Step 3 — Deploy from the admin centre

1. **admin.microsoft.com → Settings → Integrated apps**
2. **Upload custom apps**
3. App type: **Office Add-in**, then **Upload manifest file (.xml)**
4. Assignment: ⭐ **start with "Just me" or a small test group.** Do not deploy org-wide first.
5. Accept permissions and finish.

⏱️ **Deployment can take up to 24 hours to reach all users**, though it is often much quicker for a single test user. Restart Outlook.

### Step 4 — Test

Send to the watched address. Expected:

> A dialog with **your custom message**, and buttons to **Send Anyway** or **Don't Send**.

⭐ **Test on every client the branches use** — new Outlook, classic Outlook, Outlook on the web, mobile. Record each result.

---

## ⭐ Design decisions worth understanding

### `SendMode="PromptUser"` — deliberate

If the add-in cannot run (offline, slow to load, error), **the user is prompted but can still send.** It fails open.

⛔ **Do not change this to `Block` or `SoftBlock`.** Doing so means that if the add-in ever breaks, staff cannot email nursing homes at all and have no way to comply. That is exactly the trap this project exists to avoid, and it is the same sequencing principle as the billing project: never deploy a block that people cannot satisfy.

### The code also fails open

Any error in the handler results in `allowEvent: true`. A bug in this add-in must never stop a pharmacy contacting a nursing home.

### To, Cc and Bcc are all checked

Not just the To line.

---

## ⚠️ Known limitations — do not oversell these

1. **It is a confirmation, not a prevention.** Someone who clicks Send Anyway still sends. It reduces error; it does not eliminate it.
2. **It fails open by design.** If the add-in does not load, the email goes.
3. **Client support varies.** Requires Mailbox requirement set 1.12 or later. Older Outlook versions will not run it. **Test before promising.**
4. **It only covers listed addresses.** Anything not in `WATCHED_RECIPIENTS` produces no warning.

---

## ⭐ This is layer two. Do not drop layers one and three.

| Layer | Status |
|---|---|
| **MailTip ribbon** | ✅ **Already live and working.** Warns while composing, *before* send. Needs no licence, no add-in, no deployment. **Keep it.** |
| **This add-in** | Confirmation at the moment of sending. |
| ⭐ **Distinct contact names** | **The prevention layer, and still the most valuable.** The mail contact's display name is what staff read in the To line — verified from message headers. Naming the homes so they cannot be confused removes the trap rather than warning about it, and protects every home at once. |

⭐ **If the add-in ever proves too much to maintain, layers one and three alone remain a defensible answer.**
