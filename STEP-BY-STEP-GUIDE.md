# Step by step guide — getting the confirmation box working

**Written for someone who has never done this before. Follow it in order. Do not skip ahead.**

**Roughly 30 to 45 minutes.** Do it with a clear head, not late at night.

---

## What you are actually doing, in one paragraph

Outlook cannot show a custom "are you sure?" box on its own. You have to give it a tiny piece of software that does it. That software has to live on the internet somewhere so Outlook can fetch it, and then it has to be installed into Outlook by an administrator. That is the whole job:

> **Put the files online → tell the files where they live → install into Outlook → test**

---

# PART 1 — Put the files on the internet

We use **GitHub Pages** because it is free, permanent, and gives you a proper https address.

### 1.1 Get a GitHub account
1. Go to **github.com**
2. If you already have an account, sign in and skip to 1.2
3. Otherwise click **Sign up**, enter an email, a password and a username, and verify the email

⭐ **Use a work-appropriate username.** It becomes part of the web address that gets installed into everyone's Outlook.

### 1.2 Create a repository
A "repository" is just a folder that lives on GitHub.

1. Top right, click the **+** icon → **New repository**
2. **Repository name:** `stacks-addin`
3. **Description:** leave blank
4. Select **Public**
   - ⚠️ **It must be Public.** GitHub Pages does not work on private repositories with a free account.
   - ⚠️ **This means anyone can read these files.** That is fine — there is nothing secret in them. **Never put passwords, patient data or account numbers in this repository.**
5. Do **not** tick "Add a README file"
6. Click **Create repository**

### 1.3 Upload the files
1. On the page that appears, click **uploading an existing file** (blue link), or **Add file → Upload files**
2. Open the folder `smart-alerts-addin` on your computer
3. Drag these into the browser window:
   - `manifest.xml`
   - `commands.html`
   - `commands.js`
   - `icon-16.png`
   - `icon-32.png`
   - `icon-64.png`
   - `icon-80.png`
   - `icon-128.png`
4. Scroll down, click the green **Commit changes** button

You should now see all eight files listed.

### 1.4 Switch on GitHub Pages
1. Near the top of the repository page click the **Settings** tab
   - ⚠️ The repository's Settings, not your account settings
2. In the left sidebar click **Pages**
3. Under **Build and deployment**:
   - **Source:** Deploy from a branch
   - **Branch:** `main`, folder `/ (root)`
4. Click **Save**
5. **Wait one to two minutes**, then refresh the page

You should see a green box: *"Your site is live at https://YOURNAME.github.io/stacks-addin/"*

### 1.5 ⭐ Prove it worked before going further
Open a new browser tab and go to:

```
https://YOURNAME.github.io/stacks-addin/commands.js
```

replacing `YOURNAME` with your GitHub username.

- ✅ **You should see the code as plain text.**
- ❌ If you get "404 File not found", wait another two minutes and try again. If it still fails, Pages is not switched on properly. **Go back to 1.4. Do not continue until this works.**

⭐ **Write your address down. You need it in Part 2.**

---

# PART 2 — Tell the files where they live

The manifest currently says `https://REPLACE-ME`. You have to swap that for your real address.

1. On your computer, open the `smart-alerts-addin` folder
2. **Right-click `manifest.xml` → Open with → Notepad**
   - ⚠️ Notepad, not Word. Word will corrupt it.
3. Press **Ctrl + H** (Replace)
4. **Find what:** `https://REPLACE-ME`
5. **Replace with:** your address **without a trailing slash**, for example:
   ```
   https://YOURNAME.github.io/stacks-addin
   ```
6. Click **Replace All**. It should report **10 replacements**.
7. ⚠️ **Now fix one line by hand.** Press Ctrl+F and search for `AppDomain`. That line needs the bare domain only, no folder:
   ```xml
   <AppDomain>https://YOURNAME.github.io</AppDomain>
   ```
   Delete the `/stacks-addin` from that one line only.
8. **File → Save**, then close Notepad

⭐ **Check your work:** press Ctrl+F and search for `REPLACE-ME`. If anything is found, you missed one. There must be zero left.

---

# PART 3 — Install it into Outlook

### 3.1 Open the admin centre
1. Go to **admin.microsoft.com**
2. Sign in as **it@stackspharmacy.ie**

### 3.2 Upload
1. Left sidebar → **Settings** → **Integrated apps**
   - If you cannot see Settings, click **Show all** at the bottom of the sidebar
2. Click **Upload custom apps** (near the top)
3. **App type:** choose **Office Add-in**
4. Choose **Upload manifest file (.xml) from device**
5. Select your edited `manifest.xml`
6. Click **Next**

⚠️ **If it rejects the file here, stop and send me the exact error message.** Do not try to guess a fix.

### 3.3 Assign it to yourself only
1. On the users page choose **Just me** (or **Specific users/groups** and pick only your own account)
   - ⛔ **Do not choose Entire organization yet.** Test on yourself first.
2. **Next**
3. Review the permissions it requests, click **Accept**
4. **Finish deployment**, then **Done**

### 3.4 Wait
Microsoft says deployment can take **up to 24 hours**, though for a single user it is often only a few minutes.

---

# PART 4 — Test it

1. **Close Outlook completely and reopen it.** Not just the window — quit the app.
2. Start a new email
3. Put `achalnm02@gmail.com` in the To line
4. Add any subject and body
5. Press **Send**

### What should happen
A box appears containing **your own wording**, with two buttons:

- **Send Anyway**
- **Don't Send**

Click **Send Anyway** and the message goes.

### If nothing happens
1. Give it longer — deployment can lag
2. Restart Outlook again
3. Try **Outlook on the web** at outlook.office.com
4. Check your Outlook version: **File → Office Account → About Outlook**. Classic Outlook must be **version 2206 or later**.

⭐ **Record what each client does.** New Outlook, classic Outlook, web, phone. That record is the evidence for whether this can go to the branches.

---

# PART 5 — Only after it works: switch to the real address

⛔ **Do not do this until Part 4 has worked.**

1. On your computer open `commands.js` in Notepad
2. Near the top find `WATCHED_RECIPIENTS`
3. Replace the test entry with:
   ```javascript
   {
     address: "darraglynnnursinghome@healthmail.ie",
     label: "Darraglynn Nursing Home",
     message: "You are sending to DARRAGLYNN NURSING HOME. Please check this is the correct nursing home for this resident before sending."
   }
   ```
4. Save
5. Go back to GitHub → your repository → click `commands.js` → click the **pencil icon** to edit → paste the new content → **Commit changes**
6. Wait a couple of minutes for Pages to update
7. Restart Outlook and test again

⭐ **You do not need to reinstall the add-in for this.** It fetches the code fresh from GitHub, so editing `commands.js` is enough. Only changes to `manifest.xml` require redeploying.

---

# PART 6 — Only after all of that: give it to everyone

1. admin.microsoft.com → **Settings → Integrated apps**
2. Find **Recipient Confirmation** → **Edit users**
3. Change to **Entire organization**

⚠️ **Tell branch staff before you do this**, not after. A box they have never seen appearing on an email to a nursing home will otherwise generate phone calls.

⭐ Also tell Lara and Ade it is going live, so there is a written record with your name on it.

---

# If you get stuck

Send me:
1. Which part and step number
2. The exact error message, word for word
3. A screenshot

**Do not start changing things at random to see what happens.** That is how you end up not knowing which change fixed it.

---

# ⭐ Remember what you already have

The **warning ribbon is already live and working**, with no add-in, no hosting and no deployment. Ade has seen it.

If this add-in turns out to be more trouble than it is worth, the ribbon plus giving the nursing home contacts clearly distinct names is a real, defensible control. **This part is an improvement on something that already works, not a rescue of something broken.**
