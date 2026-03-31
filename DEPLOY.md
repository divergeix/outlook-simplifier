# Outlook Thread Simplifier — Desktop Add-in
## Deployment Guide (For First-Time Users)

---

## What You Are Deploying

This is a **Microsoft Office Add-in** — a panel that appears INSIDE Outlook desktop
(Windows or Mac) as a sidebar. It uses the official Microsoft extension system.

Works with:
- ✅ Outlook Desktop on Windows (Microsoft 365 / Outlook 2019+)
- ✅ Outlook Desktop on Mac (Microsoft 365)
- ✅ Outlook Web (as a bonus — same add-in works in browser too)

Does NOT work with:
- ❌ Outlook 2016 or older
- ❌ The Windows Mail app (that is not Outlook)

---

## STEP 1 — Create a Free GitHub Account (5 min)

You need GitHub to host the add-in files for free with HTTPS.
(HTTPS is required — Microsoft will not load add-ins from plain HTTP)

1. Go to https://github.com
2. Click "Sign up" — use any email
3. Choose the **Free** plan
4. Verify your email

---

## STEP 2 — Create a New Repository (3 min)

1. After logging in, click the **+** icon (top right) → **New repository**
2. Repository name: `outlook-simplifier`  (exactly this, lowercase)
3. Make sure it is set to **Public**
4. Check ✅ "Add a README file"
5. Click **Create repository**

---

## STEP 3 — Upload the Add-in Files (5 min)

You need to upload these files from the zip you downloaded:
```
taskpane.html
commands.html
manifest.xml
icon-16.png
icon-32.png
icon-64.png
icon-80.png
icon-128.png
```

How to upload:
1. On your new repository page, click **Add file** → **Upload files**
2. Drag ALL the files above into the upload area (select all in your folder)
3. At the bottom, click **Commit changes**

---

## STEP 4 — Enable GitHub Pages (3 min)

GitHub Pages turns your repository into a free HTTPS website.

1. On your repository page, click **Settings** (top menu)
2. Scroll down to **Pages** in the left sidebar
3. Under "Source", select **Deploy from a branch**
4. Branch: select **main** | Folder: select **/ (root)**
5. Click **Save**
6. Wait 2-3 minutes, then refresh the page
7. You will see: "Your site is published at **https://YOURUSERNAME.github.io/outlook-simplifier**"

Copy that URL — you will need it in the next step.
Example: `https://johndoe.github.io/outlook-simplifier`

---

## STEP 5 — Update the manifest.xml File (3 min)

The manifest.xml file tells Outlook where to find the add-in.
You need to replace `YOUR_GITHUB_PAGES_URL` with your actual URL.

### Option A: Edit on GitHub (recommended for beginners)
1. On your repository page, click on **manifest.xml**
2. Click the ✏ pencil icon (Edit this file)
3. Press **Ctrl+H** (Find and Replace) — type:
   - Find: `YOUR_GITHUB_PAGES_URL`
   - Replace: `https://YOURUSERNAME.github.io/outlook-simplifier`
   - Click **Replace All**
4. Click **Commit changes** → Commit changes

### Option B: Edit locally
1. Open manifest.xml in Notepad
2. Find all instances of `YOUR_GITHUB_PAGES_URL`
3. Replace with `https://YOURUSERNAME.github.io/outlook-simplifier`
4. Save → re-upload to GitHub

After editing, your manifest.xml URLs should look like:
```
https://johndoe.github.io/outlook-simplifier/taskpane.html
https://johndoe.github.io/outlook-simplifier/icon-16.png
```
(not YOUR_GITHUB_PAGES_URL/...)

---

## STEP 6 — Download the manifest.xml to Your Computer

After editing on GitHub:
1. Click on manifest.xml in your repository
2. Click the **⬇ Download raw file** button
3. Save it somewhere you can find (e.g., Desktop)

---

## STEP 7 — Load the Add-in into Outlook Desktop

### On Windows:

1. Open **Outlook Desktop**
2. Click **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Click **Trusted Add-in Catalogs** in the left menu
4. In the "Catalog URL" box, type the **folder path** where you saved manifest.xml
   - Example: if saved to Desktop, type: `C:\Users\YourName\Desktop`
   - Or create a dedicated folder: `C:\OutlookAddins`
5. Click **Add to List** → check ✅ **Show in Menu**
6. Click **OK** → **OK**
7. **Restart Outlook completely**

Then to activate:
1. Open any email in Outlook
2. Click the **Home** tab in the ribbon
3. Find the **Thread Simplifier** button (or click **Get Add-ins** / **My Add-ins**)
4. Click **⚡ Simplify** — the sidebar opens

### On Mac:

1. Open **Outlook Desktop**
2. Click the **Tools** menu → **Get Add-ins**
3. Click **My Add-ins** → **Add a custom add-in** → **Add from file**
4. Select your `manifest.xml` file
5. Click **Install**
6. Open any email → click **⚡ Thread Simplifier** in the ribbon

---

## STEP 8 — Add Your Anthropic API Key (2 min)

When the sidebar opens for the first time, it will show a red banner asking for your API key.

1. Go to https://console.anthropic.com/account/keys
2. Sign up if you haven't already
3. Click **Create Key** → copy the key (starts with `sk-ant-api03-`)
4. Add a small credit at https://console.anthropic.com/settings/billing ($5 is plenty — each analysis costs ~$0.02)
5. Paste the key into the red banner in the sidebar → click **Save Key**
6. The red banner disappears — you are ready

---

## STEP 9 — Using the Add-in

1. Open any email thread in Outlook
2. The sidebar should be visible on the right. If not, click **⚡ Simplify** in the ribbon
3. Click **⚡ Run Full Analysis**
4. Results appear in ~5 seconds

**Tips:**
- For best results, open the **last email in a long thread** — Outlook includes
  all the quoted replies in it, so the AI sees the full conversation
- The **Flow tab** shows how many messages were detected
- **Reply tab** → generate replies in 8 different tones
- **Export** → PDF, Text, Notes, Task List
- **Chat tab** → ask any question about the email

---

## Troubleshooting

**"Add-in not appearing in ribbon"**
→ Restart Outlook completely after sideloading
→ Check manifest.xml has the correct GitHub Pages URL (no YOUR_GITHUB_PAGES_URL left)

**"Cannot reach add-in page"**
→ Wait 5 minutes after enabling GitHub Pages — it takes time to deploy
→ Try opening https://YOURUSERNAME.github.io/outlook-simplifier/taskpane.html in a browser — if it loads, the add-in will work

**"API error / No API key"**
→ Make sure you added billing credit at console.anthropic.com
→ Re-enter the API key in the sidebar red banner

**"Only 1 message found in thread"**
→ Open the LAST email in the thread (the most recent reply)
→ Outlook includes all quoted history in the latest reply

**Add-in shows blank white panel**
→ This usually means GitHub Pages is still deploying. Wait 5 minutes and reload.

---

## Cost Reference
- Each analysis: ~$0.02–0.04 (2–4 cents)
- 100 analyses per month: ~$2–4
- API costs are billed directly to your Anthropic account, not to this tool
