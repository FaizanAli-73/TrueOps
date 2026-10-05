# Craftaire: Google Sheets + WhatsApp setup

Allow about an hour for this, once. When you're done, texting the business WhatsApp number fills in the workbook.

```
You text WhatsApp  →  Meta (WhatsApp Cloud API)  →  Cloudflare Worker (relay)  →  Google Apps Script  →  Claude reads it  →  writes to Google Sheets
                                                                                                  ↓
                                                                              WhatsApp reply: "✅ Added Time Log …  Reply UNDO to reverse"
```

What you'll need:
- the Google account the workbook will live in
- a Meta (Facebook) account
- an Anthropic API account (console.anthropic.com)
- a free Cloudflare account
- **a phone number for the bot.** The easiest option is a cheap second number (a prepaid SIM) that isn't already on WhatsApp. The team keeps texting from your own WhatsApp; the bot is just a contact you message.

---

## Step 1: Put the workbook in Google Sheets (5 min)

1. Go to **sheets.new**. A blank Google Sheet opens.
2. **File → Import → Upload**, then pick `Business_Data_Workbook_GoogleSheets.xlsx` from the `crm/` folder of this repo.
3. Choose **Replace spreadsheet** and click **Import data**.
4. Rename it (e.g. "Craftaire — Business Data").
5. **File → Settings**: set **Locale = United States** and **Time zone = (GMT-05:00) Central Time – Chicago**. Click Save.
6. Open the **Team** tab. In the **Phone (WhatsApp)** column, enter each team member's WhatsApp numbers with the country code and no symbols, e.g. `13125550142`. Only numbers listed there can write to the sheet.

Everything from the Excel version is here: all 22 tabs, formulas, dropdowns, color rules and charts, plus a new **WhatsApp Log** tab. From now on, use the Google Sheet and stop using the Excel file, so you only have one copy of your data.

## Step 2: Add the bot code to the sheet (10 min)

1. In the Google Sheet: **Extensions → Apps Script**.
2. Delete what's in `Code.gs`, paste in the entire contents of `agent/Code.gs`, and click 💾 Save.
3. Choose `makeKey` in the function dropdown at the top and click **Run**. Google will ask for permission: click *Review permissions*, choose your account, then **Advanced → Go to (your project)**, then **Allow**. You'll see that "Google hasn't verified this app" warning because it's your own script. Copy the long code it prints in the log. That is your **RELAY_KEY**.
4. **Project Settings (⚙️) → Script Properties → Add script property**. Add these (you'll get the WhatsApp values in step 3; add them when you have them):

| Property | Value |
|---|---|
| `ANTHROPIC_API_KEY` | from console.anthropic.com → API Keys |
| `WHATSAPP_TOKEN` | the permanent token from step 3 |
| `WHATSAPP_PHONE_NUMBER_ID` | from step 3 |
| `RELAY_KEY` | the code from `makeKey` |

   Optional: `CLAUDE_MODEL` (default `claude-opus-5`), `CLAUDE_EFFORT` (default `medium`), `GRAPH_VERSION` (default `v23.0`).
5. **Deploy → New deployment → ⚙️ → Web app**. Set *Execute as*: **Me**, and *Who has access*: **Anyone**. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).
   - "Anyone" is required so Cloudflare can reach it. Requests without your RELAY_KEY are rejected.
   - If you change the code later: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same.
6. Choose `setup` in the dropdown and click **Run**. The log lists anything still missing.
7. Once `ANTHROPIC_API_KEY` is set, run `testMessage`. It sends a sample message to Claude and logs the reply, without writing anything or using WhatsApp. This is your first live check that Claude is connected.

## Step 3: Create the WhatsApp number in Meta (20–30 min)

1. Go to **developers.facebook.com** → My Apps → **Create App**. Pick the WhatsApp use case ("Connect with customers through WhatsApp"), then create or choose a business portfolio.
2. In the app, open **WhatsApp → API Setup**.
   - Meta gives you a free **test number** straight away. It's good for trying everything out. Add each team member's numbers under "To" and verify them.
   - To go live, click **Add phone number** and register the bot's own number.
   - Copy the **Phone number ID** into the `WHATSAPP_PHONE_NUMBER_ID` Script Property.
3. **Permanent token.** The token on the API Setup page expires in 24 hours, so use a system user token instead:
   **business.facebook.com → Settings → Users → System users → Add** (role: Admin) → **Assign assets**: your app (full control) and your WhatsApp account → **Generate token**. Select the app, set expiry to **Never**, and tick `whatsapp_business_messaging` and `whatsapp_business_management`. Paste it into `WHATSAPP_TOKEN`.
4. Copy **App settings → Basic → App secret**. You'll need it for Cloudflare.

## Step 4: Set up the relay on Cloudflare (10 min)

Meta has to reach a web address that answers instantly. Google Apps Script can't do that on its own, so this small free Worker sits in between and checks that every message really came from Meta.

1. **dash.cloudflare.com** → sign up (free) → **Workers & Pages → Create → Create Worker**. Name it (e.g. `craftaire-whatsapp`) and click **Deploy**.
2. **Edit code**. Replace everything with the contents of `relay/worker.js` and click **Deploy**.
3. **Settings → Variables and Secrets → Add** (type **Secret**) for each of these:

| Name | Value |
|---|---|
| `VERIFY_TOKEN` | any phrase you make up, e.g. `level-true-verify-2026` |
| `APP_SECRET` | the Meta App secret from step 3.4 |
| `APPS_SCRIPT_URL` | the Web app URL from step 2.5 |
| `RELAY_KEY` | the same RELAY_KEY as in Apps Script |

4. Copy the Worker's address, e.g. `https://craftaire-whatsapp.yourname.workers.dev`.

## Step 5: Connect Meta to the relay (5 min)

1. In the Meta app: **WhatsApp → Configuration → Webhook → Edit**.
   - **Callback URL**: the Worker address from step 4.4.
   - **Verify token**: the same phrase you used for `VERIFY_TOKEN`.
   - Click **Verify and save**.
2. Under **Webhook fields**, subscribe to **messages**.
3. If messages from real numbers don't arrive, set the app to **Live** (App Mode toggle at the top). Meta requires a privacy-policy URL for that; the website can host a simple one.

## Step 6: Try it

Text the bot number from your WhatsApp:
- `help`
- `Me and Sam 8 to 4:30 on J-0001, install, 30 min lunch`
- a photo of a receipt with the caption `Johnson job`
- `stats`
- `undo`

Every message, and every cell the bot wrote, shows up on the **WhatsApp Log** tab.

---

## What the bot can and can't do

**Can:**
- **Add rows on any tab:** clients, properties, quotes, projects, change orders, hours, expenses (including reading receipt photos and PDFs), invoices and payments, team, vendors, callbacks, competitor prices.
- **Update rows that have an ID:** e.g. "Q-0012 won", "Johnson job done", "Got paid 4,600".
- **Answer questions from the data:** "How much did we make on the Johnson job?"
- **Ask you when a message is unclear** instead of guessing (e.g. which job?).
- **Adapt when you change the sheet:** it reads the column headers and dropdown lists live, so renaming a project type or adding a column doesn't break it.

**Won't:**
- write into gray (calculated) columns
- delete anything
- edit Time Log or Expenses rows that already exist (use UNDO, or fix them in the sheet)
- understand voice notes (type the message or send a photo instead)

**Safety:**
- Only numbers on the Team tab are accepted.
- Every change can be reversed with UNDO.
- Message text can never become a formula.
- Google Sheets keeps full version history (File → Version history).

## Running costs (rough)

| Piece | Cost |
|---|---|
| Google Sheets and Apps Script | Free |
| Cloudflare Worker | Free (100,000 requests a day) |
| WhatsApp | Free when you text the bot and it replies within 24 hours (a "service conversation" in Meta's pricing). Check Meta's current pricing page for your country. |
| Claude API | About 2–5¢ per message on the default model (`claude-opus-5`); roughly $10–30 a month at 10–20 messages a day. Setting `CLAUDE_MODEL` to `claude-sonnet-5` costs about 60% less, with slightly weaker reading of messy receipts. |

## If something isn't working

| Symptom | Check |
|---|---|
| No reply at all | Cloudflare Worker → Logs. Meta → WhatsApp → Configuration shows webhook errors. Apps Script → Executions shows every run and its errors. |
| "This number isn't on the Team tab" | Put the number in the Team tab's Phone (WhatsApp) column, digits only with country code. |
| "Claude API error 401" | `ANTHROPIC_API_KEY` is wrong. |
| Replies stopped after a day | You're using the 24-hour token. Make the permanent system user token (step 3.3). |
| You changed Code.gs and nothing happened | Deploy a **new version** (step 2.5). |
