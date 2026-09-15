# Investment Tracker

A simple multi-client investment tracker. Google Sheets is the database,
a Google Apps Script Web App is the API, and this static site (hosted on
GitHub Pages) is the frontend.

- **Advisor** logs in with an Admin Key, picks a client, and adds entries
  (Investment / Withdrawal / Current Value).
- **Each client** logs in with their Client ID + personal Access Code and
  sees only their own entries and totals (Total Invested, Withdrawn, Net,
  Current Value, Gain/Loss).

## 1. Open the Google Sheet

A blank sheet named **Investment Tracker** has already been created for
you in Drive. Open it, then follow step 2 — the tabs and headers get
created automatically, you don't need to type them by hand.

## 2. Deploy the Apps Script Web App

1. In the Sheet: **Extensions > Apps Script**.
2. Delete the default code, paste in [`apps-script/Code.gs`](apps-script/Code.gs).
3. Change `ADMIN_KEY` at the top to a long random string only you know —
   this is your advisor password, so keep it private.
4. In the toolbar, use the function dropdown (next to the bug icon) to
   select **setup**, then click **Run** (▶). The first time, Google will
   ask you to authorize the script — click through **Advanced > Go to
   (project name) > Allow**. This creates the `Clients` and `Entries`
   tabs in your sheet with headers and one sample client row (`C001`).
5. **Deploy > New deployment**, type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click Deploy, authorize again if asked, and copy the **Web app URL**
   (looks like `https://script.google.com/macros/s/XXXX/exec`).

Redeploy (**Deploy > Manage deployments > Edit > New version**) any time
you change the script.

### Adding real clients

In the `Clients` tab, add one row per client: a `ClientID` you make up
(e.g. `C002`), their name, and an `AccessCode` (a password you choose for
them). Share the `ClientID` + `AccessCode` with that client directly —
never share the sheet itself. Delete or edit the sample `C001` row once
you have real clients.

## 3. Point the site at your Web App

Open [`js/app.js`](js/app.js) and set:

```js
const WEB_APP_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

## 4. GitHub Pages

Already live at **https://rajeshpatidar1803.github.io/investment-tracker/**.
Any time you change a file, commit and push and the live site updates
within a minute or two:

```bash
git add .
git commit -m "your message"
git push
```

## Security notes

This is a lightweight tool, not a bank-grade system:

- Anyone with the site URL, a Client ID, and that client's Access Code can
  view that client's data. Treat Access Codes like passwords — don't reuse
  them, and rotate a client's code (edit the `Clients` tab) if you think
  it leaked.
- The Admin Key can add entries for any client. Keep it private and don't
  commit it into the repo (it lives only in the Apps Script project, not
  in this codebase).
- The Apps Script Web App is deployed as "Execute as: Me, Anyone" so the
  static GitHub Pages site can call it without a server. It only exposes
  the `login` and `addEntry` actions defined in `Code.gs` — nothing else
  in your Google account is reachable through it.
