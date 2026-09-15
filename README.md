# Investment Tracker

A simple multi-client investment tracker. Google Sheets is the database,
a Google Apps Script Web App is the API, and this static site (hosted on
GitHub Pages) is the frontend.

- **Advisor** logs in with an Admin Key, picks a client, and adds entries
  (Investment / Withdrawal / Current Value).
- **Each client** logs in with their Client ID + personal Access Code and
  sees only their own entries and totals (Total Invested, Withdrawn, Net,
  Current Value, Gain/Loss).

## 1. Create the Google Sheet

Create a new Google Sheet with two tabs, headers exactly as below (row 1):

**Clients**
| ClientID | ClientName | AccessCode |
|---|---|---|
| C001 | Jane Doe | jane-secret-1 |

**Entries**
| ClientID | Date | Type | Amount | Notes | Timestamp |
|---|---|---|---|---|---|

Leave `Entries` empty besides the header row — entries are added from the
app. `Type` should be one of: `Investment`, `Withdrawal`, `Current Value`.

- `Investment` / `Withdrawal`: money in/out of the portfolio.
- `Current Value`: a snapshot of what the portfolio is worth today (used
  to calculate gain/loss vs. net invested). Add one whenever you want to
  update a client's current value.

Give each client their own `ClientID` and a private `AccessCode` — share
those with the client instead of the sheet itself.

## 2. Deploy the Apps Script Web App

1. In the Sheet: **Extensions > Apps Script**.
2. Delete the default code, paste in [`apps-script/Code.gs`](apps-script/Code.gs).
3. Change `ADMIN_KEY` at the top to a long random string only you know.
4. **Deploy > New deployment**, type **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click Deploy, authorize the script, and copy the **Web app URL**
   (looks like `https://script.google.com/macros/s/XXXX/exec`).

Redeploy (**Deploy > Manage deployments > Edit > New version**) any time
you change the script.

## 3. Point the site at your Web App

Open [`js/app.js`](js/app.js) and set:

```js
const WEB_APP_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

## 4. Host on GitHub Pages

```bash
git init
git add .
git commit -m "Investment tracker"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

Then in the GitHub repo: **Settings > Pages > Source: Deploy from branch >
main / (root)**. Your site will be live at
`https://<you>.github.io/<repo>/` within a minute or two.

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
