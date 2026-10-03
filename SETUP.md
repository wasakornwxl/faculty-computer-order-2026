# Setting up the order site with Google Sheets

This takes about 15 minutes. You need a Google account. The account you use becomes the owner of the order spreadsheet.

## 1. Create the spreadsheet and script

1. Go to [sheets.new](https://sheets.new) and name the spreadsheet, for example **Faculty Computer Orders 2026**.
2. Open **Extensions → Apps Script**.
3. Delete the sample code, paste in everything from `apps-script/Code.gs`, and click **Save**.
4. In the function menu at the top, choose **setup** and click **Run**.
   - Google asks you to authorise the script the first time. Choose your account, click **Advanced → Go to (project name)**, then **Allow**. The warning appears because you wrote the script yourself, not a verified app.
   - Two tabs appear in the spreadsheet: **Computers** and **Tablets**.

## 2. Publish the script as a web app

1. In Apps Script, click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy** and copy the **Web app URL**. It ends in `/exec`.

"Anyone" lets faculty submit without a Google account. They can only send an order to the script. They can't open or read the spreadsheet.

## 3. Connect the website

1. Open `config.js` and paste the URL between the quotes after `endpoint:`.
2. Check `allowedDomains`. Only emails ending in these domains can submit. Use the same list as `ALLOWED_DOMAINS` in `Code.gs`.

## 4. Put the website online

Upload the whole `order-site-sheets` folder (`index.html`, `config.js`, `img/`) to any web host:

- **University web server:** ask IT to host the folder, for example at `https://…mahidol.ac.th/computer-order/`.
- **GitHub Pages (free):** create a repository, upload the files, then turn on **Settings → Pages → Deploy from branch**.
- **Netlify Drop (free, fastest):** drag the folder onto [app.netlify.com/drop](https://app.netlify.com/drop).

The `apps-script` folder and this file don't need to be uploaded.

## 5. Test before sending the link

1. Open the site, choose a computer and submit with your own university email. A row should appear on the **Computers** tab.
2. Submit a tablet with the same email. The row should move from **Computers** to **Tablets**.
3. Delete your test row.

## How the data works

- **One row per person:** each email has one row across both tabs. A new submission replaces the old one, and moving from computer to tablet moves the row between tabs.
- **Columns are matched by header name:** you can reorder or delete columns, or add your own, such as "PO number" or "Delivered". Your own columns are kept when someone updates their order. Don't rename the built-in headers. Running `setup` again only adds missing headers.
- **Prices are checked on the server:** the script works out the price and top-up from its own price list, so a tampered page can't change them.
- **Typed text is made safe:** text that starts with `=`, `+`, `-` or `@` is stored as plain text, not as a formula.
- **Returning visitors:** the page remembers what that browser last submitted. On a different device, the person just submits again and their row is replaced.

## Optional settings in Code.gs

- `NOTIFY_EMAIL`: an address that gets an email for every submission.
- `CLOSE_DATE`: a date (`"2026-10-31"`) after which submissions are refused.

After changing `Code.gs`, click **Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy**. The URL stays the same.

## Maintenance functions

Run these from the Apps Script editor: choose the function name in the menu at the top, then click **Run**.

- `markDiscontinued`: highlights rows that still name a product that's no longer offered, and explains why in the **Check** column.
- `migrateP5toB5`: changes every ASUS P5 order to the ASUS ExpertBook B5 14″, updating price, quote and item ID. Name, email, notes and your own columns are kept, and the change is noted in **Check**. Running it again does nothing once there are no P5 rows left.

## If prices or products change

Update both files:
- `index.html`: the `CATALOG` (computers) and `TABLETS` lists near the top of the script.
- `apps-script/Code.gs`: the `COMPUTERS` and `TABLETS` lists. Then redeploy as above.

## Limits to know

- **No login:** people are identified only by the email they type. Someone who knows a colleague's email could replace that colleague's row. The **Updated** column shows when each row last changed, and each row keeps its **First submitted** time.
- **Apps Script quotas:** a free Google account can handle roughly 20,000 submissions a day, far more than you need.
