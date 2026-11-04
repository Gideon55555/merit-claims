# Merit Properties – Customer Claim System

- **Staff app:** `index.html` (login: register customers, records, users, agents)
- **Sales check page:** `check.html` (no login: is a number free or claimed?)
- **Backend:** Google Apps Script bound to the claims Google Sheet. Source in `apps-script/`.

`config.js` holds the Apps Script web app URL (ends in `/exec`). After changing the Apps Script code,
use *Deploy → Manage deployments → Edit → New version* so the URL stays the same.

All data stays in the Google Sheet. This repository holds only the web pages.
