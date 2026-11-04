# Merit Properties – Customer Claim System: Setup

The system runs inside your existing claims Google Sheet (the one with the **Office Visit**, **Expecting** and **site visit** tabs).
Nothing is imported or copied. New registrations are written straight into those tabs.

You need: the Google account that owns the sheet, and the 3 files `Code.gs`, `Seed.gs`, `Index.html`.

---

## 1. Open the sheet as a Google Sheet
Open the claims sheet in Google Sheets. If it's an Excel file (.xlsx) in Drive, first do **File → Save as Google Sheets** and use the new copy.

## 2. Add the code
1. In the sheet: **Extensions → Apps Script**.
2. **Code.gs**: delete everything in the editor and paste the whole contents of `Code.gs`.
3. Click **+** next to *Files* → **Script** → name it `Seed` → paste the contents of `Seed.gs`.
4. Click **+** → **HTML** → name it exactly `Index` → paste the contents of `Index.html`.
5. Press **Ctrl + S** to save.

## 3. Run setup once
1. In the function dropdown at the top, choose **setup**, then click **Run**.
2. Google asks for permission. Click **Review permissions**, choose your account, then **Advanced → Go to (project) → Allow**.
   It shows a warning because you wrote the script yourself. That's normal.
3. Setup:
   - adds 4 columns to each category tab: *Agent Type, Supervisor, Entered By, Entered At* (existing data is not changed),
   - creates an **Agents** tab with 166 names from your Sales Data Base (Sales, Elite Sales, Supervisors, Freelancers, Agent companies),
   - creates a hidden **Users** tab and the first login: **admin / ChangeMe123**.

## 4. Publish the web app
1. **Deploy → New deployment** → click the gear icon → **Web app**.
2. **Execute as:** *Me*. **Who has access:** *Anyone*.
   (Anyone with the link only sees the login page. Staff don't need Google accounts.)
3. **Deploy**, then copy the **Web app URL**. That's the link you give to staff.

## 5. First login
1. Open the link and log in with **admin / ChangeMe123**. You'll be asked to set a new password.
2. Go to **Users → Add user** for each staff member:
   - **Data entry**: registers customers and checks numbers
   - **Viewer**: sees records and checks numbers
   - **Admin**: everything, including users and agents
   Each person sets their own password at first login.

---

## How it works
- **30-day rule:** a phone number registered in any category is locked for 30 days. On day 31 anyone can register it again. Elite Sales follow the same rule.
- **Matching:** numbers are compared on their **first 8 digits** (e.g. `0911 51 83`), so old records with hidden last digits still count. Spaces, `+251`, and a missing leading 0 are handled.
- **Supervisor:** filled in automatically from the Agents tab. Elite Sales and Agent companies have none. A supervisor registering for themselves is recorded as their own supervisor.
- **New rows** go right after the last customer row in the tab, replacing the pre-filled dates below it.
- **Old date mistakes** in the sheet are handled automatically:
  - dates saved with day and month swapped (e.g. *11 May* saved as *5 Nov*) are read correctly,
  - impossible future dates (e.g. *22/12/2035*) are ignored.

## Day-to-day
- **Add or remove agents:** in the app, go to **Agents** (admin). You can also edit the Agents tab directly. Set *Active* to FALSE to hide a name instead of deleting it.
- **Forgotten password:** an admin uses **Users → Reset password**. If the admin is locked out, open the sheet and use **Claims System → Reset admin password**. The menu appears after reloading the sheet.
- **Don't rename** the category tabs. Their names must keep *Expecting*, *Office* and *Site*. Don't delete their header rows.
- **Don't edit the hidden Users tab.** Share the Google Sheet only with admins: anyone with access to the sheet can see all records.

## Updating the code later
After pasting new code: **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The link stays the same.
