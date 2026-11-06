# MaaMadurai Clocking

Clock in / clock out app for MaaMadurai staff. It runs free on Google Apps Script, and every record is saved in a Google Sheet.

## What it does

**Staff (on their phone)**
- **Clock in**, **Start break**, **End break**, **Clock out**, with an optional note.
- GPS is checked at clock in and clock out. Within 150 m of the restaurant the shift is auto-approved. Further away, or with location off, the shift is still saved but goes to the manager to approve.
- Staff can see this week's shifts, hours and estimated pay.

**Manager**
- **Timesheet** tab: weekly view per staff member. Flagged shifts are highlighted. Tap a shift to fix times, approve or reject it, or add a forgotten shift. Overnight shifts work: if clock-out is earlier than clock-in, it counts as the next day.
- **Approve week**: saves totals to the *Timesheets* tab and emails each staff member their hours and pay. The manager gets a summary of all staff. In the 2nd week of a pay fortnight, the emails also show fortnight totals for payroll.
- **Staff** tab: add staff with name, email and hourly rate. Their sign-in link is emailed automatically. You can resend or reset links and mark people inactive.
- **Settings** tab: set the restaurant location (stand inside and tap *Use my current location*), the radius, and the pay fortnight start.

**Automatic**
- Every day at about 4am, shifts nobody clocked out of are flagged **Missing clock-out** for the manager to fix.
- Every Monday at about 7am, the manager is emailed last week's timesheet with a **Review & approve** button.

**Sign-in**: each person gets a personal link by email. They open it once and add it to their home screen. If they lose it, they enter their email on the app page and a new link is sent, but only to an email the manager registered. The manager can reset a link at any time, and the old one stops working.

## Google Sheet tabs (created automatically)

| Tab | Contents |
|---|---|
| Settings | Restaurant GPS, radius, currency, pay fortnight start |
| Staff | Name, email, **hourly rate**, role (Staff/Manager), active |
| Shifts | One row per shift: times, break minutes, hours, pay, GPS result, approval |
| Events | Every button press with raw GPS coordinates |
| Timesheets | Approved weekly totals per staff member |
| Audit log | Every manager change (old value → new value) |

---

## Setup (about 10 minutes, can be done on a phone)

> **On a phone:** open these pages in Chrome or Safari and switch to **desktop site**. Safari: **aA → Request Desktop Website**. Chrome: **⋮ → Desktop site**. The Apps Script editor doesn't work in phone view.

### 1. Create the Google Sheet
1. Go to **sheets.new** while signed in as **naveentamil@gmail.com**.
2. Name it **MaaMadurai Clocking**.

### 2. Add the code
1. In the sheet: **Extensions → Apps Script**.
2. Rename the project (top left) to **MaaMadurai Clocking**.
3. **Code.gs**: delete everything in it, then paste in all of [`Code.gs`](Code.gs). On GitHub, tap the file, then **Raw**, select all, and copy.
4. Click **＋ (Add a file) → HTML**, name it `Index` (it becomes `Index.html`), delete its contents, and paste in all of [`Index.html`](Index.html).
5. **⚙ Project Settings → Time zone → (GMT+10:00) Sydney**.
6. Click **💾 Save**.

### 3. Deploy as a web app
1. **Deploy → New deployment → ⚙ Select type → Web app**.
2. Set **Execute as: Me (naveentamil@gmail.com)** and **Who has access: Anyone**.
3. Click **Deploy**, then **Authorize access** and choose your Google account.
   - If you see *"Google hasn't verified this app"*: tap **Advanced → Go to MaaMadurai Clocking (unsafe) → Allow**. This is normal for your own script.
4. Copy the **Web app URL** (ends in `/exec`).

### 4. First run
1. Open the Web app URL. It sets up all the tabs and shows **Setup complete**.
2. Check **naveentamil@gmail.com** for the email *"your clock in / clock out link"* and open it. This is your manager link, so don't share it.
3. Add it to your home screen.

### 5. Set the restaurant location
1. **Stand inside the restaurant**, open your manager link, then go to **Settings → 📍 Use my current location** and allow location.
2. Leave the radius at 150 m, or change it.
3. Set **Pay fortnight starts** to the Monday your current pay fortnight began, then **Save settings**.

### 6. Add your 5 staff
**Staff → + Add staff member**: name, email and hourly rate (AUD). Each person gets an email with their link. Ask them to open it on their phone, tap **Clock in** once to test, allow location, and add it to their home screen.

---

## Updating the code later
After pasting new code: **Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy**. The URL and everyone's links stay the same.

## Troubleshooting
- **"Location: No location" every time**: the phone blocked location. iPhone: **Settings → Privacy → Location Services → Safari Websites → While Using**. Android: tap the lock icon in Chrome → **Permissions → Location → Allow**. The shift is still saved, but it goes to the manager to approve.
- **Lost manager link**: open the web app URL, enter your email and tap **Email me my link**. Or in Apps Script, select `resendMyManagerLink` and press **Run**.
- **Staff member left**: Staff → tap them → **Active: No**. Their link stops working.
- **Change someone's hourly rate**: Staff → tap them → edit the rate. It applies to new shifts. Past shifts keep the rate they were worked at.
- **Email limits**: free Gmail accounts can send about 100 emails a day, which is plenty for 5 staff.

## Notes
- Pay is **gross** (hours × hourly rate). It doesn't include tax, super, penalty or overtime rates.
- Break time between *Start break* and *End break* is unpaid.
