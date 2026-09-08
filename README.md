# AKUMA Food Registration — GitHub + Render + Google Apps Script

## 1. Google Apps Script backend

1. Existing Apps Script project: replace `Code.gs` and `Total.gs`.
   Add `Graphic.gs` as a new Apps Script file.
2. Run `setupSheets()` once.
3. Project Settings → Script Properties: add `ADMIN_PASSWORD` with your admin password.
4. Add `SPREADSHEET_URL` with the full Google Sheets link, or add `SPREADSHEET_ID` with only the ID between `/d/` and `/edit`.
5. Deploy → New deployment → Web app.
6. Execute as: Me. Access: Anyone with the link.
7. Copy the `/exec` URL.

## 2. Static frontend

Open `config.js` and replace:

```js
API_URL: "PASTE_YOUR_GOOGLE_APPS_SCRIPT_EXEC_URL_HERE"
```

Upload the repository to GitHub. On Render create a **Static Site**:

- Build command: leave empty
- Publish directory: `.`

Employee page: `/index.html`

Admin dashboard: `/admin.html`

## Rules

- FitFood and Хоол: after 11:30 they become Pending for admin approval.
- ServiceFood: accepted after 11:30 too.
- Monthly calculation period is the 26th through the following month’s 25th.
- Admin total money: `(FitFood total + Hool total) × 15,000 + ServiceFood total amount`.
- Breakfast (`OglooniiTsai`) stays as its own existing sheet: one click = 1 and its sheet price remains 12,000₮ each. Registration closes at 10:00 (Asia/Ulaanbaatar). Breakfast is shown in count graphics but is not added again to `Total` money because that money is already included in ServiceFood.
- `Graphic` sheet is refreshed after every successful registration and can also be rebuilt by running `createGraphicSheet()`.
- The password-protected Admin Dashboard can manually add FitFood, ServiceFood, Hool, and Breakfast for any employee and date. Manual admin entries bypass normal time cutoffs and are recorded in `Logs` as `ADMIN_MANUAL_ADD`.
- Employee IDs are read from `Employees` and also from existing workbook sheets containing `Код`, `Овог`, `Нэр`, and `Албан тушаал` headers.
- Duplicate employee + date + food-type registrations are blocked on the server for both employee and admin entry, shown as a popup, and recorded in `Logs`.
- Every login, registration, rejection and admin action is saved in `Logs`.
