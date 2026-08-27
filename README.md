# AKUMA Food Registration — GitHub + Render + Google Apps Script

## 1. Google Apps Script backend

1. Existing Apps Script project: replace `Code.gs` and `Total.gs`.
2. Run `setupSheets()` once.
3. Project Settings → Script Properties: add `ADMIN_PASSWORD` with your admin password.
4. Deploy → New deployment → Web app.
5. Execute as: Me. Access: Anyone with the link.
6. Copy the `/exec` URL.

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
- Every login, registration, rejection and admin action is saved in `Logs`.
