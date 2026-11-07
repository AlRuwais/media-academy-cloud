# Media Academy Cloud

**Team Al Ruwais** — Academy management platform starter.

## Architecture
- Frontend: HTML/CSS/JavaScript on GitHub Pages.
- Backend: Google Apps Script Web App.
- Database: private Google Sheets spreadsheet. Do not publish the spreadsheet.

## Included modules
Dashboard, users, groups, subjects, sessions, attendance, exams, results, learning content, announcements, complaints and audit log.

## Setup
1. Create a blank Google Spreadsheet and keep it private.
2. Open Extensions → Apps Script. Copy this repository's Code.gs into the Apps Script editor.
3. In Apps Script → Project Settings → Script Properties, add DATABASE_ID with the spreadsheet ID from its URL.
4. Save and run initializeDatabase from the function selector. Approve the Google permissions. The script creates the required tabs and headers.
5. Set up the first owner securely. Add two temporary Script Properties: INITIAL_OWNER_EMAIL and INITIAL_OWNER_PASSWORD. Then run createInitialOwner from the editor. The function creates the owner and deletes both temporary properties. Use a strong unique password.
6. Deploy → New deployment → Web app. Execute as Me; access Anyone. Copy the /exec URL.
7. Put the URL into config.js as window.MAC_API_URL.
8. In GitHub Settings → Pages, select GitHub Actions as the source.

## Important
This is an MVP foundation. Use dummy data first and test permissions before adding real student records. Keep the spreadsheet private. Never commit passwords, spreadsheet IDs, tokens, deployment URLs or student data to this public repository. Apps Script and Sheets have quotas and are not a high-scale database.
