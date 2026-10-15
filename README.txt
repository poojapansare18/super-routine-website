SUPER ROUTINE — ORIGINAL DESIGN + GOOGLE SHEETS

This project retains the original Claude-style design, copy, product details, pill visuals and multi-step waitlist form. It includes the expanded research questions, Instagram/WhatsApp links and a count starting at an illustrative baseline of 357 plus unique saved registrations.

IMPORTANT: The website uses a same-origin Vercel serverless function at /api/waitlist to communicate with Google Apps Script. Do not put names, phone numbers or questionnaire answers into a GET/JSONP URL.

SETUP

1. Open the existing Google Sheet. The spreadsheet ID is configured in GoogleAppsScript.gs.
2. Open Extensions > Apps Script. Back up existing code, replace it with the current GoogleAppsScript.gs, then Save.
3. Select setupSheet and click Run. Approve the permissions. Confirm that Founder Dashboard appears and Responses retains existing records.
4. In Apps Script, choose Deploy > Manage deployments > Edit > New version > Deploy (Web app). Set Execute as: Me. Access must permit the public website to submit responses. Copy the deployed URL ending in /exec.
5. Open Vercel Dashboard > the super-routine-website project > Settings > Environment Variables. Add APPS_SCRIPT_URL and set its value to the new deployed /exec URL. Select Production (and Preview if you want to test preview deployments). Save.
6. In Vercel, open Deployments and redeploy the latest production deployment after adding the environment variable. Pushing a new commit to the connected main branch can also trigger deployment.
7. Open https://super-routine-website.vercel.app/api/waitlist in a new tab. It should return JSON containing ok:true, count and displayNumber. If it says APPS_SCRIPT_URL is not configured, revisit step 5. If Apps Script does not return JSON, check the /exec URL and deployment access.
8. Test a new WhatsApp number through the website. Confirm one row is written to Responses and the counter increases. Submit the same number again; it must be rejected without adding a row.

DATA NOTES

- 357 is an illustrative baseline, not a count of real members. Display count = 357 + unique WhatsApp numbers saved in Responses.
- Historical duplicate rows are not automatically deleted. Review them before using the records for analysis.
- Instagram: https://www.instagram.com/super.routine/
- WhatsApp contact: https://wa.me/919529966963. Confirm this number is the intended public project contact before publishing.
- Product and ingredient copy/visuals are concept content provided for this site; the project does not independently verify efficacy, formulation, regulatory approval or product availability.
- Never commit private credentials or access tokens into this public repository.
