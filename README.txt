SUPER ROUTINE — ORIGINAL DESIGN + GOOGLE SHEETS

This package preserves the supplied original Claude-style design, copy, product details, pill visuals, and existing multi-step waitlist form. Added optional research questions (goal, pain point, intent, discovery channel), WhatsApp/Instagram links, and a counter with an explicitly illustrative baseline of 357 plus unique saved registrations.

SETUP
1. Open the existing Google Sheet. Its ID is configured in GoogleAppsScript.gs.
2. Open Extensions > Apps Script. Back up existing code, replace it with this package's GoogleAppsScript.gs, and Save.
3. Select setupSheet and click Run; approve permissions. This creates/refreshes Founder Dashboard and keeps the Responses tab.
4. Deploy > Manage deployments > Edit > New version > Deploy as Web app. Execute as Me. Access must allow website visitors to submit. Keep the deployed /exec URL.
5. In index.html, replace FORM_ENDPOINT and COUNT_ENDPOINT with that /exec URL if it changed. The current URL in the file is the previous deployment URL; do not assume it points to the newly deployed code.
6. Publish the website folder through Vercel. Preserve the assets directory.
7. Test a new phone number and then submit it again. The second submission must be rejected and must not add a row. Check Founder Dashboard and the live counter.

DATA NOTES
- The baseline 357 is illustrative, not real customers. The display count is 357 + unique WhatsApp numbers saved in Responses.
- Existing historical duplicate rows are not automatically deleted. Review the current rows before using the data for analysis.
- Instagram points to https://www.instagram.com/super.routine/ and WhatsApp contact to https://wa.me/919529966963. Confirm that this is the founder's correct public WhatsApp number before publishing.
- Product/ingredient copy and visuals are preserved as supplied concept content; this website does not verify efficacy, formulation, regulatory approval, or product availability.
