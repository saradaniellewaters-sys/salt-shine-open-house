# Salt & Shine Co. Open House QR System — Google Sheets Version

This version posts every guest sign-in to Google Sheets using a Google Apps Script Web App. It also keeps the local browser CRM and CSV export as a backup.

## Files

- `index.html` — your live open house QR website
- `google-apps-script.js` — paste this into Google Apps Script attached to your Google Sheet

## Set up Google Sheets lead capture

1. Create a new Google Sheet, for example: `Salt & Shine Leads`.
2. In the Sheet, go to **Extensions → Apps Script**.
3. Delete the starter code and paste everything from `google-apps-script.js`.
4. Optional but recommended: set `REQUIRED_SECRET` near the top of the script, for example:

   ```js
   const REQUIRED_SECRET = 'saltshine2026';
   ```

5. Click **Save**.
6. In Apps Script, choose the `setupSheet` function and click **Run** once. Approve the permissions.
7. Click **Deploy → New deployment**.
8. Choose **Web app**.
9. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
10. Click **Deploy** and copy the Web App URL ending in `/exec`.
11. Open your deployed `index.html` site.
12. Paste the Web App URL into **Google Sheets Lead Delivery**.
13. If you used a secret in Apps Script, paste the same value into **Optional Lead Secret**.
14. Click **Send Test Lead** and confirm a row appears in the Sheet.
15. Fill in the property details, then download or copy your new QR links.

## Important live-use note

The generated QR URLs now include your property details and Google Sheets endpoint, so buyers who scan the QR from their own phones will see the correct guest form/property info and submissions will go to your Sheet.

## Deploying the website

The simplest hosting path is Netlify manual deploy:

1. Go to Netlify.
2. Add a new site.
3. Choose manual deploy.
4. Drag in this folder or ZIP.
5. Use the live URL to set up your open house and generate QR codes.

## Testing checklist

- Open the live site.
- Paste your Apps Script URL.
- Click **Send Test Lead**.
- Confirm `Test Lead` appears in Google Sheets.
- Fill out the guest form from your phone.
- Confirm the real lead appears in Google Sheets.
- Download the QR codes after your property details and Sheets URL are filled in.
