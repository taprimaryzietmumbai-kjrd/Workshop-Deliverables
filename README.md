# Workshop Deliverables & Classroom Impact Hub — ZIET Mumbai

Banner-led responsive participant and administrator portal for workshop preparation, ongoing contributions, classroom implementation, Action Research and follow-up CPD.

## Repository files
- `index.html` — responsive interface matching the approved three-phase infographic banner.
- `Code.gs` — Google Apps Script backend for Sheets, Drive folders, participant roster import, submissions, monitoring and A4 PDF workbook generation.

## Configured Google resources
- Master spreadsheet: `1y8BQ3UsRE7xlHQFPyRVops7vOh8n7ZOvsEuI4sLXMZQ`
- Drive repository folder: `1pFe84ogk9ihbPDTmwscOVPHvIyUcYC7q`
- Nameplates source: configured to scan the master spreadsheet for a sheet whose first row contains recognizable Name and Vidyalaya/School headers. The Admin panel also accepts an explicit tab name.

## Install in the existing Apps Script project
Open the Apps Script project supplied by the administrator:
https://script.google.com/u/0/home/projects/1WzLjKwi2HTtKiRZmA6ZHpshL5YAtf3zJ52STs4r15W-r4HwO8VbZcuP0/edit

1. Add or update the server-side script using the complete contents of `Code.gs` from this repository. If the project already has an existing `doGet` or `doPost`, reconcile those handlers rather than leaving duplicate function names.
2. In Apps Script, click **+** beside Files → **HTML**, name the file exactly `index` (Apps Script adds the .html extension), and paste the complete contents of this repository's `index.html`.
3. Save all files.
4. Select `initialiseWorkshopImpactHub` in the function selector and click **Run**. Review the requested Google Sheets, Drive and Docs permissions. This creates only missing data tabs and required repository subfolders; existing sheet data is retained.
5. In **Deploy → New deployment → Web app**, configure the deployment deliberately for your organisation's identity and access requirements. Do not expose this app publicly until participant identity checks and administrator authorization have been configured and tested.
6. Open the deployed Web App URL. The HTML is served by Apps Script and uses `google.script.run` for same-origin calls; this avoids relying on cross-origin fetch from GitHub Pages.

## First-time administration
1. Open the Web App and navigate to **ZIET Admin**.
2. Create a workshop with title, batch/region and dates.
3. Choose the Nameplates tab (or leave blank for auto-detection), then click **Import Nameplates Participants**.
4. Participants can then choose their workshop and name, complete their profile and submit deliverables.
5. Generate an A4 PDF workbook from the participant workspace.

## Data model created by setup
`Workshops`, `Participants`, `Submissions`, `Daily_Proceedings`, `Classroom_Impact`, `Action_Research`, `Monitoring`, `Audit_Log`.

## Important release checks
This repository contains the source implementation, not a deployment of the Apps Script project. The supplied Apps Script project must receive the two source files and be deployed by an authorized account. Before live participant use, configure and test sign-in / participant identity verification, administrator allow-listing, upload handling, consent for any photo/GPS evidence, and access permissions for personal data (including DOB, social category and PwD status). The current prototype does not yet enforce an administrator allow-list or verify participant identity; its Admin menu is therefore not a security boundary. The current evidence fields accept Drive/document links; direct browser camera/GPS capture and direct file uploads are not yet implemented.
