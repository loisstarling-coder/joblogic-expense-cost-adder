# JobLogic – Add expense costs to jobs

A Tampermonkey userscript for JobLogic (`https://go.joblogic.com/*`).

Upload a Payhawk expense export and it checks each line against its JobLogic
job to see whether the cost is already there. Anything missing can then be
added as a material line, but only when you click **Add** and confirm.

## Install

1. Install [Tampermonkey](https://www.tampermonkey.net/) in Chrome.
2. Open this link and click **Install**:

   **https://raw.githubusercontent.com/loisstarling-coder/joblogic-expense-cost-adder/main/joblogic-expense-cost-adder.user.js**

3. Go to JobLogic and refresh. Click **Advanced Controls** on the right, then
   **➕ Add expense costs to jobs**.

If the button doesn't appear, go to `chrome://extensions` → Tampermonkey →
**Details** and turn on **Allow User Scripts**.

No GitHub account is needed. Tampermonkey checks this link for new versions
and updates the script automatically.

## Use

1. Click **Choose Excel file** and pick the Payhawk export (`.xlsx`), or copy
   the rows from Google Sheets (including the header row) and paste them into
   the box.
2. Click **Check against jobs**. This only reads JobLogic; nothing changes.
3. Each row gets a status:

   | Status | What happens |
   |---|---|
   | Already in job | Left alone |
   | Not in job / No costs | Ticked, ready to add |
   | Possible match | A similar line exists. Unticked; tick it to add anyway |
   | Incorrect | On the job but entered wrong (e.g. with VAT). Fix it by hand |
   | On undelivered PO | Deliver the PO instead |
   | No reference / Job not found | No usable job number on the row |

4. Untick anything you don't want added (e.g. £0.01 card authorisations).
5. Click **Add X ticked to jobs** and confirm. Each line is added as
   cost = Net, qty 1, 20% VAT, with the job's normal uplift. PROJ jobs and
   quoted jobs are added as non-chargeable. Each job is re-checked just
   before adding, so nothing is added twice.
6. **Copy results** copies a report you can paste into a sheet.

Job references recognised: RE…, SRE…, PROJ…, PM…, 8-digit Salesforce numbers
starting with 1, and quote numbers UP0… (which use the job the quote became).

Tip: try a small file (5–10 rows) first and check those jobs in JobLogic.
