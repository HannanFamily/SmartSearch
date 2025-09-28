# Rubberduck VBA setup (dev QA)

This repo uses Rubberduck VBA to catch issues early and to run unit tests.

Steps
- Install Rubberduck from https://rubberduckvba.com/
- Excel Options > Trust Center > Trust access to the VBA project object model: Enabled
- Open the workbook (Search Dashboard v1.2 STABLE.xlsm or a v1.3 Dev copy)
- In VBE: Rubberduck > Refresh, then Rubberduck > Inspections
- Fix any compile/inspections flagged as Critical (missing references, ambiguous names, unused duplicates)
- Optional: Rubberduck > Unit Tests to run tests once available

References
- Ensure Microsoft Forms 2.0 Object Library is referenced for dynamic UserForm scenarios

Troubleshooting
- If inspections stall, close VBE and retry
- If references are missing, open Tools > References and resolve broken links, then Refresh