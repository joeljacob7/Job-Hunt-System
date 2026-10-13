# Job-Hunt-System

Local, manual UAE job discovery and application-preparation MVP. It uses free public ATS endpoints, stores results in a CSV, and never submits applications or sends messages.

## Run

From this folder in PowerShell:

```powershell
py -m unittest discover -s tests -v
py .\job_hunt.py discover --dry-run
py .\job_hunt.py discover
py .\job_hunt.py list
py .\job_hunt.py draft JOB_ID
```

`discover --dry-run` fetches and scores listings but leaves the tracker unchanged. Normal discovery updates `Applications/jobs.csv`; repeat sightings retain the existing `status`, `next_action`, and `date_found`. Edit those columns in the CSV as needed. Runs are logged to `Logs/job_hunt.log`.

Normal discovery creates separate Markdown resume and outreach drafts for jobs at or above the configured review threshold. Lower-scoring jobs can be drafted on request with `draft JOB_ID`. Drafts are recorded in the CSV and are not regenerated on repeat sightings. They validate and read the approved DOCX master, use resume-backed facts, and are saved under `Outreach/Drafts` with unique filenames. Nothing is sent or submitted.

## Scope and behavior

- The only enabled source is Careem's public Greenhouse job-board endpoint, verified on 2026-10-13. Add other company board IDs in `job_hunt_config.json` only after independently verifying each public endpoint. This is narrow coverage, not the whole UAE market.
- UAE locations are required. Results are ordered Dubai, Sharjah, Abu Dhabi, then other UAE locations.
- The salary floor is AED 6,500 per month. Explicit AED annual values are divided by 12. Foreign currency is not converted because no reliable live exchange source is configured. Unknown or ambiguous pay remains flagged for human review; known ranges entirely below target are excluded.
- Rule-based match scores use target role family, intersections between job requirements and resume-backed evidence groups, and explicit experience, education, and tool requirements. It is a review aid, not a hiring prediction. Senior/leadership titles and experience requirements above five years are excluded.
- No schedule is configured. Network calls are public HTTPS GET requests to Greenhouse or Lever endpoints. There are no credentials, paid APIs, or third-party Python dependencies.