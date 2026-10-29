# Prashast On boarding status :: SSM, Purba Bardhaman

A web portal for monitoring PRASHAST onboarding by block, school and day. It has two modules,
switched at the top of the side menu:

- **Teacher Onboarding**: teachers onboarded per school (Total Teachers / Teachers Onboarded).
- **School Onboarding**: whether each school has completed onboarding (Status: Completed / Pending).
- **Screening (Part-1)**: per school, onboarding, whether screening has started, total students and students screened.
It is a static site: no server or database is needed, and it runs on **GitHub Pages**.

## Features

| Page | What it shows |
|---|---|
| Dashboard | District KPIs with change since the previous day, key insights, ranked block chart, school category chart, daily trend, top/bottom 5 blocks |
| Block-wise Status | Rank, schools, teachers, onboarded, pending, %, daily change, 100% / 1-99% / 0% school counts, Leading / Moderate / Lagging status. Tap a block to open its schools |
| Leading & Lagging | Top 10 and bottom 10 blocks, top gainers, least progress, most pending teachers, most 0% schools, priority schools, large 0% schools, best improving schools |
| All / 100% / 1-99% / 0% Schools | School lists filtered by block, category and search, with change since the previous day |
| Day-wise Comparison | Any two dates side by side: block-wise change, plus schools that improved, became 100%, moved out of 0%, stayed at 0% or decreased |
| Daily Trend | District summary for every date, plus a block × date matrix (%, onboarded, 0% schools, 100% schools) |
| Management / Category | Teacher onboarding by school Management, Category or School Type (optionally for one block), with daily change, 100% / 1-99% / 0% school counts, a chart, and a block × group matrix |
| Upload / Manage Data | Upload daily files for either module or the school master, delete a day, back up and restore |

The **School Onboarding** module has the same pages, using schools instead of teachers:

| Page | What it shows |
|---|---|
| Dashboard | Total / onboarded / pending schools and onboarded % with change since the previous day, highlights, status chart, blocks by onboarding level, daily trend, top/bottom 5 blocks |
| Block-wise Status | Rank, total schools, schools onboarded, pending, %, daily change, Leading / Moderate / Lagging |
| Leading & Lagging | Top 10 / bottom 10 blocks, top gainers, least progress, most pending schools |
| All / Onboarded / Pending Schools | School lists filtered by block, status and search, with the status on the previous date |
| Day-wise Comparison | Block-wise change, plus schools newly onboarded, still pending, completed on both dates, reverted to pending |
| Daily Trend | District summary for every date, plus a block × date matrix (%, onboarded, pending) |
| Management / Category | School onboarding by Management, Category or School Type, with a block × group matrix |

The **Screening (Part-1)** module has the same pages. Its main measure is the % of students screened. Each school is also put in a stage:
*Not onboarded* → *Onboarded, not started* → *Screening started* → *All students screened*.

| Page | What it shows |
|---|---|
| Dashboard | Students screened %, total / screened students, schools onboarded and screening started, schools per stage, highlights, charts, trend, top/bottom 5 blocks |
| Block-wise Status | Schools, onboarded, screening started, students, screened, not screened, %, daily change, Leading / Moderate / Lagging |
| Leading & Lagging | Top / bottom 10 blocks, top gainers, least progress, most onboarded schools not started, most students not screened, large schools not started / not onboarded |
| School lists | All / not onboarded / onboarded but not started / screening started / all students screened, with block, Management, Category and Type filters |
| Day-wise Comparison | Block-wise change, plus schools where screening newly started, newly onboarded, more students screened, still not started, still not onboarded |
| Daily Trend | District summary for every date and a block × date matrix |
| Management / Category | Screening by Management, Category or School Type, with a block × group matrix |

## School master (Management / Category / Type)

The Management / Category reports, the Management and Category columns in school lists, and the matching filters
use the **School Search** export (`UDISE Code`, `Management`, `Category`, `Type`), matched by UDISE code.

- For everyone: save it as `data/school_master.csv` in the repository (the name is fixed, it is not listed in the manifest).
- For one browser: upload the CSV on **Upload / Manage Data**. It overrides the repo copy in that browser.

The Upload page shows how many schools of the latest day are missing from the master; those appear as *Not in school master*.

Every table has **Excel** and **PDF** download buttons. Each file carries the header
*"Prashast On boarding status :: SSM, Purba Bardhaman"*, the report title, the data date and the generation time.
The layout works on phones, tablets and desktops.

**Leading / Lagging rule:** blocks are ranked by onboarded %. The top quarter are *Leading*, the bottom quarter are *Lagging*, and the rest are *Moderate*.

## Daily use

1. On the PRASHAST portal, run the Tampermonkey scripts:
   - `prashast-teacher-onboarding-all-blocks.user.js` on *Teacher Onboarding | Block Wise Report*. It downloads
     `Teacher_Onboarding_School_Wise_All_Blocks_YYYY-MM-DD.xlsx`.
   - `prashast-schools-onboarding-all-blocks.user.js` on *Schools Onboarding Status | Block Wise Report*. It downloads
     `Schools_Onboarding_Status_All_Blocks_YYYY-MM-DD.xlsx`.
   - `prashast-screening-part1-all-blocks.user.js` on *Schools Screening Status [Part-1] | Block Wise Report*. It downloads
     `Schools_Screening_Status_Part1_All_Blocks_YYYY-MM-DD.xlsx`.
2. Open the portal and go to **Upload / Manage Data**, then upload the file(s). The module (teacher, school or screening) is detected from the columns.
   The date is read from the file name. If the name has no date, choose one in the *Data date* box.
3. Uploading again for the same date replaces that day's data.

**Upload / Manage Data is password protected.** After unlocking, it stays open until the browser tab is closed or *Lock* is pressed.
Only a SHA-256 hash of the password is in `js/app.js` (`ADMIN_HASH`). To change the password, put the hash of the new one there.
This is a static site, so the lock keeps casual visitors out of the upload page; it is not server-side security.

Uploaded data is stored **in that browser only** (IndexedDB). Use *Download backup* to move it to another device.

## Publish on GitHub Pages

1. Create a repository, for example `prashast-monitor`, and upload all the files in this folder:
   `index.html`, `css/`, `js/`, `data/`, `.nojekyll`.
2. Go to **Settings → Pages → Build and deployment**. Set *Source* to **Deploy from a branch**, choose branch `main` and folder `/ (root)`, then save.
3. After about a minute the portal is live at `https://<your-username>.github.io/prashast-monitor/`.

### Optional: let everyone see the same data

So that every officer who opens the link sees the full history without uploading anything, commit the daily files to the repo:

1. Put each day's Excel file in `data/`, keeping the date in its file name.
2. List the files in `data/manifest.json`:
   ```json
   [
     "Teacher_Onboarding_School_Wise_All_Blocks_2026-10-27.xlsx",
     "Teacher_Onboarding_School_Wise_All_Blocks_2026-10-28.xlsx",
     "Schools_Onboarding_Status_All_Blocks_2026-10-29.xlsx"
   ]
   ```
   Teacher and school files go in the same list.
3. Commit. These dates appear for everyone, marked "from repo". If someone uploads a file for the same date in their own browser, their upload is used instead.

> If the repository is public, anyone with the link can download files in `data/`. Make the repo private (GitHub Pages on private repos needs a paid plan) or keep data upload-only if that matters.

## Accepted file format

The first sheet that has a **School Name** column is used. The other columns are found by name:
- Teacher Onboarding: `Block Name`, `School Name`, `UDISE Code`, `Total Teachers`, `Teachers Onboarded`.
- School Onboarding: `Block Name`, `School Name`, `UDISE Code`, `Status` (`Completed` counts as onboarded, anything else as pending).
- Screening (Part-1): `Block Name`, `School Name`, `UDISE Code`, `Schools Onboarded` (`On-boarded` / `Pending`), `Screening Started` (`Not Started` = not started), `Total Students`, `Screened Students`.

The portal's own format `School Name | UDISE` in one column also works. Percentages are recalculated from the counts.

## Local testing

Open the folder through any web server, for example XAMPP: `http://localhost/prashast/prashast-monitor/`.
Opening `index.html` directly from disk also works, but the `data/manifest.json` loading needs a web server.
