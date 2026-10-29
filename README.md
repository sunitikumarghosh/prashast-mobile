# Prashast On boarding status :: SSM, Purba Bardhaman

A web portal for monitoring PRASHAST onboarding by block, school and day. It has two modules,
switched at the top of the side menu:

- **Teacher Onboarding**: teachers onboarded per school (Total Teachers / Teachers Onboarded).
- **School Onboarding**: whether each school has completed onboarding (Status: Completed / Pending).
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
| Upload / Manage Data | Upload daily files for either module, delete a day, back up and restore |

The **School Onboarding** module has the same pages, using schools instead of teachers:

| Page | What it shows |
|---|---|
| Dashboard | Total / onboarded / pending schools and onboarded % with change since the previous day, highlights, status chart, blocks by onboarding level, daily trend, top/bottom 5 blocks |
| Block-wise Status | Rank, total schools, schools onboarded, pending, %, daily change, Leading / Moderate / Lagging |
| Leading & Lagging | Top 10 / bottom 10 blocks, top gainers, least progress, most pending schools |
| All / Onboarded / Pending Schools | School lists filtered by block, status and search, with the status on the previous date |
| Day-wise Comparison | Block-wise change, plus schools newly onboarded, still pending, completed on both dates, reverted to pending |
| Daily Trend | District summary for every date, plus a block × date matrix (%, onboarded, pending) |

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
2. Open the portal and go to **Upload / Manage Data**, then upload the file(s). The module (teacher or school) is detected from the columns.
   The date is read from the file name. If the name has no date, choose one in the *Data date* box.
3. Uploading again for the same date replaces that day's data.

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

The portal's own format `School Name | UDISE` in one column also works. Percentages are recalculated from the counts.

## Local testing

Open the folder through any web server, for example XAMPP: `http://localhost/prashast/prashast-monitor/`.
Opening `index.html` directly from disk also works, but the `data/manifest.json` loading needs a web server.
