# Galgotias University Exam Seat Finder

A small static website that lets a student open the exam **seating plan spreadsheet** (.xls or .xlsx) and find their room and seat by **admission number** or **full name**.

> Independent student-developed utility. Not an official university website.

The spreadsheet is read **inside the browser** with JavaScript. It is never uploaded to a server, never saved in browser storage, and is cleared when the tab is closed or reloaded.

---

## 1. Project structure

```
exam-seat-finder/
├── index.html      Page layout and text
├── style.css       All styling (white page, navy and red accents, responsive)
├── script.js       Excel reading, column detection, search, results
├── vercel.json     Vercel security headers (blocks the page from sending data anywhere)
├── .vercelignore   Stops Vercel from ever uploading spreadsheet files
├── .gitignore      Stops Git/GitHub from ever storing spreadsheet files
└── README.md       This guide
```

`vercel.json` and `.vercelignore` are extra files added only for Vercel deployment. The site also works without them, but keep them: they are your safety net.

## 2. Create the project folder

1. Create a folder named `exam-seat-finder` anywhere on your computer (not inside Downloads with the spreadsheet).
2. Open VS Code → **File → Open Folder…** → choose `exam-seat-finder`.
3. Create each file listed above (**File → New File**, then save with the exact name, including the dot in `.gitignore` and `.vercelignore`).
4. Paste each file's code into the file with the same name and save.

> **Keep the seating plan spreadsheet OUTSIDE this folder.** The site does not need it in the folder; students choose the file in the browser.

## 3. Run locally with Live Server

1. In VS Code, open **Extensions** (Ctrl+Shift+X), search for **Live Server** (by Ritwick Dey), and install it.
2. Right-click `index.html` → **Open with Live Server**.
3. The browser opens something like `http://127.0.0.1:5500/index.html`.

Why not double-click `index.html`? Opening files as `file://` behaves differently in some browsers. A local server behaves like the real website. An internet connection is needed because the Excel reader (SheetJS) and fonts load from their CDNs.

## 4. Test with the university spreadsheet

1. On the page, click **Choose an Excel file** (or drag the file onto the box) and select the seating plan.
2. You should see a green message such as *"Seating plan ready. 30,311 exam entries loaded."* For the 23 Oct 2026 file, all 4 shift worksheets are detected automatically (header on row 6: Admission No., Enrollment No., Name, Program, Sem, Sec, Sub Code, Room No., Seat No.).
3. Test these cases:
   - **Admission number**: type a known admission number in lower case with extra spaces. It should still match. If the student has exams in more than one shift, one card appears per shift.
   - **Enrollment number**: also works in the admission box.
   - **Full name**: switch to **Full name** and type a name. Capital letters and extra spaces don't matter.
   - **Duplicate names**: search a common name (many exist in this file). You get a list of students with only the last 4 characters of their IDs, and must enter the full admission number to see the seat.
   - **No match**: type a made-up name. You should see *"No matching student found. Please check your details."*
   - **Empty search**: press Search with nothing typed. You should see a validation message.
   - **Keyboard**: Tab through the page, press Enter to search and Escape in the search box to clear.
4. **Worksheet to search**: "All worksheets" searches every shift. Pick one shift to search only that sheet.
5. **Adjust columns**: if a different file is detected wrongly, open this section, set the header row (the Excel row number that holds the column titles) and choose the right column for each field, then **Apply columns**.
6. **Privacy check (optional)**: open DevTools (F12) → **Network** tab, then load a file and search. No new requests appear, because the data never leaves the browser.

## 5. Deploy to Vercel

The site is plain HTML/CSS/JS, so there is **no build step**.

### Option A: GitHub + Vercel (recommended)

1. Create a **new GitHub repository** (Public or Private) and upload only the project files: `index.html`, `style.css`, `script.js`, `vercel.json`, `.vercelignore`, `.gitignore`, `README.md`.
   - Using Git: `git init`, `git add .`, `git commit -m "Initial version"`, then push. `.gitignore` will skip any spreadsheet files automatically.
   - Before pushing, run `git status` and make sure no `.xls`, `.xlsx` or `.csv` file is listed.
2. Go to <https://vercel.com>, sign in with GitHub, click **Add New → Project**, and import the repository.
3. Settings: **Framework Preset: Other**, **Build Command: empty**, **Output Directory: empty** (root). Click **Deploy**.
4. You get a URL like `https://exam-seat-finder.vercel.app`. Every later push to GitHub redeploys automatically.

### Option B: Vercel CLI (no GitHub)

1. Install Node.js, then run `npm i -g vercel`.
2. Inside the `exam-seat-finder` folder run `vercel` and follow the prompts (choose "Other", no build command).
3. Run `vercel --prod` to publish. `.vercelignore` prevents any spreadsheet in the folder from being uploaded.

### What `vercel.json` does

- **Content-Security-Policy** with `connect-src 'none'`: the browser itself blocks the page from making any network request (fetch, XHR, beacons). Even if the code were changed by mistake, the student data could not be sent anywhere.
- Scripts are allowed only from this site and `cdn.sheetjs.com`; styles and fonts only from Google Fonts.
- `X-Frame-Options: DENY` and `frame-ancestors 'none'`: other sites cannot embed this page to trick students.
- `Referrer-Policy: no-referrer`, `X-Content-Type-Options: nosniff`, and a restrictive `Permissions-Policy`.

If you later add another script or font source, add its domain to the CSP in `vercel.json`, otherwise the browser will block it.

## 6. Keeping student data out of GitHub and Vercel

The seating plan contains names and admission numbers of thousands of students. Treat it as confidential.

- Never put the spreadsheet inside the project folder. If it is there by accident, `.gitignore` and `.vercelignore` block common spreadsheet types, but don't rely on that alone.
- Never convert the spreadsheet into a `.js` or `.json` file inside the project, and never paste student rows into the code. Anything in the project folder becomes **public** once deployed, even if the GitHub repository is private.
- If a spreadsheet was ever committed, deleting it in a new commit is **not enough** (it stays in Git history). Make the repository private immediately, remove it from history (for example with `git filter-repo`), and redeploy.
- Do not share screenshots of the page with other students' details.

## 7. Deploying the website vs. uploading the spreadsheet

| | Deploying the website | Opening the spreadsheet |
|---|---|---|
| What moves | Only the code files (`index.html`, `style.css`, `script.js`, config) | Nothing leaves the device |
| Where it goes | Vercel's servers, public on the internet | Into the memory of that student's browser tab |
| Who does it | You, once (and on every update) | Each student, on their own phone or laptop |
| Contains student data? | **No, never** | Yes, but only locally, and cleared on reload |

So the public website is an empty tool. Each student needs the seating plan file themselves (for example, the one shared by the exam cell or department) and opens it in the tool.

## 8. How it works (for developers)

- **SheetJS 0.20.3** is loaded from the official SheetJS CDN with a pinned version. The `xlsx` package on npm is stuck at 0.18.5, which has known security issues, so it is not used.
- **Header detection**: the first 30 rows are scored, and the row whose cells look most like "Admission", "Name", "Seat", "Room", etc. is used as the header. This handles sheets with title rows above the table.
- **Session details** (exam date, shift timings) are copied word-for-word from the title rows above the header. Nothing is invented; missing fields are simply not shown, and a missing seat is shown as "Not listed in the spreadsheet".
- **Search**: admission and enrollment numbers are matched exactly, ignoring case and spaces. Names match exactly first (ignoring case, dots and extra spaces). If there is no exact match, names containing every typed word are shown.
- **Data cleaning**: blank rows and repeated header rows are skipped. Rows with neither a name nor an admission number are counted and skipped. Large numeric IDs shown by Excel as `2.4E+11` are restored to full digits.
- **Security**: spreadsheet text is always inserted with `textContent`, never `innerHTML`, so a crafted cell cannot inject code.

## 9. Troubleshooting

| Problem | Fix |
|---|---|
| "The spreadsheet reader did not load" | Check internet connection; `cdn.sheetjs.com` must be reachable. |
| "Password-protected" | Remove the password in Excel and save again. |
| "Could not be read as an Excel workbook" | Open in Excel → Save As → Excel Workbook (.xlsx). |
| Wrong or missing columns | Open **Adjust columns**, fix the header row and column choices, then Apply. |
| Works locally but blank on Vercel | Make sure `index.html` is at the repository root, not inside a sub-folder (or set Vercel's Root Directory to that folder). |
