# TTG MIR / WIR Tracker — Project Features

Project reference notes for maintenance, handover, and future updates.

## Project at a glance

- Browser based tracker for Traffic Tech MIR and WIR submissions, statuses, and related documents.
- The same project supports two separate modes:
  - **Laptop editor:** a local Python service reads and updates the Excel workbooks and attachment folders.
  - **GitHub Pages site:** a public static snapshot generated from the committed workbook and document folders.
- The Excel workbook remains the data source; there is no SQL database.
- Main tracker workbook: `TTG worksheet for wir and mir.xlsx`.
- User workbook: `users.xlsx`.
- `Form.xlsx` is excluded from the GitHub Pages build/repository by `.gitignore`.

## Main pages and files

- `index.html` — sign-in page for the local editor and the published static viewer.
- `home.html` — searchable tracker table, filters, pagination, status controls, and signed-in user display.
- `add_record.html` — Level 1 form for adding a MIR/WIR tracker row and optional attachments (local editor only).
- `recordlist.html` — Level 1 list for searching, editing, or deleting tracker rows (local editor only).
- `register.html` — Level 1 form for adding a user to `users.xlsx` (local editor only).
- `userlist.html` — Level 1 user list for searching, editing, or deleting users (local editor only).
- `tracker_server.py` — local HTTP service and Excel/file operations.
- `update_tracker.py` — reads the workbook, indexes attachment folders, and builds/refreshes tracker data.
- `Update TTG Tracker.bat` — Windows launcher for the local editor/service.
- `.github/workflows/pages.yml` and `build_github_site.py` — build and publish the static GitHub Pages site after a push to `main` (or a manual workflow run).

## Login and access levels

- User fields are read from `users.xlsx`: number, username, password, first name, last name, email, and level.
- The signed-in user's first and last name and an `L-1`, `L-2`, or `L-3` label appear in the tracker header when those name fields are present.
- **Level 1 — Super Admin:** can stage MIR/WIR status changes and save them to Excel; add, edit, and delete tracker records; register users; and manage the user list.
- **Level 2 — Standard user:** read-only tracker access.
- **Level 3 — Read-only user:** read-only tracker access.
- Users can change their own password in the local editor; the new password is written to that user's row in `users.xlsx`.
- Sign out returns to the sign-in page. Local restricted pages and write operations check the signed-in session and Level 1 access.

## Tracker and Excel features

- Tracker rows are read from `Sheet1` of the main workbook; the row serial number is retained when a row is edited.
- Search covers description, item/notes, MIR and WIR references, delivery note, and comments/CRS text.
- Filters include record group, status, and the Items / Notes dropdown.
- Record number can be sorted in either direction: newest/highest first or oldest/lowest first. Newest/highest is the default.
- The home table shows 10 records per page and provides numbered pagination with next/last navigation when more records match.
- A top-right Notifications menu lists the three most recently modified/committed attachments across MIR, WIR, Delivery Notes, and CRS. Each entry opens the attachment; long MIR/WIR names use a short label and retain the full filename on hover.
- MIR/WIR references are presented with compact labels such as `MIR-ITS-0018` and `WIR-ITS-0094`; hovering a link reveals its full reference, and matching PDFs open in a new tab.
- Descriptions remain visible in the table. Item/notes are available from the workbook for searching and filtering and as description hover context.
- Level 1 can change MIR/WIR statuses in the page, review staged changes, then choose **Update Excel** to write those status changes to the workbook.
- If the workbook is open/locked in Excel, close it and retry the save.

## Record entry, editing, and attachments

- New records and edits are submitted through the local service; file inputs are optional unless the user chooses to attach a document.
- MIR and WIR PDFs are stored in `TTG MIR` and `TTG WIR` respectively.
- Delivery note files are stored in `Delivery Notes`.
- CRS/comments attachments are stored in `CRS`; supported types include PDF, Excel, Word, PNG, and JPG/JPEG.
- Attachment size is limited to 50 MB per file.
- When a selected attachment's filename already exists in its destination folder (including a matching revision/timestamp family), the page asks whether to reuse an existing attachment.
- Confirming reuses the existing file and saves its name/reference in the workbook; it does not upload a second copy or create a new `Rev` file. Canceling leaves the report unchanged.
- If a new filename does not match an existing file, the attachment is saved in its destination folder and the workbook reference is filled from its filename without the extension.
- Editing/deleting a record does not automatically remove its existing attachment files from the folders.

## Local laptop workflow

- Keep the project files, workbooks, and attachment folders together in the project directory.
- Close `users.xlsx` or the tracker workbook in Excel before saving changes through the app.
- Double-click `Update TTG Tracker.bat` to start the local service, then use the sign-in page in the browser.
- Make record/status/user changes through the local app. The local service runs at `127.0.0.1:8767`.
- After changing workbook data or adding/removing PDFs outside the app, refresh/rebuild the tracker data with the supplied launcher/update process.
- Review changes locally before committing or pushing them to GitHub.

## GitHub Pages deployment and limitations

- The GitHub Actions workflow builds a static site from the committed workbook and document folders, then publishes it under the repository's GitHub Pages URL.
- The static build uses Git history to order recent attachment notifications; the laptop editor uses each file's local modified time.
- Pushing updated source files to `main` starts the Pages workflow automatically; a successful workflow run indicates deployment completed. A manual workflow run is also available.
- The published site is a read-only snapshot: users can search, filter, paginate, and open published attachments, but cannot save status changes or write to Excel.
- Static hosting cannot run `tracker_server.py` or the Windows `.bat` launcher. Add/edit/delete, user administration, and password changes therefore work only in the local editor.
- The Pages sign-in checks a generated `users.json` list in the browser. This is a convenience gate, **not secure authentication**; the published account data can be viewed or bypassed by visitors.
- A public repository/site exposes published workbook-derived data and attachments. Treat all committed files as public and do not reuse passwords from other services.
- To publish manual Excel/PDF updates: commit and push the changed workbook, user workbook (if changed), relevant attachments, and site source to `main`; then confirm the Pages workflow succeeds and open the published URL.

## Routine maintenance checklist

- Keep workbook sheet/header names compatible with the tracker reader, especially the `Sheet1` MIR/WIR table.
- Keep attachment folder names consistent: `TTG MIR`, `TTG WIR`, `Delivery Notes`, and `CRS`.
- Keep the published `users.xlsx` account list intentional and current; remember that the generated static account list is public.
- After changing application files, restart the local tracker and check the local pages before pushing.
- After pushing, check **Actions → Publish static TTG viewer** for a successful deployment, then check the GitHub Pages site.
- Do not commit temporary Excel lock files (`~$*.xlsx`), Python caches, generated `_site` output, or `Form.xlsx`.
-   
## Newly added features to add Notification pannel and Dec Asc Order filter.

