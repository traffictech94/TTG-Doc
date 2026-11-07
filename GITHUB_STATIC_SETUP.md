# Local editor and GitHub Pages sign-in viewer

This project has two separate modes:

- **Laptop editor:** `index.html` signs in against local `users.xlsx`; `home.html` uses the local Python helper to save MIR/WIR statuses into the Excel workbook. Start this local editor with `Update TTG Tracker.bat`.
- **GitHub Pages viewer:** a public, read-only snapshot built from `Sheet1` in the tracker workbook and the PDF folders. Its sign-in page checks the username, password, and level from `users.xlsx` in the browser. The account list is published as `users.json`, so all usernames, passwords, and levels are visible to anyone. This is only a convenience gate, not secure authentication. The viewer does not write status changes to Excel.

GitHub Pages serves static files and cannot securely check an Excel password or write into a local Excel file. The hosted page, published login list, workbook source, and PDFs are public to anyone. Do not use passwords that protect other accounts. `Form.xlsx` remains excluded from the Pages build.

## Publish

1. Commit `users.xlsx` to the public repository. The workflow needs it to generate the sign-in account list. Do not commit `Form.xlsx`.
2. Include `TTG worksheet for wir and mir.xlsx`, `TTG MIR`, `TTG WIR`, `Delivery Notes`, `CRS`, `Logo/TR.jpg`, `build_github_site.py`, `index.html`, `github_static_viewer.html`, and `update_tracker.py` in the GitHub repository. The files used as static viewer inputs will be public.
3. In repository Settings → Pages, select **GitHub Actions** as the source.
4. Push changes to the `main` branch. The workflow rebuilds the searchable site from the latest Excel data and PDFs and publishes it. If your default branch has another name, change the branch in `.github/workflows/pages.yml`.

The build includes only the fields used by the tracker page, the account credentials from `users.xlsx`, and PDF files in the `TTG MIR`, `TTG WIR`, `Delivery Notes`, and `CRS` folders. The site checks the account list in the browser; visitors can inspect or bypass that check. All hosted status values remain read-only.

## Updating

On your laptop, edit workbook rows and attachments manually, then use the local editor to change MIR/WIR statuses. Push the updated tracker workbook, `users.xlsx`, and attachment folders to GitHub; the Actions workflow refreshes the public snapshot. Visitors to the GitHub site can sign in, search, filter, paginate, and open attachments, but status edits must be made through the laptop editor and published with the updated workbook.
