# Local editor and GitHub Pages viewer

This project has two separate modes:

- **Laptop editor:** `index.html` signs in against local `users.xlsx`; `home.html` uses the local Python helper to save MIR/WIR statuses into the Excel workbook. Start this local editor with `Update TTG Tracker.bat`.
- **GitHub Pages viewer:** a public, read-only snapshot built from `Sheet1` in the tracker workbook and the PDF folders. Opening the hosted page never needs a BAT file. It does not read `users.xlsx`, offer login, or change statuses.

GitHub Pages serves static files and cannot securely check an Excel password or write into a local Excel file. The hosted page and PDFs are public to anyone who has the URL. Do not publish `users.xlsx` or `Form.xlsx`. The main tracker workbook and the PDF folders are inputs to the Pages build, so their contents will be visible in the public viewer and repository.

## Publish

1. Keep `users.xlsx` and `Form.xlsx` on the laptop. `.gitignore` excludes these files when using Git locally; do not upload them manually in the GitHub browser.
2. Include `TTG worksheet for wir and mir.xlsx`, `TTG MIR`, `TTG WIR`, `Delivery Notes`, `CRS`, `Logo/TR.jpg`, `build_github_site.py`, `github_static_viewer.html`, and `update_tracker.py` in the GitHub repository. The files used as static viewer inputs will be public.
3. In repository Settings → Pages, select **GitHub Actions** as the source.
4. Push changes to the `main` branch. The workflow rebuilds the searchable site from the latest Excel data and PDFs and publishes it. If your default branch has another name, change the branch in `.github/workflows/pages.yml`.

The build includes only the fields used by the tracker page and PDF files in the `TTG MIR`, `TTG WIR`, `Delivery Notes`, and `CRS` folders. It does not place the source workbook or the user-password workbook in the published site artifact.

## Updating

On your laptop, edit workbook rows and attachments manually, then use the local editor to change MIR/WIR statuses. Push the updated tracker workbook and attachment folders to GitHub; the Actions workflow refreshes the public snapshot. Visitors to the GitHub site can search, filter, paginate, and open attachments, but status edits must be made through the laptop editor and published with the updated workbook.
