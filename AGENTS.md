# AGENTS.md

## Project overview

- This project is a Streamlit survey app for Meiwa company travel voting.
- Users enter employee information, choose whether they will join, and if joining they must choose exactly one destination:
  - `Nha Trang`
  - `Da Lat`
- The app writes results into one consolidated Excel file and can optionally sync that file to Google Drive.
- The main user-facing app is in [app.py](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\app.py).
- Excel generation logic is in [excel_export.py](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\excel_export.py).

## Setup commands

- Install dependencies:
  ```powershell
  pip install -r requirements.txt
  ```
- Run the app locally:
  ```powershell
  streamlit run app.py
  ```
- Run the app with explicit network binding:
  ```powershell
  streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true
  ```

## Important files

- [app.py](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\app.py): Streamlit UI, state, dashboard, Google Drive sync, save flow
- [excel_export.py](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\excel_export.py): Excel workbook layout and export logic
- [render.yaml](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\render.yaml): Render deployment config
- [requirements.txt](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\requirements.txt): Python dependencies
- [README.md](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\README.md): human-oriented project notes
- [exports](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\exports): local output folder for the consolidated Excel file

## Development tips

- This is a single-app Streamlit project, not a package-based monorepo.
- Prefer small, targeted edits in `app.py`; it contains most of the business logic.
- Keep the survey flow simple. The user explicitly wants the app to open and be usable immediately.
- The app is bilingual Vietnamese/Japanese in many visible places. Preserve both languages when editing user-facing text.
- Mobile usability matters. Always consider compact layout and readable labels on phones.
- Only two destination choices are valid. Do not introduce extra destinations unless the user asks.
- If `Không / 不参加` is selected, destination selection must stay hidden or cleared.
- If `Có / 参加` is selected, destination selection must appear immediately and require one choice.

## Code style

- Use Python 3.11-compatible syntax.
- Keep code straightforward and readable over clever abstractions.
- Follow the existing normalization helpers and record-cleaning patterns rather than duplicating logic.
- Preserve the current naming style for record fields:
  - `msnv`
  - `ho_ten`
  - `bo_phan`
  - `cong_doan`
  - `tham_gia`
  - `dia_diem`
- Prefer keeping user-visible labels in Vietnamese first, Japanese second.
- When editing UI text, avoid changing business meaning accidentally.

## Testing instructions

- At minimum, run Python syntax validation after every code change:
  ```powershell
  py -3 -m py_compile app.py
  ```
- If `excel_export.py` changes, validate both files:
  ```powershell
  py -3 -m py_compile app.py excel_export.py
  ```
- For UI-related changes, manually verify in Streamlit:
  - form input flow
  - join / not join toggle behavior
  - destination visibility behavior
  - save button behavior
  - dashboard rendering
  - download button behavior
- For Excel-related changes, confirm that the exported workbook still matches the existing company format.

## Data and Excel rules

- The project should maintain one consolidated Excel output file, not many separate result files.
- New survey entries should update the consolidated file without breaking the existing layout.
- The export file name is:
  - `Form khao sat Du lich Cong ty meiwa nam 2026.xlsx`
- The local saved file normally lives under:
  - [exports\Form khao sat Du lich Cong ty meiwa nam 2026.xlsx](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\exports\Form khao sat Du lich Cong ty meiwa nam 2026.xlsx)

## Google Drive and secrets

- Google Drive sync is optional but important in production.
- Never hardcode secrets into source files.
- Use environment variables or Streamlit secrets for:
  - `GOOGLE_DRIVE_FOLDER_ID`
  - `GOOGLE_DRIVE_SHARED_URL`
  - `GOOGLE_SERVICE_ACCOUNT_JSON`
- Treat all service account JSON content as sensitive.
- Do not log secrets, paste secrets into committed files, or echo them into output unnecessarily.

## Render deployment

- Render is the production hosting target for this project.
- Deployment config is defined in [render.yaml](D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa\render.yaml).
- The expected Render build command is:
  ```powershell
  pip install -r requirements.txt
  ```
- The expected Render start command is:
  ```powershell
  streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false
  ```
- The app expects persistent storage on Render with:
  - `DATA_DIR=/var/data`
- If a change affects deployment behavior, verify that `render.yaml` still matches the app’s real needs.

## UI and product expectations

- The user prefers a simple interface over a technical one.
- The app must work on both desktop and mobile.
- Survey actions should feel obvious:
  - save button should be prominent
  - download action should be easy to find
  - dashboard should be readable
- Keep bilingual headings compact so mobile screens do not become too tall.
- Avoid UI regressions where raw HTML appears in the app.

## Agent instructions

- Before editing, check whether the change affects:
  - survey form logic
  - Excel export layout
  - Google Drive sync
  - Render deployment
  - bilingual UI text
- After editing, validate syntax and then do the smallest reasonable manual verification.
- If you fix a production UI bug, prefer also pushing and redeploying so the online version matches local changes.
- Do not create extra result files or alternative export flows unless explicitly requested.
- Do not remove bilingual labels unless explicitly requested.
