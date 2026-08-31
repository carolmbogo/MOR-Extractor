
# MOR Column Extractor V2

This version replaces the original PDF extraction logic.

## Why V2 was needed

The original app relied on generic PDF table extraction. On multi-row MOR forms,
that can split headings into fragments such as "Fin", "al Efflu", and "Column 18".
It can also shift values into neighboring fields.

V2 uses actual PDF word coordinates.

For the Fourth Creek / KUB MOR layout, the app recognizes the form and maps the
printed column positions to clean field names. This includes:

- BOD Raw / Intermediate / Final
- Suspended Solids Raw / Intermediate / Final
- Set Solids Raw / Intermediate / Final
- DO
- pH
- NH3-N
- Grease
- Total N / Total P
- E. coli
- chlorine fields
- effluent flow
- Page 2 Secondary System DO
- MLSS
- SVI
- Waste GPD
- metals

The page-1 DO and page-2 Secondary System DO are deliberately separate fields.

## Run it

Open Command Prompt in this folder and run:

    python -m pip install -r requirements.txt

Then:

    python -m streamlit run app.py

Or double-click `run_app.bat` after the packages are installed.

## Testing the Fourth Creek DO column

Upload a Fourth Creek MOR or the ZIP of monthly MORs.

Choose:

    Daily MOR - Page 1

Then select:

    Date
    DO (mg/L)

The sample values for January 2021 should begin:

    5.9, 6.4, 6.2, 6.0, 6.2, 6.6 ...

For the files previously tested, these spot checks are expected:

- 01/01/2021 = 5.90
- 03/01/2021 = 6.80
- 04/01/2021 = 8.70
- 09/01/2021 = 6.70
- 01/01/2022 = 5.80
- 02/01/2022 = 6.40
- 03/01/2022 = 7.40

## Unknown MOR layouts

If a text-based PDF does not match the saved Fourth Creek/KUB layout, V2 uses a
generic positional detector. Those results are labeled "Review required".

The app shows sample values before export so the user can verify that a detected
field is the intended column.

Scanned/image-only PDFs are not OCR'd in this build.
