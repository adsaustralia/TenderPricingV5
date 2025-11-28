# Tender Pricing App (Full UI – Header Select)

Based on your original `app (2).py` with these changes:

- **Header row selection** after you choose the sheet:  
  You can pick which Excel row is the header (e.g. 2 if your headings are on row 2).

- **Size parsing ("other considerations")**:
  - Single values & Ø / Diameter (`260mm Ø`, `333mm Ø`, `520mm Ø`, `50mm Diameter`, `290mm`, etc.) are treated as circles (diameter).
  - Multi-size strings with quantities (`2 x 2547mm x 755mm 2 x 967mm x 755mm`, `2 x 355mm x 355mm`, etc.) are summed as qty × width × height.
  - “Various – N decals ...” lines without explicit qty are handled by summing each `width x height` pair as 1 each, ignoring pure cap-height bits.
  - Strings like `25,000mm (w) x 75mm (h) 10 strips x 2,500mm (w)` take the first full `w x h` pair (25,000 × 75).
  - Cap-height-only lines (`450mm Cap Height`, `40mm high text`) produce no area.

- **Double-sided loading default = 20%** (still editable in the UI).
- **Tier count default = 1** (you can add more tiers from the UI).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then deploy this same folder to GitHub / Streamlit Cloud.
