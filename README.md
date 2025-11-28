# Tender Pricing App (Full UI – Fresh)

This build is based on your original `app (2).py` with these changes:

- **New size logic** (all your "other considerations"):
  - Single values & Ø / Diameter (`260mm Ø`, `50mm Diameter`, `290mm`, etc.) treated as circles.
  - "Various - N decals ..." and other multi-size lines are summed properly.
  - `2 x 2547mm x 755mm 2 x 967mm x 755mm` style lines are handled.
  - `25,000mm (w) x 75mm (h) 10 strips x 2,500mm (w)` uses 25,000 × 75.
  - Pure cap-height lines like `450mm Cap Height`, `40mm high text` produce **no area**.

- **Double-sided loading default = 20%** (was 25%).  
- **Tier count default = 1** (you can still add tiers in the UI).  
- **Header row picker** right after selecting the sheet:
  - You choose which Excel row is the header (e.g. 2 if your headings are on row 2).

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then deploy this folder to GitHub / Streamlit Cloud if needed.
