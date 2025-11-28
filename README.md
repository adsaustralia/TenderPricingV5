# Tender Pricing App

Streamlit app for automating tender pricing from Excel workbooks.

## Features

- Upload an Excel pricing template.
- Map columns for **Size**, **Material**, **Quantity**, **Side (SS/DS)** and **Group**.
- Automatically:
  - Parse sizes to square metres (sqm).
  - Apply material rates and DS loadings.
  - Write results back into a copy of the original workbook.
- Supports complex size strings, including:
  - Rectangles: `841mm x 1189mm`, `1.2m x 2m`, `25,000mm (w) x 75mm (h)`.
  - Diameters: `260mm Ø`, `50mm Diameter`, or a single value like `290mm`.
  - Multi-sizes with quantities:
    - `2 x 2547mm x 755mm 2 x 967mm x 755mm`
    - `Various - 7 Decals 1. 2 x 355mm x 355mm ... 7. 4 x 30mm x 30mm`.

## Getting started

1. Create a new GitHub repo and add these files:
   - `app.py`
   - `requirements.txt`
   - `README.md`
2. Install dependencies (locally):

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
streamlit run app.py
```

4. Deploy to Streamlit Cloud by connecting the GitHub repo and choosing `app.py` as the entrypoint.
