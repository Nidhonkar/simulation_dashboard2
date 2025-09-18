
# Ganga Jamuna — VP Dashboard (Fresh Connection) — v3

**No more "file not found":** the app now searches `data/` for expected files *and* lets you upload `.xlsx` files from the sidebar.  
It also accepts both `FinanceReport (6).xlsx` and `FinanceReport_6.xlsx` automatically.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Cloud)
- Push to GitHub; set **Main file path** to `app.py`.
- If your data files aren't in the repo, upload them via the app's sidebar after deploy.
