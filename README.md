# ASCVD survival model comparison
This repository contains the Python analysis code used for the manuscript:
Comparing Cox regression with regularised and machine-learning survival models for 10-year atherosclerotic cardiovascular disease prediction in a Middle Eastern cohort. The data are not included because they contain confidential clinical information.

Use 64-bit CPython 3.11.7. Install the packages with:

```bash
python -m pip install -r requirements.txt
```

Run the full analysis with:

```bash
python analysis.py --csv path/to/analysis_data.csv --master path/to/source_data.xlsx --sheet "sheet name" --outdir results
```

The script checks the packages and input data before fitting. It runs the main, sensitivity, secondary, and final stages in order. Results are saved in separate folders under the selected output folder.
