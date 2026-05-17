# ASCVD survival model comparison

This repository contains the Python analysis code used for the manuscript:

10-year atherosclerotic cardiovascular disease prediction in an Emirati cohort: Cox proportional hazards regression versus alternative survival models.

Files:
- primary_analysis.py: runs the main analysis, including model fitting, repeated internal validation, multiple imputation, performance metrics, calibration plots, SHAP outputs, and tables.
- sensitivity_analysis.py: runs the sensitivity analysis in which test-set imputation does not use event status, follow-up time, or the Nelson-Aalen cumulative hazard.

The individual-level dataset is not included because it cannot be shared publicly because of institutional and ethical restrictions.

To run the primary analysis:
python primary_analysis.py --csv data.csv --outdir outputs

To run the sensitivity analysis:
python sensitivity_analysis.py --csv data.csv --original-script primary_analysis.py --original-output-dir outputs --outdir outputs_sensitivity

The scripts were developed for Python 3.11.7. Package versions used in the analysis are reported in the manuscript and saved by the primary analysis script.
