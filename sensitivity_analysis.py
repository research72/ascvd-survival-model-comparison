from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

import numpy as np
import pandas as pd
from sklearn.base import clone
from sklearn.experimental import enable_iterative_imputer
from sklearn.impute import IterativeImputer
from sklearn.linear_model import BayesianRidge


KEY_METRICS = [
    "harrell_c",
    "uno_c_tau",
    "brier_score_at_horizon",
    "integrated_brier_score",
    "calibration_slope",
]

PRIMARY_COL_MAP = {
    "Harrell_C": "harrell_c",
    "Uno_C_tau": "uno_c_tau",
    "Brier_120m": "brier_score_at_horizon",
    "IBS_12_120m": "integrated_brier_score",
    "Calibration_slope_120m": "calibration_slope",
}

HIGHER_IS_BETTER = {"harrell_c", "uno_c_tau"}
LOWER_IS_BETTER = {"brier_score_at_horizon", "integrated_brier_score"}
DEFAULT_CSV_PATH = "data.csv"
DEFAULT_ORIGINAL_SCRIPT_PATH = "primary_analysis.py"
DEFAULT_ORIGINAL_OUTPUT_DIR = "outputs"
DEFAULT_OUTDIR = "outputs_sensitivity"
DEFAULT_SKIP_REPEATED = False

DEFAULT_MASTER_SEED = 20260320
DEFAULT_TEST_SIZE = 0.2
DEFAULT_N_REPEATS = 30
DEFAULT_M_IMPUTATIONS = 10
DEFAULT_IMPUTE_MAX_ITER = 20
DEFAULT_HORIZON_MONTHS = 120.0
DEFAULT_IBS_START_MONTH = 12.0
DEFAULT_CALIBRATION_GROUPS = 5
DEFAULT_N_JOBS = -1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", default=DEFAULT_CSV_PATH)
    parser.add_argument("--original-script", default=DEFAULT_ORIGINAL_SCRIPT_PATH)
    parser.add_argument("--original-output-dir", default=DEFAULT_ORIGINAL_OUTPUT_DIR)
    parser.add_argument("--outdir", default=DEFAULT_OUTDIR)
    parser.add_argument("--seed", type=int, default=DEFAULT_MASTER_SEED)
    parser.add_argument("--test-size", type=float, default=DEFAULT_TEST_SIZE)
    parser.add_argument("--n-repeats", type=int, default=DEFAULT_N_REPEATS)
    parser.add_argument("--m-imputations", type=int, default=DEFAULT_M_IMPUTATIONS)
    parser.add_argument("--impute-max-iter", type=int, default=DEFAULT_IMPUTE_MAX_ITER)
    parser.add_argument("--horizon-months", type=float, default=DEFAULT_HORIZON_MONTHS)
    parser.add_argument("--ibs-start-month", type=float, default=DEFAULT_IBS_START_MONTH)
    parser.add_argument("--calibration-groups", type=int, default=DEFAULT_CALIBRATION_GROUPS)
    parser.add_argument("--n-jobs", type=int, default=DEFAULT_N_JOBS)
    parser.add_argument(
        "--skip-repeated",
        action="store_true",
        default=DEFAULT_SKIP_REPEATED,
    )
    return parser.parse_args()


def load_original_module(script_path: Path):
    module_name = "orig_pipeline"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise ImportError("Could not import module.")
    module = importlib.util.module_from_spec(spec)
    import sys
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return module


def locate_file(base_dir: Path, filename: str) -> Path:
    candidates = [
        base_dir / filename,
        base_dir / "tables" / filename,
        base_dir / "metadata" / filename,
        base_dir / "figures" / filename,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"Could not find {filename}.")


def load_original_outputs(base_dir: Path) -> Dict[str, pd.DataFrame]:
    files = {
        "primary": "table3_primary_split_metrics.csv",
        "repeat_perf": "performance_by_repeat.csv",
        "repeat_summary": "performance_summary_repeated_resampling.csv",
        "paired": "supplementary_table_s3_paired_differences_vs_coxph.csv",
        "tuning_primary": "tuning_records_primary_split.csv",
        "tuning_repeat": "tuning_records_resampling.csv",
    }
    out: Dict[str, pd.DataFrame] = {}
    for key, filename in files.items():
        out[key] = pd.read_csv(locate_file(base_dir, filename))
    return out


def build_imputation_matrix(df: pd.DataFrame, orig: Any, include_aux_outcomes: bool) -> pd.DataFrame:
    out = df[orig.PREDICTOR_COLUMNS].copy()
    if include_aux_outcomes:
        out["_aux_event"] = df[orig.EVENT_COL].astype(float)
        out["_aux_time"] = np.log1p(df[orig.TIME_COL].astype(float))
        out["_aux_nelson_aalen"] = orig.nelson_aalen_feature(df)
    else:
        out["_aux_event"] = np.nan
        out["_aux_time"] = np.nan
        out["_aux_nelson_aalen"] = np.nan
    return out


def generate_imputed_split_datasets_sensitivity(
    train_df: pd.DataFrame,
    test_df: pd.DataFrame,
    orig: Any,
    m: int,
    seed: int,
    max_iter: int,
) -> List[Dict[str, pd.DataFrame]]:
    train_matrix = build_imputation_matrix(train_df, orig, include_aux_outcomes=True)
    test_matrix = build_imputation_matrix(test_df, orig, include_aux_outcomes=False)
    cols = list(train_matrix.columns)
    out: List[Dict[str, pd.DataFrame]] = []
    for i in range(m):
        imputer = IterativeImputer(
            estimator=BayesianRidge(),
            max_iter=max_iter,
            sample_posterior=True,
            random_state=seed + i,
            initial_strategy="median",
            imputation_order="ascending",
        )
        train_completed = pd.DataFrame(imputer.fit_transform(train_matrix[cols]), columns=cols, index=train_df.index)
        test_completed = pd.DataFrame(imputer.transform(test_matrix[cols]), columns=cols, index=test_df.index)
        train_pred = orig.postprocess_imputed_predictors(train_completed[orig.PREDICTOR_COLUMNS])
        test_pred = orig.postprocess_imputed_predictors(test_completed[orig.PREDICTOR_COLUMNS])
        imputed_train = pd.concat(
            [train_df[[orig.STUDY_ID_COL, orig.TIME_COL, orig.EVENT_COL]].reset_index(drop=True), train_pred.reset_index(drop=True)],
            axis=1,
        )
        imputed_test = pd.concat(
            [test_df[[orig.STUDY_ID_COL, orig.TIME_COL, orig.EVENT_COL]].reset_index(drop=True), test_pred.reset_index(drop=True)],
            axis=1,
        )
        out.append({"train": imputed_train, "test": imputed_test})
    return out


def tuning_lookup(tuning_df: pd.DataFrame) -> Dict[Tuple[int, int, str], Dict[str, Any]]:
    lookup: Dict[Tuple[int, int, str], Dict[str, Any]] = {}
    for _, row in tuning_df.iterrows():
        repeat = int(row["repeat"])
        imputation = int(row["imputation"])
        model = str(row["model"])
        params = json.loads(row["best_params_json"])
        lookup[(repeat, imputation, model)] = params
    return lookup


def fit_classical_from_saved_params(
    orig: Any,
    model_name: str,
    train_df: pd.DataFrame,
    split_seed: int,
    n_jobs: int,
    params: Dict[str, Any],
):
    specs = orig.get_classical_model_specs(random_state=split_seed, n_jobs=n_jobs)
    estimator = clone(specs[model_name]["pipeline"])
    estimator.set_params(**params)
    X_train = train_df[orig.PREDICTOR_COLUMNS].copy()
    y_train = orig.get_surv_array(train_df)
    estimator.fit(X_train, y_train)
    return estimator


def find_config_index(config: Dict[str, Any], config_list: Sequence[Dict[str, Any]]) -> int:
    target = json.dumps(config, sort_keys=True)
    for idx, candidate in enumerate(config_list, start=1):
        if json.dumps(candidate, sort_keys=True) == target:
            return idx
    return 1


def fit_deepsurv_fixed(orig: Any, train_df: pd.DataFrame, random_state: int, cfg: Dict[str, Any]):
    import torchtuples as tt                
    from pycox.models import CoxPH as PyCoxCoxPH                

    cfg_index = find_config_index(cfg, orig.get_deepsurv_configs())
    orig.set_global_seed(random_state)
    feature_names = list(orig.PREDICTOR_COLUMNS)
    prep = orig.TabularSurvivalPreprocessor(orig.CONTINUOUS_COLUMNS, orig.BINARY_COLUMNS, scale_continuous=True)
    prep.fit(train_df[feature_names])
    splitter = orig.StratifiedShuffleSplit(n_splits=1, test_size=0.2, random_state=random_state)
    strata = orig.make_strata(train_df)
    tr_idx, val_idx = next(splitter.split(train_df, strata))
    tr_df = train_df.iloc[tr_idx].copy()
    val_df = train_df.iloc[val_idx].copy()

    x_tr = prep.transform(tr_df[feature_names]).to_numpy(dtype="float32")
    x_val = prep.transform(val_df[feature_names]).to_numpy(dtype="float32")
    y_tr = (tr_df[orig.TIME_COL].astype("float32").to_numpy(), tr_df[orig.EVENT_COL].astype("int64").to_numpy())
    y_val = (val_df[orig.TIME_COL].astype("float32").to_numpy(), val_df[orig.EVENT_COL].astype("int64").to_numpy())

    orig.set_global_seed(random_state + cfg_index)
    net = tt.practical.MLPVanilla(
        in_features=x_tr.shape[1],
        num_nodes=cfg["num_nodes"],
        out_features=1,
        batch_norm=True,
        dropout=cfg["dropout"],
        output_bias=False,
    )
    model = PyCoxCoxPH(net, tt.optim.Adam)
    model.optimizer.set_lr(cfg["lr"])
    callbacks = [tt.callbacks.EarlyStopping()]
    model.fit(
        x_tr,
        y_tr,
        batch_size=cfg["batch_size"],
        epochs=cfg["epochs"],
        callbacks=callbacks,
        verbose=False,
        val_data=(x_val, y_val),
        val_batch_size=cfg["batch_size"],
    )
    model.compute_baseline_hazards()
    return orig.PyCoxAdapter(model=model, preprocessor=prep, feature_names=feature_names)


def fit_coxtime_fixed(orig: Any, train_df: pd.DataFrame, random_state: int, cfg: Dict[str, Any]):
    import torchtuples as tt                
    from pycox.models import CoxTime                
    from pycox.models.cox_time import MLPVanillaCoxTime                

    cfg_index = find_config_index(cfg, orig.get_coxtime_configs())
    orig.set_global_seed(random_state)
    feature_names = list(orig.PREDICTOR_COLUMNS)
    prep = orig.TabularSurvivalPreprocessor(orig.CONTINUOUS_COLUMNS, orig.BINARY_COLUMNS, scale_continuous=True)
    prep.fit(train_df[feature_names])
    splitter = orig.StratifiedShuffleSplit(n_splits=1, test_size=0.2, random_state=random_state)
    strata = orig.make_strata(train_df)
    tr_idx, val_idx = next(splitter.split(train_df, strata))
    tr_df = train_df.iloc[tr_idx].copy()
    val_df = train_df.iloc[val_idx].copy()

    x_tr = prep.transform(tr_df[feature_names]).to_numpy(dtype="float32")
    x_val = prep.transform(val_df[feature_names]).to_numpy(dtype="float32")
    durations_tr = tr_df[orig.TIME_COL].astype("float32").to_numpy()
    events_tr = tr_df[orig.EVENT_COL].astype("float32").to_numpy()
    durations_val = val_df[orig.TIME_COL].astype("float32").to_numpy()
    events_val = val_df[orig.EVENT_COL].astype("float32").to_numpy()

    orig.set_global_seed(random_state + 100 + cfg_index)
    labtrans = CoxTime.label_transform()
    y_tr = labtrans.fit_transform(durations_tr, events_tr)
    y_val = labtrans.transform(durations_val, events_val)
    net = MLPVanillaCoxTime(
        in_features=x_tr.shape[1],
        num_nodes=cfg["num_nodes"],
        batch_norm=True,
        dropout=cfg["dropout"],
    )
    model = CoxTime(net, tt.optim.Adam, labtrans=labtrans)
    model.optimizer.set_lr(cfg["lr"])
    callbacks = [tt.callbacks.EarlyStopping()]
    val_data = tt.tuplefy(x_val, y_val)
    model.fit(
        x_tr,
        y_tr,
        batch_size=cfg["batch_size"],
        epochs=cfg["epochs"],
        callbacks=callbacks,
        verbose=False,
        val_data=val_data,
        val_batch_size=cfg["batch_size"],
    )
    model.compute_baseline_hazards()
    return orig.PyCoxAdapter(
        model=model,
        preprocessor=prep,
        feature_names=feature_names,
        predict_from_surv=True,
        risk_horizon=120.0,
    )


def fit_model_from_saved_params(
    orig: Any,
    model_name: str,
    train_df: pd.DataFrame,
    split_seed: int,
    n_jobs: int,
    params: Dict[str, Any],
):
    if model_name in {"CoxPH", "ElasticNetCox", "RSF", "GBSA", "XGBoost", "SVM"}:
        return fit_classical_from_saved_params(orig, model_name, train_df, split_seed, n_jobs, params)
    if model_name == "DeepSurv":
        return fit_deepsurv_fixed(orig, train_df, split_seed, params)
    if model_name == "CoxTime":
        return fit_coxtime_fixed(orig, train_df, split_seed, params)
    raise ValueError(f"Unsupported model: {model_name}")


def run_primary_sensitivity(
    orig: Any,
    df: pd.DataFrame,
    tuning_df: pd.DataFrame,
    out_dir: Path,
    seed: int,
    test_size: float,
    m_imputations: int,
    impute_max_iter: int,
    horizon_months: float,
    ibs_start_month: float,
    calibration_groups: int,
    n_jobs: int,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    train_idx, test_idx = orig.primary_split(df, seed=seed, test_size=test_size)
    raw_train_df = df.iloc[train_idx].copy().reset_index(drop=True)
    raw_test_df = df.iloc[test_idx].copy().reset_index(drop=True)
    imputed_splits = generate_imputed_split_datasets_sensitivity(
        raw_train_df,
        raw_test_df,
        orig=orig,
        m=m_imputations,
        seed=seed,
        max_iter=impute_max_iter,
    )
    lookup = tuning_lookup(tuning_df)
    records: List[Dict[str, Any]] = []

    for imp_idx, split in enumerate(imputed_splits, start=1):
        train_df = split["train"]
        test_df = split["test"]
        split_seed = seed + imp_idx
        for model_name in orig.MODEL_ORDER:
            params = lookup.get((0, imp_idx, model_name))
            if params is None:
                raise KeyError(f"Missing saved parameters for primary split imputation {imp_idx}, model {model_name}.")
            fitted_model = fit_model_from_saved_params(orig, model_name, train_df, split_seed, n_jobs, params)
            metrics, _, _ = orig.evaluate_model(
                model_name,
                fitted_model,
                train_df,
                test_df,
                horizon_months,
                ibs_start_month,
                calibration_groups,
            )
            records.append({"imputation": imp_idx, "model": model_name, **metrics})

    by_imp = pd.DataFrame(records)
    by_imp.to_csv(out_dir / "tables" / "sensitivity_table3_primary_split_by_imputation.csv", index=False)

    avg_rows = []
    for model_name, sub in by_imp.groupby("model"):
        row = {"Model": model_name}
        for col in KEY_METRICS:
            vals = sub[col].dropna()
            row[col] = float(vals.mean()) if not vals.empty else np.nan
        avg_rows.append(row)
    primary = pd.DataFrame(avg_rows).set_index("Model").loc[orig.MODEL_ORDER].reset_index()
    primary = primary.rename(
        columns={
            "harrell_c": "Harrell_C",
            "uno_c_tau": "Uno_C_tau",
            "brier_score_at_horizon": "Brier_120m",
            "integrated_brier_score": "IBS_12_120m",
            "calibration_slope": "Calibration_slope_120m",
        }
    )
    primary.to_csv(out_dir / "tables" / "sensitivity_table3_primary_split_metrics.csv", index=False)
    return by_imp, primary


def run_repeated_sensitivity(
    orig: Any,
    df: pd.DataFrame,
    tuning_df: pd.DataFrame,
    out_dir: Path,
    seed: int,
    n_repeats: int,
    test_size: float,
    m_imputations: int,
    impute_max_iter: int,
    horizon_months: float,
    ibs_start_month: float,
    calibration_groups: int,
    n_jobs: int,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    splitter = orig.StratifiedShuffleSplit(n_splits=n_repeats, test_size=test_size, random_state=seed)
    strata = orig.make_strata(df)
    lookup = tuning_lookup(tuning_df)
    split_level_records: List[Dict[str, Any]] = []

    for repeat, (train_idx, test_idx) in enumerate(splitter.split(df, strata), start=1):
        raw_train_df = df.iloc[train_idx].copy().reset_index(drop=True)
        raw_test_df = df.iloc[test_idx].copy().reset_index(drop=True)
        imputed_splits = generate_imputed_split_datasets_sensitivity(
            raw_train_df,
            raw_test_df,
            orig=orig,
            m=m_imputations,
            seed=seed + 1000 * repeat,
            max_iter=impute_max_iter,
        )
        imp_metrics: List[Dict[str, Any]] = []
        for imp_idx, split in enumerate(imputed_splits, start=1):
            train_df = split["train"]
            test_df = split["test"]
            split_seed = seed + 1000 * repeat + imp_idx
            for model_name in orig.MODEL_ORDER:
                params = lookup.get((repeat, imp_idx, model_name))
                if params is None:
                    raise KeyError(f"Missing saved parameters for repeat {repeat}, imputation {imp_idx}, model {model_name}.")
                fitted_model = fit_model_from_saved_params(orig, model_name, train_df, split_seed, n_jobs, params)
                metrics, _, _ = orig.evaluate_model(
                    model_name,
                    fitted_model,
                    train_df,
                    test_df,
                    horizon_months,
                    ibs_start_month,
                    calibration_groups,
                )
                imp_metrics.append({"repeat": repeat, "imputation": imp_idx, "model": model_name, **metrics})
        imp_metrics_df = pd.DataFrame(imp_metrics)
        for model_name, sub in imp_metrics_df.groupby("model"):
            row = {"repeat": repeat, "model": model_name}
            for col in KEY_METRICS:
                vals = sub[col].dropna()
                row[col] = float(vals.mean()) if not vals.empty else np.nan
            split_level_records.append(row)

    perf_df = pd.DataFrame(split_level_records)
    perf_df.to_csv(out_dir / "tables" / "sensitivity_performance_by_repeat.csv", index=False)
    summary_df, formatted_df = orig.summarise_performance(perf_df)
    summary_df.to_csv(out_dir / "tables" / "sensitivity_performance_summary_repeated_resampling.csv", index=False)
    formatted_df.to_csv(out_dir / "tables" / "sensitivity_table4_repeated_resampling_formatted.csv", index=False)
    comparisons_df = orig.compare_models_vs_reference(perf_df, reference_model="CoxPH")
    comparisons_df.to_csv(out_dir / "tables" / "sensitivity_paired_differences_vs_coxph.csv", index=False)
    return perf_df, summary_df, comparisons_df


def standardize_primary(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out = out.rename(columns=PRIMARY_COL_MAP)
    return out[["Model", *KEY_METRICS]].rename(columns={"Model": "model"})


def standardize_repeat_summary(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rename = {
        "model": "model",
        "harrell_c_mean": "harrell_c",
        "uno_c_tau_mean": "uno_c_tau",
        "brier_score_at_horizon_mean": "brier_score_at_horizon",
        "integrated_brier_score_mean": "integrated_brier_score",
        "calibration_slope_mean": "calibration_slope",
    }
    cols = [
        "model",
        "harrell_c_mean",
        "uno_c_tau_mean",
        "brier_score_at_horizon_mean",
        "integrated_brier_score_mean",
        "calibration_slope_mean",
    ]
    return out[cols].rename(columns=rename)


def compare_primary(original_primary: pd.DataFrame, sensitivity_primary: pd.DataFrame) -> pd.DataFrame:
    base = standardize_primary(original_primary).set_index("model")
    sens = standardize_primary(sensitivity_primary).set_index("model")
    rows = []
    for model in sens.index:
        for metric in KEY_METRICS:
            old = float(base.loc[model, metric])
            new = float(sens.loc[model, metric])
            rows.append(
                {
                    "model": model,
                    "metric": metric,
                    "original": old,
                    "sensitivity": new,
                    "absolute_change": new - old,
                }
            )
    return pd.DataFrame(rows)


def compare_repeat_summary(original_summary: pd.DataFrame, sensitivity_summary: pd.DataFrame) -> pd.DataFrame:
    base = standardize_repeat_summary(original_summary).set_index("model")
    sens = standardize_repeat_summary(sensitivity_summary).set_index("model")
    rows = []
    for model in sens.index:
        for metric in KEY_METRICS:
            rows.append(
                {
                    "model": model,
                    "metric": metric,
                    "original_mean": float(base.loc[model, metric]),
                    "sensitivity_mean": float(sens.loc[model, metric]),
                    "absolute_change": float(sens.loc[model, metric] - base.loc[model, metric]),
                }
            )
    return pd.DataFrame(rows)


def classify_paired_difference(metric: str, mean_diff: float, ci_low: float, ci_high: float) -> str:
    if metric in HIGHER_IS_BETTER:
        if ci_low > 0:
            return "clear_advantage_over_coxph"
        if ci_high < 0:
            return "clear_disadvantage_vs_coxph"
        return "no_clear_difference"
    if metric in LOWER_IS_BETTER:
        if ci_high < 0:
            return "clear_advantage_over_coxph"
        if ci_low > 0:
            return "clear_disadvantage_vs_coxph"
        return "no_clear_difference"
    return "descriptive_only"


def compare_paired_tables(original_paired: pd.DataFrame, sensitivity_paired: pd.DataFrame) -> pd.DataFrame:
    base = original_paired.copy()
    sens = sensitivity_paired.copy()
    base["status_original"] = base.apply(
        lambda r: classify_paired_difference(r["Metric"], r["Mean_difference"], r["CI_low"], r["CI_high"]), axis=1
    )
    sens["status_sensitivity"] = sens.apply(
        lambda r: classify_paired_difference(r["Metric"], r["Mean_difference"], r["CI_low"], r["CI_high"]), axis=1
    )
    merged = base.merge(
        sens[["Model", "Metric", "Mean_difference", "CI_low", "CI_high", "status_sensitivity"]],
        on=["Model", "Metric"],
        suffixes=("_original", "_sensitivity"),
    )
    merged["status_changed"] = merged["status_original"] != merged["status_sensitivity"]
    return merged


def rank_models(metric: str, values: pd.Series) -> List[str]:
    sub = values.dropna().copy()
    if metric in HIGHER_IS_BETTER:
        return list(sub.sort_values(ascending=False).index)
    if metric in LOWER_IS_BETTER:
        return list(sub.sort_values(ascending=True).index)
    return list((sub - 1.0).abs().sort_values(ascending=True).index)


def build_winner_table(primary_df: pd.DataFrame, repeat_summary_df: pd.DataFrame) -> pd.DataFrame:
    primary_std = standardize_primary(primary_df).set_index("model")
    repeat_std = standardize_repeat_summary(repeat_summary_df).set_index("model")
    rows = []
    for metric in KEY_METRICS:
        primary_rank = rank_models(metric, primary_std[metric])
        repeat_rank = rank_models(metric, repeat_std[metric])
        rows.append(
            {
                "metric": metric,
                "primary_best_model": primary_rank[0] if primary_rank else np.nan,
                "repeated_best_model": repeat_rank[0] if repeat_rank else np.nan,
                "primary_ranking": " > ".join(primary_rank),
                "repeated_ranking": " > ".join(repeat_rank),
            }
        )
    return pd.DataFrame(rows)


def save_metadata(out_dir: Path, args: argparse.Namespace) -> None:
    payload = {
        "seed": args.seed,
        "test_size": args.test_size,
        "n_repeats": args.n_repeats,
        "m_imputations": args.m_imputations,
        "impute_max_iter": args.impute_max_iter,
        "horizon_months": args.horizon_months,
        "ibs_start_month": args.ibs_start_month,
        "calibration_groups": args.calibration_groups,
        "n_jobs": args.n_jobs,
        "skip_repeated": args.skip_repeated,
    }
    with open(out_dir / "metadata" / "sensitivity_analysis_config.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)


def main() -> None:
    args = parse_args()
    out_dir = Path(args.outdir)
    (out_dir / "tables").mkdir(parents=True, exist_ok=True)
    (out_dir / "metadata").mkdir(parents=True, exist_ok=True)
    original_script = Path(args.original_script)
    original_output_dir = Path(args.original_output_dir)
    csv_path = Path(args.csv)
    orig = load_original_module(original_script)
    orig.set_global_seed(args.seed)
    original_outputs = load_original_outputs(original_output_dir)
    df = orig.load_dataset(str(csv_path))
    _, sensitivity_primary = run_primary_sensitivity(
        orig=orig,
        df=df,
        tuning_df=original_outputs["tuning_primary"],
        out_dir=out_dir,
        seed=args.seed,
        test_size=args.test_size,
        m_imputations=args.m_imputations,
        impute_max_iter=args.impute_max_iter,
        horizon_months=args.horizon_months,
        ibs_start_month=args.ibs_start_month,
        calibration_groups=args.calibration_groups,
        n_jobs=args.n_jobs,
    )

    sensitivity_perf = None
    sensitivity_summary = None
    sensitivity_paired = None
    if not args.skip_repeated:
        sensitivity_perf, sensitivity_summary, sensitivity_paired = run_repeated_sensitivity(
            orig=orig,
            df=df,
            tuning_df=original_outputs["tuning_repeat"],
            out_dir=out_dir,
            seed=args.seed,
            n_repeats=args.n_repeats,
            test_size=args.test_size,
            m_imputations=args.m_imputations,
            impute_max_iter=args.impute_max_iter,
            horizon_months=args.horizon_months,
            ibs_start_month=args.ibs_start_month,
            calibration_groups=args.calibration_groups,
            n_jobs=args.n_jobs,
        )

    primary_comparison = compare_primary(original_outputs["primary"], sensitivity_primary)
    primary_comparison.to_csv(out_dir / "tables" / "comparison_original_vs_sensitivity_primary.csv", index=False)

    if sensitivity_summary is not None and sensitivity_paired is not None:
        repeat_summary_comparison = compare_repeat_summary(original_outputs["repeat_summary"], sensitivity_summary)
        repeat_summary_comparison.to_csv(out_dir / "tables" / "comparison_original_vs_sensitivity_repeated_summary.csv", index=False)

        paired_comparison = compare_paired_tables(original_outputs["paired"], sensitivity_paired)
        paired_comparison.to_csv(out_dir / "tables" / "comparison_original_vs_sensitivity_paired_vs_coxph.csv", index=False)

        winner_original = build_winner_table(original_outputs["primary"], original_outputs["repeat_summary"])
        winner_sensitivity = build_winner_table(sensitivity_primary, sensitivity_summary)
        winner_original.to_csv(out_dir / "tables" / "model_winners_original.csv", index=False)
        winner_sensitivity.to_csv(out_dir / "tables" / "model_winners_sensitivity.csv", index=False)

    save_metadata(out_dir, args)


if __name__ == "__main__":
    main()
