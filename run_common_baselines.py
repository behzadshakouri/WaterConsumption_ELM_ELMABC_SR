#!/usr/bin/env python3
"""
Run common baseline regressors on the same raw case workbooks used by the
water-consumption paper.

Models
------
MLR       Multiple linear regression
RF        Random forest regression
XGBoost   XGBoost regression (optional dependency: xgboost)
SVR       Support-vector regression with an RBF kernel
GWR       Geographically weighted regression (optional dependency: mgwr)

Expected project layout
-----------------------
PROJECT_ROOT/
    data/                         raw case workbooks, recursively discovered
    prepare_paper_results_v8.py   optional; not imported by this script

Each raw workbook is expected to have a header row. By default:
    first column = grid-cell identifier
    columns 7-15 = LST, land-value, and elevation summaries (resolved by name)
    column 16 / last column = case-specific NHWSCC or annual-HWC target

Column positions can also be selected explicitly. Positions are one-based, as
shown in Excel. For example:
    --input-columns 1,4,6 --output-column 8

IMPORTANT: use --predictors to explicitly identify the predictors used in the
paper whenever possible. This prevents accidental inclusion of unrelated
columns. For GWR, provide --x-column and --y-column or use recognizable names
such as x/y, lon/lat, longitude/latitude, easting/northing.

Examples
--------
python run_common_baselines.py --root . --predictors LST LandValue Elevation

python run_common_baselines.py --root . \
    --predictors LST LandValue Elevation \
    --models MLR RF XGBoost SVR GWR \
    --x-column X --y-column Y

Reproduce the original row-order 70/30 split and also save 600/700 m results:
python run_common_baselines.py --root . --split original --best-meshes 600 700

Optional spatial-block split:
python run_common_baselines.py --root . --split spatial-block \
    --x-column X --y-column Y --spatial-blocks 5

Outputs
-------
baseline_results/
    all_cases_metrics.xlsx
    all_cases_metrics.csv
    model_comparison_summary.csv
    audit.csv
    run_configuration.json
    all_meshes/<MODEL>/<CASE>/
        Values.xlsx
        Statistics.xlsx
        residuals.csv
        split_assignments.csv
    best_meshes_600_700/<MODEL>/<CASE>/
        copies of the same case outputs for 600 m and 700 m only

The script never uses ELM or ELM-ABC predictions as inputs. It uses only the
raw predictors and observed targets, thereby producing genuine baselines.
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import math
import re
import shutil
import sys
import traceback
import warnings
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Iterable, Optional

import numpy as np
import pandas as pd

EPS = 1e-12
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
MODEL_ORDER = ["MLR", "RF", "XGBoost", "SVR", "GWR"]
RANDOM_STATE = 42
PAPER_MEAN_PREDICTORS=["T96_mean","P_mean","Elv_mean"]
PAPER_SUMMARY_PREDICTORS=["T96_mean","T96_min","T96_max","P_mean","P_min","P_max","Elv_mean","Elv_min","Elv_max"]
UNSAFE_PREDICTORS={"id","fid","objectid","grid_id","cell_id","left","top","right","bottom","x","y","lon","lat","longitude","latitude","easting","northing","centroid_x","centroid_y","xcoord","ycoord","x_coord","y_coord"}


@dataclass
class CaseResult:
    case_id: str
    raw_file: str
    relative_path: str
    grid_m: Optional[int]
    period: str
    target: str
    target_variable: str
    model: str
    split_strategy: str
    split_random_state: int
    test_fraction_requested: float
    n_total: int
    n_train: int
    n_test: int
    predictors: str
    coordinate_x_source: str
    coordinate_y_source: str
    unsafe_predictors_allowed: bool
    status: str
    error: str = ""


def natural_key(value: Any) -> list[Any]:
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", str(value))]


def safe_name(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value).strip())
    return value.strip("._") or "case"


def infer_grid(path: Path) -> Optional[int]:
    m = re.search(r"mesh\s*[-_ ]?(\d+)", str(path), flags=re.I)
    return int(m.group(1)) if m else None


def infer_period(path: Path) -> str:
    text = str(path).lower()
    if any(t in text for t in ["2-ybp", "2_ybp", "2ybp", "before", "normal", "2018", "2019"]):
        return "2-YBP"
    if any(t in text for t in ["1-yap", "1_yap", "1yap", "pandemic", "covid", "2020", "2021", "after"]):
        return "1-YAP"
    return "Unspecified"


def infer_paper_target(target_col: str, path: Path) -> str:
    """Classify the paper's case-specific output without treating every case as HWC."""
    text = f"{target_col} {path.stem}".strip().lower()
    if re.search(r"(?:^|[^a-z])nc\s*\d+", text) or "nhwscc" in text:
        return "NHWSCC"
    if any(token in text for token in ("annual", "yearly", "hwc", "consumption", "sum")):
        return "Annual HWC"
    # In the paper workbooks, NC-prefixed targets are NHWSCC cases and the
    # remaining case-specific final-column targets are annual HWC.
    return "Annual HWC"


def read_first_nonempty_sheet(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    last_error: Optional[Exception] = None
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=0)
            if not df.empty:
                return df
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise ValueError(f"No non-empty worksheet found in {path}")


def resolve_column(df: pd.DataFrame, requested: Optional[str], aliases: Iterable[str]) -> Optional[str]:
    lookup = {str(c).strip().lower(): str(c) for c in df.columns}
    if requested:
        if requested in df.columns:
            return requested
        hit = lookup.get(requested.strip().lower())
        if hit:
            return hit
        raise KeyError(f"Column {requested!r} not found. Available columns: {list(df.columns)}")
    for alias in aliases:
        if alias.lower() in lookup:
            return lookup[alias.lower()]
    return None


def parse_column_positions(values: Optional[Iterable[str]], option_name: str) -> Optional[list[int]]:
    if not values:
        return None
    positions: list[int] = []
    for value in values:
        for token in str(value).split(","):
            token = token.strip()
            if not token:
                continue
            try:
                position = int(token)
            except ValueError as exc:
                raise ValueError(
                    f"{option_name} must contain one-based integer column numbers; got {token!r}."
                ) from exc
            if position < 1:
                raise ValueError(f"{option_name} positions must be at least 1; got {position}.")
            positions.append(position)
    if not positions:
        raise ValueError(f"{option_name} did not contain any column numbers.")
    return list(dict.fromkeys(positions))


def column_at_position(df: pd.DataFrame, position: int, option_name: str) -> str:
    if position > len(df.columns):
        raise ValueError(
            f"{option_name}={position} is outside this workbook, which has "
            f"{len(df.columns)} columns."
        )
    return str(df.columns[position - 1])


def select_columns(
    df,
    predictor_names,
    input_positions,
    id_name,
    target_name,
    output_position,
    predictor_set="paper-summary",
    allow_unsafe=False,
):
    id_col=resolve_column(df,id_name,["fid","id","objectid","grid_id","cell_id"]) or str(df.columns[0])
    if output_position is not None:
        target_col=column_at_position(df,output_position,"--output-column")
    else:
        target_col=resolve_column(df,target_name,["target","observed","observation","y"]) or str(df.columns[-1])
    if input_positions:
        requested=[column_at_position(df,n,"--input-columns") for n in input_positions]
    elif predictor_names: requested=predictor_names
    elif predictor_set=="paper-mean": requested=PAPER_MEAN_PREDICTORS
    elif predictor_set=="paper-summary": requested=PAPER_SUMMARY_PREDICTORS
    else: requested=[str(c) for c in df.columns if str(c).lower() not in UNSAFE_PREDICTORS|{id_col.lower(),target_col.lower()} and pd.to_numeric(df[c],errors="coerce").notna().any()]
    predictors=[]
    for n in requested: predictors.append(resolve_column(df,n,[]))
    predictors=list(dict.fromkeys(predictors))
    if target_col in predictors:
        raise ValueError(
            f"Output column {target_col!r} cannot also be an input. "
            "Choose different --input-columns/--output-column values."
        )
    unsafe=[x for x in predictors if x.lower() in UNSAFE_PREDICTORS|{id_col.lower()}]
    if unsafe and not allow_unsafe and not input_positions:
        raise ValueError(f"Unsafe predictor leakage prevented: {unsafe}")
    if not predictors: raise ValueError("No valid predictors")
    return id_col,target_col,predictors

def derive_coordinates(df,x_requested,y_requested):
    x=resolve_column(df,x_requested,["centroid_x","x","easting","longitude","lon","xcoord","x_coord"]); y=resolve_column(df,y_requested,["centroid_y","y","northing","latitude","lat","ycoord","y_coord"])
    if x and y: return df[[x,y]].apply(pd.to_numeric,errors="coerce").to_numpy(float),x,y
    l=resolve_column(df,None,["left"]); r=resolve_column(df,None,["right"]); t=resolve_column(df,None,["top"]); bot=resolve_column(df,None,["bottom"])
    if l and r and t and bot:
        cx=(pd.to_numeric(df[l],errors="coerce")+pd.to_numeric(df[r],errors="coerce"))/2; cy=(pd.to_numeric(df[t],errors="coerce")+pd.to_numeric(df[bot],errors="coerce"))/2
        return np.column_stack([cx,cy]),"derived_centroid_x","derived_centroid_y"
    return None,"",""


def calculate_metrics(y: np.ndarray, pred: np.ndarray, prefix: str) -> dict[str, float]:
    mask = np.isfinite(y) & np.isfinite(pred)
    y = np.asarray(y, dtype=float)[mask]
    pred = np.asarray(pred, dtype=float)[mask]
    names = [
        "n", "r2_corr", "r2_standard", "rmse", "mae", "nrmse_mean",
        "nrmse_range", "nrmse_sd", "bias", "pbias", "nse", "pearson_r",
        "residual_sd", "within10_all", "within10_nonzero", "observed_mean",
        "predicted_mean",
    ]
    if not len(y):
        return {f"{prefix}_{n}": np.nan for n in names}
    err = pred - y
    sse = float(np.sum(err ** 2))
    sst = float(np.sum((y - np.mean(y)) ** 2))
    rmse = float(np.sqrt(np.mean(err ** 2)))
    mae = float(np.mean(np.abs(err)))
    ymean = float(np.mean(y))
    yrange = float(np.max(y) - np.min(y))
    ysd = float(np.std(y, ddof=1)) if len(y) > 1 else np.nan
    if len(y) > 1 and np.std(y) > EPS and np.std(pred) > EPS:
        pearson = float(np.corrcoef(y, pred)[0, 1])
        r2corr = pearson ** 2
    else:
        pearson = np.nan
        r2corr = np.nan
    nonzero = np.abs(y) > EPS
    within_all = np.abs(err) <= 0.10 * np.abs(y)
    return {
        f"{prefix}_n": int(len(y)),
        f"{prefix}_r2_corr": r2corr,
        f"{prefix}_r2_standard": 1.0 - sse / sst if sst > EPS else np.nan,
        f"{prefix}_rmse": rmse,
        f"{prefix}_mae": mae,
        f"{prefix}_nrmse_mean": rmse / abs(ymean) if abs(ymean) > EPS else np.nan,
        f"{prefix}_nrmse_range": rmse / yrange if yrange > EPS else np.nan,
        f"{prefix}_nrmse_sd": rmse / ysd if np.isfinite(ysd) and ysd > EPS else np.nan,
        f"{prefix}_bias": float(np.mean(err)),
        f"{prefix}_pbias": 100.0 * float(np.sum(err)) / float(np.sum(y)) if abs(np.sum(y)) > EPS else np.nan,
        f"{prefix}_nse": 1.0 - sse / sst if sst > EPS else np.nan,
        f"{prefix}_pearson_r": pearson,
        f"{prefix}_residual_sd": float(np.std(err, ddof=1)) if len(err) > 1 else np.nan,
        f"{prefix}_within10_all": float(np.mean(within_all)),
        f"{prefix}_within10_nonzero": float(np.mean(within_all[nonzero])) if np.any(nonzero) else np.nan,
        f"{prefix}_observed_mean": ymean,
        f"{prefix}_predicted_mean": float(np.mean(pred)),
    }


def make_split(
    n: int,
    strategy: str,
    test_size: float,
    random_state: int,
    coords: Optional[np.ndarray],
    spatial_blocks: int,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    from sklearn.model_selection import train_test_split, GroupShuffleSplit
    indices = np.arange(n)
    groups = np.full(n, -1, dtype=int)
    if strategy == "original":
        n_train = int(math.ceil((1.0 - test_size) * n))
        train_idx, test_idx = indices[:n_train], indices[n_train:]
    elif strategy == "random":
        train_idx, test_idx = train_test_split(indices, test_size=test_size, random_state=random_state)
    elif strategy == "spatial-block":
        if coords is None:
            raise ValueError("spatial-block split requires coordinate columns")
        x, y = coords[:, 0], coords[:, 1]
        # Quantile blocks are robust to irregular city extents and uneven point density.
        xb = pd.qcut(pd.Series(x), q=min(spatial_blocks, len(np.unique(x))), labels=False, duplicates="drop").to_numpy()
        yb = pd.qcut(pd.Series(y), q=min(spatial_blocks, len(np.unique(y))), labels=False, duplicates="drop").to_numpy()
        groups = xb * (np.nanmax(yb) + 1) + yb
        splitter = GroupShuffleSplit(n_splits=1, test_size=test_size, random_state=random_state)
        train_idx, test_idx = next(splitter.split(indices, groups=groups))
    else:
        raise ValueError(f"Unsupported split strategy: {strategy}")
    return np.asarray(train_idx), np.asarray(test_idx), groups


def build_model(name: str, random_state: int, n_jobs: int, args: argparse.Namespace):
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer
    from sklearn.preprocessing import StandardScaler
    from sklearn.linear_model import LinearRegression
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.svm import SVR

    if name == "MLR":
        return Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("model", LinearRegression()),
        ])
    if name == "RF":
        return Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", RandomForestRegressor(
                n_estimators=args.rf_trees,
                max_depth=args.rf_max_depth,
                min_samples_leaf=args.rf_min_samples_leaf,
                max_features=args.rf_max_features,
                random_state=random_state,
                n_jobs=n_jobs,
            )),
        ])
    if name == "SVR":
        return Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("model", SVR(kernel="rbf", C=args.svr_c, epsilon=args.svr_epsilon, gamma=args.svr_gamma)),
        ])
    if name == "XGBoost":
        try:
            from xgboost import XGBRegressor
        except ImportError as exc:
            raise ImportError("XGBoost requested but xgboost is not installed. Run: pip install xgboost") from exc
        return Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("model", XGBRegressor(
                n_estimators=args.xgb_trees,
                learning_rate=args.xgb_learning_rate,
                max_depth=args.xgb_max_depth,
                min_child_weight=args.xgb_min_child_weight,
                subsample=args.xgb_subsample,
                colsample_bytree=args.xgb_colsample,
                objective="reg:squarederror",
                eval_metric="rmse",
                random_state=random_state,
                n_jobs=n_jobs,
                reg_lambda=args.xgb_reg_lambda,
            )),
        ])
    raise ValueError(name)


def fit_predict_gwr(
    X_train: np.ndarray,
    y_train: np.ndarray,
    coords_train: np.ndarray,
    X_test: np.ndarray,
    coords_test: np.ndarray,
    args: argparse.Namespace,
) -> tuple[np.ndarray, np.ndarray, dict[str, Any]]:
    try:
        from mgwr.gwr import GWR
        from mgwr.sel_bw import Sel_BW
    except ImportError as exc:
        raise ImportError("GWR requested but mgwr is not installed. Run: pip install mgwr") from exc
    from sklearn.impute import SimpleImputer
    from sklearn.preprocessing import StandardScaler

    imp = SimpleImputer(strategy="median")
    scaler = StandardScaler()
    Xtr = scaler.fit_transform(imp.fit_transform(X_train))
    Xte = scaler.transform(imp.transform(X_test))
    ytr = np.asarray(y_train, dtype=float).reshape((-1, 1))
    ctr = np.asarray(coords_train, dtype=float)
    cte = np.asarray(coords_test, dtype=float)

    if len(ytr) < max(8, Xtr.shape[1] + 3):
        raise ValueError("Too few training observations for GWR")

    if args.gwr_bandwidth is not None:
        bw = float(args.gwr_bandwidth)
    else:
        selector = Sel_BW(ctr, ytr, Xtr, fixed=args.gwr_fixed, spherical=args.gwr_spherical)
        bw = selector.search(criterion=args.gwr_criterion)
    model = GWR(ctr, ytr, Xtr, bw=bw, fixed=args.gwr_fixed, kernel=args.gwr_kernel, spherical=args.gwr_spherical)
    fit = model.fit()
    train_pred = np.asarray(fit.predy).reshape(-1)
    prediction = model.predict(cte, Xte, exog_scale=fit.scale, exog_resid=fit.resid_response)
    test_pred = np.asarray(prediction.predictions).reshape(-1)
    info = {"gwr_bandwidth": float(bw), "gwr_aicc": float(fit.aicc), "gwr_enp": float(fit.ENP)}
    return train_pred, test_pred, info


def run_kfold_cross_validation(model_name,Xtr,ytr,ctr,args):
    from sklearn.model_selection import KFold
    k=min(args.cv_folds,len(ytr));
    if k<2: raise ValueError("Cross-validation requires at least 2 training observations")
    rows=[]; oof=np.full(len(ytr),np.nan); splitter=KFold(k,shuffle=True,random_state=args.random_state)
    for fold,(fi,vi) in enumerate(splitter.split(Xtr),1):
        extra={}
        if model_name=="GWR":
            if ctr is None: raise ValueError("GWR CV requires coordinates")
            fp,vp,extra=fit_predict_gwr(Xtr.iloc[fi].to_numpy(),ytr[fi],ctr[fi],Xtr.iloc[vi].to_numpy(),ctr[vi],args)
        else:
            m=build_model(model_name,args.random_state+fold,args.n_jobs,args); m.fit(Xtr.iloc[fi],ytr[fi]); fp=np.asarray(m.predict(Xtr.iloc[fi]),float); vp=np.asarray(m.predict(Xtr.iloc[vi]),float)
        if args.clip_negative_predictions: fp=np.maximum(fp,0); vp=np.maximum(vp,0)
        oof[vi]=vp; rows.append({"fold":fold,"n_fold_train":len(fi),"n_fold_validation":len(vi),**extra,**calculate_metrics(ytr[fi],fp,"fold_train"),**calculate_metrics(ytr[vi],vp,"fold_validation")})
    folds=pd.DataFrame(rows); summary={"cv_folds_requested":args.cv_folds,"cv_folds_used":k,"cv_shuffle":True,"cv_random_state":args.random_state,**calculate_metrics(ytr,oof,"cv_oof")}
    for metric in ["r2_corr","r2_standard","rmse","mae","nrmse_mean","nrmse_range","nrmse_sd","bias","pbias","nse","pearson_r","residual_sd","within10_all","within10_nonzero"]:
        z=pd.to_numeric(folds[f"fold_validation_{metric}"],errors="coerce"); summary[f"cv_fold_{metric}_mean"]=z.mean(); summary[f"cv_fold_{metric}_sd"]=z.std(ddof=1); summary[f"cv_fold_{metric}_median"]=z.median()
    return folds,summary

def model_parameters(model_name: str, fitted: Any, feature_names: list[str]) -> pd.DataFrame:
    try:
        est = fitted.named_steps["model"] if hasattr(fitted, "named_steps") else fitted
        if model_name == "MLR":
            rows = [{"parameter": "intercept_scaled_space", "value": float(est.intercept_)}]
            rows += [{"parameter": f"coefficient_scaled_{n}", "value": float(v)} for n, v in zip(feature_names, est.coef_)]
            return pd.DataFrame(rows)
        if model_name in {"RF", "XGBoost"} and hasattr(est, "feature_importances_"):
            return pd.DataFrame({"feature": feature_names, "importance": est.feature_importances_}).sort_values("importance", ascending=False)
        if model_name == "SVR":
            return pd.DataFrame([{"parameter": "n_support_vectors", "value": int(np.sum(est.n_support_))}])
    except Exception:
        pass
    return pd.DataFrame()


def write_case_outputs(
    case_dir: Path,
    ids: pd.Series,
    observed: np.ndarray,
    all_pred: np.ndarray,
    split_labels: np.ndarray,
    groups: np.ndarray,
    X: pd.DataFrame,
    metrics: dict[str, Any],
    params: pd.DataFrame,
    cv_folds: pd.DataFrame,
) -> None:
    case_dir.mkdir(parents=True, exist_ok=True)
    values = pd.DataFrame({
        "id": ids,
        **{c: X[c].to_numpy() for c in X.columns},
        "observed": observed,
        "predicted": all_pred,
        "residual_pred_minus_obs": all_pred - observed,
        "set": split_labels,
        "spatial_block": groups,
    })
    values.to_csv(case_dir / "residuals.csv", index=False, encoding="utf-8-sig")
    values[["id", "set", "spatial_block"]].to_csv(case_dir / "split_assignments.csv", index=False, encoding="utf-8-sig")
    stats_df = pd.DataFrame([metrics])
    with pd.ExcelWriter(case_dir / "Values.xlsx", engine="openpyxl") as writer:
        values.to_excel(writer, index=False, sheet_name="Values")
    with pd.ExcelWriter(case_dir / "Statistics.xlsx", engine="openpyxl") as writer:
        stats_df.to_excel(writer, index=False, sheet_name="Statistics")
        if not params.empty: params.to_excel(writer,index=False,sheet_name="Model Parameters")
        if not cv_folds.empty: cv_folds.to_excel(writer,index=False,sheet_name="10-Fold CV")
    if not cv_folds.empty: cv_folds.to_csv(case_dir/"cross_validation_folds.csv",index=False,encoding="utf-8-sig")


def copy_to_best_meshes(src: Path, best_root: Path, model_name: str, case_name: str, grid: Optional[int], best_meshes: set[int]) -> None:
    if grid not in best_meshes:
        return
    dst = best_root / model_name / case_name
    if dst.exists():
        shutil.rmtree(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(src, dst)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run MLR, RF, XGBoost, SVR, and GWR baselines on raw paper workbooks.")
    p.add_argument("--root", type=Path, default=Path.cwd())
    p.add_argument("--data-folder", default="data")
    p.add_argument("--output", type=Path, default=Path("baseline_results"))
    p.add_argument("--file-pattern", action="append", default=[],
                   help="Process only relative paths or filenames matching this glob; repeat as needed.")
    p.add_argument("--models", nargs="+", default=MODEL_ORDER, choices=MODEL_ORDER)
    p.add_argument("--predictors", nargs="+", default=None, help="Exact predictor column names. Strongly recommended.")
    p.add_argument(
        "--input-columns", nargs="+", default=None, metavar="N",
        help="One-based Excel column numbers used as inputs. Accepts spaces and/or commas; overrides --predictors and --predictor-set.",
    )
    p.add_argument("--predictor-set",choices=["paper-mean","paper-summary","numeric-safe"],default="paper-summary")
    p.add_argument("--allow-unsafe-predictors",action="store_true")
    p.add_argument("--clip-negative-predictions",action="store_true")
    p.add_argument("--id-column", default=None)
    p.add_argument("--target-column", default=None, help="Use only if every workbook has the same target header. Default: final column.")
    p.add_argument(
        "--output-column", type=int, default=None, metavar="N",
        help="One-based Excel column number used as the output; overrides --target-column. Default: final column.",
    )
    p.add_argument("--x-column", default=None, help="Coordinate x/easting/longitude column for GWR or spatial-block split.")
    p.add_argument("--y-column", default=None, help="Coordinate y/northing/latitude column for GWR or spatial-block split.")
    p.add_argument("--split", choices=["original", "random", "spatial-block"], default="random")
    p.add_argument("--test-size",type=float,default=0.30)
    p.add_argument("--cv-folds",type=int,default=10)
    p.add_argument("--skip-cross-validation",action="store_true")
    p.add_argument("--random-state", type=int, default=RANDOM_STATE)
    p.add_argument("--spatial-blocks", type=int, default=5)
    p.add_argument("--best-meshes", nargs="+", type=int, default=[600, 700])
    p.add_argument("--n-jobs", type=int, default=-1)
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--fail-on-missing-model", action="store_true")
    # RF
    p.add_argument("--rf-trees", type=int, default=500)
    p.add_argument("--rf-max-depth", type=int, default=None)
    p.add_argument("--rf-min-samples-leaf", type=int, default=1)
    p.add_argument("--rf-max-features", default="sqrt")
    # XGBoost
    p.add_argument("--xgb-trees", type=int, default=300)
    p.add_argument("--xgb-learning-rate", type=float, default=0.03)
    p.add_argument("--xgb-max-depth", type=int, default=2)
    p.add_argument("--xgb-min-child-weight", type=float, default=5.0)
    p.add_argument("--xgb-subsample", type=float, default=0.8)
    p.add_argument("--xgb-colsample", type=float, default=0.8)
    p.add_argument("--xgb-reg-lambda", type=float, default=5.0)
    # SVR
    p.add_argument("--svr-c", type=float, default=10.0)
    p.add_argument("--svr-epsilon", type=float, default=0.1)
    p.add_argument("--svr-gamma", default="scale")
    # GWR
    p.add_argument("--gwr-bandwidth", type=float, default=None)
    p.add_argument("--gwr-fixed", action="store_true", help="Use distance bandwidth. Default is adaptive nearest-neighbor bandwidth.")
    p.add_argument("--gwr-spherical", action="store_true", help="Coordinates are longitude/latitude in degrees.")
    p.add_argument("--gwr-kernel", choices=["gaussian", "bisquare", "exponential"], default="bisquare")
    p.add_argument("--gwr-criterion", choices=["AICc", "AIC", "BIC", "CV"], default="AICc")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    args.input_columns = parse_column_positions(args.input_columns, "--input-columns")
    if args.output_column is not None and args.output_column < 1:
        raise ValueError("--output-column must be at least 1.")
    root = args.root.expanduser().resolve()
    data_root = root / args.data_folder
    output = args.output if args.output.is_absolute() else root / args.output
    all_root = output / "all_meshes"
    best_root = output / "best_meshes_600_700"
    if not data_root.exists():
        raise FileNotFoundError(f"Raw data folder not found: {data_root}")
    if output.exists() and args.overwrite:
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)

    files = sorted([
        p for p in data_root.rglob("*")
        if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
    ], key=natural_key)
    if args.file_pattern:
        files = [
            path for path in files
            if any(
                fnmatch.fnmatch(path.relative_to(data_root).as_posix(), pattern)
                or fnmatch.fnmatch(path.name, pattern)
                for pattern in args.file_pattern
            )
        ]
    if not files:
        pattern_note = f" matching {args.file_pattern}" if args.file_pattern else ""
        raise FileNotFoundError(f"No Excel workbooks found under {data_root}{pattern_note}")

    config = vars(args).copy()
    config["root"] = str(root); config["output"] = str(output)
    (output / "run_configuration.json").write_text(json.dumps(config, indent=2, default=str), encoding="utf-8")

    metric_rows: list[dict[str, Any]] = []
    audit_rows: list[dict[str, Any]] = []
    best_meshes = set(args.best_meshes)

    for file_no, raw_path in enumerate(files, 1):
        rel = raw_path.relative_to(data_root)
        case_name = safe_name(str(rel.with_suffix("")))
        grid = infer_grid(raw_path)
        print(f"[{file_no}/{len(files)}] {rel}")
        try:
            df = read_first_nonempty_sheet(raw_path)
            id_col,target_col,predictors=select_columns(
                df,args.predictors,args.input_columns,args.id_column,args.target_column,
                args.output_column,args.predictor_set,args.allow_unsafe_predictors
            )
            target_variable = infer_paper_target(str(target_col), raw_path)
            y_series = pd.to_numeric(df[target_col], errors="coerce")
            X = df[predictors].apply(pd.to_numeric, errors="coerce")
            ids = df[id_col]

            coords,coordinate_x_source,coordinate_y_source=derive_coordinates(df,args.x_column,args.y_column)
            valid = y_series.notna().to_numpy() & X.notna().any(axis=1).to_numpy()
            if coords is not None and ("GWR" in args.models or args.split == "spatial-block"):
                valid &= np.isfinite(coords).all(axis=1)
            if valid.sum() < 5:
                raise ValueError(f"Only {int(valid.sum())} usable observations")
            df = df.loc[valid].reset_index(drop=True)
            ids = ids.loc[valid].reset_index(drop=True)
            X = X.loc[valid].reset_index(drop=True)
            y = y_series.loc[valid].to_numpy(dtype=float)
            if coords is not None:
                coords = coords[valid]

            train_idx, test_idx, groups = make_split(len(y), args.split, args.test_size, args.random_state, coords, args.spatial_blocks)
            split_labels = np.array(["Unused"] * len(y), dtype=object)
            split_labels[train_idx] = "Train"; split_labels[test_idx] = "Test"

            for model_name in args.models:
                base_result = CaseResult(
                    case_id=case_name,
                    raw_file=str(raw_path),
                    relative_path=str(rel),
                    grid_m=grid,
                    period=infer_period(raw_path),
                    target=str(target_col),
                    target_variable=target_variable,
                    model=model_name,
                    split_strategy=args.split,
                    split_random_state=args.random_state,
                    test_fraction_requested=args.test_size,
                    n_total=len(y),
                    n_train=len(train_idx),
                    n_test=len(test_idx),
                    predictors=" | ".join(predictors),
                    coordinate_x_source=coordinate_x_source,
                    coordinate_y_source=coordinate_y_source,
                    unsafe_predictors_allowed=args.allow_unsafe_predictors,
                    status="OK",
                )
                case_dir = all_root / model_name / case_name
                try:
                    extra={}; params=pd.DataFrame(); cv_folds=pd.DataFrame(); cv_summary={"cv_folds_requested":args.cv_folds,"cv_folds_used":0,"cv_skipped":args.skip_cross_validation}
                    if not args.skip_cross_validation:
                        cv_folds,cv_summary=run_kfold_cross_validation(model_name,X.iloc[train_idx].reset_index(drop=True),y[train_idx],coords[train_idx] if coords is not None else None,args); cv_summary["cv_skipped"]=False
                    if model_name == "GWR":
                        if coords is None:
                            raise ValueError("GWR requires coordinate columns. Use --x-column and --y-column.")
                        train_pred, test_pred, extra = fit_predict_gwr(
                            X.iloc[train_idx].to_numpy(), y[train_idx], coords[train_idx],
                            X.iloc[test_idx].to_numpy(), coords[test_idx], args,
                        )
                    else:
                        fitted = build_model(model_name, args.random_state, args.n_jobs, args)
                        fitted.fit(X.iloc[train_idx], y[train_idx])
                        train_pred = np.asarray(fitted.predict(X.iloc[train_idx]), dtype=float)
                        test_pred = np.asarray(fitted.predict(X.iloc[test_idx]), dtype=float)
                        params = model_parameters(model_name, fitted, predictors)

                    if args.clip_negative_predictions: train_pred=np.maximum(train_pred,0); test_pred=np.maximum(test_pred,0)
                    all_pred = np.full(len(y), np.nan)
                    all_pred[train_idx] = train_pred; all_pred[test_idx] = test_pred
                    metrics = {
                        **asdict(base_result),
                        **extra,
                        **cv_summary,
                        **calculate_metrics(y[train_idx], train_pred, "train"),
                        **calculate_metrics(y[test_idx], test_pred, "test"),
                        **calculate_metrics(y, all_pred, "all"),
                    }
                    metrics["generalization_gap_r2_corr"]=metrics["train_r2_corr"]-metrics["test_r2_corr"]; metrics["test_train_rmse_ratio"]=metrics["test_rmse"]/metrics["train_rmse"] if metrics["train_rmse"]>EPS else np.nan; metrics["overfitting_flag"]=bool(metrics["generalization_gap_r2_corr"]>0.25 or metrics["test_train_rmse_ratio"]>3)
                    write_case_outputs(case_dir,ids,y,all_pred,split_labels,groups,X,metrics,params,cv_folds)
                    copy_to_best_meshes(case_dir, best_root, model_name, case_name, grid, best_meshes)
                    metric_rows.append(metrics)
                except Exception as exc:
                    base_result.status = "ERROR"
                    base_result.error = f"{type(exc).__name__}: {exc}"
                    metric_rows.append(asdict(base_result))
                    audit_rows.append({
                        "severity": "ERROR", "case": case_name, "model": model_name,
                        "file": str(raw_path), "message": base_result.error,
                        "traceback": traceback.format_exc(),
                    })
                    print(f"  {model_name}: ERROR: {exc}", file=sys.stderr)
                    if args.fail_on_missing_model and isinstance(exc, ImportError):
                        raise
        except Exception as exc:
            audit_rows.append({
                "severity": "ERROR", "case": case_name, "model": "ALL",
                "file": str(raw_path), "message": f"{type(exc).__name__}: {exc}",
                "traceback": traceback.format_exc(),
            })
            print(f"  CASE ERROR: {exc}", file=sys.stderr)

    metrics_df = pd.DataFrame(metric_rows)
    audit_df = pd.DataFrame(audit_rows, columns=["severity", "case", "model", "file", "message", "traceback"])
    metrics_df.to_csv(output / "all_cases_metrics.csv", index=False, encoding="utf-8-sig")
    audit_df.to_csv(output / "audit.csv", index=False, encoding="utf-8-sig")

    if not metrics_df.empty:
        ok = metrics_df[metrics_df.get("status", "") == "OK"].copy()
        summary_metrics = [c for c in [
            "test_r2_corr", "test_r2_standard", "test_rmse", "test_mae",
            "test_nrmse_mean", "test_nrmse_range", "test_nrmse_sd", "test_bias",
            "test_pbias", "test_nse", "test_pearson_r", "test_within10_nonzero",
        ] if c in ok.columns]
        if len(ok):
            summary = ok.groupby(["model", "grid_m"], dropna=False)[summary_metrics].agg(["count", "mean", "median", "std"]).reset_index()
            summary.columns = ["_".join([str(x) for x in c if str(x)]) if isinstance(c, tuple) else str(c) for c in summary.columns]
        else:
            summary = pd.DataFrame()
        summary.to_csv(output / "model_comparison_summary.csv", index=False, encoding="utf-8-sig")
        best = ok[ok["grid_m"].isin(best_meshes)].copy() if "grid_m" in ok else pd.DataFrame()
        best.to_csv(output / "best_meshes_600_700_metrics.csv", index=False, encoding="utf-8-sig")
        with pd.ExcelWriter(output / "all_cases_metrics.xlsx", engine="openpyxl") as writer:
            metrics_df.to_excel(writer, index=False, sheet_name="All Case Metrics")
            summary.to_excel(writer, index=False, sheet_name="Model Grid Summary")
            best.to_excel(writer, index=False, sheet_name="Meshes 600 700")
            audit_df.to_excel(writer, index=False, sheet_name="Audit")

    readme = f"""Common baseline modelling output

Raw files discovered: {len(files)}
Models requested: {', '.join(args.models)}
Split strategy: {args.split}
Best meshes copied separately: {sorted(best_meshes)}
Successful model-case runs: {sum(1 for r in metric_rows if r.get('status') == 'OK')}
Failed model-case runs: {len(audit_rows)}

Fair-comparison rules
- Only raw predictors and observed targets were used.
- The default 'original' split preserves the first 70% training / last 30% test convention.
- Scalers and imputers were fitted on training data only.
- GWR and spatial-block validation require coordinates.
- Explicit --predictors is strongly recommended to avoid predictor leakage.

Metric definitions
- r2_corr: squared Pearson correlation, matching the manuscript convention.
- r2_standard / NSE: 1 - SSE/SST.
- PBIAS: 100 * sum(predicted - observed) / sum(observed).
- within10_nonzero excludes zero observations from relative-error evaluation.
"""
    (output / "README.txt").write_text(readme, encoding="utf-8")
    print(f"\nFinished. Results: {output}")
    print(f"Successful model-case runs: {sum(1 for r in metric_rows if r.get('status') == 'OK')}")
    print(f"Audit errors: {len(audit_rows)}")
    return 0 if metric_rows else 1


if __name__ == "__main__":
    warnings.filterwarnings("ignore", category=RuntimeWarning)
    raise SystemExit(main())
