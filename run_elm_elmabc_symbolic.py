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

Run the paper's reproducible random 70/30 split and also save 600/700 m results:
python run_common_baselines.py --root . --split random --best-meshes 600 700

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

The script uses only the
raw predictors and observed targets. ELM and ELM-ABC follow the attached MATLAB design; symbolic regression is independently fitted.
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

import warnings

warnings.filterwarnings(
    "ignore",
    message=r"`BaseEstimator\._validate_data` is deprecated",
    category=FutureWarning,
)

EPS = 1e-12
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
MODEL_ORDER = ["ELM", "ELMABC", "SymbolicRegression"]
RANDOM_STATE = 42
PAPER_MEAN_PREDICTORS=["T96_mean","P_mean","Elv_mean"]
PAPER_SUMMARY_PREDICTORS=["T96_mean","T96_min","T96_max","P_mean","P_min","P_max","Elv_mean","Elv_min","Elv_max"]
UNSAFE_PREDICTORS={"id","fid","objectid","grid_id","cell_id","left","top","right","bottom","x","y","lon","lat","longitude","latitude","easting","northing","centroid_x","centroid_y","xcoord","ycoord","x_coord","y_coord"}
SR_ARITY = {
    "add": 2, "sub": 2, "mul": 2, "div": 2, "max": 2, "min": 2,
    "sqrt": 1, "log": 1, "abs": 1, "neg": 1,
}


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


class ELMRegressor:
    """Single-hidden-layer Extreme Learning Machine regressor.

    Inputs and target are scaled to [-1, 1] using training data only. Hidden
    weights and biases are random and fixed; output weights are obtained with
    a ridge-stabilized Moore-Penrose solution.
    """
    def __init__(self, hidden_neurons=10, activation="sigmoid", alpha=1e-8,
                 random_state=42, trials=100, validation_size=0.20):
        self.hidden_neurons = int(hidden_neurons)
        self.activation = activation
        self.alpha = float(alpha)
        self.random_state = int(random_state)
        self.trials = int(trials)
        self.validation_size = float(validation_size)

    def _activate(self, z):
        z = np.clip(z, -60.0, 60.0)
        if self.activation == "sigmoid": return 1.0 / (1.0 + np.exp(-z))
        if self.activation == "tanh": return np.tanh(z)
        if self.activation == "sin": return np.sin(z)
        if self.activation == "radbas": return np.exp(-(z ** 2))
        if self.activation == "hardlim": return (z >= 0.0).astype(float)
        if self.activation == "tribas": return np.maximum(1.0 - np.abs(z), 0.0)
        raise ValueError(f"Unsupported ELM activation: {self.activation}")

    def _solve_beta(self, H, y):
        eye = np.eye(H.shape[1])
        return np.linalg.pinv(H.T @ H + self.alpha * eye) @ H.T @ y

    def _fit_scaled_realization(self, Xs, ys, seed):
        rng = np.random.default_rng(seed)
        weights = rng.uniform(-1.0, 1.0, (Xs.shape[1], self.hidden_neurons))
        biases = rng.uniform(-1.0, 1.0, self.hidden_neurons)
        hidden = self._activate(Xs @ weights + biases)
        beta = self._solve_beta(hidden, ys)
        return weights, biases, beta

    def fit(self, X, y):
        from sklearn.impute import SimpleImputer
        from sklearn.model_selection import train_test_split
        from sklearn.preprocessing import MinMaxScaler
        if self.trials < 1:
            raise ValueError("ELM trials must be at least 1")
        if not 0.05 <= self.validation_size <= 0.50:
            raise ValueError("ELM validation_size must be between 0.05 and 0.50")
        self.imputer_ = SimpleImputer(strategy="median")
        self.x_scaler_ = MinMaxScaler(feature_range=(-1.0, 1.0))
        self.y_scaler_ = MinMaxScaler(feature_range=(-1.0, 1.0))
        Xs = self.x_scaler_.fit_transform(self.imputer_.fit_transform(X))
        ys = self.y_scaler_.fit_transform(np.asarray(y, float).reshape(-1, 1)).ravel()
        indices = np.arange(len(ys))
        fit_idx, validation_idx = train_test_split(
            indices, test_size=self.validation_size, shuffle=True,
            random_state=self.random_state
        )
        trial_rows = []
        best_seed = self.random_state
        best_rmse = np.inf
        for trial in range(self.trials):
            seed = self.random_state + trial
            weights, biases, beta = self._fit_scaled_realization(
                Xs[fit_idx], ys[fit_idx], seed
            )
            prediction = self._activate(
                Xs[validation_idx] @ weights + biases
            ) @ beta
            rmse = float(np.sqrt(np.mean((prediction - ys[validation_idx]) ** 2)))
            trial_rows.append({
                "trial": trial + 1,
                "random_state": seed,
                "validation_normalized_rmse": rmse,
            })
            if rmse < best_rmse:
                best_rmse = rmse
                best_seed = seed
        # The trial/seed is selected using only the internal validation subset.
        # Refit its analytical output layer on the entire outer-training set.
        self.input_weights_, self.hidden_biases_, self.output_weights_ = (
            self._fit_scaled_realization(Xs, ys, best_seed)
        )
        self.selected_seed_ = int(best_seed)
        self.selection_validation_normalized_rmse_ = float(best_rmse)
        self.trial_history_ = pd.DataFrame(trial_rows)
        self.n_features_in_ = Xs.shape[1]
        return self

    def predict(self, X):
        Xs = self.x_scaler_.transform(self.imputer_.transform(X))
        H = self._activate(Xs @ self.input_weights_ + self.hidden_biases_)
        scaled = (H @ self.output_weights_).reshape(-1, 1)
        return self.y_scaler_.inverse_transform(scaled).ravel()


class ELMABCRegressor(ELMRegressor):
    """ELM whose hidden weights and biases are optimized by reproducible ABC.

    Output weights are solved analytically for every food source. They are not
    incorrectly constrained to the hidden-parameter [-1, 1] search interval.
    """
    def __init__(self, hidden_neurons=10, activation="sigmoid", alpha=1e-8,
                 random_state=42, trials=100, validation_size=0.20,
                 population_size=30, onlooker_count=20, max_iterations=100,
                 limit=None, max_acceleration=0.4):
        super().__init__(
            hidden_neurons, activation, alpha, random_state,
            trials=trials, validation_size=validation_size
        )
        self.requested_elm_trials = int(trials)
        self.population_size = int(population_size)
        self.onlooker_count = int(onlooker_count)
        self.max_iterations = int(max_iterations)
        self.limit = limit
        self.max_acceleration = float(max_acceleration)

    def _decode(self, vector, n_features):
        a = n_features * self.hidden_neurons
        W = vector[:a].reshape(n_features, self.hidden_neurons)
        bias = vector[a:a + self.hidden_neurons]
        return W, bias

    def fit(self, X, y):
        from sklearn.model_selection import train_test_split
        super().fit(X, y)
        Xs = self.x_scaler_.transform(self.imputer_.transform(X))
        ys = self.y_scaler_.transform(np.asarray(y, float).reshape(-1, 1)).ravel()
        n_features = Xs.shape[1]
        indices = np.arange(len(ys))
        fit_idx, validation_idx = train_test_split(
            indices, test_size=self.validation_size, shuffle=True,
            random_state=self.random_state
        )
        initial = np.concatenate([self.input_weights_.ravel(), self.hidden_biases_])
        dimensions = len(initial)
        rng = np.random.default_rng(self.random_state)
        population = rng.uniform(-1.0, 1.0, (self.population_size, dimensions))
        population[0] = np.clip(initial, -1.0, 1.0)
        trials = np.zeros(self.population_size, dtype=int)
        limit = int(self.limit) if self.limit is not None else max(5, dimensions * self.population_size // 2)

        def objective(v):
            W, bias = self._decode(v, n_features)
            hidden_fit = self._activate(Xs[fit_idx] @ W + bias)
            beta = self._solve_beta(hidden_fit, ys[fit_idx])
            pred = self._activate(Xs[validation_idx] @ W + bias) @ beta
            return float(np.sqrt(np.mean((pred - ys[validation_idx]) ** 2)))

        costs = np.array([objective(v) for v in population])
        best_i = int(np.argmin(costs)); best = population[best_i].copy(); best_cost = float(costs[best_i])
        history = []

        def greedy_candidate(i):
            choices = np.delete(np.arange(self.population_size), i)
            k = int(rng.choice(choices))
            phi = rng.uniform(-self.max_acceleration, self.max_acceleration, dimensions)
            candidate = np.clip(population[i] + phi * (population[i] - population[k]), -1.0, 1.0)
            candidate_cost = objective(candidate)
            if candidate_cost < costs[i]:
                population[i] = candidate; costs[i] = candidate_cost; trials[i] = 0
            else:
                trials[i] += 1

        for iteration in range(self.max_iterations):
            for i in range(self.population_size): greedy_candidate(i)
            fitness = 1.0 / (1.0 + np.maximum(costs, 0.0))
            probabilities = fitness / np.sum(fitness)
            for _ in range(self.onlooker_count): greedy_candidate(int(rng.choice(self.population_size, p=probabilities)))
            for i in np.where(trials >= limit)[0]:
                population[i] = rng.uniform(-1.0, 1.0, dimensions)
                costs[i] = objective(population[i]); trials[i] = 0
            i = int(np.argmin(costs))
            if costs[i] < best_cost: best = population[i].copy(); best_cost = float(costs[i])
            history.append({"iteration": iteration + 1, "best_normalized_rmse": best_cost})

        self.input_weights_, self.hidden_biases_ = self._decode(best, n_features)
        hidden = self._activate(Xs @ self.input_weights_ + self.hidden_biases_)
        self.output_weights_ = self._solve_beta(hidden, ys)
        self.abc_best_validation_normalized_rmse_ = best_cost
        # Backward-compatible field name retained for existing result readers.
        self.abc_best_normalized_rmse_ = best_cost
        self.abc_history_ = pd.DataFrame(history)
        self.abc_function_evaluations_ = self.population_size + self.max_iterations * (self.population_size + self.onlooker_count)
        return self


class SymbolicRegressionWrapper:
    """Auditable, Eureqa-inspired symbolic regression using gplearn."""
    def __init__(self, population_size=1000, generations=30, tournament_size=20,
                 stopping_criteria=0.0, const_range=(-1.0, 1.0), init_depth=(2, 6),
                 function_set=("add", "sub", "mul", "div", "sqrt", "log", "abs", "neg", "max", "min"),
                 parsimony_coefficient=0.001, max_samples=0.9, random_state=42,
                 n_jobs=-1, verbose=0, validation_size=0.20,
                 accuracy_tolerance=0.10, max_length=40, max_depth=8,
                 search_runs=3, min_validation_r2=0.0,
                 min_test_r2=0.30, max_abs_pbias=20.0,
                 min_features=1, max_features=0, min_length=1,
                 semantic_threshold=0.01):
        self.engine_ = "gplearn"
        self.population_size=population_size; self.generations=generations
        self.tournament_size=tournament_size; self.stopping_criteria=stopping_criteria
        self.const_range=const_range; self.init_depth=init_depth; self.function_set=function_set
        self.parsimony_coefficient=parsimony_coefficient; self.max_samples=max_samples
        self.random_state=random_state; self.n_jobs=n_jobs; self.verbose=verbose
        self.validation_size=validation_size
        self.accuracy_tolerance=accuracy_tolerance
        self.max_length=max_length
        self.max_depth=max_depth
        self.search_runs=search_runs
        self.min_validation_r2=min_validation_r2
        self.min_test_r2=min_test_r2
        self.max_abs_pbias=max_abs_pbias
        self.min_features=min_features
        self.max_features=max_features
        self.min_length=min_length
        self.semantic_threshold=semantic_threshold

    def fit(self, X, y):
        try:
            # gplearn 0.4.2 still calls this estimator method. Restore a narrow
            # compatibility shim only for scikit-learn versions that removed it.
            from sklearn.base import BaseEstimator
            if not hasattr(BaseEstimator, "_validate_data"):
                from sklearn.utils.validation import validate_data

                def _validate_data(estimator, *arrays, **check_params):
                    return validate_data(estimator, *arrays, **check_params)

                BaseEstimator._validate_data = _validate_data
            from gplearn.genetic import SymbolicRegressor
        except ImportError as exc:
            raise ImportError("SymbolicRegression requires gplearn. Install it with: pip install gplearn") from exc
        from sklearn.impute import SimpleImputer
        from sklearn.model_selection import train_test_split
        self.imputer_ = SimpleImputer(strategy="median")
        y_array = np.asarray(y, float)
        indices = np.arange(len(y_array))
        fit_idx, validation_idx = train_test_split(
            indices,
            test_size=self.validation_size,
            shuffle=False,
        )
        Xi_fit = self.imputer_.fit_transform(X.iloc[fit_idx] if hasattr(X, "iloc") else np.asarray(X)[fit_idx])
        Xi_validation = self.imputer_.transform(
            X.iloc[validation_idx] if hasattr(X, "iloc") else np.asarray(X)[validation_idx]
        )
        self.feature_names_in_ = [
            str(c) for c in getattr(
                X, "columns", [f"X{i}" for i in range(Xi_fit.shape[1])]
            )
        ]
        candidates = []
        seen = set()
        models = []
        for search_run in range(self.search_runs):
            model = SymbolicRegressor(
                population_size=self.population_size, generations=self.generations,
                tournament_size=self.tournament_size, stopping_criteria=self.stopping_criteria,
                const_range=self.const_range, init_depth=self.init_depth,
                function_set=self.function_set, metric="rmse",
                parsimony_coefficient=self.parsimony_coefficient,
                p_crossover=0.7, p_subtree_mutation=0.1, p_hoist_mutation=0.05,
                p_point_mutation=0.1, max_samples=self.max_samples,
                feature_names=self.feature_names_in_, warm_start=False,
                low_memory=True, n_jobs=self.n_jobs, verbose=self.verbose,
                random_state=self.random_state + search_run)
            model.fit(Xi_fit, y_array[fit_idx])
            models.append(model)
            for program in model._programs[-1]:
                if program is None:
                    continue
                expression = str(program)
                if expression in seen:
                    continue
                seen.add(expression)
                # Eureqa refits coefficients of linear components.  Apply the
                # reproducible one-component analogue y = intercept + slope*f(x)
                # using only the internal fitting subset, never validation/test.
                fit_raw = np.asarray(program.execute(Xi_fit), dtype=float)
                val_raw = np.asarray(program.execute(Xi_validation), dtype=float)
                if not np.all(np.isfinite(fit_raw)) or not np.all(np.isfinite(val_raw)):
                    continue
                design = np.column_stack([np.ones(len(fit_raw)), fit_raw])
                intercept, slope = np.linalg.lstsq(
                    design, y_array[fit_idx], rcond=None
                )[0]
                if not np.isfinite(intercept) or not np.isfinite(slope):
                    continue
                prediction = intercept + slope * val_raw
                residual = prediction - y_array[validation_idx]
                rmse = float(np.sqrt(np.mean(residual ** 2)))
                mae = float(np.mean(np.abs(residual)))
                bias = float(np.mean(residual))
                obs_mean = float(np.mean(y_array[validation_idx]))
                pbias = 100.0 * bias / obs_mean if abs(obs_mean) > EPS else np.nan
                denominator = float(np.sum((y_array[validation_idx] - obs_mean) ** 2))
                r2 = 1.0 - float(np.sum(residual ** 2)) / denominator if denominator > EPS else np.nan
                complexity = eureqa_weighted_complexity(expression)
                textual_features = [
                    name for name in self.feature_names_in_
                    if re.search(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", expression)
                ]
                used_features, effects = semantic_feature_activity(
                    lambda values: intercept + slope * np.asarray(program.execute(values), dtype=float),
                    Xi_validation, self.feature_names_in_, self.semantic_threshold,
                    self.random_state + search_run,
                )
                candidates.append({
                    "program": program,
                    "expression": expression,
                    "search_run": search_run + 1,
                    "calibration_intercept": float(intercept),
                    "calibration_slope": float(slope),
                    "validation_rmse": rmse,
                    "validation_mae": mae,
                    "validation_bias": bias,
                    "validation_pbias": pbias,
                    "validation_r2_standard": r2,
                    "complexity": complexity,
                    "length": int(program.length_),
                    "depth": int(program.depth_),
                    "feature_count": len(used_features),
                    "features_used": " | ".join(used_features),
                    "textual_features": " | ".join(textual_features),
                    "semantic_effects": json.dumps(effects, sort_keys=True),
                    "semantic_feature_constraint": True,
                })
        if not candidates:
            raise RuntimeError("Symbolic Regression produced no final-population candidates")

        for row in candidates:
            row["pareto"] = not any(
                (other["validation_rmse"] <= row["validation_rmse"]
                 and other["complexity"] <= row["complexity"]
                 and (other["validation_rmse"] < row["validation_rmse"]
                      or other["complexity"] < row["complexity"]))
                for other in candidates
            )
            row["within_complexity_limits"] = (
                self.min_length <= row["length"] <= self.max_length
                and row["depth"] <= self.max_depth
                and row["feature_count"] >= self.min_features
                and (self.max_features == 0 or row["feature_count"] <= self.max_features)
            )
            row["passes_validation_accuracy"] = bool(
                np.isfinite(row["validation_r2_standard"])
                and row["validation_r2_standard"] >= self.min_validation_r2
                and np.isfinite(row["validation_pbias"])
                and abs(row["validation_pbias"]) <= self.max_abs_pbias
            )

        pareto = [row for row in candidates if row["pareto"]]
        feasible = [
            row for row in pareto
            if row["within_complexity_limits"] and row["passes_validation_accuracy"]
        ]
        pool = feasible or [
            row for row in pareto if row["within_complexity_limits"]
        ] or pareto
        best_rmse = min(row["validation_rmse"] for row in pool)
        accurate = [
            row for row in pool
            if row["validation_rmse"] <= best_rmse * (1.0 + self.accuracy_tolerance)
        ]
        selected = min(
            accurate,
            key=lambda row: (row["complexity"], row["length"], row["validation_rmse"]),
        )
        selected["selected"] = True
        for row in candidates:
            row.setdefault("selected", False)

        self.model_ = models[selected["search_run"] - 1]
        self.model_._program = selected["program"]
        self.program_ = selected["expression"]
        self.calibration_intercept_ = selected["calibration_intercept"]
        self.calibration_slope_ = selected["calibration_slope"]
        self.eureqa_complexity_ = selected["complexity"]
        self.program_length_ = selected["length"]
        self.program_depth_ = selected["depth"]
        self.validation_rmse_ = selected["validation_rmse"]
        self.validation_r2_standard_ = selected["validation_r2_standard"]
        self.validation_pbias_ = selected["validation_pbias"]
        self.semantic_features_ = selected["features_used"]
        self.semantic_feature_count_ = selected["feature_count"]
        self.structurally_ready_ = bool(selected["within_complexity_limits"])
        self.validation_ready_ = bool(selected["passes_validation_accuracy"])
        self.publication_ready_ = False
        self.publication_status_ = "NOT ASSESSED ON HOLDOUT"
        self.pareto_candidates_ = pd.DataFrame([
            {key: value for key, value in row.items() if key != "program"}
            for row in candidates
        ]).sort_values(
            ["pareto", "complexity", "validation_rmse"],
            ascending=[False, True, True],
        )
        return self

    def predict(self, X):
        raw = np.asarray(
            self.model_._program.execute(self.imputer_.transform(X)), float
        )
        return self.calibration_intercept_ + self.calibration_slope_ * raw

    def assess_publication_readiness(self, metrics):
        test_r2 = metrics.get("test_r2_standard", np.nan)
        test_pbias = metrics.get("test_pbias", np.nan)
        test_ready = bool(
            np.isfinite(test_r2) and test_r2 >= self.min_test_r2
            and np.isfinite(test_pbias) and abs(test_pbias) <= self.max_abs_pbias
        )
        self.publication_ready_ = bool(
            self.structurally_ready_ and self.validation_ready_ and test_ready
        )
        failed = []
        if not self.structurally_ready_: failed.append("complexity")
        if not self.validation_ready_: failed.append("internal validation")
        if not test_ready: failed.append("independent holdout")
        self.publication_status_ = (
            "READY" if not failed else "NOT READY: " + ", ".join(failed)
        )
        return self.publication_ready_


def semantic_feature_activity(predict_function, X, feature_names, threshold=0.01, random_state=42):
    """Return predictors whose permutation materially changes equation output.

    The normalized RMS change is measured against the standard deviation of the
    original equation output. This rejects decorative/canceling variables such
    as x/x or a-b+b that do not affect predictions on the validation data.
    """
    values = np.asarray(X, dtype=float)
    baseline = np.asarray(predict_function(values), dtype=float).reshape(-1)
    scale = max(float(np.nanstd(baseline)), EPS)
    rng = np.random.default_rng(random_state)
    active, effects = [], {}
    for column, name in enumerate(feature_names):
        permuted = values.copy()
        permuted[:, column] = rng.permutation(permuted[:, column])
        changed = np.asarray(predict_function(permuted), dtype=float).reshape(-1)
        effect = float(np.sqrt(np.nanmean((changed - baseline) ** 2)) / scale)
        effects[str(name)] = effect
        if np.isfinite(effect) and effect >= threshold:
            active.append(str(name))
    return active, effects


class PySRSymbolicRegressionWrapper:
    """Automatic evolutionary symbolic regression using PySR/SymbolicRegression.jl."""
    def __init__(self, args, random_state=42):
        self.engine_ = "pysr"
        self.args = args
        self.random_state = random_state
        self.validation_size = args.sr_validation_size
        self.accuracy_tolerance = args.sr_accuracy_tolerance
        self.max_length = args.sr_max_length
        self.max_depth = args.sr_max_depth
        self.min_validation_r2 = args.sr_min_validation_r2
        self.min_test_r2 = args.sr_min_test_r2
        self.max_abs_pbias = args.sr_max_abs_pbias
        self.min_features = args.sr_min_features
        self.max_features = args.sr_max_features
        self.min_length = args.sr_min_length
        self.semantic_threshold = args.sr_semantic_threshold

    def fit(self, X, y):
        try:
            from pysr import PySRRegressor
        except ImportError as exc:
            raise ImportError(
                "PySR engine requires PySR. Install it with: pip install pysr"
            ) from exc
        from sklearn.impute import SimpleImputer
        from sklearn.model_selection import train_test_split

        self.imputer_ = SimpleImputer(strategy="median")
        y_array = np.asarray(y, dtype=float)
        indices = np.arange(len(y_array))
        fit_idx, validation_idx = train_test_split(
            indices, test_size=self.validation_size, shuffle=False
        )
        source = X if hasattr(X, "iloc") else np.asarray(X)
        Xi_fit = self.imputer_.fit_transform(
            source.iloc[fit_idx] if hasattr(source, "iloc") else source[fit_idx]
        )
        Xi_validation = self.imputer_.transform(
            source.iloc[validation_idx] if hasattr(source, "iloc") else source[validation_idx]
        )
        self.feature_names_in_ = [
            str(c) for c in getattr(X, "columns", [f"X{i}" for i in range(Xi_fit.shape[1])])
        ]
        operators = {"+": "+", "add": "+", "-": "-", "sub": "-",
                     "*": "*", "mul": "*", "/": "/", "div": "/"}
        binary = list(dict.fromkeys(
            operators[name] for name in self.args.sr_functions if name in operators
        ))
        if not binary:
            binary = ["+", "-", "*", "/"]
        output_dir = Path(self.args.sr_pysr_run_directory)
        model = PySRRegressor(
            niterations=self.args.sr_pysr_iterations,
            populations=self.args.sr_pysr_populations,
            population_size=self.args.sr_pysr_population_size,
            ncycles_per_iteration=self.args.sr_pysr_cycles,
            maxsize=self.max_length,
            maxdepth=self.max_depth,
            binary_operators=binary,
            unary_operators=[],
            complexity_of_operators={"/": 2},
            model_selection="accuracy",
            parsimony=self.args.sr_parsimony,
            early_stop_condition=self.args.sr_pysr_early_stop or None,
            timeout_in_seconds=(
                self.args.sr_pysr_timeout if self.args.sr_pysr_timeout > 0 else None
            ),
            batching=self.args.sr_pysr_batching,
            batch_size=self.args.sr_pysr_batch_size,
            warm_start=self.args.sr_pysr_warm_start,
            output_directory=str(output_dir),
            run_id=self.args.sr_pysr_run_id,
            random_state=self.random_state,
            deterministic=self.args.sr_pysr_deterministic,
            parallelism="serial" if self.args.sr_pysr_deterministic else "multithreading",
            verbosity=self.args.sr_verbose,
            progress=self.args.sr_verbose > 0,
        )
        model.fit(pd.DataFrame(Xi_fit, columns=self.feature_names_in_), y_array[fit_idx])
        candidates = []
        for equation_index, row in model.equations_.reset_index(drop=True).iterrows():
            function = row.get("lambda_format")
            if not callable(function):
                continue
            fit_raw = np.asarray(function(Xi_fit), dtype=float).reshape(-1)
            val_raw = np.asarray(function(Xi_validation), dtype=float).reshape(-1)
            if len(fit_raw) != len(fit_idx) or not np.all(np.isfinite(fit_raw)) or not np.all(np.isfinite(val_raw)):
                continue
            intercept, slope = np.linalg.lstsq(
                np.column_stack([np.ones(len(fit_raw)), fit_raw]),
                y_array[fit_idx], rcond=None
            )[0]
            predict_fn = lambda values, f=function, a=intercept, b=slope: (
                a + b * np.asarray(f(values), dtype=float).reshape(-1)
            )
            prediction = predict_fn(Xi_validation)
            residual = prediction - y_array[validation_idx]
            obs_mean = float(np.mean(y_array[validation_idx]))
            denominator = float(np.sum((y_array[validation_idx] - obs_mean) ** 2))
            r2 = 1.0 - float(np.sum(residual ** 2)) / denominator if denominator > EPS else np.nan
            pbias = 100.0 * float(np.sum(residual)) / float(np.sum(y_array[validation_idx])) if abs(np.sum(y_array[validation_idx])) > EPS else np.nan
            active, effects = semantic_feature_activity(
                predict_fn, Xi_validation, self.feature_names_in_,
                self.semantic_threshold, self.random_state + equation_index,
            )
            complexity = int(row.get("complexity", 1))
            candidates.append({
                "equation_index": int(equation_index),
                "expression": str(row.get("equation", "")),
                "function": function,
                "calibration_intercept": float(intercept),
                "calibration_slope": float(slope),
                "validation_rmse": float(np.sqrt(np.mean(residual ** 2))),
                "validation_mae": float(np.mean(np.abs(residual))),
                "validation_bias": float(np.mean(residual)),
                "validation_pbias": pbias,
                "validation_r2_standard": r2,
                "complexity": complexity,
                "length": complexity,
                "depth": np.nan,
                "feature_count": len(active),
                "features_used": " | ".join(active),
                "semantic_effects": json.dumps(effects, sort_keys=True),
                "semantic_feature_constraint": True,
            })
        if not candidates:
            raise RuntimeError("PySR produced no finite candidate equations")
        for row in candidates:
            row["pareto"] = not any(
                other["validation_rmse"] <= row["validation_rmse"]
                and other["complexity"] <= row["complexity"]
                and (other["validation_rmse"] < row["validation_rmse"]
                     or other["complexity"] < row["complexity"])
                for other in candidates
            )
            row["within_complexity_limits"] = (
                self.min_length <= row["length"] <= self.max_length
                and row["feature_count"] >= self.min_features
                and (not self.max_features or row["feature_count"] <= self.max_features)
            )
            row["passes_validation_accuracy"] = bool(
                np.isfinite(row["validation_r2_standard"])
                and row["validation_r2_standard"] >= self.min_validation_r2
                and np.isfinite(row["validation_pbias"])
                and abs(row["validation_pbias"]) <= self.max_abs_pbias
            )
        pareto = [row for row in candidates if row["pareto"]]
        feasible = [row for row in pareto if row["within_complexity_limits"] and row["passes_validation_accuracy"]]
        pool = feasible or [row for row in pareto if row["within_complexity_limits"]] or pareto
        best_rmse = min(row["validation_rmse"] for row in pool)
        accurate = [row for row in pool if row["validation_rmse"] <= best_rmse * (1 + self.accuracy_tolerance)]
        selected = min(accurate, key=lambda row: (row["complexity"], row["validation_rmse"]))
        for row in candidates:
            row["selected"] = row is selected
        self.model_ = model
        self.function_ = selected["function"]
        self.program_ = selected["expression"]
        self.calibration_intercept_ = selected["calibration_intercept"]
        self.calibration_slope_ = selected["calibration_slope"]
        self.eureqa_complexity_ = selected["complexity"]
        self.program_length_ = selected["length"]
        self.program_depth_ = selected["depth"]
        self.validation_rmse_ = selected["validation_rmse"]
        self.validation_r2_standard_ = selected["validation_r2_standard"]
        self.validation_pbias_ = selected["validation_pbias"]
        self.semantic_features_ = selected["features_used"]
        self.semantic_feature_count_ = selected["feature_count"]
        self.structurally_ready_ = bool(selected["within_complexity_limits"])
        self.validation_ready_ = bool(selected["passes_validation_accuracy"])
        self.publication_ready_ = False
        self.publication_status_ = "NOT ASSESSED ON HOLDOUT"
        self.pareto_candidates_ = pd.DataFrame([
            {key: value for key, value in row.items() if key != "function"}
            for row in candidates
        ])
        return self

    def predict(self, X):
        values = self.imputer_.transform(X)
        raw = np.asarray(self.function_(values), dtype=float).reshape(-1)
        return self.calibration_intercept_ + self.calibration_slope_ * raw

    def assess_publication_readiness(self, metrics):
        test_r2 = metrics.get("test_r2_standard", np.nan)
        test_pbias = metrics.get("test_pbias", np.nan)
        test_ready = bool(
            np.isfinite(test_r2) and test_r2 >= self.min_test_r2
            and np.isfinite(test_pbias) and abs(test_pbias) <= self.max_abs_pbias
        )
        failed = []
        if not self.structurally_ready_: failed.append("complexity/semantic features")
        if not self.validation_ready_: failed.append("internal validation")
        if not test_ready: failed.append("independent holdout")
        self.publication_ready_ = not failed
        self.publication_status_ = "READY" if not failed else "NOT READY: " + ", ".join(failed)
        return self.publication_ready_


def _split_sr_arguments(text: str) -> list[str]:
    """Split a gplearn argument list without splitting nested calls."""
    parts, start, depth = [], 0, 0
    for i, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth == 0:
            parts.append(text[start:i].strip())
            start = i + 1
    parts.append(text[start:].strip())
    return parts


def parse_sr_expression(text: str):
    """Parse gplearn's prefix expression into a small, dependency-free tree."""
    text = text.strip()
    first = text.find("(")
    if first < 0 or not text.endswith(")"):
        try:
            return ("const", float(text))
        except ValueError:
            return ("var", text)
    op = text[:first].strip()
    if op not in SR_ARITY:
        return ("var", text)
    args = _split_sr_arguments(text[first + 1:-1])
    if len(args) != SR_ARITY[op]:
        raise ValueError(f"Invalid SR expression for {op}: {text}")
    return (op, *(parse_sr_expression(arg) for arg in args))


def eureqa_weighted_complexity(expression: str) -> int:
    """Operator-weighted expression complexity inspired by Eureqa."""
    weights = {
        "add": 1, "sub": 1, "neg": 1, "abs": 1,
        "mul": 2, "div": 2, "max": 2, "min": 2,
        "sqrt": 3, "log": 3,
    }
    tree = parse_sr_expression(expression)

    def score(node):
        if node[0] in {"const", "var"}:
            return 1
        return weights.get(node[0], 2) + sum(score(arg) for arg in node[1:])

    return int(score(tree))


def _is_const(node, value: Optional[float] = None) -> bool:
    return node[0] == "const" and (value is None or abs(node[1] - value) <= EPS)


def simplify_sr_tree(node):
    """Apply conservative identities that preserve gplearn protected operators."""
    kind = node[0]
    if kind in {"const", "var"}:
        return node
    args = tuple(simplify_sr_tree(arg) for arg in node[1:])
    if all(_is_const(arg) for arg in args):
        values = [arg[1] for arg in args]
        try:
            if kind == "add": value = values[0] + values[1]
            elif kind == "sub": value = values[0] - values[1]
            elif kind == "mul": value = values[0] * values[1]
            elif kind == "div": value = values[0] / values[1] if abs(values[1]) > 0.001 else 1.0
            elif kind == "sqrt": value = math.sqrt(abs(values[0]))
            elif kind == "log": value = math.log(abs(values[0])) if abs(values[0]) > 0.001 else 0.0
            elif kind == "abs": value = abs(values[0])
            elif kind == "neg": value = -values[0]
            elif kind == "max": value = max(values)
            elif kind == "min": value = min(values)
            else: value = None
            if value is not None and math.isfinite(value):
                return ("const", float(value))
        except (ArithmeticError, ValueError):
            pass
    if kind == "add":
        if _is_const(args[0], 0): return args[1]
        if _is_const(args[1], 0): return args[0]
    elif kind == "sub":
        if _is_const(args[1], 0): return args[0]
        if args[0] == args[1]: return ("const", 0.0)
    elif kind == "mul":
        if _is_const(args[0], 0) or _is_const(args[1], 0): return ("const", 0.0)
        if _is_const(args[0], 1): return args[1]
        if _is_const(args[1], 1): return args[0]
    elif kind == "div":
        if _is_const(args[1], 1): return args[0]
        # Do not simplify x/x: gplearn's protected division returns 1 when
        # |x| <= 0.001, so ordinary cancellation can change predictions.
    elif kind == "neg" and args[0][0] == "neg":
        return args[0][1]
    return (kind, *args)


def sr_tree_to_infix(node) -> str:
    kind = node[0]
    if kind == "const":
        return f"{node[1]:.10g}"
    if kind == "var":
        return str(node[1])
    args = [sr_tree_to_infix(arg) for arg in node[1:]]
    binary = {"add": "+", "sub": "-", "mul": r"\times", "div": "/"}
    if kind in binary:
        return f"({args[0]} {binary[kind]} {args[1]})"
    if kind == "neg":
        return f"(-{args[0]})"
    if kind == "sqrt":
        return f"sqrt(abs({args[0]}))"
    if kind == "log":
        return f"protected_log({args[0]})"
    return f"{kind}({', '.join(args)})"


def sr_tree_complexity(node) -> tuple[int, int, int]:
    """Return node count, operator count, and depth."""
    if node[0] in {"const", "var"}:
        return 1, 0, 0
    child = [sr_tree_complexity(arg) for arg in node[1:]]
    return 1 + sum(x[0] for x in child), 1 + sum(x[1] for x in child), 1 + max(x[2] for x in child)


def symbolic_equation_record(
    fitted,
    base_result: CaseResult,
    metrics: dict[str, Any],
) -> dict[str, Any]:
    if fitted.engine_ == "gplearn":
        raw_tree = parse_sr_expression(fitted.program_)
        simplified_tree = simplify_sr_tree(raw_tree)
        nodes, operators, depth = sr_tree_complexity(simplified_tree)
        raw_expression = sr_tree_to_infix(raw_tree)
        simple_expression = sr_tree_to_infix(simplified_tree)
    else:
        raw_expression = fitted.program_
        simple_expression = fitted.program_
        nodes, operators, depth = fitted.program_length_, np.nan, fitted.program_depth_
    calibrated_raw = (
        f"({fitted.calibration_intercept_:.10g}) + "
        f"({fitted.calibration_slope_:.10g}) * ({raw_expression})"
    )
    calibrated_simple = (
        f"({fitted.calibration_intercept_:.10g}) + "
        f"({fitted.calibration_slope_:.10g}) * ({simple_expression})"
    )
    return {
        "case": base_result.case_id,
        "relative_path": base_result.relative_path,
        "grid_m": base_result.grid_m,
        "period": base_result.period,
        "target": base_result.target,
        "predictors": base_result.predictors,
        "sr_engine": fitted.engine_,
        "raw_symbolic_expression": fitted.program_,
        "raw_infix_equation": f"{base_result.target} = {calibrated_raw}",
        "simplified_equation": f"{base_result.target} = {calibrated_simple}",
        "calibration_intercept": fitted.calibration_intercept_,
        "calibration_slope": fitted.calibration_slope_,
        "eureqa_weighted_complexity": fitted.eureqa_complexity_,
        "raw_program_length": fitted.program_length_,
        "raw_program_depth": fitted.program_depth_,
        "simplified_node_count": nodes,
        "simplified_operator_count": operators,
        "simplified_depth": depth,
        "selection_validation_rmse": fitted.validation_rmse_,
        "selection_validation_r2_standard": fitted.validation_r2_standard_,
        "selection_validation_pbias": fitted.validation_pbias_,
        "publication_ready": fitted.publication_ready_,
        "publication_status": fitted.publication_status_,
        "preferred_max_length": fitted.max_length,
        "preferred_max_depth": fitted.max_depth,
        "required_min_features": fitted.min_features,
        "allowed_max_features": fitted.max_features or "all",
        "required_min_length": fitted.min_length,
        "semantic_features_used": fitted.semantic_features_,
        "semantic_feature_count": fitted.semantic_feature_count_,
        "semantic_activity_threshold": fitted.semantic_threshold,
        "train_r2_standard": metrics.get("train_r2_standard"),
        "test_r2_standard": metrics.get("test_r2_standard"),
        "test_r2_corr": metrics.get("test_r2_corr"),
        "test_rmse": metrics.get("test_rmse"),
        "test_mae": metrics.get("test_mae"),
        "split_strategy": base_result.split_strategy,
        "random_state": base_result.split_random_state,
    }


def build_model(name: str, random_state: int, n_jobs: int, args: argparse.Namespace):
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer
    from sklearn.preprocessing import StandardScaler
    from sklearn.linear_model import LinearRegression
    from sklearn.ensemble import RandomForestRegressor
    from sklearn.svm import SVR
    if name == "ELM":
        return ELMRegressor(
            args.elm_hidden_neurons, args.elm_activation, args.elm_alpha,
            random_state, args.elm_trials, args.elm_validation_size
        )
    if name == "ELMABC":
        return ELMABCRegressor(args.elm_hidden_neurons, args.elm_activation, args.elm_alpha,
            random_state, args.elm_trials, args.elm_validation_size,
            args.abc_population_size, args.abc_onlookers,
            args.abc_iterations, args.abc_limit, args.abc_max_acceleration)
    if name == "SymbolicRegression":
        if args.sr_engine == "pysr":
            return PySRSymbolicRegressionWrapper(args, random_state)
        return SymbolicRegressionWrapper(args.sr_population_size, args.sr_generations,
            args.sr_tournament_size, args.sr_stopping_criteria, (-1.0, 1.0),
            (args.sr_init_depth_min, args.sr_init_depth_max), tuple(args.sr_functions),
            args.sr_parsimony, args.sr_max_samples, random_state, n_jobs, args.sr_verbose,
            args.sr_validation_size, args.sr_accuracy_tolerance,
            args.sr_max_length, args.sr_max_depth, args.sr_search_runs,
            args.sr_min_validation_r2, args.sr_min_test_r2,
            args.sr_max_abs_pbias, args.sr_min_features,
            args.sr_max_features, args.sr_min_length,
            args.sr_semantic_threshold)

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
        if model_name in {"ELM", "ELMABC"}:
            rows = [
                {"parameter": "hidden_neurons", "value": est.hidden_neurons},
                {"parameter": "activation", "value": est.activation},
                {"parameter": "ridge_alpha", "value": est.alpha},
                {"parameter": "random_state", "value": est.random_state},
                {"parameter": "internal_validation_fraction", "value": est.validation_size},
            ]
            if model_name == "ELM":
                rows += [
                    {"parameter": "elm_trials", "value": est.trials},
                    {"parameter": "selected_seed", "value": est.selected_seed_},
                    {"parameter": "selection_validation_normalized_rmse",
                     "value": est.selection_validation_normalized_rmse_},
                ]
            else:
                rows += [
                    {"parameter": "abc_searches_hidden_parameters_only", "value": True},
                    {"parameter": "abc_output_weights_solved_analytically", "value": True},
                    {"parameter": "abc_population_size", "value": est.population_size},
                    {"parameter": "abc_onlooker_count", "value": est.onlooker_count},
                    {"parameter": "abc_iterations", "value": est.max_iterations},
                    {"parameter": "abc_best_validation_normalized_rmse",
                     "value": est.abc_best_validation_normalized_rmse_},
                    {"parameter": "abc_function_evaluations", "value": est.abc_function_evaluations_},
                ]
            return pd.DataFrame(rows)
        if model_name == "SymbolicRegression":
            return pd.DataFrame([
                {"parameter": "sr_engine", "value": est.engine_},
                {"parameter": "symbolic_expression", "value": est.program_},
                {"parameter": "program_length", "value": est.program_length_},
                {"parameter": "program_depth", "value": est.program_depth_},
                {"parameter": "population_size", "value": getattr(est, "population_size", getattr(getattr(est, "args", None), "sr_pysr_population_size", np.nan))},
                {"parameter": "generations_or_iterations", "value": getattr(est, "generations", getattr(getattr(est, "args", None), "sr_pysr_iterations", np.nan))},
                {"parameter": "parsimony_coefficient", "value": getattr(est, "parsimony_coefficient", getattr(getattr(est, "args", None), "sr_parsimony", np.nan))},
                {"parameter": "selection_validation_size", "value": est.validation_size},
                {"parameter": "selection_validation_rmse", "value": est.validation_rmse_},
                {"parameter": "selection_validation_r2_standard", "value": est.validation_r2_standard_},
                {"parameter": "selection_validation_pbias", "value": est.validation_pbias_},
                {"parameter": "calibration_intercept", "value": est.calibration_intercept_},
                {"parameter": "calibration_slope", "value": est.calibration_slope_},
                {"parameter": "eureqa_weighted_complexity", "value": est.eureqa_complexity_},
                {"parameter": "search_runs", "value": getattr(est, "search_runs", 1)},
                {"parameter": "semantic_features_used", "value": est.semantic_features_},
                {"parameter": "semantic_feature_count", "value": est.semantic_feature_count_},
                {"parameter": "semantic_activity_threshold", "value": est.semantic_threshold},
                {"parameter": "publication_ready", "value": est.publication_ready_},
                {"parameter": "publication_status", "value": est.publication_status_},
                {"parameter": "preferred_max_length", "value": est.max_length},
                {"parameter": "preferred_max_depth", "value": est.max_depth},
                {"parameter": "random_state", "value": est.random_state},
            ])
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
    p = argparse.ArgumentParser(description="Run ELM, ELM-ABC, and symbolic regression on the same paper workbooks and validation design.")
    p.add_argument("--root", type=Path, default=Path.cwd())
    p.add_argument("--data-folder", default="data")
    p.add_argument("--output", type=Path, default=Path("elm_symbolic_results"))
    p.add_argument("--file-pattern", action="append", default=[],
                   help="Process only relative paths matching this glob; repeat as needed (example: *Mesh600*).")
    p.add_argument("--max-files", type=int, default=0,
                   help="Process at most this many matched workbooks; 0 means all.")
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
    p.add_argument("--split", choices=["original", "random", "spatial-block"], default="original")
    p.add_argument("--test-size",type=float,default=0.30)
    p.add_argument("--cv-folds",type=int,default=10)
    p.add_argument("--skip-cross-validation",action="store_true")
    p.add_argument("--random-state", type=int, default=RANDOM_STATE)
    p.add_argument("--spatial-blocks", type=int, default=5)
    p.add_argument("--best-meshes", nargs="+", type=int, default=[600, 700])
    p.add_argument("--n-jobs", type=int, default=-1)
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--fail-on-missing-model", action="store_true")
    p.add_argument(
        "--fail-on-error", action="store_true",
        help="Stop immediately if any requested model or workbook fails.",
    )
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
    # ELM / ELM-ABC
    p.add_argument("--elm-hidden-neurons", type=int, default=10)
    p.add_argument("--elm-activation", choices=["sigmoid", "tanh", "sin", "radbas", "hardlim", "tribas"], default="sigmoid")
    p.add_argument("--elm-alpha", type=float, default=1e-8)
    p.add_argument("--elm-trials", type=int, default=100,
                   help="Independent ELM initializations selected on an internal training-only validation subset.")
    p.add_argument("--elm-validation-size", type=float, default=0.20,
                   help="Fraction of each outer-training set used only to select the ELM realization.")
    p.add_argument("--abc-population-size", type=int, default=30)
    p.add_argument("--abc-onlookers", type=int, default=20)
    p.add_argument("--abc-iterations", type=int, default=100)
    p.add_argument("--abc-limit", type=int, default=None)
    p.add_argument("--abc-max-acceleration", type=float, default=0.4)
    # Symbolic regression
    p.add_argument("--sr-engine", choices=["gplearn", "pysr"], default="gplearn",
                   help="Evolutionary SR engine. PySR is faster/stronger for final searches; gplearn preserves legacy results.")
    p.add_argument("--sr-population-size", type=int, default=1000)
    p.add_argument("--sr-generations", type=int, default=30)
    p.add_argument("--sr-tournament-size", type=int, default=20)
    p.add_argument("--sr-stopping-criteria", type=float, default=0.0)
    p.add_argument("--sr-init-depth-min", type=int, default=2)
    p.add_argument("--sr-init-depth-max", type=int, default=6)
    p.add_argument("--sr-functions", nargs="+", default=["add", "sub", "mul", "div", "sqrt", "log", "abs", "neg", "max", "min"])
    p.add_argument("--sr-parsimony", type=float, default=0.001)
    p.add_argument("--sr-max-samples", type=float, default=0.9)
    p.add_argument("--sr-verbose", type=int, default=0)
    p.add_argument("--sr-publication-mode", action="store_true",
                   help="Use publication-safe operators (add/sub/mul/protected div), "
                        "initial depth 2-4, and max_samples 1.0; explicit search-size settings are preserved.")
    p.add_argument("--sr-equations-file", default="SR_Equations.xlsx",
                   help="Consolidated publication-equation workbook written under --output.")
    p.add_argument("--sr-validation-size", type=float, default=0.20,
                   help="Fraction of outer training data reserved for SR candidate selection.")
    p.add_argument("--sr-accuracy-tolerance", type=float, default=0.10,
                   help="Select the shortest candidate within this fraction of the best validation RMSE.")
    p.add_argument("--sr-max-length", type=int, default=40,
                   help="Preferred maximum gplearn program length for a publication-ready equation.")
    p.add_argument("--sr-max-depth", type=int, default=8,
                   help="Preferred maximum gplearn program depth for a publication-ready equation.")
    p.add_argument("--sr-search-runs", type=int, default=3,
                   help="Independent GP searches pooled into one error-complexity frontier.")
    p.add_argument("--sr-min-validation-r2", type=float, default=0.0,
                   help="Minimum internal-validation standard R2 for publication readiness.")
    p.add_argument("--sr-min-test-r2", type=float, default=0.30,
                   help="Minimum independent-holdout standard R2 for publication readiness only.")
    p.add_argument("--sr-max-abs-pbias", type=float, default=20.0,
                   help="Maximum absolute validation and holdout PBIAS for publication readiness.")
    p.add_argument("--sr-min-features", type=int, default=1,
                   help="Minimum distinct predictors required in an eligible equation.")
    p.add_argument("--sr-max-features", type=int, default=0,
                   help="Maximum distinct predictors in an eligible equation; 0 means no limit.")
    p.add_argument("--sr-min-length", type=int, default=1,
                   help="Minimum program length for an eligible equation.")
    p.add_argument("--sr-semantic-threshold", type=float, default=0.01,
                   help="Minimum normalized prediction change after permuting a predictor for it to count as active.")
    p.add_argument("--sr-fast-mode", action="store_true",
                   help="Fast screening preset: 600 population, 20 generations, 2 searches.")
    # PySR / SymbolicRegression.jl
    p.add_argument("--sr-pysr-iterations", type=int, default=100)
    p.add_argument("--sr-pysr-populations", type=int, default=8)
    p.add_argument("--sr-pysr-population-size", type=int, default=50)
    p.add_argument("--sr-pysr-cycles", type=int, default=300)
    p.add_argument("--sr-pysr-timeout", type=float, default=0,
                   help="Per-case PySR time limit in seconds; 0 disables the limit.")
    p.add_argument("--sr-pysr-early-stop", default="",
                   help="Optional Julia early-stop expression accepted by PySR.")
    p.add_argument("--sr-pysr-batching", action="store_true")
    p.add_argument("--sr-pysr-batch-size", type=int, default=128)
    p.add_argument("--sr-pysr-warm-start", action="store_true",
                   help="Continue the same PySR run; use only with unchanged data and settings.")
    p.add_argument("--sr-pysr-run-directory", default="pysr_runs")
    p.add_argument("--sr-pysr-run-id", default=None)
    p.add_argument("--sr-pysr-deterministic", action="store_true",
                   help="Use deterministic serial PySR execution for exact reproducibility.")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    args.input_columns = parse_column_positions(args.input_columns, "--input-columns")
    if args.output_column is not None and args.output_column < 1:
        raise ValueError("--output-column must be at least 1.")
    if not 0.05 <= args.sr_validation_size <= 0.50:
        raise ValueError("--sr-validation-size must be between 0.05 and 0.50")
    if args.elm_trials < 1:
        raise ValueError("--elm-trials must be at least 1")
    if not 0.05 <= args.elm_validation_size <= 0.50:
        raise ValueError("--elm-validation-size must be between 0.05 and 0.50")
    if args.sr_accuracy_tolerance < 0:
        raise ValueError("--sr-accuracy-tolerance must be nonnegative")
    if args.sr_max_length < 1 or args.sr_max_depth < 1:
        raise ValueError("--sr-max-length and --sr-max-depth must be positive")
    if args.sr_search_runs < 1:
        raise ValueError("--sr-search-runs must be positive")
    if args.sr_max_abs_pbias < 0:
        raise ValueError("--sr-max-abs-pbias must be nonnegative")
    if args.sr_min_features < 1 or args.sr_max_features < 0:
        raise ValueError("--sr-min-features must be >= 1 and --sr-max-features must be >= 0")
    if args.sr_max_features and args.sr_max_features < args.sr_min_features:
        raise ValueError("--sr-max-features cannot be smaller than --sr-min-features")
    if args.sr_min_length < 1 or args.max_files < 0:
        raise ValueError("--sr-min-length must be positive and --max-files cannot be negative")
    if args.sr_semantic_threshold < 0:
        raise ValueError("--sr-semantic-threshold must be nonnegative")
    if args.sr_publication_mode:
        args.sr_functions = ["add", "sub", "mul", "div"]
        args.sr_init_depth_min = 2
        args.sr_init_depth_max = 4
        args.sr_max_samples = 1.0
    if args.sr_fast_mode:
        if args.sr_engine == "pysr":
            args.sr_pysr_iterations = min(args.sr_pysr_iterations, 30)
            args.sr_pysr_populations = min(args.sr_pysr_populations, 4)
            args.sr_pysr_population_size = min(args.sr_pysr_population_size, 40)
            args.sr_pysr_cycles = min(args.sr_pysr_cycles, 200)
        else:
            args.sr_population_size = 600
            args.sr_generations = 20
            args.sr_search_runs = 2
            args.sr_tournament_size = min(args.sr_tournament_size, 15)
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
    if args.max_files:
        files = files[:args.max_files]
    if not files:
        raise FileNotFoundError(f"No Excel workbooks found under {data_root}")

    config = vars(args).copy()
    config["root"] = str(root); config["output"] = str(output)
    (output / "run_configuration.json").write_text(json.dumps(config, indent=2, default=str), encoding="utf-8")

    metric_rows: list[dict[str, Any]] = []
    audit_rows: list[dict[str, Any]] = []
    sr_equation_rows: list[dict[str, Any]] = []
    sr_range_rows: list[dict[str, Any]] = []
    sr_candidate_rows: list[dict[str, Any]] = []
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
                    if model_name == "SymbolicRegression":
                        fitted.assess_publication_readiness(metrics)
                        params = model_parameters(model_name, fitted, predictors)
                        sr_equation_rows.append(symbolic_equation_record(fitted, base_result, metrics))
                        candidate_table = fitted.pareto_candidates_.copy()
                        candidate_table.insert(0, "target", str(target_col))
                        candidate_table.insert(0, "grid_m", grid)
                        candidate_table.insert(0, "case", case_name)
                        sr_candidate_rows.extend(candidate_table.to_dict("records"))
                        for predictor in predictors:
                            values = pd.to_numeric(X[predictor], errors="coerce")
                            train_values = values.iloc[train_idx]
                            sr_range_rows.append({
                                "case": case_name,
                                "grid_m": grid,
                                "target": str(target_col),
                                "predictor": predictor,
                                "unit": "Specify in manuscript",
                                "all_min": values.min(),
                                "all_max": values.max(),
                                "training_min": train_values.min(),
                                "training_max": train_values.max(),
                                "training_median": train_values.median(),
                            })
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
                    if args.fail_on_error or (
                        args.fail_on_missing_model and isinstance(exc, ImportError)
                    ):
                        raise
        except Exception as exc:
            audit_rows.append({
                "severity": "ERROR", "case": case_name, "model": "ALL",
                "file": str(raw_path), "message": f"{type(exc).__name__}: {exc}",
                "traceback": traceback.format_exc(),
            })
            print(f"  CASE ERROR: {exc}", file=sys.stderr)
            if args.fail_on_error or (
                args.fail_on_missing_model and isinstance(exc, ImportError)
            ):
                raise

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

    if sr_equation_rows:
        equation_df = pd.DataFrame(sr_equation_rows).sort_values(
            ["grid_m", "case"], na_position="last"
        )
        ranges_df = pd.DataFrame(sr_range_rows).sort_values(
            ["grid_m", "case", "predictor"], na_position="last"
        )
        functions_df = pd.DataFrame([
            {"function": "div(a,b)", "publication_definition": "a/b when |b| > 0.001; otherwise 1", "protected": True},
            {"function": "sqrt(a)", "publication_definition": "sqrt(|a|)", "protected": True},
            {"function": "log(a)", "publication_definition": "ln(|a|) when |a| > 0.001; otherwise 0", "protected": True},
            {"function": "abs(a)", "publication_definition": "|a|", "protected": False},
            {"function": "max(a,b)", "publication_definition": "maximum of a and b", "protected": False},
            {"function": "min(a,b)", "publication_definition": "minimum of a and b", "protected": False},
        ])
        settings_df = pd.DataFrame([
            {"setting": "sr_engine", "value": args.sr_engine},
            {"setting": "population_size", "value": args.sr_population_size},
            {"setting": "generations", "value": args.sr_generations},
            {"setting": "tournament_size", "value": args.sr_tournament_size},
            {"setting": "parsimony_coefficient", "value": args.sr_parsimony},
            {"setting": "function_set", "value": " ".join(args.sr_functions)},
            {"setting": "max_samples", "value": args.sr_max_samples},
            {"setting": "selection_validation_size", "value": args.sr_validation_size},
            {"setting": "accuracy_tolerance", "value": args.sr_accuracy_tolerance},
            {"setting": "preferred_max_length", "value": args.sr_max_length},
            {"setting": "preferred_max_depth", "value": args.sr_max_depth},
            {"setting": "independent_search_runs", "value": args.sr_search_runs},
            {"setting": "minimum_validation_r2_standard", "value": args.sr_min_validation_r2},
            {"setting": "minimum_holdout_r2_standard", "value": args.sr_min_test_r2},
            {"setting": "maximum_absolute_pbias_percent", "value": args.sr_max_abs_pbias},
            {"setting": "minimum_distinct_features", "value": args.sr_min_features},
            {"setting": "maximum_distinct_features", "value": args.sr_max_features or "unlimited"},
            {"setting": "minimum_program_length", "value": args.sr_min_length},
            {"setting": "semantic_activity_threshold", "value": args.sr_semantic_threshold},
            {"setting": "pysr_iterations", "value": args.sr_pysr_iterations},
            {"setting": "pysr_populations", "value": args.sr_pysr_populations},
            {"setting": "pysr_population_size", "value": args.sr_pysr_population_size},
            {"setting": "pysr_timeout_seconds", "value": args.sr_pysr_timeout},
            {"setting": "split", "value": args.split},
            {"setting": "cross_validation_skipped", "value": args.skip_cross_validation},
            {"setting": "simplification", "value": "Conservative protected-operator identities only"},
        ])
        with pd.ExcelWriter(output / args.sr_equations_file, engine="openpyxl") as writer:
            equation_df.to_excel(writer, index=False, sheet_name="Equations")
            if sr_candidate_rows:
                pd.DataFrame(sr_candidate_rows).to_excel(
                    writer, index=False, sheet_name="Pareto Candidates"
                )
            ranges_df.to_excel(writer, index=False, sheet_name="Predictor Ranges")
            functions_df.to_excel(writer, index=False, sheet_name="Function Definitions")
            settings_df.to_excel(writer, index=False, sheet_name="SR Settings")
        equation_df.to_csv(output / "SR_Equations.csv", index=False, encoding="utf-8-sig")

    readme = f"""ELM, ELM-ABC, and symbolic-regression modelling output

Raw files discovered: {len(files)}
Models requested: {', '.join(args.models)}
Split strategy: {args.split}
Best meshes copied separately: {sorted(best_meshes)}
Successful model-case runs: {sum(1 for r in metric_rows if r.get('status') == 'OK')}
Failed model-case runs: {len(audit_rows)}
Consolidated SR equations: {len(sr_equation_rows)}

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
