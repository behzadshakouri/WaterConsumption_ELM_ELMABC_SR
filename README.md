# Water Consumption Modeling with ELM, ELM-ABC, Symbolic Regression, and Common Baselines

Python workflows for evaluating spatial water-consumption models across multiple grid resolutions and time periods. The repository provides:

- Extreme Learning Machine (**ELM**)
- Artificial Bee Colony–optimized ELM (**ELM-ABC**)
- Symbolic Regression (**SR**) with publication-ready equation export
- Multiple Linear Regression (**MLR**)
- Random Forest (**RF**)
- Extreme Gradient Boosting (**XGBoost**)
- Support Vector Regression (**SVR**)
- Geographically Weighted Regression (**GWR**)

All models use the same raw Excel workbooks, predictor definitions, data partitions, performance metrics, and output structure to support fair and reproducible comparisons.

## Repository files

| File | Purpose |
|---|---|
| `run_elm_elmabc_symbolic.py` | Runs ELM, ELM-ABC, and Symbolic Regression. It also simplifies and consolidates SR equations. |
| `run_common_baselines.py` | Runs MLR, RF, XGBoost, SVR, and GWR baseline models. |

> The downloaded files may contain suffixes such as `(1)` or `(2)`. Renaming them to the filenames above is recommended before running the commands in this README.

## Expected project structure

```text
WaterConsumption_ELM_ELMABC_SR/
├── data/
│   ├── Mesh600-C4-150-300.xlsx
│   ├── Mesh600-C4-300-450.xlsx
│   └── ...
├── run_elm_elmabc_symbolic.py
├── run_common_baselines.py
└── README.md
```

The scripts recursively discover `.xlsx`, `.xlsm`, and `.xls` files under `data/`. Each workbook must contain a header row. Unless explicitly overridden:

- the first column is treated as the observation or grid-cell identifier;
- the last column is treated as the observed target;
- selected numeric columns between them are used as predictors;
- the first non-empty worksheet is analyzed.

## Predictor sets

Three predictor-selection modes are available:

| Option | Predictors |
|---|---|
| `paper-mean` | `T96_mean`, `P_mean`, `Elv_mean` |
| `paper-summary` | Mean, minimum, and maximum values of `T96`, `P`, and `Elv` |
| `numeric-safe` | Eligible numeric columns, excluding IDs, coordinates, the target, and other potentially unsafe fields |

`paper-summary` is the default. For maximum reproducibility, explicitly provide either a named predictor set or the exact predictor columns.

## Installation

Python 3.10 or newer is recommended.

```bash
git clone https://github.com/behzadshakouri/WaterConsumption_ELM_ELMABC_SR.git
cd WaterConsumption_ELM_ELMABC_SR

python -m venv .venv
```

Activate the environment on Windows:

```powershell
.\.venv\Scripts\Activate.ps1
```

Activate it on Linux or macOS:

```bash
source .venv/bin/activate
```

Install the core and optional model dependencies:

```bash
python -m pip install --upgrade pip
pip install numpy pandas openpyxl scikit-learn gplearn pysr xgboost mgwr
```

Notes:

- `gplearn` is required only for Symbolic Regression.
- `pysr` enables the recommended high-performance evolutionary SR engine. Its
  Julia backend is installed/configured automatically on first import, so the
  first run takes longer than later runs.
- `xgboost` is required only for XGBoost.
- `mgwr` is required only for GWR.
- Reading legacy `.xls` files may additionally require `xlrd`.

## Quick start

### ELM and ELM-ABC only

The following command uses the original sequential 70/30 split, runs ELM and ELM-ABC, performs 10-fold cross-validation on the training subset, and replaces an existing output directory:

```powershell
python run_elm_elmabc_symbolic.py `
  --models ELM ELMABC `
  --split original `
  --overwrite
```

To skip the additional cross-validation:

```powershell
python run_elm_elmabc_symbolic.py `
  --models ELM ELMABC `
  --split original `
  --skip-cross-validation `
  --overwrite
```

### Symbolic Regression only

Two evolutionary engines are available:

| Engine | Use |
|---|---|
| `--sr-engine gplearn` | Backward-compatible reproduction of earlier runs |
| `--sr-engine pysr` | Recommended automatic final search using PySR/SymbolicRegression.jl |

Both engines use the same internal validation, independent holdout, Pareto
selection, coefficient refit, and publication gates. Feature requirements are
now semantic: a variable counts only when permuting it changes the equation
output by at least `--sr-semantic-threshold` (default 0.01 normalized RMS).
Decorative variables in expressions such as `Elv/Elv` or canceling terms do not
satisfy `--sr-min-features`.

### Recommended PySR fast screen

```powershell
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --sr-engine pysr `
  --predictor-set paper-mean `
  --split original `
  --sr-publication-mode `
  --sr-fast-mode `
  --sr-min-features 1 `
  --sr-max-length 40 `
  --sr-max-depth 8 `
  --sr-pysr-timeout 180 `
  --file-pattern "Mesh600*.xlsx" `
  --skip-cross-validation `
  --output sr_pysr_screen `
  --overwrite
```

Fast mode caps PySR at 30 iterations, four populations, 40 individuals per
population, and 200 cycles per iteration. `--sr-pysr-timeout 180` imposes a
three-minute limit per workbook. Use `--max-files 1` when checking installation.

### Recommended PySR extended final search

Run this only for cases that passed screening:

```powershell
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --sr-engine pysr `
  --predictor-set paper-mean `
  --split original `
  --sr-publication-mode `
  --file-pattern "Mesh600-C4-150-300.xlsx" `
  --sr-validation-size 0.20 `
  --sr-accuracy-tolerance 0.02 `
  --sr-min-features 2 `
  --sr-max-features 3 `
  --sr-semantic-threshold 0.01 `
  --sr-min-length 5 `
  --sr-max-length 40 `
  --sr-max-depth 7 `
  --sr-parsimony 0.001 `
  --sr-pysr-iterations 150 `
  --sr-pysr-populations 8 `
  --sr-pysr-population-size 50 `
  --sr-pysr-cycles 300 `
  --sr-pysr-timeout 1800 `
  --sr-min-validation-r2 0.30 `
  --sr-min-test-r2 0.30 `
  --sr-max-abs-pbias 20 `
  --skip-cross-validation `
  --output sr_pysr_extended `
  --overwrite
```

The independent holdout remains report-only and is never used to choose an
equation. If no candidate has two semantically active predictors, the best
fallback is exported as `NOT READY`.

For an interrupted search with identical data and settings, add a stable
`--sr-pysr-run-directory`, `--sr-pysr-run-id`, and `--sr-pysr-warm-start`.
Never warm-start after changing predictors, operators, or data.

For a fast first-pass screen based on the three mean predictors:

```powershell
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --predictor-set paper-mean `
  --split original `
  --sr-publication-mode `
  --sr-fast-mode `
  --sr-validation-size 0.20 `
  --sr-accuracy-tolerance 0.10 `
  --sr-max-length 40 `
  --sr-max-depth 8 `
  --sr-min-validation-r2 0.0 `
  --sr-min-test-r2 0.30 `
  --sr-max-abs-pbias 20 `
  --skip-cross-validation `
  --output sr_equation_results `
  --overwrite
```

`--sr-fast-mode` uses 600 individuals, 20 generations, and two independent
searches. It is intended to identify promising cases quickly. `--sr-publication-mode`
restricts the operators to addition, subtraction, multiplication, and protected
division, uses shallower initial trees, and evaluates all samples. Candidate equations
are compared on an operator-weighted error-versus-complexity Pareto frontier
using an internal 20% validation subset drawn only from the outer training
data. Linear intercept and slope coefficients are refitted on the internal
fitting subset before validation, analogous to Eureqa's coefficient-refit
step. The shortest equation within 10% of the best validation RMSE is selected
from the eligible Pareto candidates.

The independent 30% holdout is never used to choose the equation. It is used
only for final reporting. `publication_ready` becomes true only when the
selected equation:

- stays within 40 nodes and depth 8;
- has internal-validation standard \(R^2 \ge 0\);
- has holdout standard \(R^2 \ge 0.30\); and
- has absolute PBIAS no greater than 20% on validation and holdout data.

If no equation passes these gates, the script still exports the best available
candidate but marks it `NOT READY`, preventing a compact yet ineffective
formula from being presented as publication-ready.

### Extended equations with two or three predictors

To investigate whether a moderately richer equation improves validation
accuracy or holdout bias, require two to three distinct predictors:

```powershell
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --predictor-set paper-mean `
  --split original `
  --sr-publication-mode `
  --sr-validation-size 0.20 `
  --sr-accuracy-tolerance 0.02 `
  --sr-min-features 2 `
  --sr-max-features 3 `
  --sr-min-length 5 `
  --sr-max-length 60 `
  --sr-max-depth 10 `
  --sr-search-runs 3 `
  --sr-population-size 1800 `
  --sr-generations 40 `
  --sr-tournament-size 20 `
  --sr-parsimony 0.003 `
  --sr-functions add sub mul div `
  --sr-min-validation-r2 0.0 `
  --sr-min-test-r2 0.30 `
  --sr-max-abs-pbias 20 `
  --skip-cross-validation `
  --n-jobs -1 `
  --output sr_equation_results_extended `
  --overwrite
```

Feature constraints apply to equation eligibility, not genetic creation. If the
search finds no eligible two-predictor Pareto candidate, the exported result is
an audit fallback and is marked `NOT READY`. Compare any extended equation with
the unconstrained parsimonious equation and retain it only when it materially
improves validation RMSE, independent-holdout performance, or bias.

### Reducing runtime

Processing hundreds of workbooks with several full genetic-programming searches
can take many hours. Use a staged workflow:

1. Screen all cases with `--sr-fast-mode`.
2. Identify promising `READY` cases in `SR_Equations.xlsx`.
3. Rerun only those cases with the extended/final settings.

Use glob filters to run one mesh, case, or workbook:

```powershell
# One exact workbook
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --predictor-set paper-mean `
  --sr-publication-mode `
  --sr-fast-mode `
  --file-pattern "Mesh600-C4-150-300.xlsx" `
  --skip-cross-validation `
  --output sr_screen_one `
  --overwrite

# All workbooks whose filename contains Mesh600
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --predictor-set paper-mean `
  --sr-publication-mode `
  --sr-fast-mode `
  --file-pattern "*Mesh600*" `
  --skip-cross-validation `
  --output sr_screen_mesh600 `
  --overwrite

# A quick smoke test using only the first matched workbook
python run_elm_elmabc_symbolic.py `
  --models SymbolicRegression `
  --predictor-set paper-mean `
  --sr-publication-mode `
  --sr-fast-mode `
  --max-files 1 `
  --skip-cross-validation `
  --output sr_smoke_test `
  --overwrite
```

`--n-jobs -1` lets gplearn use the available CPU cores. Avoid running multiple
copies of the script simultaneously when each copy already uses `--n-jobs -1`,
because CPU and memory contention can make the total run slower.

The consolidated equations are written to:

```text
sr_equation_results/SR_Equations.xlsx
sr_equation_results/SR_Equations.csv
```

The Excel workbook includes:

- raw and simplified symbolic expressions;
- equation length, depth, operation count, and operator-weighted complexity;
- candidates pooled from all searches and their Pareto status;
- internal-validation metrics, holdout gates, and an explicit publication status;
- training and testing metrics;
- predictor validity ranges;
- protected-function definitions;
- Symbolic Regression settings.

### Common baseline models

Run all baseline models:

```powershell
python run_common_baselines.py `
  --models MLR RF XGBoost SVR GWR `
  --split original `
  --x-column X `
  --y-column Y `
  --overwrite
```

GWR requires coordinate columns. To run the nonspatial baselines without coordinates:

```powershell
python run_common_baselines.py `
  --models MLR RF XGBoost SVR `
  --split original `
  --overwrite
```

## Data-partition options

| Strategy | Description |
|---|---|
| `original` | First 70% of rows for training and final 30% for testing. This is the default in the ELM/ELM-ABC/SR script. |
| `random` | Reproducible random train/test split controlled by `--random-state`. This is the default in the baseline script. |
| `spatial-block` | Group-based spatial split requiring x and y coordinate columns. |

For direct comparison among all models, always pass the same split explicitly. For example:

```text
--split original --test-size 0.30 --random-state 42
```

Cross-validation is performed only on the training portion. Use `--skip-cross-validation` to disable it.

## Output structure

The ELM/ELM-ABC/SR workflow writes to `elm_symbolic_results/` by default, while the baseline workflow writes to `baseline_results/`.

```text
<output>/
├── all_cases_metrics.xlsx
├── all_cases_metrics.csv
├── model_comparison_summary.csv
├── audit.csv
├── run_configuration.json
├── README.txt
├── all_meshes/
│   └── <MODEL>/
│       └── <CASE>/
│           ├── Values.xlsx
│           ├── Statistics.xlsx
│           ├── residuals.csv
│           └── split_assignments.csv
└── best_meshes_600_700/
    └── <MODEL>/
        └── <CASE>/
            └── ...
```

The ELM/ELM-ABC/SR workflow additionally creates `SR_Equations.xlsx` and `SR_Equations.csv` when Symbolic Regression completes successfully.

`Statistics.xlsx` contains model performance and, when applicable, model parameters and fold-level cross-validation results. `Values.xlsx` contains observed and predicted values for the complete case.

## Performance metrics

The workflows report a common set of regression diagnostics, including:

- standard coefficient of determination (`R²`);
- squared Pearson correlation;
- RMSE and MAE;
- normalized RMSE;
- bias and percentage bias;
- Nash–Sutcliffe efficiency;
- Pearson correlation;
- residual standard deviation;
- predictions within ±10% of observations.

Always use the standard test-set `R²`, RMSE, and MAE as the primary out-of-sample comparison metrics. The squared correlation is also reported for compatibility but should not be confused with the standard `R²`.

## Reproducibility and safeguards

- The default random seed is `42`.
- Predictor leakage is reduced by excluding identifiers, coordinates, the target, and recognized spatial-index fields from automatic predictor selection.
- Raw ELM or ELM-ABC predictions are never used as inputs to the baseline models.
- Every run records its arguments in `run_configuration.json`.
- `split_assignments.csv` records the exact training and testing membership.
- `audit.csv` records successful cases, skipped models, and processing errors.
- `--overwrite` deletes and recreates only the selected output directory. Use it carefully.
- Use `--fail-on-missing-model` when a missing optional dependency should stop the complete run instead of being recorded in the audit.

## Useful command options

```text
--root PATH
--data-folder NAME
--output PATH
--file-pattern GLOB
--max-files N
--models MODEL [MODEL ...]
--predictor-set {paper-mean,paper-summary,numeric-safe}
--predictors COLUMN [COLUMN ...]
--id-column COLUMN
--target-column COLUMN
--split {original,random,spatial-block}
--test-size 0.30
--cv-folds 10
--skip-cross-validation
--random-state 42
--best-meshes 600 700
--clip-negative-predictions
--overwrite
--sr-fast-mode
--sr-min-features N
--sr-max-features N
--sr-min-length N
```

Display the complete options for either workflow:

```bash
python run_elm_elmabc_symbolic.py --help
python run_common_baselines.py --help
```

## Citation

If you use this repository in a publication, please cite the associated water-consumption study and this software repository. The formal paper citation can be added here after publication.

```bibtex
@software{shakouri_waterconsumption_models,
  author  = {Behzad Shakouri},
  title   = {Water Consumption Modeling with ELM, ELM-ABC, Symbolic Regression, and Common Baselines},
  url     = {https://github.com/behzadshakouri/WaterConsumption_ELM_ELMABC_SR},
  year    = {2026}
}
```

## License

No license has been specified yet. Add a `LICENSE` file before redistribution or third-party reuse. A permissive license such as MIT or BSD-3-Clause is commonly used for research software.
