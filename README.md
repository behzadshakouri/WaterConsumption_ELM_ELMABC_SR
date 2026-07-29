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

Two PowerShell runners are included. `run_all_paper_methods.ps1` processes all
supported Excel workbooks under `data`, while
`run_mesh600_700_paper_methods.ps1` processes only filenames beginning with
`Mesh600` or `Mesh700`.

## Repository files

| File | Purpose |
|---|---|
| `run_elm_elmabc_symbolic.py` | Runs ELM, ELM-ABC, and Symbolic Regression. It also simplifies and consolidates SR equations. |
| `run_common_baselines.py` | Runs MLR, RF, XGBoost, SVR, and GWR baseline models. |
| `run_all_paper_methods.ps1` | Runs every workbook: all non-SR methods with 10-fold CV and SR without CV. |
| `run_mesh600_700_paper_methods.ps1` | Runs only `Mesh600*` and `Mesh700*` with the same CV design. |
| `summarize_supplementary_results.py` | Creates separate non-SR and SR supplementary workbooks for Mesh600/700 after successful runs. |

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
├── run_all_paper_methods.ps1
├── run_mesh600_700_paper_methods.ps1
└── README.md
```

The scripts recursively discover `.xlsx`, `.xlsm`, and `.xls` files under `data/`. Each workbook must contain a header row. Unless explicitly overridden:

- the first column is treated as the observation or grid-cell identifier;
- the paper-summary variables are used as predictors;
- the last column is treated as the case-specific paper target: either NHWSCC
  or annual HWC;
- the first non-empty worksheet is analyzed.

Column numbers used by the command-line options are **one-based**, matching
Excel (`A = 1`, `B = 2`, and so on).

## Predictor sets

Three predictor-selection modes are available:

| Option | Predictors |
|---|---|
| `paper-mean` | `T96_mean`, `P_mean`, `Elv_mean` |
| `paper-summary` | Mean, minimum, and maximum values of `T96`, `P`, and `Elv` |
| `numeric-safe` | Eligible numeric columns, excluding IDs, coordinates, the target, and other potentially unsafe fields |

`paper-summary` is the default because these are the inputs used in the paper:
land-surface temperature (`T96`), land value (`P`), and elevation (`Elv`),
each represented by its mean, minimum, and maximum. In the supplied 16-column
data layout, these nine inputs are columns **7–15**. The output is the final
column, **16**, but its meaning is case-specific: **NHWSCC** for consumption-
class subscriber-count workbooks and **annual HWC** for annual-consumption
workbooks.

| Column | Workbook header | Default role |
|---:|---|---|
| 6 | `bottom` | Coordinate; excluded |
| 7 | `T96_mean` | Input |
| 8 | `T96_min` | Input |
| 9 | `T96_max` | Input |
| 10 | `P_mean` | Input |
| 11 | `P_min` | Input |
| 12 | `P_max` | Input |
| 13 | `Elv_mean` | Input |
| 14 | `Elv_min` | Input |
| 15 | `Elv_max` | Input |
| 16 | Case-specific NHWSCC or annual-HWC field | Output |

The scripts resolve the paper inputs by header name rather than fixed position,
so the default remains correct if the columns are reordered. The target
defaults to the final column because its header changes among cases. Targets
whose names use the paper's `NC...` convention (for example, `NC4-150-300`,
`NC8-390-455`, and `NC10-L50`) are recorded as **NHWSCC**; annual/yearly HWC
cases are recorded as **Annual HWC**. The audit and metrics outputs include
both the exact workbook header (`target`) and this conceptual classification
(`target_variable`).

`P_mean`, `P_min`, and `P_max` mean **land value**. `P` does not mean
precipitation in this study.

## Selecting inputs and output

Both runners support exact column names and Excel-style column numbers.

Use one-based column numbers:

```powershell
python run_elm_elmabc_symbolic.py `
  --root . `
  --input-columns 1,4,6 `
  --output-column 8 `
  --overwrite
```

Spaces are also accepted:

```powershell
python run_common_baselines.py `
  --root . `
  --input-columns 1 4 6 `
  --output-column 8 `
  --models MLR RF XGBoost SVR `
  --overwrite
```

Selection precedence is:

1. `--input-columns` overrides `--predictors` and `--predictor-set`.
2. `--predictors` overrides `--predictor-set`.
3. Without either override, `--predictor-set paper-summary` is used.
4. `--output-column` overrides `--target-column`; otherwise the final column
   is the output.

The output column is never permitted as an input. Explicit numerical input
selection is treated as intentional, so any valid workbook column can be used,
including column 1. Every selected position is checked separately for every
workbook.

To state the paper configuration by positions instead of using its name-based
default:

```powershell
python run_elm_elmabc_symbolic.py `
  --root . `
  --input-columns 7 8 9 10 11 12 13 14 15 `
  --output-column 16 `
  --split random `
  --overwrite
```

For maximum reproducibility across reordered files, the default
`--predictor-set paper-summary` plus the default final-column target is
recommended.

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

Install the shared dependencies required by **both** Python files:

```bash
python -m pip install --upgrade pip
python -m pip install numpy pandas openpyxl scikit-learn
```

Then install the optional packages for the models you intend to run:

```bash
# Symbolic Regression with the legacy gplearn engine
python -m pip install gplearn

# Symbolic Regression with the recommended PySR engine
python -m pip install pysr

# XGBoost baseline
python -m pip install xgboost

# Geographically Weighted Regression baseline
python -m pip install mgwr

# Only needed when reading legacy .xls workbooks
python -m pip install xlrd
```

To install everything used by either script in one command:

```bash
python -m pip install numpy pandas openpyxl scikit-learn gplearn pysr xgboost mgwr xlrd
```

### Dependency requirements by Python file

| Python file / requested model | Required packages |
|---|---|
| Both scripts: data handling, Excel output, splitting, preprocessing, metrics, MLR, RF, and SVR | `numpy`, `pandas`, `openpyxl`, `scikit-learn` |
| `run_elm_elmabc_symbolic.py`: ELM and ELM-ABC | Shared packages only |
| `run_elm_elmabc_symbolic.py`: `--sr-engine gplearn` | Shared packages + `gplearn` |
| `run_elm_elmabc_symbolic.py`: `--sr-engine pysr` | Shared packages + `pysr` |
| `run_common_baselines.py`: XGBoost | Shared packages + `xgboost` |
| `run_common_baselines.py`: GWR | Shared packages + `mgwr` |
| Either script: legacy `.xls` input | Add `xlrd` |

Notes:

- `gplearn` is required only when Symbolic Regression uses
  `--sr-engine gplearn`.
- `pysr` enables the recommended high-performance evolutionary SR engine. Its
  Julia backend is installed/configured automatically on first import, so the
  first run takes longer than later runs.
- `xgboost` is required only for XGBoost.
- `mgwr` is required only for GWR.
- Standard `.xlsx` and `.xlsm` workbooks use `openpyxl`; legacy `.xls`
  workbooks additionally require `xlrd`.
- Python standard-library modules such as `argparse`, `json`, `math`, `pathlib`,
  and `shutil` are included with Python and must not be installed separately.

Verify the shared installation and both command-line interfaces:

```bash
python -c "import numpy, pandas, openpyxl, sklearn; print('Shared dependencies: OK')"
python run_elm_elmabc_symbolic.py --help
python run_common_baselines.py --help
```

Verify the optional packages after installing all models:

```bash
python -c "import gplearn, pysr, xgboost, mgwr, xlrd; print('Optional dependencies: OK')"
```

## Quick start

### Run every paper method with the requested CV design

On Windows PowerShell, the included runner executes:

- ELM and ELM-ABC with 10-fold cross-validation;
- MLR, RF, XGBoost, SVR, and GWR with 10-fold cross-validation;
- Symbolic Regression separately, without cross-validation.

ELM evaluates 100 reproducible random initializations inside each
outer-training set and selects the realization using a 20% training-only
validation subset. ELM-ABC searches hidden-layer weights and biases while
recomputing its output weights analytically; output weights are not forced
into the hidden-parameter bounds. Neither procedure uses the independent test
subset for selection.

It resolves the nine paper inputs by header name (`T96`, land value `P`, and
elevation; mean/min/max), uses each workbook's final column as its case-specific
NHWSCC or Annual-HWC output, applies the reproducible random 70/30 paper split
with seed 42, and writes separate output folders. In the supplied 16-column layout,
this is equivalent to input columns 7–15 and output column 16:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\run_all_paper_methods.ps1
```

To run only Mesh600 and Mesh700 workbooks:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\run_mesh600_700_paper_methods.ps1
```

PySR is the default SR engine. To use the legacy gplearn engine:

```powershell
.\run_all_paper_methods.ps1 -SrEngine gplearn
```

The same options are accepted by the restricted runner:

```powershell
.\run_mesh600_700_paper_methods.ps1 -SrEngine gplearn
```

The runner uses `.venv\Scripts\python.exe` when available and otherwise uses
`python` from `PATH`. Existing result folders are replaced by default; pass
`-NoOverwrite` to keep them.

### Create the two supplementary-material workbooks

After the Mesh600/700 runner completes, run:

```powershell
python summarize_supplementary_results.py --overwrite
```

This creates:

```text
supplementary_material\Supplementary_NonSR_Mesh600_700.xlsx
supplementary_material\Supplementary_SR_Mesh600_700.xlsx
```

The summarizer requires successful results for all seven non-SR methods:
ELM, ELM-ABC, MLR, RF, XGBoost, SVR, and GWR. If one is absent, it stops and
points to the baseline audit instead of silently producing an incomplete
supplementary workbook.

### ELM and ELM-ABC only

The following command uses the reproducible random 70/30 paper split, runs ELM and ELM-ABC, performs 10-fold cross-validation on the training subset, and replaces an existing output directory:

```powershell
python run_elm_elmabc_symbolic.py `
  --models ELM ELMABC `
  --split random `
  --overwrite
```

To skip the additional cross-validation:

```powershell
python run_elm_elmabc_symbolic.py `
  --models ELM ELMABC `
  --split random `
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
  --split random `
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
  --split random `
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
  --split random `
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
  --split random `
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
  --split random `
  --x-column X `
  --y-column Y `
  --overwrite
```

GWR requires coordinate columns. To run the nonspatial baselines without coordinates:

```powershell
python run_common_baselines.py `
  --models MLR RF XGBoost SVR `
  --split random `
  --overwrite
```

## Data-partition options

| Strategy | Description |
|---|---|
| `original` | First 70% of rows for training and final 30% for testing. Retained only for reproducing legacy row-order runs. |
| `random` | Reproducible random train/test split controlled by `--random-state`. This is the paper workflow used by both PowerShell runners. |
| `spatial-block` | Group-based spatial split requiring x and y coordinate columns. |

For direct comparison among all models, always pass the same split explicitly. For example:

```text
--split random --test-size 0.30 --random-state 42
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
- The included PowerShell runners use `--fail-on-error`, so a missing package
  or any failed requested model (including GWR) stops the run instead of
  printing a misleading success message.
- Use `--fail-on-missing-model` in custom commands when only missing optional
  dependencies should be fatal.

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
--input-columns N [N ...]
--id-column COLUMN
--target-column COLUMN
--output-column N
--split {original,random,spatial-block}
--test-size 0.30
--cv-folds 10
--skip-cross-validation
--random-state 42
--best-meshes 600 700
--clip-negative-predictions
--overwrite
--fail-on-error
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
