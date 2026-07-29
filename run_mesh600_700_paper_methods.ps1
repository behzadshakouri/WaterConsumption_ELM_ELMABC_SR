[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [string]$DataFolder = "data",
    [string]$Python = "",
    [ValidateSet("pysr", "gplearn")]
    [string]$SrEngine = "pysr",
    [switch]$NoOverwrite
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Root = (Resolve-Path $Root).Path
$MainScript = Join-Path $Root "run_elm_elmabc_symbolic.py"
$BaselineScript = Join-Path $Root "run_common_baselines.py"
$DataPath = Join-Path $Root $DataFolder

if (-not (Test-Path $MainScript -PathType Leaf)) {
    throw "Missing script: $MainScript"
}
if (-not (Test-Path $BaselineScript -PathType Leaf)) {
    throw "Missing script: $BaselineScript"
}
if (-not (Test-Path $DataPath -PathType Container)) {
    throw "Missing data folder: $DataPath"
}

if ([string]::IsNullOrWhiteSpace($Python)) {
    $VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
    if (Test-Path $VenvPython -PathType Leaf) {
        $Python = $VenvPython
    }
    else {
        $Python = "python"
    }
}

$CommonArguments = @(
    "--root", $Root,
    "--data-folder", $DataFolder,
    "--predictor-set", "paper-summary",
    "--split", "original",
    "--test-size", "0.30",
    "--random-state", "42",
    "--fail-on-error",
    "--file-pattern", "Mesh600*",
    "--file-pattern", "Mesh700*",
    "--best-meshes", "600", "700"
)

$OverwriteArgument = @()
if (-not $NoOverwrite) {
    $OverwriteArgument = @("--overwrite")
}

function Invoke-PaperStage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$Script,
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Host ""
    Write-Host "===== $Title =====" -ForegroundColor Cyan
    & $Python $Script @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Title failed with exit code $LASTEXITCODE."
    }
}

Invoke-PaperStage `
    -Title "ELM and ELM-ABC: 10-fold CV" `
    -Script $MainScript `
    -Arguments (@(
        "--models", "ELM", "ELMABC",
        "--cv-folds", "10",
        "--output", "results_10fold_elm_elmabc"
    ) + $CommonArguments + $OverwriteArgument)

Invoke-PaperStage `
    -Title "MLR, RF, XGBoost, SVR, and GWR: 10-fold CV" `
    -Script $BaselineScript `
    -Arguments (@(
        "--models", "MLR", "RF", "XGBoost", "SVR", "GWR",
        "--cv-folds", "10",
        "--output", "results_10fold_baselines"
    ) + $CommonArguments + $OverwriteArgument)

Invoke-PaperStage `
    -Title "Symbolic Regression: no cross-validation" `
    -Script $MainScript `
    -Arguments (@(
        "--models", "SymbolicRegression",
        "--sr-engine", $SrEngine,
        "--sr-publication-mode",
        "--skip-cross-validation",
        "--output", "results_sr_no_cv"
    ) + $CommonArguments + $OverwriteArgument)

Write-Host ""
Write-Host "Mesh600/Mesh700 paper methods completed successfully." -ForegroundColor Green
Write-Host "Processed files: Mesh600* and Mesh700* only"
Write-Host "ELM/ELM-ABC: $(Join-Path $Root 'results_10fold_elm_elmabc')"
Write-Host "Baselines:   $(Join-Path $Root 'results_10fold_baselines')"
Write-Host "SR (no CV):  $(Join-Path $Root 'results_sr_no_cv')"
