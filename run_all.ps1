param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $ProjectRoot
try {
    & $Python "src/run_experiment.py"
    if ($LASTEXITCODE -ne 0) { throw "Experiment failed." }

    & $Python "src/evaluate_checkpoints.py"
    if ($LASTEXITCODE -ne 0) { throw "Checkpoint verification failed." }

    & $Python "src/build_paper.py"
    if ($LASTEXITCODE -ne 0) { throw "Paper build failed." }

    & $Python "src/make_manifest.py"
    if ($LASTEXITCODE -ne 0) { throw "Manifest build failed." }
}
finally {
    Pop-Location
}
