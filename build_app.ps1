$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$python = Join-Path $projectRoot "venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw "Project virtual environment not found at $python. Install requirements first."
}

Push-Location $projectRoot
try {
    & $python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name "SpecificationFormatter" `
        --paths $projectRoot `
        --collect-all customtkinter `
        --add-data "master_prompt.txt;." `
        --add-data "run_instruction_prompt.txt;." `
        --add-data "spec_formatter\style_application\core\prompts;spec_formatter\style_application\core\prompts" `
        gui.py
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

$executable = Join-Path $projectRoot "dist\SpecificationFormatter.exe"
if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
    throw "PyInstaller completed without creating $executable."
}
Write-Host "Built $executable"
