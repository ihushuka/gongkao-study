$ErrorActionPreference = "Stop"

if (-not (Test-Path ".git")) {
  throw "This script must be run inside the gongkao-study project folder."
}
if (-not (Get-Command npm.cmd -ErrorAction SilentlyContinue)) {
  throw "Node.js was not found. Reopen PowerShell after installing Node.js."
}

npm.cmd run build
if ($LASTEXITCODE -ne 0) {
  throw "Build failed. Do not push until the error is fixed."
}

git add src/App.tsx src/index.css update.ps1
if ($LASTEXITCODE -ne 0) {
  throw "git add failed."
}

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
  Write-Host "No new changes were found. The v5 files may already be installed."
  exit 0
}

git commit -m "Improve editable practice options idiom verification and daily timeline"
if ($LASTEXITCODE -ne 0) {
  throw "git commit failed."
}

git push
if ($LASTEXITCODE -ne 0) {
  throw "git push failed. The commit is safe locally; fix the network and run git push again."
}

Write-Host "Update uploaded. GitHub Pages will refresh automatically in a few minutes."
