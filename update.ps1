$ErrorActionPreference = "Stop"

if (-not (Test-Path ".git")) {
  throw "Run this script from the original project folder."
}
if (-not (Get-Command npm.cmd -ErrorAction SilentlyContinue)) {
  throw "Node.js was not found. Reopen PowerShell after installing Node.js."
}

npm.cmd run build
if ($LASTEXITCODE -ne 0) { throw "Build failed. Do not push until the error is fixed." }

git add src/App.tsx src/index.css update.ps1
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
  Write-Host "No new source changes to commit. Checking remote sync." -ForegroundColor Yellow
} else {
  git commit -m "Add weekly monthly reports and active anti-delay supervision"
  if ($LASTEXITCODE -ne 0) { throw "git commit failed." }
}

git push
if ($LASTEXITCODE -ne 0) { throw "git push failed. Check the network or proxy, then run git push again." }

Write-Host "Update uploaded. GitHub Pages will refresh automatically in a few minutes." -ForegroundColor Green
