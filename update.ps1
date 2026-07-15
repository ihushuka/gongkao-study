$ErrorActionPreference = "Stop"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js was not found. Reopen PowerShell after installing Node.js."
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw "Git was not found. Reopen PowerShell after installing Git."
}

npm.cmd run build
if ($LASTEXITCODE -ne 0) {
  throw "Build failed. Do not push until the error is fixed."
}

git add src/App.tsx src/index.css public/daily-news.json scripts/fetch-news.mjs .github/workflows/update-daily-news.yml update.ps1
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
  Write-Host "No new changes to upload." -ForegroundColor Yellow
  exit 0
}

git commit -m "Add themed controls fresh daily news OCR review and data validation"
if ($LASTEXITCODE -ne 0) {
  throw "Git commit failed."
}

git push
if ($LASTEXITCODE -ne 0) {
  throw "Git push failed. Check the proxy or network, then run git push again."
}

if (Get-Command gh -ErrorAction SilentlyContinue) {
  gh workflow run update-daily-news.yml
  if ($LASTEXITCODE -ne 0) {
    Write-Host "Code uploaded, but the first news refresh was not started. Run the workflow manually in GitHub Actions." -ForegroundColor Yellow
  }
}

Write-Host "Update uploaded. GitHub Pages will refresh automatically in a few minutes." -ForegroundColor Green
