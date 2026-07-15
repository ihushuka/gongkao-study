param(
  [string]$RepoName = "gongkao-study"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
  throw "GitHub CLI was not found. Close PowerShell, open it again, and retry."
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw "Git was not found. Run: winget install --id Git.Git --source winget. Then reopen PowerShell."
}
gh auth status
$owner = gh api user --jq .login

git init

if (-not (git config user.name)) {
  git config user.name $owner
}
if (-not (git config user.email)) {
  git config user.email "$owner@users.noreply.github.com"
}

git add .
git commit -m "Create gongkao study site"
git branch -M main
gh repo create $RepoName --public --source=. --remote=origin
gh api --method POST "repos/$owner/$RepoName/pages" -f build_type=workflow
git push -u origin main

Write-Host "Source uploaded. Waiting for GitHub Pages deployment..." -ForegroundColor Cyan
Start-Sleep -Seconds 6
gh run watch --exit-status

Write-Host "Published: https://$owner.github.io/$RepoName/" -ForegroundColor Green
