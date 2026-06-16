# Custom Video Request add-in — one-shot GitHub setup (Windows PowerShell)
# Run from this folder:  .\init-repo.ps1
# Prereq: create an EMPTY repo named "custom-video-request-addin" under your
# GitHub account first (no README/license), OR install gh and uncomment the
# gh line below to create it automatically.

$ErrorActionPreference = "Stop"
$RepoUrl = "https://github.com/Austin-GPSFMS/custom-video-request-addin.git"

# 1. Remove the partial .git left by the sandbox (safe if it doesn't exist)
if (Test-Path ".git") { Remove-Item -Recurse -Force ".git" }

# 2. Fresh repo
git init
git add -A
git commit -m "Custom Video Request add-in: page + Trips History map variants, direct media-services, GitHub Pages ready"
git branch -M main

# 3. Optional: auto-create the GitHub repo (needs GitHub CLI: https://cli.github.com)
# gh repo create Austin-GPSFMS/custom-video-request-addin --public --source=. --remote=origin --push; exit

# 4. Point at the remote and push
git remote add origin $RepoUrl
git push -u origin main

Write-Host ""
Write-Host "Done. Now enable GitHub Pages: repo Settings > Pages > Branch: main / (root)."
Write-Host "Then paste config.hosted.json (page) and/or trips-history/configuration.json (Trips History) into MyGeotab once."
