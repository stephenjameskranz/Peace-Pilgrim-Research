# Downloads every clipping scan listed in data\links.csv into clippings\ (skips files already present).
# The scans are hosted by Friends of Peace Pilgrim at peacepilgrim.org and are not stored in the repo.
param($root = (Split-Path $PSScriptRoot -Parent))
New-Item -ItemType Directory -Force "$root\clippings" | Out-Null
$urls = Import-Csv "$root\data\links.csv" | ForEach-Object url | Where-Object { $_ -match '/s/' } | Sort-Object -Unique
foreach ($u in $urls) {
  $f = Join-Path "$root\clippings" ($u -split '/')[-1]
  if (-not (Test-Path $f)) {
    try { Invoke-WebRequest -UseBasicParsing $u -OutFile $f -TimeoutSec 60 } catch { Write-Warning "failed: $u" }
  }
}
"$((Get-ChildItem "$root\clippings").Count) files in clippings\"
