# Geocodes the places in peace-pilgrim-location-events.csv using OpenStreetMap Nominatim.
# Results are cached in data\geocode_cache.csv so each place is looked up only once.
# Usage: scripts\geocode.ps1 [-root <project folder>]
param($root = (Split-Path $PSScriptRoot -Parent))
. "$PSScriptRoot\placekey.ps1"

$cachePath = "$root\data\geocode_cache.csv"
$cache = @{}
if (Test-Path $cachePath) { Import-Csv $cachePath | ForEach-Object { $cache[$_.key] = $_ } }

$events = Import-Csv "$root\peace-pilgrim-location-events.csv"
$todo = @{}
foreach ($e in $events) {
  $p = Get-PlaceKey $e.city $e.state_province $e.country
  if ($p -and -not $cache.ContainsKey($p.key)) { $todo[$p.key] = $p }
}
"$($todo.Count) places to look up ($($cache.Count) already cached)"

$headers = @{ 'User-Agent' = 'peace-pilgrim-itinerary-research/1.0'; 'Accept-Language' = 'en' }
foreach ($p in $todo.Values) {
  $url = 'https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=' + [uri]::EscapeDataString($p.query)
  $r = $null
  try { $r = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 30 } catch { Write-Warning "failed: $($p.query)" }
  $hit = @($r)[0]
  $cache[$p.key] = [pscustomobject][ordered]@{
    key = $p.key; query = $p.query; level = $p.level
    lat = $(if ($hit) { '{0:N4}' -f [double]$hit.lat -replace ',', '' } else { '' })
    lon = $(if ($hit) { '{0:N4}' -f [double]$hit.lon -replace ',', '' } else { '' })
    osm_name = $(if ($hit) { $hit.display_name } else { 'NOT FOUND' })
    osm_type = $(if ($hit) { "$($hit.category)/$($hit.type)" } else { '' })
  }
  Start-Sleep -Milliseconds 1100
}
$cache.Values | Sort-Object key | Export-Csv -NoTypeInformation -Encoding UTF8 $cachePath
"cache now $($cache.Count) places; not found: $(@($cache.Values | ? { -not $_.lat }).Count)"
