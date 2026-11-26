# Writes peace-pilgrim-events.geojson (one point per geocoded event) for the web map.
param($root = (Split-Path $PSScriptRoot -Parent))
$ev = Import-Csv "$root\peace-pilgrim-location-events.csv" | Where-Object { $_.lat -and $_.lon }
$features = foreach ($e in $ev) {
  $props = [ordered]@{}
  foreach ($p in $e.PSObject.Properties) { if ($p.Name -notin 'lat', 'lon') { $props[$p.Name] = $p.Value } }
  [ordered]@{ type = 'Feature'; geometry = [ordered]@{ type = 'Point'; coordinates = @([double]$e.lon, [double]$e.lat) }; properties = $props }
}
$fc = [ordered]@{ type = 'FeatureCollection'; features = @($features) }
[IO.File]::WriteAllText("$root\peace-pilgrim-events.geojson", (ConvertTo-Json $fc -Depth 6 -Compress), (New-Object Text.UTF8Encoding $false))
"$(@($features).Count) features -> peace-pilgrim-events.geojson"
