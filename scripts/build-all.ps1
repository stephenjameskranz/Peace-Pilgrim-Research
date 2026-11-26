# Rebuilds every output from the files in data\.
#   peace-pilgrim-sources.csv, peace-pilgrim-location-events.csv, peace-pilgrim-itinerary.xlsx, peace-pilgrim-events.geojson
# Add -Geocode to look up any new places with OpenStreetMap Nominatim first.
param([switch]$Geocode)
$root = Split-Path $PSScriptRoot -Parent
& "$PSScriptRoot\build.ps1" "$root\data\links.csv" "$root\data\index_parsed.csv" | Select-Object -First 1
& "$PSScriptRoot\assemble.ps1" $root $root
if ($Geocode) { & "$PSScriptRoot\geocode.ps1" $root; & "$PSScriptRoot\assemble.ps1" $root $root }
& "$PSScriptRoot\export-geojson.ps1" $root
