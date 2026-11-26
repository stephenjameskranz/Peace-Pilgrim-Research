param($dir, $outDir)
$src = @(Import-Csv "$dir\data\index_parsed.csv")
if (Test-Path "$dir\data\extra_sources.csv") {
  $cols = $src[0].PSObject.Properties.Name
  foreach ($x in Import-Csv "$dir\data\extra_sources.csv") {
    $o = [ordered]@{}; foreach ($c in $cols) { $o[$c] = '' }
    foreach ($p in $x.PSObject.Properties) { $o[$p.Name] = $p.Value }
    $o['date_meaning'] = $src[0].date_meaning; $o['clipping_read'] = 'no'
    $src += [pscustomobject]$o
  }
}
$reads = @{}; Import-Csv "$dir\data\reads.csv" | ForEach-Object { $reads[$_.id] = $_ }
foreach ($s in $src) {
  $r = $reads[$s.id]
  if ($r) {
    $s.publication = $r.publication; $s.miles_reported = $r.miles_reported; $s.route_before = $r.route_before
    $s.route_planned = $r.route_planned; $s.people_mentioned = $r.people_mentioned; $s.notable_details = $r.notable_details
    $s.clipping_read = 'yes'
    if ($r.date_note) { $s.notes = ("Clipping: " + $r.date_note + ' ' + $s.notes).Trim() }
  }
}
# auto events from index-only sources
$manual = Import-Csv "$dir\data\events_manual.csv"
$covered = @{}; $manual | ForEach-Object { $_.source_ids -split ';' | ForEach-Object { $covered[$_] = 1 } }
# sequence date: for loosely dated entries, use the last day-precise date before it in index order (index is roughly chronological)
$seq = @{}; $last = ''
foreach ($s in $src) {
  $y = if ($s.date -match '^(\d{4})') { $Matches[1] } else { '' }
  if ($s.date_precision -eq 'day') { $last = $s.date }
  elseif ($y -and -not $last.StartsWith($y)) { $last = "$y-01-01" }
  $seq[$s.id] = $last
}
$auto = @{}
foreach ($s in $src) {
  if ($reads[$s.id] -or $covered[$s.id]) { continue }
  if ($s.location_confidence -in 'No location', 'Not a location of Peace', 'Country only') { continue }
  if (-not $s.city -and -not $s.state_province) { continue }
  $d = $s.date; $ds = ''; $de = ''
  switch ($s.date_precision) {
    'day' { $ds = $d; $de = $d }
    'month' { $ds = "$d-01"; $de = ([datetime]"$d-01").AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd') }
    'year' { $ds = "$d-01-01"; $de = "$d-12-31" }
  }
  $cert = switch -wildcard ($s.location_confidence) {
    'Documented' { 'Known' } 'Documented (date approximate)' { 'Known place / approx. date' }
    'Documented (place of death)' { 'Known' } 'Probable' { 'Probable' } default { $s.location_confidence }
  }
  $key = "$($s.city)|$($s.state_province)|$ds|$de"
  if ($auto[$key]) { $auto[$key].source_ids += ";$($s.id)"; continue }
  $act = if ($s.record_type -eq 'Death report') { 'Killed in a car crash near Knox, Indiana' } else { 'Reported in local newspaper (from index title; clipping not yet read)' }
  $auto[$key] = [pscustomobject][ordered]@{
    date_start = $ds; date_end = $de; date_basis = 'Publication date (index)'; city = $s.city; state_province = $s.state_province; country = $s.country
    lat = ''; lon = ''; certainty = $cert; evidence = 'Newspaper report'; activity = $act; source_ids = $s.id; notes = $s.notes
  }
}
function SortKey($e) {
  if (-not $e.date_start) { return '9999' }
  $span = ([datetime]$e.date_end - [datetime]$e.date_start).TotalDays
  if ($span -le 62) { return $e.date_start + '|' + $e.date_end }
  $first = ($e.source_ids -split ';')[0]
  if ($seq[$first] -and $seq[$first].StartsWith($e.date_start.Substring(0, 4))) { return $seq[$first] + '|z' + $first }
  return $e.date_start + '|' + $e.date_end
}
$events = @($manual) + @($auto.Values) | Sort-Object { SortKey $_ }
# coordinates from the geocode cache (scripts\geocode.ps1)
. "$PSScriptRoot\placekey.ps1"
$geo = @{}; if (Test-Path "$dir\data\geocode_cache.csv") { Import-Csv "$dir\data\geocode_cache.csv" | ForEach-Object { $geo[$_.key] = $_ } }
$n = 0
$events = $events | ForEach-Object {
  $n++; $o = [ordered]@{ event_id = 'E{0:D4}' -f $n }
  foreach ($p in $_.PSObject.Properties) { $o[$p.Name] = $p.Value }
  $k = Get-PlaceKey $_.city $_.state_province $_.country
  $g = if ($k) { $geo[$k.key] } else { $null }
  $o['lat'] = if ($g) { $g.lat } else { '' }
  $o['lon'] = if ($g) { $g.lon } else { '' }
  $out = [ordered]@{}
  foreach ($key in $o.Keys) {
    $out[$key] = $o[$key]
    if ($key -eq 'lon') {
      $out['geo_level'] = if ($g) { $g.level } elseif ($k) { 'not geocoded yet' } else { 'none (route or no place)' }
      $out['geo_place'] = if ($g) { $g.osm_name } else { '' }
    }
  }
  [pscustomobject]$out
}

$src | Export-Csv -NoTypeInformation -Encoding UTF8 "$outDir\peace-pilgrim-sources.csv"
$events | Export-Csv -NoTypeInformation -Encoding UTF8 "$outDir\peace-pilgrim-location-events.csv"

# ---------- xlsx ----------
function Esc($v) { if ($null -eq $v) { return '' }; [Security.SecurityElement]::Escape([string]$v) -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '' }
function Col($i) { $s = ''; $i++; while ($i -gt 0) { $m = ($i - 1) % 26; $s = [char](65 + $m) + $s; $i = [int][math]::Floor(($i - $m) / 26) }; $s }
function Sheet($rows, $widths) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>')
  for ($c = 0; $c -lt $widths.Count; $c++) { [void]$sb.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$($widths[$c])`" customWidth=`"1`"/>") }
  [void]$sb.Append('</cols><sheetData>')
  for ($r = 0; $r -lt $rows.Count; $r++) {
    [void]$sb.Append("<row r=`"$($r+1)`">")
    $row = $rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
      $v = $row[$c]; if ($null -eq $v -or $v -eq '') { continue }
      $ref = (Col $c) + ($r + 1); $st = if ($r -eq 0) { ' s="1"' } else { ' s="2"' }
      [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"$st><is><t xml:space=`"preserve`">$(Esc $v)</t></is></c>")
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>'); $sb.ToString()
}
function ToRows($objs) {
  $props = @($objs[0].PSObject.Properties.Name)
  $rows = New-Object System.Collections.ArrayList
  [void]$rows.Add($props)
  foreach ($o in $objs) { [void]$rows.Add(@($props | ForEach-Object { $o.$_ })) }
  , $rows
}
$about = @(
 @('Peace Pilgrim - Itinerary Research Workbook', ''),
 @('Purpose', 'Documented record of where and when Peace Pilgrim (Mildred Norman, 1908-1981) was during her pilgrimage (1953-1981), to drive a map on moop.earth.'),
 @('Compiled', (Get-Date -Format 'yyyy-MM-dd')),
 @('Primary source list', 'https://www.peacepilgrim.org/articles-about-peace-pilgrim (Friends of Peace Pilgrim newspaper clipping archive)'),
 @('Tab: Location events', 'One row per claim about where she was and when. A single clipping can produce several rows (e.g. the route she says she walked). This is the map dataset.'),
 @('Tab: Sources', 'One row per clipping listed on the archive page, with date, place, headline, link and notes. Columns miles_reported to notable_details are filled only where clipping_read = yes.'),
 @('certainty: Known', 'A source places her at this place on this date (or within a day or two of it).'),
 @('certainty: Known place / approx. date', 'A source places her here, but the date is only known to the month, year or a window.'),
 @('certainty: Probable', 'Named in a route she reported, or an obvious fix to an index error; date estimated from route order.'),
 @('certainty: Estimated', '(Reserved) Locations interpolated between known points - to be added when building the map.'),
 @('date_basis', 'Publication date = the date the newspaper ran the story; she was usually there that day or 1-2 days earlier. Stated in article = the article gives the day of the event.'),
 @('Source IDs', 'source_ids in Location events refer to the id column in Sources.'),
 @('Index caveats', 'The archive index has some errors (state abbreviations, dates, broken links pointing at a placeholder file). These are flagged in the notes column, never silently changed: index_title keeps the original wording.'),
 @('Status', "Clippings read so far: $(@($src | ? clipping_read -eq 'yes').Count) of $($src.Count). Coordinates: 0 of 0 events geocoded with OpenStreetMap Nominatim."),
 @('Coordinates', 'lat/lon come from OpenStreetMap Nominatim (data/geocode_cache.csv). geo_level says what the point is: city = the named town; city (first of several) = the row names a route or several towns and the point is the first one; state / region = only a state or region is known, so the point is its centre and should be drawn as approximate; landmark = a specific site (bridge, peak, building).')
)
$sheets = @(
  @{ name = 'About'; xml = (Sheet $about @(34, 140)) },
  @{ name = 'Location events'; xml = (Sheet (ToRows $events) @(9, 11, 11, 24, 26, 8, 8, 9, 10, 18, 40, 24, 24, 60, 14, 70)) },
  @{ name = 'Sources'; xml = (Sheet (ToRows $src) @(6, 11, 9, 30, 18, 7, 8, 20, 22, 24, 50, 26, 7, 14, 40, 40, 30, 70, 8, 60, 50, 11, 70)) }
)
$tmp = Join-Path $env:TEMP ("xlsx_" + [guid]::NewGuid()); New-Item -ItemType Directory "$tmp\_rels", "$tmp\xl\_rels", "$tmp\xl\worksheets" -Force | Out-Null
$enc = New-Object Text.UTF8Encoding $false
$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
$wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
$wbr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
for ($i = 0; $i -lt $sheets.Count; $i++) {
  $k = $i + 1
  [IO.File]::WriteAllText("$tmp\xl\worksheets\sheet$k.xml", $sheets[$i].xml, $enc)
  $ct += "<Override PartName=`"/xl/worksheets/sheet$k.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
  $wb += "<sheet name=`"$($sheets[$i].name)`" sheetId=`"$k`" r:id=`"rId$k`"/>"
  $wbr += "<Relationship Id=`"rId$k`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$k.xml`"/>"
}
$ct += '</Types>'; $wb += '</sheets></workbook>'
$wbr += "<Relationship Id=`"rId99`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
$styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF4"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment vertical="top" wrapText="1"/></xf><xf applyAlignment="1"><alignment vertical="top"/></xf></cellXfs></styleSheet>'
[IO.File]::WriteAllText("$tmp\[Content_Types].xml", $ct, $enc)
[IO.File]::WriteAllText("$tmp\_rels\.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', $enc)
[IO.File]::WriteAllText("$tmp\xl\workbook.xml", $wb, $enc)
[IO.File]::WriteAllText("$tmp\xl\_rels\workbook.xml.rels", $wbr, $enc)
[IO.File]::WriteAllText("$tmp\xl\styles.xml", $styles, $enc)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$xlsx = "$outDir\peace-pilgrim-itinerary.xlsx"; if (Test-Path $xlsx) { Remove-Item $xlsx }
$zip = [IO.Compression.ZipFile]::Open($xlsx, 'Create')
Get-ChildItem $tmp -Recurse -File | ForEach-Object { $rel = $_.FullName.Substring($tmp.Length + 1) -replace '\\', '/'; [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) }
$zip.Dispose(); Remove-Item $tmp -Recurse -Force
"sources: $($src.Count)  events: $($events.Count)  ->  $xlsx"
