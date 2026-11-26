param($in, $out)
$raw = Import-Csv $in
$placeholder = 'PPclip_1953Jan12_FullertonCA-2.pdf'

# 1. merge fragments (index splits some entries over several links)
$items = New-Object System.Collections.ArrayList
foreach ($r in $raw) {
  $t = $r.title.Trim() -replace '^-\s*', ''
  $file = ($r.url -split '/')[-1]
  $isPh = ($file -eq $placeholder) -and ($t -notmatch 'Fullerton')
  if ($file -eq 'PeacePilgrimBook-69ct.pdf') { continue }
  $isStart = $t -match '^(19\d\d|Unknown)'
  if ($isStart -or $items.Count -eq 0) {
    [void]$items.Add([ordered]@{ title = $t; file = $(if ($isPh) { '' } else { $file }); ph = $isPh })
  } else {
    $cur = $items[$items.Count - 1]
    if ($t -and $t -ne '.') { $cur.title = ($cur.title.TrimEnd() + ' ' + $t).Trim() }
    if (-not $cur.file -and -not $isPh) { $cur.file = $file }
    if ($isPh) { $cur.ph = $true }
  }
}
# drop duplicate entries pointing to same title+file
$seen = @{}
$items = $items | Where-Object { $k = $_.title + '|' + $_.file; if ($seen[$k]) { $false } else { $seen[$k] = 1; $true } }

$months = @{ jan=1; feb=2; geb=2; mar=3; ap=4; apr=4; may=5; jun=6; june=6; jul=7; july=7; aug=8; ag=8; sep=9; sept=9; oct=10; nov=11; de=12; dec=12; december=12; august=8; september=9 }
function Get-Mon($s) { if (-not $s) { return $null }; $k = $s.ToLower().TrimEnd('.', ','); foreach ($c in @($k, $k.Substring(0, [Math]::Min(4, $k.Length)), $k.Substring(0, [Math]::Min(3, $k.Length)), $k.Substring(0, [Math]::Min(2, $k.Length)))) { if ($months.ContainsKey($c)) { return $months[$c] } }; $null }

# 2. manual corrections / context, keyed by file name
$fix = @{
 'PPclip_1952Dec_WPEN-Radio-Phili-Excerpts.pdf' = @{ city='Philadelphia'; st='PA'; type='Media appearance (radio)'; note='Radio interview, pre-pilgrimage (she walked the Appalachian Trail in 1952).' }
 'PPclip_1952Dec6_WPEN-PhiliRadioExcerpts2.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Peacemaker'; note='Newsletter item; no location in index title. File name suggests more WPEN Philadelphia radio excerpts.' }
 'PPclip_1953Jan31_PeacemakerNewsltr.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Peacemaker' }
 'PPclip_1953Ap13_Peacemaker21.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='Peacemaker Newsletter' }
 'PPclip_1953Nov26_PeacemakerNewsletter.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='Peacemaker Newsletter'; note='Reports her approaching Washington DC.' }
 'PPclip_1954Geb1_PeaceMakerNewsletter.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Peacemaker' }
 'PPclip_1954Dec6_PeaceMakerNewsletter.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Peacemaker'; note='Announces planned second pilgrimage.' }
 'PPclip_1955Dec5_PeaceMakerNewsletter.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Peacemaker'; note='Reports her in the South (no specific place in title).' }
 'PPclip_1953Sept_TheBlazeNwsletter.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='The Blaze Newsletter' }
 'PPclip_1950s_BritishNews.pdf' = @{ city=''; st=''; country=''; type='Wire/foreign coverage'; pub='British News'; note='Foreign coverage; does not place her.' }
 'PPclip_1950s_EpiscopalBulletin.pdf' = @{ city=''; st=''; type='Newsletter report'; pub='Episcopal Bulletin' }
 'PPclip_1953_ag.pdf' = @{ city=''; st=''; pub='"Ag" (unidentified)'; note='Index title begins "Ag:" - unclear if a place or publication; may be "Aug". Needs clipping read.' }
 'PPclip_1953_California.pdf' = @{ city=''; st='CA'; note='State only.' }
 'PPclip_1953_Illinois-Heat.pdf' = @{ city=''; st='IL'; note='State only.' }
 'PPclip_1953Aug25_KalamazooMI-w-photo.pdf' = @{ city='Kalamazoo'; st='MI'; note='Index title has no place; place taken from file name (same file as the Kalamazoo Aug 25 entry).' }
 'PPclip_1953Sept2_Michigan' = @{}
 'PPclip_1955Jan14_SacramentoCA-Union-w-photoi.jpg' = @{ note='LIKELY INDEX ERROR: listed under 1953 Jan 14 but links a 1955 file, duplicating the 1955 Jan 14 Sacramento entry. Sacramento is not on her Jan 1953 southern route. Treat as 1955.'; date='1955-01-14'; conf='Doubtful (probable duplicate)' }
 'PPclip_1953-Cumberland.pdf' = @{ city='Cumberland'; st='MD?'; note='State not given. Cumberland MD fits the late-1953 route through PA/MD to Washington DC - unconfirmed.' ; conf='Probable' }
 'PPclip_1953-N-Pennsylvania.pdf' = @{ city=''; st='PA'; place='North PA'; note='Region only ("North PA").' }
 'PPclip_1953Dec3_CecilMD-w-photo.pdf' = @{ city='Cecil County (Elkton?)'; st='MD'; note='"Cecil" probably the Cecil Whig (Elkton) or Cecil County.' }
 'PPclip_1950s_GlobeDemocrat_WashDC.pdf' = @{ city='Washington'; st='DC'; pub='Globe-Democrat'; note='Index dates "1950s"; content (petition to Eisenhower) indicates late 1953. Globe-Democrat was a St. Louis paper - place in title may be dateline, not publisher.' ; date='1953'; conf='Probable' }
 'PPclip_1953_ChicagoSunTimes.pdf' = @{ city='Chicago'; st='IL'; pub='Chicago Sun-Times' }
 'PPclip_1953-AtlanticCityNJ-w-photo.pdf' = @{ city='Atlantic City'; st='NJ'; pub='Atlantic City Press'; note='Visiting her mother.' }
 'PPclip_1953Dec-NewYork.pdf' = @{ city='New York'; st='NY'; note='End of first (1953) pilgrimage - United Nations.' }
 'PPclip_1950s_LeesSummitMO-photo-PINK.pdf' = @{ city="Lee's Summit"; st='MO'; date='1953'; conf='Probable'; note='Index dates "1950s". Fits 1953 route (Kansas City area, Jun-Jul 1953) - unconfirmed.' }
 'PPclip_1953Jan_ElCajonCA.pdf' = @{ city='El Cajon'; st='CA' }
 'PPclip_1955HuronSD.pdf' = @{ city='Huron'; st='SD'; note='Walking Huron to Pierre.' }
 'PPclip_1955_LittleRockAK-Gazette.pdf' = @{ st='AR'; pub='Arkansas Gazette'; note='Index says "AK"; Little Rock is in Arkansas (AR).' }
 'PPclip_1955-WacoAustinTX-w-photo.pdf' = @{ note='File name mentions Waco and Austin.' }
 'PPclip_1955MisoulaMT-w-Photo.pdf' = @{ note='Index lists at end of 1955; ~6,200 miles reported suggests mid-1955 (before Mitchell SD, Jul 2) - check clipping.' }
 'PPclip_1956-JohnsonMI-w-photo.pdf' = @{ note='File name says "MI"; index says MS. "Johnson" unclear - possibly a newspaper name.'; conf='Uncertain place' }
 'PPclip_1956Jan11_JacksonMI-w-photo.pdf' = @{ note='File name says "MI"; Jackson MS fits route (Natchez Jan 8 -> Montgomery Jan 25).' }
 'PPclip_1956Jan23_MontgomeryAL-Times-w-photosi.pdf' = @{ note='Index says Jan 25; file name says Jan 23.' }
 'PPclip_1955Jan17_VallegoCATimesHerald-w-photoi.pdf' = @{ note='Index says Jan 12; file name says Jan 17.' }
 'PPclip_1956June11_OttowaCAN-w-photo.pdf' = @{ city='Ottawa'; st='ON'; country='Canada' }
 'PPclip_1957Jan26-WindsorONT-w-photos.pdf' = @{ city='Windsor'; st='ON'; country='Canada' }
 'PPclip_1957May21VancouverBC-photo.pdf' = @{ city='Vancouver'; st='BC'; country='Canada' }
 'PPclip_1957June7CalgaryAlbertanCAN-w-photo.pdf' = @{ city='Calgary'; st='AB'; country='Canada'; pub='The Albertan' }
 'PPclip_1957June12-Calgary-CAN.pdf' = @{ city='Calgary'; st='AB'; country='Canada' }
 'PPclip_1957June24MoosejawCAN-w-photos.pdf' = @{ city='Moose Jaw'; st='SK'; country='Canada' }
 'PPclip_1957July22CharlottetownCAN-w-photo-d9mx.pdf' = @{ city='Charlottetown'; st='PE'; country='Canada' }
 'PPclip_1957StJohnsCAN.pdf' = @{ city="St. John's"; st='NL'; country='Canada'; note='Could be St. John''s NL or Saint John NB - check clipping.' }
 'PPclip_1957StJohnsCAN-photo.pdf' = @{ city="St. John's"; st='NL'; country='Canada'; note='Could be St. John''s NL or Saint John NB - check clipping.' }
 'PPclip_1957Sept1-Quebec-french-headine.pdf' = @{ city=''; st='QC'; country='Canada'; note='French-language paper; Quebec province (city unclear).' }
 'PPclip_1959_LasVegasNV-2.pdf' = @{ city='Las Vegas'; st='NV' }
 'PPclip_1959-FellowshipMag.pdf' = @{ city=''; st=''; type='Magazine feature'; pub='Fellowship Magazine' }
 'PPclip_1959-FellowshipMag-2.pdf' = @{ city=''; st=''; type='Magazine feature'; pub='Fellowship Magazine' }
 'PPclip_1959-FellowshipMag-3.pdf' = @{ city=''; st=''; type='Magazine feature'; pub='Fellowship Magazine' }
 'PPclip_1960Ap18-GrantsPA-w-photo.pdf' = @{ st='NM'; note='Index and file say "PA", but Grants NM is ~80 mi west of Albuquerque (Mar 26, 1960). Almost certainly Grants, New Mexico.'; conf='Probable' }
 'PPclip_1960Dec3_TylerTX-w-photo.pdf' = @{ note='Listed as Fort Worth Dec 12, but the link points to a Tyler TX file - possible wrong link on index.' }
 'PPclip_1961May18_CearVallyIA-w-photo.pdf' = @{ city='Vinton'; st='IA'; note='Index: "Cedar Valley" - headline says she walked through Vinton IA.' }
 'PPclip_1961May18_CearVallyIA-2.pdf' = @{ city='Vinton'; st='IA'; note='Index: "Cedar Valley" - headline says she walked through Vinton IA.' }
 'PPclip_1961June9_StPPaulMS-w-photo.pdf' = @{ st='MN' }
 'PPclip_1962-LewisvilleARK-w-photo.pdf' = @{ note='Index dates 1962 between Apr 26 (Roanoke VA) and May 5 (Indianapolis) - Arkansas does not fit that sequence; true date likely different. Check clipping.' }
 'PPclip_1967July21_MilwaukeeWI-w-photo.pdf' = @{ city='Milwaukee'; st='WI'; note='Index title has no place; place from file name.' }
 'PPclip_1969-Seattle-Times.pdf' = @{ city='Seattle'; st='WA'; pub='Seattle Times' }
 'PPclip_Unknown_photos-sneakers.pdf' = @{ city=''; st=''; date=''; type='Undated photo'; note='Date and place unknown.' }
 'PPclip_1971Feb27_PineBluffAK-w-photo.pdf' = @{ st='AR'; note='Index says "AK"; Pine Bluff is in Arkansas (AR).' }
 'PPclip_1973EastOregon-w-photo.pdf' = @{ city=''; st='OR'; pub='East Oregonian (Pendleton)?'; note='"East Oregon" is probably the East Oregonian newspaper, Pendleton OR.' }
 'PPclip_1973June18-SalemOR-w-photo.pdf' = @{ note='Same headline as the 1959 June 28 Salem entry - one may be misdated. Check clippings.' }
 'PPclip_1959June28_-SalemOR-w-photo.pdf' = @{ note='Same headline as the 1973 June 18 Salem entry - one may be misdated. Check clippings.' }
 'PPclip_1973Canada-w-photo.pdf' = @{ city=''; st=''; country='Canada'; note='Country only.' }
 'PPclip_1973Dec16-PasadenaCA-w-photo.pdf' = @{ city='Pasadena'; st='CA'; note='Index title has no place; place from file name.' }
 'PPclip_1975-unknown-w-photos.pdf' = @{ city=''; st=''; note='Place unknown.' }
 'PPclip_1975Dec7_FamiyWeeklyMag-w-photo.pdf' = @{ city=''; st=''; type='Magazine feature'; pub='Family Weekly Magazine' }
 'PPclip_1976Jan30-EdinbregTX-w-photo.pdf' = @{ city='Edinburg' }
 'PPclip_1976June2_EvansvilleIA-w-photo.pdf' = @{ st='IN?'; note='Index says "IA"; Evansville IN is far more likely (fits Chattanooga/Knoxville TN -> Springfield IL sequence).'; conf='Probable' }
 'PPclip_1977s-UNKNOWN-photo.pdf' = @{ city=''; st=''; note='Place unknown.' }
 'PPclip_1977Dec7_SuncoastFL-w-photo.pdf' = @{ city=''; st='FL'; note='"Suncoast" = Florida Gulf coast (Sarasota/Tampa area).' }
 'PPclip_1977_HeadShotwithquote.pdf' = @{ city=''; st=''; note='Place unknown.' }
 'PPclip_1979_NorthArizona.pdf' = @{ city=''; st='AZ'; place='North Arizona'; note='Region only. Headline: through all 50 states.' }
 'PPclip_1979_NewMexico.pdf' = @{ city=''; st='NM'; note='State only.' }
 'PPclip_1980Arlington-photo.pdf' = @{ city='Arlington'; st=''; note='State not given (VA or TX?).' ; conf='Uncertain place' }
 'PPclip_1980Aug15_Honolulu-w-photo.pdf' = @{ city='Honolulu'; st='HI' }
 'PPclip_1981Jul7_KnoxTN-w-photo.pdf' = @{ st='IN'; type='Death report'; note='Index says "TN", but she died 7 July 1981 in a car crash near Knox, INDIANA (corroborated by Elkhart, Valparaiso and Hammond IN reports).' }
 'PPclip_1981Jul8_KnoxTX-w-photo.pdf' = @{ st='IN'; type='Death report'; note='Index says "TN", file says "TX"; she died near Knox, Indiana.' }
 'PPclip_1981_KnoxTN.pdf' = @{ st='IN'; type='Death report'; note='Index says "TN"; she died near Knox, Indiana.' }
 'PPclip_1981Jul8_ElkhartIN-w-photo.pdf' = @{ type='Death report'; note='Newspaper location, not hers.' }
 'PPclip_1981Jul9_ValparaisoIN-w-photo.pdf' = @{ type='Death report'; note='Newspaper location, not hers.' }
 'PPclip_1981July_ChattanoogaTN-w-photo.pdf' = @{ type='Death report'; note='Local friends react; newspaper location, not hers.' }
 'PPclip_1981Jul10_HammondIN.pdf' = @{ type='Death report'; note='Newspaper location, not hers.' }
 'PPclip_1992Jan26_EnidOH-WindowOnPast-nfj2.pdf' = @{ st='OK'; note='Index says "OH"; Enid is in Oklahoma. Retrospective "Window on the Past" column - may recall her 1975 Enid visit.' }
 'PPclip_1981Aug15_YanktonSD-3acf.jpg' = @{ city='Yankton'; st='SD'; type='Death report'; note='Place from file name; newspaper location, not hers.' }
}
$retro = 'PPclip_198[4-9]|PPclip_199'

$rows = @(); $i = 0
foreach ($it in $items) {
  $i++
  $t = $it.title; $file = $it.file; $f = if ($file -and $fix.ContainsKey($file)) { $fix[$file] } else { @{} }
  $m = [regex]::Match($t, '^(?<d>(?:19\d\ds?|Unknown)(?:\s+[A-Za-z]+\.?,?)?(?:\s+\d{1,2})?)\s*(?:[\u2013-]\s*)?(?<rest>.*)$')
  $dtxt = $m.Groups['d'].Value.Trim(); $rest = $m.Groups['rest'].Value.Trim()
  $dm = [regex]::Match($dtxt, '^(?<y>19\d\d)(?<s>s?)(?:\s+(?<mo>[A-Za-z]+)\.?,?)?(?:\s+(?<dd>\d{1,2}))?')
  $date = ''; $prec = ''
  if ($dm.Success) {
    $y = $dm.Groups['y'].Value; $mo = Get-Mon $dm.Groups['mo'].Value; $dd = $dm.Groups['dd'].Value
    if ($dm.Groups['s'].Value) { $date = "${y}s"; $prec = 'decade' }
    elseif ($mo -and $dd) { $date = '{0}-{1:D2}-{2:D2}' -f $y, $mo, [int]$dd; $prec = 'day' }
    elseif ($mo) { $date = '{0}-{1:D2}' -f $y, $mo; $prec = 'month' }
    else { $date = $y; $prec = 'year' }
  } else { $prec = 'unknown' }
  # place / headline split
  $place = ''; $head = $rest
  $cm = [regex]::Match($rest, '^(?<p>[^:]{2,45}):\s*(?<h>.*)$')
  if ($cm.Success) { $place = $cm.Groups['p'].Value.Trim(); $head = $cm.Groups['h'].Value.Trim() }
  elseif ($rest -match '^(?<p>[A-Z][A-Za-z\. ]+ [A-Z]{2})$') { $place = $rest; $head = '' }
  $head = ($head -replace '\s+', ' ') -replace '^[\u2013-]\s*', ''; $city = $null
  $city = $place; $st = ''; $country = 'USA'
  $pm = [regex]::Match($place, '^(?<c>.+?),?\s+(?<s>[A-Z]{2})$')
  if ($pm.Success) { $city = $pm.Groups['c'].Value; $st = $pm.Groups['s'].Value }
  elseif ($place -match '^(?<c>.+?),?\s+Canada$') { $city = $Matches['c']; $country = 'Canada' }
  if ($place -eq 'Michigan') { $city = ''; $st = 'MI' }
  if ($place -eq '') { $country = '' }
  $pub = ''
  if ($place -match '^(.+?) ([A-Z]{2}) (Sun Times|Globe Democrat|Press)$') { $city = $Matches[1]; $st = $Matches[2] }
  # file-name date for cross-check
  $fdate = ''
  $fm = [regex]::Match($file, 'PPclip_(?<y>19\d\d)[-_]?(?<mo>[A-Za-z]{2,5})?(?<d>\d{1,2})?')
  if ($fm.Success) { $fmo = Get-Mon $fm.Groups['mo'].Value; $fdate = $fm.Groups['y'].Value; if ($fmo) { $fdate += '-{0:D2}' -f $fmo; if ($fm.Groups['d'].Value) { $fdate += '-{0:D2}' -f [int]$fm.Groups['d'].Value } } }
  $type = 'Newspaper report (local)'
  if ($file -match $retro) { $type = 'Posthumous / retrospective' }
  $conf = if ($prec -eq 'day') { 'Documented' } elseif ($prec -in 'month', 'year') { 'Documented (date approximate)' } else { 'Date uncertain' }
  $note = ''
  if ($f.city -ne $null) { $city = $f.city }
  if ($f.st -ne $null) { $st = $f.st }
  if ($f.country -ne $null) { $country = $f.country }
  if ($f.place) { $place = $f.place }
  if ($f.pub) { $pub = $f.pub }
  if ($f.type) { $type = $f.type }
  if ($f.conf) { $conf = $f.conf }
  if ($f.date -ne $null) { $date = $f.date; if ($date -eq '') { $prec = 'unknown' } elseif ($date.Length -eq 4) { $prec = 'year' } elseif ($date.Length -eq 10) { $prec = 'day' } }
  if ($f.note) { $note = $f.note }
  if (-not $city -and -not $st) { if ($country -eq 'Canada') { $conf = 'Country only' } else { $conf = 'No location'; $country = '' } }
  if ($type -match 'Posthumous|Death report' -and $file -notmatch 'Knox') { $conf = 'Not a location of Peace' }
  if ($file -match '1981.*Knox') { $city = 'Knox'; $conf = 'Documented (place of death)' }
  if ($fdate -and $date -and ($date -notmatch 's$') -and ($fdate.Length -gt 4) -and ($date.Length -gt 4) -and -not $fdate.StartsWith($date) -and -not $date.StartsWith($fdate)) {
    if ($note -notmatch 'file name') { $note = ($note + " Index date $date differs from file-name date $fdate.").Trim() }
  }
  if (-not $file) { $note = ($note + ' NO WORKING LINK on index page (link points to an unrelated placeholder file).').Trim() }
  elseif ($it.ph) { $note = ($note + ' Part of this index entry links to an unrelated placeholder file.').Trim() }
  if ($file -match 'Jan12_FullertonCA-2' -and $t -notmatch 'con') { }
  $rows += [pscustomobject][ordered]@{
    id = '{0:D3}' -f $i
    date = $date
    date_precision = $prec
    date_meaning = 'Publication date (per index) - she was usually there that day or 1-2 days before'
    city = $city.Trim()
    state_province = $st
    country = $country
    place_as_listed = $place
    record_type = $type
    location_confidence = $conf
    headline = ($head -replace '\s*\((with photos?|with photo|photos? & caption|with illustration|with photos)\)\s*$', '').Trim()
    publication = $pub
    has_photo = $(if ($t -match 'photo') { 'yes' } else { '' })
    miles_reported = ''
    route_before = ''
    route_planned = ''
    people_mentioned = ''
    notable_details = ''
    clipping_read = 'no'
    index_title = $t
    source_url = $(if ($file) { "https://www.peacepilgrim.org/s/$file" } else { '' })
    file_date = $fdate
    notes = $note
  }
}
$rows | Export-Csv -NoTypeInformation -Encoding UTF8 $out
"$($rows.Count) rows"
$rows | Where-Object { -not $_.city -and -not $_.state_province } | ForEach-Object { "$($_.id) | $($_.index_title)" }
