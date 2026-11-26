# Turns an event's city / state / country into one geocodable place.
# Rows that name several places (a route like "A - B - C" or "A / B") are mapped to the FIRST place;
# rows naming only a state or region are mapped to that state (level = state).
$script:StateNames = @{ AL='Alabama';AK='Alaska';AZ='Arizona';AR='Arkansas';CA='California';CO='Colorado';CT='Connecticut';DE='Delaware';DC='District of Columbia';FL='Florida';GA='Georgia';HI='Hawaii';ID='Idaho';IL='Illinois';IN='Indiana';IA='Iowa';KS='Kansas';KY='Kentucky';LA='Louisiana';ME='Maine';MD='Maryland';MA='Massachusetts';MI='Michigan';MN='Minnesota';MS='Mississippi';MO='Missouri';MT='Montana';NE='Nebraska';NV='Nevada';NH='New Hampshire';NJ='New Jersey';NM='New Mexico';NY='New York';NC='North Carolina';ND='North Dakota';OH='Ohio';OK='Oklahoma';OR='Oregon';PA='Pennsylvania';RI='Rhode Island';SC='South Carolina';SD='South Dakota';TN='Tennessee';TX='Texas';UT='Utah';VT='Vermont';VA='Virginia';WA='Washington';WV='West Virginia';WI='Wisconsin';WY='Wyoming';AB='Alberta';BC='British Columbia';MB='Manitoba';NB='New Brunswick';NL='Newfoundland and Labrador';NS='Nova Scotia';ON='Ontario';PE='Prince Edward Island';QC='Quebec';SK='Saskatchewan';CDMX='Mexico City' }

# Places that need a hand-made answer: name -> "query|level" ("" = do not map)
$script:PlaceOverrides = @{
  'Appalachian Trail (Maine to Georgia)' = '|none'
  'Camps Farthest Out camp' = 'Arkansas, USA|state'
  'Iowa (statewide)' = 'Iowa, USA|state'
  'Hawaii (island tour)' = 'Hawaii, USA|state'
  'Michigan' = 'Michigan, USA|state'
  'Minnesota' = 'Minnesota, USA|state'
  'North Carolina' = 'North Carolina, USA|state'
  'Florida' = 'Florida, USA|state'
  'California' = 'California, USA|state'
  'Southern California' = 'Los Angeles, California, USA|region'
  'New England' = 'Massachusetts, USA|region'
  'Maryland and Delaware' = 'Maryland, USA|region'
  'Alaska and Hawaii' = 'Hawaii, USA|region'
  'Texas and Louisiana' = 'Texas, USA|region'
  'South Carolina (Columbia)' = 'Columbia, South Carolina, USA|city'
  'Upper Suncoast (Hudson / Pasco County)' = 'Hudson, Florida, USA|city'
  'US 89 to Phoenix' = 'Phoenix, Arizona, USA|city'
  'Washington state line' = 'Washington State, USA|state'
  'Yankton and other South Dakota towns' = 'Yankton, South Dakota, USA|city'
  'Philadelphia area / Bucks-Montgomery Co.' = 'Horsham, Pennsylvania, USA|city'
  'Goerke''s Corners (Brookfield)' = 'Brookfield, Wisconsin, USA|city'
  'Stamps (Lafayette County)' = 'Stamps, Arkansas, USA|city'
  'Spokane to Helena' = 'Spokane, Washington, USA|city (first of several)'
  'Detroit - Ambassador Bridge' = 'Ambassador Bridge, Detroit, Michigan, USA|landmark'
  'San Francisco (Golden Gate Bridge)' = 'Golden Gate Bridge, San Francisco, USA|landmark'
  'Washington (White House)' = 'White House, Washington, DC, USA|landmark'
  'Mount Greylock' = 'Mount Greylock, Massachusetts, USA|landmark'
  'Kootenay Lake' = 'Kootenay Lake, British Columbia, Canada|landmark'
  'Knox (Indiana 23 north of town)' = 'Knox, Indiana, USA|city'
  'LaManche to St. John''s' = 'St. John''s, Newfoundland and Labrador, Canada|city'
  'Tulsa to Oklahoma City (Route 66)' = 'Tulsa, Oklahoma, USA|city'
  'Durham (outskirts)' = 'Durham, North Carolina, USA|city'
  'Cecil County (Elkton?)' = 'Elkton, Maryland, USA|city'
}

function Get-PlaceKey($city, $state, $country) {
  $city = "$city".Trim(); $state = "$state".Trim().TrimEnd('?'); $country = "$country".Trim()
  $st = ($state -split '/')[0]
  $stName = if ($script:StateNames.ContainsKey($st)) { $script:StateNames[$st] } else { $st }
  if (-not $country) { $country = 'USA' }
  if ($script:PlaceOverrides.ContainsKey($city)) {
    $q, $lvl = $script:PlaceOverrides[$city] -split '\|'
    if (-not $q) { return $null }
    return @{ key = "$city|$state"; query = $q; level = $lvl }
  }
  if (-not $city) {
    if (-not $stName) { return $null }
    $sq = if ($stName -eq 'Washington') { 'Washington State' } else { $stName }; return @{ key = "|$state|$country"; query = "$sq, $country"; level = 'state' }
  }
  $first = ($city -replace '\s*\(.*?\)', '') -split ' / | - | to ' | Select-Object -First 1
  $first = $first.Trim()
  if ($first -match ' area$') { $first = $first -replace ' area$', '' }
  $level = if ($city -match ' / | - | to ') { 'city (first of several)' } else { 'city' }
  $q = (@($first, $stName, $country) | Where-Object { $_ }) -join ', '
  return @{ key = "$first|$st|$country"; query = $q; level = $level }
}
