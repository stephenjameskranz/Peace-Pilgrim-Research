param($in, $out)
Add-Type -AssemblyName System.Web
$html = [IO.File]::ReadAllText($in)
$ms = [regex]::Matches($html, '<a[^>]+href="(/s/[^"]+)"[^>]*>(.*?)</a>', 'Singleline')
$rows = foreach ($m in $ms) {
  $t = [regex]::Replace($m.Groups[2].Value, '<[^>]+>', '')
  $t = [System.Web.HttpUtility]::HtmlDecode($t) -replace '\s+', ' '
  [pscustomobject]@{ title = $t.Trim(); url = 'https://www.peacepilgrim.org' + $m.Groups[1].Value }
}
$rows | ConvertTo-Csv -NoTypeInformation | Out-File -Encoding utf8 $out
"$($rows.Count) rows"
