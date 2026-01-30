$utf8 = [Text.Encoding]::UTF8
$sjis = [Text.Encoding]::GetEncoding(932)
$targets = @(
  (Join-Path (Get-Location) 'index.html')
) + (Get-ChildItem .\games -Recurse -Filter *.html -ErrorAction SilentlyContinue | Select-Object -Expand FullName)

foreach($p in $targets){
  if(!(Test-Path $p)){ continue }

  $bytes = [IO.File]::ReadAllBytes($p)

  # UTF-8として読んだ版 / SJIS(CP932)として読んだ版を両方作る
  $u = $utf8.GetString($bytes)
  $s = $sjis.GetString($bytes)

  # 置換文字(�)の少ない方を採用（より“正しく読めてる”可能性が高い）
  $ur = ($u.ToCharArray() | Where-Object { $_ -eq [char]0xFFFD }).Count
  $sr = ($s.ToCharArray() | Where-Object { $_ -eq [char]0xFFFD }).Count
  $text = if($ur -le $sr){ $u } else { $s }

  # 念のためバックアップ
  Copy-Item $p ($p + '.bak') -Force

  # meta charset を utf-8 に統一（無ければhead直後に挿入）
  if($text -match '(?i)<meta[^>]*charset'){
    $text = [regex]::Replace($text,'(?i)<meta[^>]*charset\s*=\s*[''"]?[^''"">\s]+[''"]?[^>]*>','<meta charset="utf-8">')
  } elseif($text -match '(?i)<head[^>]*>'){
    $text = [regex]::Replace($text,'(?i)<head[^>]*>','$0'+"`r`n"+'  <meta charset="utf-8">',1)
  } else {
    $text = '<meta charset="utf-8">'+"`r`n"+$text
  }

  # UTF-8(BOMなし)で保存
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [IO.File]::WriteAllText($p,$text,$utf8NoBom)

  Write-Host ("Fixed: " + $p)
}
