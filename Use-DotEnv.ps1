function Use-DotEnv {
  param([string]$Path = ".\.env")
  if (-not (Test-Path $Path)) { Write-Error "not found: $Path"; return }
  $lines = Get-Content -Path $Path -Encoding UTF8
  foreach ($line in $lines) {
    if ($line -match '^\s*#' -or -not $line.Trim()) { continue }
    if ($line -match '^\s*([A-Za-z0-9_]+)\s*=\s*"?(.+?)"?\s*$') {
      $k = $matches[1]
      $v = $matches[2].Trim().Trim('"')
      Set-Item -Path "Env:$k" -Value $v
    }
  }
}
