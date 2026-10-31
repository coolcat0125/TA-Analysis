$p = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -match 'media_fill' }
if ($p) { exit 0 } else { exit 1 }
