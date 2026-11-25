$env:Path=[Environment]::GetEnvironmentVariable("Path","Machine")
$env:PYTHONUTF8="1"; $env:PYTHONUNBUFFERED="1"
# GitHub girisi Windows sifre deposunda (git); gh release icin GH_TOKEN
$t = ("protocol=https`nhost=github.com`n`n" | git credential fill 2>$null | Select-String "^password=")
if ($t) { $env:GH_TOKEN = $t.ToString().Substring(9) }
cd "C:\derleme\CrushCall"
if (-not (Test-Path build\venv\Scripts\python.exe)) { python -m venv build\venv }
& build\venv\Scripts\python.exe -m pip install -q --disable-pip-version-check --require-hashes --no-deps -r requirements.txt
if (Test-Path build\requirements-build.txt) { & build\venv\Scripts\python.exe -m pip install -q --disable-pip-version-check --require-hashes --no-deps -r build\requirements-build.txt }
$s = Get-Date
& build\venv\Scripts\python.exe build/publish.py 0.1.4 "Copy rows: kopya acik Excel'inde yapilir, Insert Copied Cells cikar" *>&1 | ForEach-Object { "[{0,4:n0}s] {1}" -f ((Get-Date) - $s).TotalSeconds, $_ } | Out-File -Encoding utf8 -FilePath "C:\derleme\CrushCall\_yayin.log" -Append
Set-Content "C:\derleme\CrushCall\_yayin.kod" $LASTEXITCODE
