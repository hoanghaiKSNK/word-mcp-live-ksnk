param(
    [string]$PythonExe = 'D:\ChatGPT\mcp\word-mcp-live\.venv\Scripts\python.exe',
    [string]$ConfigPath = 'D:\ChatGPT\CodexHome\config.toml',
    [string]$InstallRoot = 'D:\ChatGPT\mcp\word-mcp-live-ksnk\releases'
)
$ErrorActionPreference = 'Stop'
$bundleRoot = $PSScriptRoot
if (!(Test-Path -LiteralPath $PythonExe -PathType Leaf)) {throw 'Supply an existing Python 3.11+ executable using -PythonExe'}
if (!(Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {throw 'Existing Codex config.toml required'}
$manifest = Get-Content -LiteralPath (Join-Path $bundleRoot 'manifest.json') -Raw | ConvertFrom-Json
if ($manifest.version -ne '1.6.21+ksnk.2') {throw 'Unexpected bundle version'}
foreach ($item in $manifest.files) {
    $file = Join-Path $bundleRoot $item.path
    if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -ne $item.sha256) {throw "Package hash mismatch: $($item.path)"}
}
$releasePath = Join-Path $InstallRoot ('1.6.21-ksnk2-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
if (Test-Path -LiteralPath $releasePath) {throw 'Release directory already exists'}
New-Item -ItemType Directory -Path $releasePath | Out-Null
$venvPath = Join-Path $releasePath '.venv'
& $PythonExe -m venv $venvPath
if ($LASTEXITCODE -ne 0) {throw 'venv creation failed'}
$runtimePython = Join-Path $venvPath 'Scripts\python.exe'
& $runtimePython -m pip install -r (Join-Path $bundleRoot 'requirements-runtime.lock') --disable-pip-version-check
if ($LASTEXITCODE -ne 0) {throw 'Dependency installation failed; existing MCP unchanged'}
& $runtimePython -m pip install --no-deps (Join-Path $bundleRoot 'word_mcp_live-1.6.21+ksnk.2-py3-none-any.whl') --disable-pip-version-check
if ($LASTEXITCODE -ne 0) {throw 'Wheel installation failed; existing MCP unchanged'}
& $runtimePython -m pip check
if ($LASTEXITCODE -ne 0) {throw 'Dependency validation failed; existing MCP unchanged'}
& $runtimePython (Join-Path $bundleRoot 'verify_runtime.py') --runtime $releasePath
if ($LASTEXITCODE -ne 0) {throw 'MCP protocol check failed; existing MCP unchanged'}
$receiptPath = Join-Path $releasePath 'installation-receipt.json'
& $runtimePython (Join-Path $bundleRoot 'update_mcp_config.py') --config $ConfigPath --command (Join-Path $venvPath 'Scripts\word_mcp_server.exe') --cwd $releasePath --backups (Join-Path $releasePath 'private-config-backups') --receipt $receiptPath
if ($LASTEXITCODE -ne 0) {throw 'Configuration validation failed'}
Write-Output "Installed KSNK Word MCP ksnk.2. Receipt: $receiptPath"
Write-Output 'Restart old Codex/MCP sessions to load the coordinated runtime. User Word documents need not be closed.'
