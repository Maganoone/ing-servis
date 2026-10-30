# Установка моста «1С ↔ Claude». Запускайте через setup.bat — он попросит права администратора.
# Continue: иначе PowerShell 5.1 обрывает скрипт на любом выводе программ в stderr; ошибки проверяем сами.
$ErrorActionPreference = 'Continue'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

function Step($text) { Write-Host "`n=== $text ===" -ForegroundColor Cyan }
function Done($text) { Write-Host "  OK: $text" -ForegroundColor Green }
function Fail($text) {
    Write-Host "`nОШИБКА: $text" -ForegroundColor Red
    Read-Host "`nНажмите Enter, чтобы закрыть окно"
    exit 1
}

$identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail 'Нужны права администратора. Запустите setup.bat двойным щелчком и нажмите «Да».'
}

# ------------------------------------------------------------------ 1. Платформа 1С
Step '1. Ищу платформу 1С и COM-коннектор'
$platforms = @()
foreach ($pair in @(@($env:ProgramFiles, 64), @(${env:ProgramFiles(x86)}, 32))) {
    $root = $pair[0]; $bits = $pair[1]
    if (-not $root -or -not (Test-Path "$root\1cv8")) { continue }
    foreach ($dir in Get-ChildItem "$root\1cv8" -Directory) {
        $dll = Join-Path $dir.FullName 'bin\comcntr.dll'
        if (Test-Path $dll) {
            $platforms += [pscustomobject]@{ Version = $dir.Name; Dll = $dll; Bits = $bits }
        }
    }
}
if (-not $platforms) { Fail 'Не нашёл платформу 1С (comcntr.dll) в Program Files.' }

# Предпочитаем ту же версию, что открывает базу (8.3.13.1513), и 64-битную, если она есть.
$platform = $platforms | Sort-Object `
    @{ Expression = { $_.Version -eq '8.3.13.1513' }; Descending = $true },
    @{ Expression = { try { [version]$_.Version } catch { [version]'0.0' } }; Descending = $true },
    @{ Expression = { $_.Bits }; Descending = $true } | Select-Object -First 1
Done "1С $($platform.Version), $($platform.Bits)-бит"

$reg = Start-Process regsvr32.exe -ArgumentList "/s `"$($platform.Dll)`"" -Wait -PassThru
if ($reg.ExitCode -ne 0) { Fail "Не удалось зарегистрировать $($platform.Dll) (код $($reg.ExitCode))." }
Done 'COM-коннектор V83.COMConnector зарегистрирован'

# ------------------------------------------------------------------ 2. Python той же разрядности
Step "2. Ищу Python $($platform.Bits)-бит"
if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    Fail 'Python не установлен. Скачайте Python 3.12 с python.org (см. README) и запустите setup.bat снова.'
}
$pyTag = $null
foreach ($ver in '3.13', '3.12', '3.11', '3.10') {
    $tag = "-$ver-$($platform.Bits)"
    & py $tag -c 'import sys' 2>$null
    if ($LASTEXITCODE -eq 0) { $pyTag = $tag; break }
}
if (-not $pyTag) {
    $kind = if ($platform.Bits -eq 32) { 'Windows installer (32-bit)' } else { 'Windows installer (64-bit)' }
    Fail "Нужен Python $($platform.Bits)-бит, как у 1С. Скачайте на python.org Python 3.12 → «$kind» и запустите setup.bat снова."
}
Done "Python $pyTag"

# ------------------------------------------------------------------ 3. Окружение и библиотеки
Step '3. Ставлю библиотеки (минуту-две)'
$python = Join-Path $here '.venv\Scripts\python.exe'
if (-not (Test-Path $python)) {
    & py $pyTag -m venv (Join-Path $here '.venv')
    if ($LASTEXITCODE -ne 0) { Fail 'Не удалось создать окружение Python.' }
}
& $python -m pip install --disable-pip-version-check -q -r (Join-Path $here 'requirements.txt')
if ($LASTEXITCODE -ne 0) { Fail 'Не удалось установить библиотеки. Проверьте интернет.' }
Done 'Библиотеки установлены'

# ------------------------------------------------------------------ 4. Параметры базы
Step '4. Подключение к базе'
$defaultBase = 'C:\Users\ИНГ-СЕРВИС\YandexDisk-Ingservice24\Bases'
$base = Read-Host "Папка базы 1С [Enter — $defaultBase]"
if (-not $base) { $base = $defaultBase }
if (-not (Test-Path (Join-Path $base '1Cv8.1CD'))) {
    Write-Host '  Внимание: в этой папке нет файла 1Cv8.1CD — проверьте путь.' -ForegroundColor Yellow
}
$defaultUser = 'Магомед Леймоев'
$user = Read-Host "Пользователь 1С [Enter — $defaultUser]"
if (-not $user) { $user = $defaultUser }
$secure = Read-Host 'Пароль этого пользователя (если пароля нет — просто Enter)' -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))

$env:ONEC_BASE = $base
$env:ONEC_USER = $user
$env:ONEC_PASSWORD = $password

Step '5. Пробное подключение к 1С'
& $python (Join-Path $here 'onec_mcp.py') --test
if ($LASTEXITCODE -ne 0) {
    Fail 'Подключиться не удалось. Частые причины: неверный пароль, база открыта монопольно (конфигуратор), нет свободной лицензии. Подробности — в README.'
}

# ------------------------------------------------------------------ 6. Конфиг Claude Desktop
Step '6. Подключаю мост к Claude Desktop'
$configs = @()
$configs += Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
$store = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Packages') -Directory -Filter 'Claude_*' -ErrorAction SilentlyContinue
foreach ($pkg in $store) { $configs += Join-Path $pkg.FullName 'LocalCache\Roaming\Claude\claude_desktop_config.json' }

$server = [pscustomobject]@{
    command = $python
    args    = @((Join-Path $here 'onec_mcp.py'))
    env     = [pscustomobject]@{ ONEC_BASE = $base; ONEC_USER = $user; ONEC_PASSWORD = $password; PYTHONIOENCODING = 'utf-8' }
}

$ErrorActionPreference = 'Stop'  # дальше любая ошибка — повод остановиться, а не затереть конфиг
foreach ($path in $configs) { try {
    $dir = Split-Path -Parent $path
    if (-not (Test-Path $dir)) {
        if ($path -ne $configs[0]) { continue }  # папку Store-версии не создаём, если её нет
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $config = [pscustomobject]@{}
    if (Test-Path $path) {
        Copy-Item $path "$path.bak" -Force
        $raw = [IO.File]::ReadAllText($path)
        if ($raw.Trim()) { $config = $raw | ConvertFrom-Json }
    }
    if (-not $config.PSObject.Properties['mcpServers']) {
        $config | Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{})
    }
    $config.mcpServers | Add-Member -NotePropertyName '1c' -NotePropertyValue $server -Force
    [IO.File]::WriteAllText($path, ($config | ConvertTo-Json -Depth 20), (New-Object Text.UTF8Encoding $false))
    Done "Записано в $path"
} catch { Fail "Не удалось записать настройки Claude Desktop в ${path}: $_" } }

Write-Host "`nГотово!" -ForegroundColor Green
Write-Host 'Полностью закройте Claude Desktop (значок у часов → Quit) и откройте снова.'
Write-Host 'Потом спросите Claude, например: «Сколько регистраторов DS-7608 на складе?»'
Read-Host "`nНажмите Enter, чтобы закрыть окно"
