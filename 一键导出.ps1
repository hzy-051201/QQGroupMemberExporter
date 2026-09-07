$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$napcatDir = "$env:USERPROFILE\NapCatShell"
if (-not (Test-Path "$napcatDir\bootmain\QQ.exe")) { $napcatDir = "$root\NapCatShell" }
$qqExe = Join-Path $napcatDir 'bootmain\QQ.exe'
$configDir = Join-Path $napcatDir 'config'
$qrPng = Join-Path $napcatDir 'cache\qrcode.png'
$runLog = Join-Path $napcatDir 'run.log'
$pyScript = Join-Path $root '导出群成员.py'
$wsHost = '127.0.0.1'
$wsPort = 3001

function Test-WsPort {
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $r = $c.BeginConnect($wsHost, $wsPort, $null, $null)
        if ($r.AsyncWaitHandle.WaitOne(1500)) { $c.EndConnect($r); return $true }
        return $false
    } catch { return $false } finally { $c.Close() }
}

function Stop-NapCatQQ {
    Get-CimInstance Win32_Process -Filter "Name='QQ.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.ExecutablePath -like '*NapCatShell*' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 3
}

# 1. 查找已登录的QQ号(通过onebot11配置文件)
$onebotCfg = Get-ChildItem $configDir -Filter 'onebot11_*.json' -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $onebotCfg) {
    Write-Host '未检测到登录信息,正在启动 NapCat 登录...' -ForegroundColor Yellow
    Stop-NapCatQQ
    Start-Process -FilePath $qqExe -ArgumentList '--enable-logging' -WorkingDirectory (Split-Path $qqExe) -RedirectStandardOutput $runLog -RedirectStandardError ($runLog + '.err') -WindowStyle Hidden
    Write-Host '请用手机QQ扫描下方二维码完成登录:' -ForegroundColor Yellow
    Write-Host ('二维码图片: ' + $qrPng) -ForegroundColor Green
    if (Test-Path $qrPng) { Invoke-Item $qrPng }
    Write-Host '等待扫码登录(最多5分钟)...' -ForegroundColor Yellow
    $deadline = (Get-Date).AddMinutes(5)
    while ((Get-Date) -lt $deadline) {
        $onebotCfg = Get-ChildItem $configDir -Filter 'onebot11_*.json' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($onebotCfg) { break }
        Start-Sleep -Seconds 3
    }
    if (-not $onebotCfg) {
        Write-Host '5分钟内未检测到登录成功。请确认已扫码后重新运行。' -ForegroundColor Red
        exit 1
    }
}

$uin = $onebotCfg.BaseName -replace '^onebot11_', ''
Write-Host ('检测到已登录QQ号: ' + $uin) -ForegroundColor Green

# 2. 写入 WebSocket 服务配置
$wsCfg = [ordered]@{
    network = [ordered]@{
        httpServers = @()
        httpSseServers = @()
        httpClients = @()
        websocketServers = @(
            [ordered]@{
                name = 'export'
                enable = $true
                host = $wsHost
                port = $wsPort
                messagePostFormat = 'array'
                reportSelfMessage = $false
                token = ''
                enableForcePushEvent = $true
                debug = $false
                heartInterval = 30000
            }
        )
        websocketClients = @()
        plugins = @()
    }
    musicSignUrl = ''
    enableLocalFile2Url = $false
    parseMultMsg = $false
    imageDownloadProxy = ''
    timeout = [ordered]@{
        baseTimeout = 10000
        uploadSpeedKBps = 256
        downloadSpeedKBps = 256
        maxTimeout = 1800000
    }
}
$json = $wsCfg | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($onebotCfg.FullName, $json, (New-Object System.Text.UTF8Encoding $false))
Write-Host ('已开启本地接口 ws://{0}:{1}' -f $wsHost, $wsPort) -ForegroundColor Green

# 3. 重启 NapCat (免扫码快速登录)
Write-Host '正在重启 NapCat(免扫码自动登录)...'
Stop-NapCatQQ
Start-Process -FilePath $qqExe -ArgumentList '--enable-logging','-q',$uin -WorkingDirectory (Split-Path $qqExe) -RedirectStandardOutput $runLog -RedirectStandardError ($runLog + '.err') -WindowStyle Hidden

# 4. 等待接口就绪
Write-Host '等待 NapCat 连接就绪(最多2分钟)...'
$deadline = (Get-Date).AddSeconds(120)
$ready = $false
while ((Get-Date) -lt $deadline) {
    if (Test-WsPort) { $ready = $true; break }
    Start-Sleep -Seconds 2
}
if (-not $ready) {
    Write-Host '本地接口未就绪,请查看日志: ' -ForegroundColor Red
    Write-Host ('  ' + $runLog) -ForegroundColor Yellow
    Get-Content $runLog -Tail 15 -ErrorAction SilentlyContinue
    exit 1
}
Write-Host '连接成功,开始导出群成员...' -ForegroundColor Green

# 5. 执行导出脚本(交互选择群)
python $pyScript
$exitCode = $LASTEXITCODE

# 6. 询问是否关闭 NapCat
$answer = Read-Host '导出完成。是否关闭 NapCat? (Y=关闭 / N=保持运行)'
if ($answer -match '^[Yy]') {
    Write-Host '正在关闭 NapCat...'
    Stop-NapCatQQ
}
exit $exitCode
