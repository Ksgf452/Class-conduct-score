$ErrorActionPreference = "Stop"
$src = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $src

$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    Write-Host "找不到 .NET Framework 4.x 的 csc.exe。" -ForegroundColor Red
    Write-Host "请安装 .NET Framework 4.8 后重试。" -ForegroundColor Red
    exit 1
}

$out = Join-Path (Split-Path -Parent $src) "班级操行分管理.exe"
$files = @(
    "Program.cs", "MainForm.cs", "BatchAddForm.cs", "Dialogs.cs", "Theme.cs", "AppConfig.cs",
    "Model.cs", "ChineseNumber.cs", "CxLog.cs", "XlsxCommon.cs",
    "XlsxReader.cs", "XlsxTemplateWriter.cs", "XlsxCreator.cs",
    "SaveService.cs", "RosterReader.cs"
)

$compression = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\System.IO.Compression.dll"
if (-not (Test-Path $compression)) {
    $compression = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\System.IO.Compression.dll"
}

$cscArgs = @("/nologo", "/langversion:5", "/target:winexe", "/out:$out", "/win32icon:app.ico", "/win32manifest:app.manifest")
$cscArgs += $files
$cscArgs += @("/r:System.Windows.Forms.dll", "/r:System.Drawing.dll", "/r:$compression", "/r:System.Xml.dll")

& $csc $cscArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "编译失败。请确认程序没有被占用。" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "编译成功：" -ForegroundColor Green
Write-Host "  $out" -ForegroundColor Green


