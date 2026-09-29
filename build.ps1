param(
    [switch]$RunTests
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$frameworkDirectory = Split-Path -Parent $compiler
$wpfDirectory = Join-Path $frameworkDirectory 'WPF'
$outputDirectory = Join-Path $projectRoot 'dist'
$outputFile = Join-Path $outputDirectory '离线文档对比.exe'
$manifest = Join-Path $projectRoot 'src\app.manifest'
$icon = Join-Path $projectRoot 'assets\app-icon.ico'

if (-not (Test-Path -LiteralPath $compiler)) {
    throw '未找到 Windows 自带的 C# 编译器。'
}

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$references = @(
    (Join-Path $frameworkDirectory 'System.dll'),
    (Join-Path $frameworkDirectory 'System.Core.dll'),
    (Join-Path $frameworkDirectory 'System.Xml.dll'),
    (Join-Path $frameworkDirectory 'System.Xml.Linq.dll'),
    (Join-Path $frameworkDirectory 'System.IO.Compression.dll'),
    (Join-Path $frameworkDirectory 'System.IO.Compression.FileSystem.dll'),
    (Join-Path $wpfDirectory 'WindowsBase.dll'),
    (Join-Path $wpfDirectory 'PresentationCore.dll'),
    (Join-Path $wpfDirectory 'PresentationFramework.dll'),
    (Join-Path $frameworkDirectory 'System.Xaml.dll')
)

$sourceFiles = Get-ChildItem -LiteralPath (Join-Path $projectRoot 'src') -Filter '*.cs' |
    Sort-Object Name |
    ForEach-Object FullName

$arguments = @(
    '/nologo',
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    '/langversion:5',
    '/utf8output',
    ('/out:' + $outputFile),
    ('/win32manifest:' + $manifest),
    ('/win32icon:' + $icon)
)

foreach ($reference in $references) {
    $arguments += '/reference:' + $reference
}
$arguments += $sourceFiles

& $compiler @arguments
if ($LASTEXITCODE -ne 0) {
    throw "编译失败，退出代码：$LASTEXITCODE"
}

Copy-Item -LiteralPath (Join-Path $projectRoot 'src\OfflineDocxCompare.exe.config') -Destination ($outputFile + '.config') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination (Join-Path $outputDirectory '使用说明.md') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'docs\调研与设计.md') -Destination (Join-Path $outputDirectory '调研与设计.md') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'docs\压力测试报告.md') -Destination (Join-Path $outputDirectory '压力测试报告.md') -Force

Write-Host "已生成：$outputFile" -ForegroundColor Green

if ($RunTests) {
    & (Join-Path $projectRoot 'test.ps1')
}
