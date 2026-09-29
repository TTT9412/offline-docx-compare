$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$testDirectory = Join-Path $projectRoot 'artifacts\test-run'
$testExecutable = Join-Path $testDirectory 'OfflineDocxCompare.Tests.exe'

New-Item -ItemType Directory -Force -Path $testDirectory | Out-Null

$arguments = @(
    '/nologo',
    '/target:exe',
    '/platform:anycpu',
    '/optimize+',
    '/langversion:5',
    '/utf8output',
    ('/out:' + $testExecutable),
    '/reference:System.dll',
    '/reference:System.Core.dll',
    '/reference:System.Xml.dll',
    '/reference:System.Xml.Linq.dll',
    '/reference:System.IO.Compression.dll',
    '/reference:System.IO.Compression.FileSystem.dll',
    (Join-Path $projectRoot 'src\Models.cs'),
    (Join-Path $projectRoot 'src\DocumentComparer.cs'),
    (Join-Path $projectRoot 'tests\TestRunner.cs')
)

& $compiler @arguments
if ($LASTEXITCODE -ne 0) {
    throw "测试程序编译失败，退出代码：$LASTEXITCODE"
}

& $testExecutable $testDirectory
if ($LASTEXITCODE -ne 0) {
    throw "测试失败，退出代码：$LASTEXITCODE"
}
