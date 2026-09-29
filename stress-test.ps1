$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$compiler = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$stressDirectory = Join-Path $projectRoot 'artifacts\stress-run'
$stressExecutable = Join-Path $stressDirectory 'OfflineDocxCompare.StressTests.exe'

New-Item -ItemType Directory -Force -Path $stressDirectory | Out-Null

$arguments = @(
    '/nologo',
    '/target:exe',
    '/platform:anycpu',
    '/optimize+',
    '/langversion:5',
    '/utf8output',
    ('/out:' + $stressExecutable),
    '/reference:System.dll',
    '/reference:System.Core.dll',
    '/reference:System.Xml.dll',
    '/reference:System.Xml.Linq.dll',
    '/reference:System.IO.Compression.dll',
    '/reference:System.IO.Compression.FileSystem.dll',
    '/reference:System.Security.dll',
    (Join-Path $projectRoot 'src\Models.cs'),
    (Join-Path $projectRoot 'src\DocumentComparer.cs'),
    (Join-Path $projectRoot 'tests\StressTestRunner.cs')
)

& $compiler @arguments
if ($LASTEXITCODE -ne 0) {
    throw "压力测试程序编译失败，退出代码：$LASTEXITCODE"
}

& $stressExecutable $stressDirectory
if ($LASTEXITCODE -ne 0) {
    throw "压力测试失败，退出代码：$LASTEXITCODE"
}
