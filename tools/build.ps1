param([string]$PythonExe = 'python')
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Set-Location -LiteralPath $projectRoot
$env:PYINSTALLER_CONFIG_DIR = Join-Path $projectRoot '.build\pyinstaller-cache'
$iconPath = Join-Path $projectRoot 'assets\app.ico'
$assetData = (Join-Path $projectRoot 'assets') + ';assets'
$sampleData = (Join-Path $projectRoot '示例文件') + ';示例文件'
& $PythonExe -m PyInstaller --noconfirm --clean --windowed --onedir --name '模板文档批量生成' --icon $iconPath --add-data $assetData --add-data $sampleData --paths tools --hidden-import 'pythoncom' --hidden-import 'pywintypes' --hidden-import 'win32timezone' --hidden-import 'win32com.client' --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module pygame --exclude-module IPython --exclude-module pytest --exclude-module win32ui --distpath '.build\dist' --workpath '.build\pyinstaller' --specpath '.build' 'src\app.py'
if ($LASTEXITCODE -ne 0) { throw '打包失败' }
$release = Join-Path $projectRoot '桌面版'
New-Item -ItemType Directory -Path $release -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot '.build\dist\模板文档批量生成\模板文档批量生成.exe') -Destination $release -Force
$releaseInternal = [IO.Path]::GetFullPath((Join-Path $release '_internal'))
if ($releaseInternal -ne [IO.Path]::GetFullPath((Join-Path $projectRoot '桌面版\_internal'))) { throw '运行依赖清理路径越界' }
if (Test-Path -LiteralPath $releaseInternal) { Remove-Item -LiteralPath $releaseInternal -Recurse -Force }
Copy-Item -LiteralPath (Join-Path $projectRoot '.build\dist\模板文档批量生成\_internal') -Destination $release -Recurse -Force
Copy-Item -LiteralPath (Join-Path $projectRoot '示例文件') -Destination $release -Recurse -Force
Copy-Item -LiteralPath (Join-Path $projectRoot '使用说明.md') -Destination $release -Force
Copy-Item -LiteralPath (Join-Path $projectRoot '快速开始.md') -Destination $release -Force
& $PythonExe 'tools\export_licenses.py'
if ($LASTEXITCODE -ne 0) { throw '许可文件收集失败' }
Write-Output "已生成桌面版：$release"
