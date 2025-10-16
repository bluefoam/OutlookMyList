# OutlookAddIn3 搜索辅助脚本
# 用于替代内置正则表达式搜索工具

param(
    [Parameter(Mandatory=$true)]
    [string]$Pattern,
    
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    
    [switch]$ShowLineNumbers = $true,
    [switch]$IgnoreCase = $true,
    [switch]$ShowContext = $false,
    [int]$ContextLines = 3
)

# 获取rg.exe路径
$rgPath = "C:\Users\nxa08267\AppData\Local\Programs\Trae\resources\app\node_modules\@vscode\ripgrep\bin\rg.exe"

if (-not (Test-Path $rgPath)) {
    Write-Error "rg.exe not found at $rgPath"
    exit 1
}

# 构建rg命令参数
$rgArgs = @()
if ($ShowLineNumbers) { $rgArgs += "-n" }
if ($IgnoreCase) { $rgArgs += "-i" }
if ($ShowContext) { 
    $rgArgs += "-C"
    $rgArgs += $ContextLines.ToString()
}
$rgArgs += $Pattern
$rgArgs += $FilePath

# 执行搜索
& $rgPath @rgArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "No matches found for pattern: $Pattern"
}