# -----------------------------------------------------------------------------
# winget 一括インストールスクリプト
# 使い方:
#   1. winget-list.txt にインストールしたいアプリを1行1アプリで記述
#   2. PowerShellを管理者権限で起動
#   3. .\winget_install.ps1
#      または .\winget_install.ps1 -ListFile "別のファイル名.txt"
# -----------------------------------------------------------------------------

param(
    [string]$ListFile = "winget-list.txt"
)

# ========================
# 設定
# ========================
$LogPath = "$env:TEMP\winget_install_log.txt"

# ========================
# 関数
# ========================
function Write-Log {
    param([string]$Message, [string]$Color = "White", [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"
    if ($IsError) {
        Write-Host $Message -ForegroundColor Red
        Add-Content -Path $LogPath -Value "[ERROR] $logLine"
    } else {
        Write-Host $Message -ForegroundColor $Color
        Add-Content -Path $LogPath -Value $logLine
    }
}

# ========================
# 事前チェック
# ========================

# 管理者権限チェック
$currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "管理者権限で実行してちょうだい！" -ForegroundColor Red
    exit 1
}

# winget の存在チェック
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "winget が見つからないわ。App Installer をインストールして。" -ForegroundColor Red
    exit 1
}

# リストファイルの存在チェック
if (-not (Test-Path $ListFile)) {
    Write-Host "リストファイルが見つからないわ: $ListFile" -ForegroundColor Red
    exit 1
}

# ========================
# メイン処理
# ========================

# ログ初期化
"" | Out-File -FilePath $LogPath -Encoding UTF8
Write-Log "===== winget 一括インストール 開始 =====" -Color Cyan
Write-Log "リストファイル: $ListFile" -Color Cyan

# ファイルを読み込んで処理
$lines = Get-Content -Path $ListFile -Encoding UTF8

$successList = [System.Collections.Generic.List[string]]::new()
$errorList   = [System.Collections.Generic.List[string]]::new()
$skipList    = [System.Collections.Generic.List[string]]::new()
$lineNumber  = 0

foreach ($line in $lines) {
    $lineNumber++

    # 空行・コメント行をスキップ
    $trimmed = $line.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) {
        continue
    }

    Write-Log ""
    Write-Log "[$lineNumber] インストール: $trimmed" -Color Cyan

    # ID形式か名前形式かを判定（ドットが含まれていればIDとみなす）
    if ($trimmed -match "^[\w][\w\-]*\.[\w][\w\-\.]*$") {
        # ID形式 (例: Git.Git, Microsoft.VisualStudioCode)
        Write-Log "  -> ID形式で検索するわ"
        $result = winget install --id $trimmed --silent --accept-package-agreements --accept-source-agreements 2>&1
    } else {
        # 名前形式
        Write-Log "  -> 名前形式で検索するわ"
        $result = winget install --name $trimmed --silent --accept-package-agreements --accept-source-agreements 2>&1
    }

    $output = $result | Out-String

    # 結果判定
    if ($LASTEXITCODE -eq 0) {
        Write-Log "  [成功] $trimmed" -Color Green
        $successList.Add($trimmed)
    } elseif ($output -match "already installed") {
        Write-Log "  [スキップ] 既にインストール済みよ: $trimmed" -Color Yellow
        $skipList.Add($trimmed)
    } else {
        Write-Log "  [失敗] $trimmed (終了コード: $LASTEXITCODE)" -IsError
        # エラー詳細をログに記録
        Add-Content -Path $LogPath -Value "  詳細: $output"
        $errorList.Add($trimmed)
    }
}

# ========================
# 結果サマリー
# ========================
Write-Log ""
Write-Log "===== インストール結果 =====" -Color Cyan
Write-Log "成功:   $($successList.Count) 件" -Color Green
Write-Log "スキップ（導入済み）: $($skipList.Count) 件" -Color Yellow
Write-Log "失敗:   $($errorList.Count) 件" -Color $(if ($errorList.Count -gt 0) { "Red" } else { "White" })

if ($errorList.Count -gt 0) {
    Write-Log ""
    Write-Log "--- 失敗したアプリ一覧 ---" -Color Red
    foreach ($app in $errorList) {
        Write-Log "  x $app" -Color Red
    }
    Write-Log ""
    Write-Log "手動でインストールするか、winget search でIDを確認してみて。" -Color Yellow
}

Write-Log ""
Write-Log "ログファイル: $LogPath" -Color Cyan
Write-Log "===== 完了 =====" -Color Cyan
