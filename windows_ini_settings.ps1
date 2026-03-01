# -----------------------------------------------------------------------------
# Windows 初期設定スクリプト
# 目的：Windowsの初期設定を自動化する。
# https://references.package-inc.com/archives/pd_report_engineer/windows-pc%e5%88%9d%e6%9c%9f%e8%a8%ad%e5%ae%9a%e3%83%a1%e3%83%a2
#
# 注意：このスクリプトはレジストリを直接変更します。
# 管理者権限で実行してください。
# -----------------------------------------------------------------------------
# 実行方法:
#   1. PowerShellを管理者権限で起動
#   2. Set-ExecutionPolicy RemoteSigned -Scope Process -Force
#   3. .\windows_ini_settings.ps1
# -----------------------------------------------------------------------------

# ========================
# ユーザー設定ブロック
# ここの値を変えればカスタマイズできるわよ
# ========================
$Config = @{
    # マウス
    ScrollLines       = "1"          # スクロール行数
    # MouseSpeed      = "12"         # ポインター速度（コメントアウト中）

    # 電源プラン GUID
    UltimateGuid      = "e9a42b02-d5df-448d-aa00-03f14749eb61"  # 究極のパフォーマンス
    HighPerfGuid      = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"  # 高パフォーマンス

    # ログファイルの保存先
    LogPath           = "$env:TEMP\ini_settings_log.txt"

    # スクリーンショット保存先
    ScreenshotPath    = "$env:USERPROFILE\Desktop"

    # God Mode フォルダパス
    GodModePath       = "$env:USERPROFILE\Desktop\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
}

# ========================
# 内部関数
# ========================

# ログ出力（コンソール + ファイル）
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White",
        [switch]$IsError
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] $Message"

    if ($IsError) {
        Write-Host $Message -ForegroundColor Red
        Add-Content -Path $Config.LogPath -Value "[ERROR] $logLine"
    } else {
        Write-Host $Message -ForegroundColor $Color
        Add-Content -Path $Config.LogPath -Value $logLine
    }
}

# レジストリ値を設定（なければ作成、あれば更新）
function Set-RegValue {
    param(
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Type = "DWord"
    )
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        if (-not (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue)) {
            Write-Log "  [新規] $Name" -Color Yellow
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        } else {
            Write-Log "  [更新] $Name"
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force
        }
    } catch {
        Write-Log "  レジストリ設定に失敗したわ: $Path\$Name -> $_" -IsError
    }
}

# ========================
# 事前チェック
# ========================

# 管理者権限チェック
$currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "管理者権限で実行してちょうだい！PowerShellを右クリック→「管理者として実行」よ。" -ForegroundColor Red
    exit 1
}

# Windowsバージョンチェック（Windows 10/11前提）
$winVer = [System.Environment]::OSVersion.Version
if ($winVer.Major -lt 10) {
    Write-Host "このスクリプトはWindows 10/11専用よ。バージョンが古すぎるわ。" -ForegroundColor Red
    exit 1
}

# ログファイル初期化
"" | Out-File -FilePath $Config.LogPath -Encoding UTF8
Write-Log "===== Windows 初期設定スクリプト 開始 =====" -Color Cyan

# 復元ポイント作成
Write-Log "[事前準備] システムの復元ポイントを作成するわ…" -Color Cyan
try {
    Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue
    Checkpoint-Computer -Description "windows_ini_settings.ps1 実行前" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
    Write-Log "  復元ポイントを作成したわ。何かあっても戻れるから安心しなさい。" -Color Green
} catch {
    Write-Log "  復元ポイントの作成に失敗したけど、処理は続けるわ: $_" -Color Yellow
}

Write-Log ""
Write-Log "貴方の命令に従い、Windowsの初期設定を開始するわ！" -Color Magenta
Start-Sleep -Seconds 1

# ========================
# 1. マウス設定
# ========================
Write-Log ""
Write-Log "[1] マウス設定" -Color Cyan

# ポインターの精度（加速）をオフ
Set-RegValue -Path "HKCU:\Control Panel\Mouse" -Name "MouseThreshold1" -Value "0" -Type String
Set-RegValue -Path "HKCU:\Control Panel\Mouse" -Name "MouseThreshold2" -Value "0" -Type String

# スクロール行数
Set-RegValue -Path "HKCU:\Control Panel\Desktop" -Name "WheelScrollLines" -Value $Config.ScrollLines -Type String

# ========================
# 2. タスクバー設定
# ========================
Write-Log ""
Write-Log "[2] タスクバー設定" -Color Cyan

$explorerAdvPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0  # 検索ボックス非表示
Set-RegValue -Path $explorerAdvPath -Name "TaskbarDa"           -Value 0  # ウィジェット非表示
Set-RegValue -Path $explorerAdvPath -Name "ShowTaskViewButton"  -Value 0  # タスクビュー非表示
Set-RegValue -Path $explorerAdvPath -Name "TaskbarMn"           -Value 0  # チャット非表示 (24H2等だと効かない場合あり)
Set-RegValue -Path $explorerAdvPath -Name "ShowCopilotButton"   -Value 0  # Copilot非表示
Set-RegValue -Path $explorerAdvPath -Name "TaskbarAl"           -Value 0  # タスクバーを左寄せ

# ========================
# 3. エクスプローラー設定
# ========================
Write-Log ""
Write-Log "[3] エクスプローラー設定" -Color Cyan

Set-RegValue -Path $explorerAdvPath                                                       -Name "HideFileExt"        -Value 0  # 拡張子を表示
Set-RegValue -Path $explorerAdvPath                                                       -Name "Hidden"             -Value 1  # 隠しファイルを表示
Set-RegValue -Path $explorerAdvPath                                                       -Name "HideDrivesWithNoMedia" -Value 0  # 空ドライブ表示
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Name "FullPath"          -Value 1  # タイトルバーにフルパスを表示

# ========================
# 4. スクリーンショット保存先
# ========================
Write-Log ""
Write-Log "[4] スクリーンショット保存先 → $($Config.ScreenshotPath)" -Color Cyan
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" `
             -Name "{B7BEDE81-DF94-4682-A7D8-57A52620B86F}" `
             -Value $Config.ScreenshotPath `
             -Type ExpandString

# ========================
# 5. 個人用設定（テーマ・アニメーション）
# ========================
Write-Log ""
Write-Log "[5] 個人用設定" -Color Cyan

$themePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
Set-RegValue -Path $themePath -Name "AppsUseLightTheme"    -Value 0  # ダークモード（アプリ）
Set-RegValue -Path $themePath -Name "SystemUsesLightTheme" -Value 0  # ダークモード（システム）
Set-RegValue -Path $themePath -Name "EnableTransparency"   -Value 0  # 透明効果オフ

# 壁紙を黒の単色に
Set-RegValue -Path "HKCU:\Control Panel\Colors"  -Name "Background" -Value "0 0 0" -Type String
Set-RegValue -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper"  -Value ""       -Type String

# アニメーション効果オフ
Set-RegValue -Path "HKCU:\Control Panel\Desktop\WindowMetrics"                             -Name "MinAnimate"       -Value "0" -Type String
Set-RegValue -Path $explorerAdvPath                                                         -Name "TaskbarAnimations" -Value 0
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"     -Name "Animations"       -Value 0

# ========================
# 5.5 スタートメニュー・右クリックメニュー (Win11向け)
# ========================
Write-Log ""
Write-Log "[5.5] スタートメニュー・右クリックメニュー設定" -Color Cyan

# スタートメニューの「おすすめ」を非表示
Set-RegValue -Path $explorerAdvPath -Name "Start_IrisRecommendations" -Value 0

# 右クリックメニュー (コンテキストメニュー) をWindows 10仕様に戻す
$cmdContextMenuPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
if (-not (Test-Path $cmdContextMenuPath)) {
    New-Item -Path $cmdContextMenuPath -Force | Out-Null
}
try {
    Set-ItemProperty -Path $cmdContextMenuPath -Name "(default)" -Value "" -Force
    Write-Log "  [新規/更新] 右クリックメニューの旧仕様化" -Color Yellow
} catch {
    Write-Log "  レジストリ設定に失敗したわ: $cmdContextMenuPath -> $_" -IsError
}

# ========================
# 6. 視覚効果カスタム設定
# ========================
Write-Log ""
Write-Log "[6] 視覚効果カスタム設定（3項目のみ有効化）" -Color Cyan

# UserPreferencesMask: ドラッグ中にウィンドウ内容を表示 + フォント滑らか を有効化
$maskValue = [byte[]](0xA2, 0x12, 0x03, 0x80, 0x10, 0x00, 0x00, 0x00)
try {
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value $maskValue -Type Binary -Force
    Write-Log "  UserPreferencesMask を設定したわ。"
} catch {
    Write-Log "  UserPreferencesMask の設定に失敗したわ: $_" -IsError
}

# アイコンの代わりに縮小版を表示（0=縮小版表示）
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "IconsOnly" -Value 0

# カスタム設定として認識させる
Set-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFxSetting" -Value 2

# ========================
# 7. 電源設定
# ========================
Write-Log ""
Write-Log "[7] 電源設定" -Color Cyan

$powerPlans = powercfg /list
if (-not ($powerPlans -match $Config.UltimateGuid)) {
    Write-Log "  究極のパフォーマンスプランが見つからないから、システムに召喚（復元）するわね。" -Color Yellow
    powercfg -duplicatescheme $Config.UltimateGuid | Out-Null
    $powerPlans = powercfg /list
}

if ($powerPlans -match $Config.UltimateGuid) {
    Write-Log "  究極のパフォーマンスプランを有効にするわよ！" -Color Green
    powercfg /setactive $Config.UltimateGuid
} elseif ($powerPlans -match $Config.HighPerfGuid) {
    Write-Log "  高パフォーマンスプランを有効にするわ。"
    powercfg /setactive $Config.HighPerfGuid
} else {
    Write-Log "  パフォーマンスプランの設定に失敗したみたいね。" -Color Red
}

# 画面タイムアウト（AC・DC両方）を無効化
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0
Write-Log "  画面タイムアウトを無効化したわ。"

# ========================
# 8. プライバシー設定
# ========================
Write-Log ""
Write-Log "[8] プライバシー設定" -Color Cyan

$privacyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy"
Set-RegValue -Path $privacyPath -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0
Set-RegValue -Path $privacyPath -Name "LetAppsUseAdvertisingId"                      -Value 0
Set-RegValue -Path $privacyPath -Name "EnableWebsitesAccessToLanguageList"            -Value 0
Set-RegValue -Path $privacyPath -Name "LetAppsTrackAppUsage"                          -Value 0

# ========================
# 9. クリップボード履歴
# ========================
Write-Log ""
Write-Log "[9] クリップボード履歴を有効化" -Color Cyan
Set-RegValue -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -Value 1

# ========================
# 10. God Mode
# ========================
Write-Log ""
Write-Log "[10] God Mode フォルダをデスクトップに召喚" -Color Cyan
if (-not (Test-Path $Config.GodModePath)) {
    New-Item -Path $Config.GodModePath -ItemType Directory | Out-Null
    Write-Log "  God Mode フォルダを作成したわ。"
} else {
    Write-Log "  God Mode フォルダは既に存在するわ。"
}

# ========================
# 仕上げ：エクスプローラー再起動
# ========================
Write-Log ""
Write-Log "仕上げ！エクスプローラーを再起動するわ。" -Color Yellow
Write-Log "（画面が一瞬ちらつくけど、世界の再構築だから驚かないで）"

rundll32.exe User32.dll, UpdatePerUserSystemParameters 1, True
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Start-Process -FilePath "explorer.exe"

# ========================
# 完了
# ========================
Write-Log ""
Write-Log "===== 全ての設定が完了したわ =====" -Color Green
Write-Log "ログファイルは $($Config.LogPath) に保存してあるわよ。" -Color Cyan
Write-Log "念のため再起動を推奨するわ。感謝しなさいよね！" -Color Magenta
Start-Sleep -Seconds 5