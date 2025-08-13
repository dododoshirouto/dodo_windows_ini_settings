# -----------------------------------------------------------------------------
# Windows 初期設定スクリプト (Webページ参照版)
# 目的：指定されたWebページの内容に基づき、Windowsの初期設定を自動化する。
# https://references.package-inc.com/archives/pd_report_engineer/windows-pc%e5%88%9d%e6%9c%9f%e8%a8%ad%e5%ae%9a%e3%83%a1%e3%83%a2
#
# 注意：このスクリプトはレジストリを直接変更します。
# 実行前にシステムの復元ポイントを作成することを強く推奨します。
# -----------------------------------------------------------------------------
# 先にこれを実行しておく
# Set-ExecutionPolicy RemoteSigned -Scope Process -Force
# -----------------------------------------------------------------------------

# --- 儀式の開始宣言 ---
Write-Host "貴方の命令に従い、古文書の封印を解き放つわ…。" -ForegroundColor Magenta
Write-Host "Windowsの初期設定を開始するから、刮目して見てなさい！"
Start-Sleep -Seconds 2

# --- レジストリ操作の汎用関数を定義 ---
# 存在しないなら創造、存在するなら更新する魔法よ。
function Set-OrNew-RegValue {
    param(
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Type = "DWord"
    )
    if (-not (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue)) {
        Write-Host "...項目'$Name'が存在しないため、新規に創造するわ。" -ForegroundColor Yellow
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force
    } else {
        Write-Host "...項目'$Name'を更新するわ。"
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force
    }
}

# --- 1-3. マウス＆スクロール設定 ---
Write-Host "[マウス] ポインター速度とスクロール行数を最適化するわ…"
# ポインターの速度 (既定値の少し速め: 12)
# Set-OrNew-RegValue -Path "HKCU:\Control Panel\Mouse" -Name "MouseSpeed" -Value "12" -Type String
# ポインターの精度を高める（加速）をオフ
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Mouse" -Name "MouseThreshold1" -Value "0" -Type String
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Mouse" -Name "MouseThreshold2" -Value "0" -Type String
# 一度にスクロールする行数 (1行)
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Desktop" -Name "WheelScrollLines" -Value "1" -Type String

# --- 4. タスクバーの整理 ---
Write-Host "[タスクバー] 不要なアイコンを闇に葬るわ…"
# 検索ボックスを非表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0
# ウィジェットを非表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Value 0
# タスクビューを非表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0
# チャットを非表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Value 0
# タスクバーを左寄せにする
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAl" -Value 0

# --- 5. エクスプローラーの最適化 ---
Write-Host "[エクスプローラー] 設定を見やすく、使いやすく変更するわ…"
# ファイル名拡張子を表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0
# 隠しファイルを表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 1
# 空のドライブを非表示にするチェックを外す
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideDrivesWithNoMedia" -Value 0
# タイトルバーに完全なパスを表示
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Name "FullPath" -Value 1

# --- 6. スクリーンショットの保存先変更 ---
Write-Host "[スクリーンショット] 保存先をデスクトップに変更するわ…"
$screenshotPath = "$($env:USERPROFILE)\Desktop"
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{B7BEDE81-DF94-4682-A7D8-57A52620B86F}" -Value $screenshotPath -Type ExpandString

# --- 7. 個人用設定の変更 ---
Write-Host "[個人用設定] ダークモードを基調とした漆黒のテーマを適用するわ…"
# ダークモード設定
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0
# 透明効果をオフ
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "EnableTransparency" -Value 0
# 背景を黒の単色に
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Colors" -Name "Background" -Value "0 0 0"
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value ""
# アニメーション効果をオフ
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFxSetting" -Value 3
Set-OrNew-RegValue -Path "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "MinAnimate" -Value "0"
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAnimations" -Value 0
Set-OrNew-RegValue -Path "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Animations" -Value 0

# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
# ★★★        視覚効果カスタム設定（指定項目のみ有効化）        ★★★
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
Write-Host "[視覚効果] 貴方の願い通り、3つの効果だけを残し、他は全て闇に葬るわ…"

# --- ステップ1: UserPreferencesMaskで主要な2項目を設定 ---
# 「ドラッグ中にウィンドウの内容を表示」「スクリーンフォントを滑らかに」を有効化し、
# 他のアニメーションを無効化する『運命の石版』の値を定義するわ。
$path = "HKCU:\Control Panel\Desktop"
$name = "UserPreferencesMask"

# 0x90 (パフォーマンス) をベースに、Bit1(ドラッグ内容)とBit4(フォント)を立てた値が 0xA2 よ。
# これが貴方だけのカスタム値。
$customMask = [byte[]](0xA2, 0x12, 0x03, 0x80, 0x10, 0x00, 0x00, 0x00)

# 『運命の石版』を貴方専用の値で強制的に上書きする！
Set-ItemProperty -Path $path -Name $name -Value $customMask -Type Binary -Force

# --- ステップ2: 「アイコンの代わりに縮小版を表示」を個別に設定 ---
# この設定は別の場所に封印されているから、個別に呪文を唱えるわ。
$thumbPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$thumbName = "IconsOnly"
# 0で縮小版表示、1でアイコンのみ。だから0を設定するの。
Set-OrNew-RegValue -Path $thumbPath -Name $thumbName -Value 0 -Type DWord

# --- ステップ3: 念のためのダメ押し ---
# システムに「カスタム設定にしたぞ」と明確に知らしめるための呪文よ。
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFxSetting" -Value 2

# --- 9. 電源設定の最適化 ---
Write-Host "[電源] パフォーマンスを最大化し、画面が消えないようにするわ…"

# パフォーマンス優先の電源プランのGUIDを探す
$ultimateGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61" # 究極のパフォーマンス
$highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c" # 高パフォーマンス
$powerPlans = powercfg /list

# 「究極のパフォーマンス」があればそれを、なければ「高パフォーマンス」を有効にする
if ($powerPlans -match $ultimateGuid) {
    Write-Host "...究極のパフォーマンスプランを有効にするわ！"
    powercfg /setactive $ultimateGuid
}
elseif ($powerPlans -match $highPerfGuid) {
    Write-Host "...高パフォーマンスプランを有効にするわ。"
    powercfg /setactive $highPerfGuid
}
else {
    Write-Host "...警告：パフォーマンスプランが見つからないけど、設定は続けるわ。" -ForegroundColor Yellow
}

# 現在有効なプランに対して、画面タイムアウトを無効（「なし」= 0分）にする
Write-Host "...画面が勝手に消えないよう、封印を施すわ。"
# AC電源（接続時）
powercfg /change monitor-timeout-ac 0
# DC電源（バッテリー時）
powercfg /change monitor-timeout-dc 0

# --- 8. IME設定 ---
# Write-Host "[IME] スペースキーで英数モードを切り替えられるようにするわ…"
# Microsoft IMEのスペースキーでの英数変換を有効化
# Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\IME\15.0\IMEJP\MSIME" -Name "Enable Eisuu Jienkan" -Value 1

# --- 10. プライバシー設定 ---
Write-Host "[プライバシー] 不要な追跡設定を無効にするわ…"
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "LetAppsUseAdvertisingId" -Value 0
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "EnableWebsitesAccessToLanguageList" -Value 0
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "LetAppsTrackAppUsage" -Value 0

# --- 11. クリップボードの履歴 ---
Write-Host "[クリップボード] 便利な履歴機能を有効にするわ…"
Set-OrNew-RegValue -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -Value 1

# --- 12. 神のモード (God Mode) ---
Write-Host "[おまけ] 全てを司る禁断のフォルダー『God Mode』をデスクトップに召喚するわ…"
$godModePath = "$($env:USERPROFILE)\Desktop\GodMode.{ED7BA470-8E54-465E-825C-99712043E01C}"
if (-not (Test-Path $godModePath)) {
    New-Item -Path $godModePath -ItemType Directory | Out-Null
}

# --- 仕上げ ---
Write-Host ""
Write-Host "仕上げよ！ エクスプローラーの再起動を命じる！" -ForegroundColor Yellow
Write-Host "（画面が一瞬ちらつくけど、世界の再構築だから驚かないで）"
rundll32.exe User32.dll, UpdatePerUserSystemParameters 1, True
Stop-Process -Name explorer -Force
Start-Process -FilePath "explorer.exe"

Write-Host ""
Write-Host "全ての儀式は完了したわ。貴方のPCは私の力で生まれ変わったのよ。" -ForegroundColor Green
Write-Host "感謝しなさいよね！"
Start-Sleep -Seconds 10