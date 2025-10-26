# Execution: powershell -ExecutionPolicy Bypass -File .\open-user-dictionary.ps1
# 執行方式：powershell -ExecutionPolicy Bypass -File .\open-user-dictionary.ps1

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# 設定 PowerShell 輸出編碼為 UTF-8
chcp 65001 > $null
# 設定命令提示字元(CMD)的編碼為 UTF-8，隱藏輸出

Add-Type -AssemblyName System.Windows.Forms
# 載入 System.Windows.Forms，用於模擬按鍵輸入
# Load System.Windows.Forms assembly to simulate keyboard input

# 關閉現有設定視窗
# Close any existing Settings window
Get-Process -Name SystemSettings -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 100
# 等待 100 毫秒，確保視窗已關閉
# Wait 100 ms to ensure the window is closed

# 開啟語言與地區設定
# Open Language & Region settings
Start-Process "ms-settings:regionlanguage"
Start-Sleep -Milliseconds 1500
# 等待 1.5 秒讓設定頁面加載完成
# Wait 1.5 seconds for the settings page to load

# 按鍵間隔 & 加載頁面時間 (越低越快，電腦性能高者可進一步調低)
# Key interval & page load wait time (lower is faster; adjust for faster PCs)
$interval = 10
$pageWait = 1200

# 第一組按鍵操作：切換到第一個選項並進入
# First key sequence: navigate to first option and press Enter
for ($i=1; $i -le 3; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds $interval }
# 按 3 次 Tab，移動焦點
# Press Tab 3 times to move focus
for ($i=1; $i -le 2; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{ENTER}"); Start-Sleep -Milliseconds $interval }
# 按 2 次 Enter，打開選項
# Press Enter 2 times to open the option
Start-Sleep -Milliseconds $pageWait
# 等待頁面加載
# Wait for page to load

# 第二組按鍵操作：切換到第二個選項並進入
# Second key sequence: navigate to second option and press Enter
for ($i=1; $i -le 3; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds $interval }
for ($i=1; $i -le 2; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{ENTER}"); Start-Sleep -Milliseconds $interval }
Start-Sleep -Milliseconds $pageWait

# 第三組按鍵操作：切換到第三個選項並進入
# Third key sequence: navigate to third option and press Enter
for ($i=1; $i -le 2; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds $interval }
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Milliseconds $pageWait

# 第四組按鍵操作：切換到使用者造詞工具按鈕並啟用
# Fourth key sequence: navigate to "User Dictionary Tool" button and activate
for ($i=1; $i -le 7; $i++) { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); Start-Sleep -Milliseconds $interval }
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
