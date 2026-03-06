# UPS_POD_Downloader.ps1
# UPS Proof of Delivery automatizált letöltő
# Futtatás: Jobb klikk -> Run with PowerShell

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Proxy beallitas
$env:HTTP_PROXY = "http://cloudproxy.dhl.com:10123"
$env:HTTPS_PROXY = "http://cloudproxy.dhl.com:10123"
$env:NO_PROXY = "127.0.0.1,localhost"

$form = New-Object System.Windows.Forms.Form
$form.Text = "UPS PoD Letöltő"
$form.Size = New-Object System.Drawing.Size(650, 780)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

# --- Fejléc ---
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(10, 10)
$headerLabel.Size = New-Object System.Drawing.Size(600, 30)
$headerLabel.Text = "UPS Proof of Delivery automatizált letöltő"
$headerLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = "DarkBlue"
$form.Controls.Add($headerLabel)

# --- Útmutató ---
$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(10, 50)
$infoPanel.Size = New-Object System.Drawing.Size(600, 90)
$infoPanel.BorderStyle = "FixedSingle"
$infoPanel.BackColor = "LightYellow"
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(10, 5)
$infoLabel.Size = New-Object System.Drawing.Size(580, 80)
$infoLabel.Text = "Használat:`n1. Kattints a 'PoD Chrome indítása' gombra - megnyílik egy Chrome ablak`n2. Jelentkezz be UPS fiókodba ebben a Chrome-ban! És utána nagyon fontos, hogy ne zárd be!`n3. Válaszd ki az Excel fájlt és a letöltési mappát, majd kattints az Indítás gombra!"
$infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$infoPanel.Controls.Add($infoLabel)
$form.Controls.Add($infoPanel)

# --- 1. lépés: POD Chrome ---
$step1Label = New-Object System.Windows.Forms.Label
$step1Label.Location = New-Object System.Drawing.Point(10, 152)
$step1Label.Size = New-Object System.Drawing.Size(600, 20)
$step1Label.Text = "1. lépés: PoD Chrome indítása"
$step1Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$step1Label.ForeColor = "DarkBlue"
$form.Controls.Add($step1Label)

$launchChromeButton = New-Object System.Windows.Forms.Button
$launchChromeButton.Location = New-Object System.Drawing.Point(10, 175)
$launchChromeButton.Size = New-Object System.Drawing.Size(200, 32)
$launchChromeButton.Text = "PoD Chrome indítása"
$launchChromeButton.BackColor = "SteelBlue"
$launchChromeButton.ForeColor = "White"
$launchChromeButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$launchChromeButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1] Kéz kurzor
$launchChromeButton.Add_Click({
    $portCheck = Test-NetConnection -ComputerName 127.0.0.1 -Port 9222 -WarningAction SilentlyContinue -InformationLevel Quiet
    if ($portCheck) {
        Write-Log "POD Chrome mar fut a 9222-es porton."
        $chromeStatus.Text = "✓ POD Chrome fut"
        $chromeStatus.ForeColor = "DarkGreen"
        return
    }

    $chromePaths = @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $chromePath = $null
    foreach ($p in $chromePaths) { if (Test-Path $p) { $chromePath = $p; break } }

    if (-not $chromePath) {
        [System.Windows.Forms.MessageBox]::Show("Chrome nem található!", "Chrome nem található", "OK", "Warning")
        return
    }

    $profileDir = "$env:USERPROFILE\SeleniumProfile"
    $upsUrl = $urlBox.Text.Trim()  # [3] UPS URL-re navigálás induláskor
    Start-Process $chromePath -ArgumentList "--remote-debugging-port=9222 --user-data-dir=`"$profileDir`" `"$upsUrl`""
    Write-Log "POD Chrome elindítva -> $upsUrl"
    Write-Log ">>> Jelentkezz be az UPS-be, majd kattints az Indítás gombra!"
    $chromeStatus.Text = "✓ Chrome elindult - jelentkezz be az UPS-be!"
    $chromeStatus.ForeColor = "DarkGreen"
})
$form.Controls.Add($launchChromeButton)

$chromeStatus = New-Object System.Windows.Forms.Label
$chromeStatus.Location = New-Object System.Drawing.Point(220, 183)
$chromeStatus.Size = New-Object System.Drawing.Size(400, 18)
$chromeStatus.Text = "Chrome még nem indult el"
$chromeStatus.Font = New-Object System.Drawing.Font("Arial", 9)
$chromeStatus.ForeColor = "Gray"
$form.Controls.Add($chromeStatus)

# --- Elválasztó ---
$sep = New-Object System.Windows.Forms.Label
$sep.Location = New-Object System.Drawing.Point(10, 215)
$sep.Size = New-Object System.Drawing.Size(610, 2)
$sep.BorderStyle = "Fixed3D"
$form.Controls.Add($sep)

# --- 2. lépés ---
$step2Label = New-Object System.Windows.Forms.Label
$step2Label.Location = New-Object System.Drawing.Point(10, 222)
$step2Label.Size = New-Object System.Drawing.Size(600, 20)
$step2Label.Text = "2. lépés: Fájlok és beállítások"
$step2Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$step2Label.ForeColor = "DarkBlue"
$form.Controls.Add($step2Label)

# UPS URL
$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Location = New-Object System.Drawing.Point(10, 250)
$urlLabel.Size = New-Object System.Drawing.Size(120, 25)
$urlLabel.Text = "UPS URL:"
$urlLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($urlLabel)

$urlBox = New-Object System.Windows.Forms.TextBox
$urlBox.Location = New-Object System.Drawing.Point(140, 250)
$urlBox.Size = New-Object System.Drawing.Size(470, 25)
$urlBox.Text = "https://www.ups.com/track?loc=en_US&requester=ST/"
$urlBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($urlBox)

# Excel fájl
$excelLabel = New-Object System.Windows.Forms.Label
$excelLabel.Location = New-Object System.Drawing.Point(10, 285)
$excelLabel.Size = New-Object System.Drawing.Size(120, 25)
$excelLabel.Text = "Excel fájl:"
$excelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($excelLabel)

$excelBox = New-Object System.Windows.Forms.TextBox
$excelBox.Location = New-Object System.Drawing.Point(140, 285)
$excelBox.Size = New-Object System.Drawing.Size(370, 25)
$excelBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($excelBox)

$excelButton = New-Object System.Windows.Forms.Button
$excelButton.Location = New-Object System.Drawing.Point(520, 285)
$excelButton.Size = New-Object System.Drawing.Size(90, 25)
$excelButton.Text = "Tallózás"
$excelButton.Font = New-Object System.Drawing.Font("Arial", 9)
$excelButton.BackColor = "LightGray"
$excelButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1]
$excelButton.Add_Click({
    $fb = New-Object System.Windows.Forms.OpenFileDialog
    $fb.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
    if ($fb.ShowDialog() -eq "OK") { $excelBox.Text = $fb.FileName }
})
$form.Controls.Add($excelButton)

# Letöltési mappa
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Location = New-Object System.Drawing.Point(10, 320)
$folderLabel.Size = New-Object System.Drawing.Size(120, 25)
$folderLabel.Text = "Letöltési mappa:"
$folderLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($folderLabel)

$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Location = New-Object System.Drawing.Point(140, 320)
$folderBox.Size = New-Object System.Drawing.Size(370, 25)
$folderBox.Font = New-Object System.Drawing.Font("Arial", 10)
$folderBox.Text = [Environment]::GetFolderPath("Desktop")
$form.Controls.Add($folderBox)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Location = New-Object System.Drawing.Point(520, 320)
$folderButton.Size = New-Object System.Drawing.Size(90, 25)
$folderButton.Text = "Tallózás"
$folderButton.Font = New-Object System.Drawing.Font("Arial", 9)
$folderButton.BackColor = "LightGray"
$folderButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1]
$folderButton.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.ShowNewFolderButton = $true
    if ($fb.ShowDialog() -eq "OK") { $folderBox.Text = $fb.SelectedPath }
})
$form.Controls.Add($folderButton)

# Excel oszlopok info
$checkLabel = New-Object System.Windows.Forms.Label
$checkLabel.Location = New-Object System.Drawing.Point(10, 358)
$checkLabel.Size = New-Object System.Drawing.Size(600, 20)
$checkLabel.Text = "Az Excel-ben szükséges oszlopok:"
$checkLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($checkLabel)

$checkList = New-Object System.Windows.Forms.ListBox
$checkList.Location = New-Object System.Drawing.Point(10, 378)
$checkList.Size = New-Object System.Drawing.Size(600, 40)
$checkList.Font = New-Object System.Drawing.Font("Arial", 9)
$checkList.Items.AddRange(@(
    "✓ 'Tracking Number' - a nyomkövetési szám",
    "✓ 'összefűz' - a letöltött fájl végső neve (ű-vel!)"
))
$checkList.Enabled = $false
$checkList.BackColor = "White"
$form.Controls.Add($checkList)

# --- Napló ---
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(10, 428)
$logLabel.Size = New-Object System.Drawing.Size(600, 20)
$logLabel.Text = "Folyamat napló:"
$logLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10, 448)
$logBox.Size = New-Object System.Drawing.Size(610, 200)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = "Black"
$logBox.ForeColor = "Lime"
$form.Controls.Add($logBox)

# --- Progress bar ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 658)
$progressBar.Size = New-Object System.Drawing.Size(310, 25)
$form.Controls.Add($progressBar)

# --- Gombok ---
$script:stopRequested = $false
$script:pythonProcess = $null

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(330, 658)
$stopButton.Size = New-Object System.Drawing.Size(80, 25)
$stopButton.Text = "STOP"
$stopButton.BackColor = "Orange"
$stopButton.ForeColor = "White"
$stopButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$stopButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1]
$stopButton.Enabled = $false
$stopButton.Add_Click({
    $script:stopRequested = $true
    Write-Log "LEALLAS kérve..."
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        Set-Content -Path (Join-Path $env:TEMP "ups_pod_stop.txt") -Value "stop" -Force
        Start-Sleep -Seconds 3
        if (!$script:pythonProcess.HasExited) { $script:pythonProcess.Kill(); Write-Log "Folyamat leállítva." }
    }
})
$form.Controls.Add($stopButton)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(420, 658)
$startButton.Size = New-Object System.Drawing.Size(100, 25)
$startButton.Text = "Indítás"  # [2] Rövidebb felirat
$startButton.BackColor = "ForestGreen"
$startButton.ForeColor = "White"
$startButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$startButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1]
$form.Controls.Add($startButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(530, 658)
$exitButton.Size = New-Object System.Drawing.Size(90, 25)
$exitButton.Text = "Kilépés"
$exitButton.BackColor = "DarkRed"
$exitButton.ForeColor = "White"
$exitButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$exitButton.Cursor = [System.Windows.Forms.Cursors]::Hand  # [1]
$exitButton.Add_Click({
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        Set-Content -Path (Join-Path $env:TEMP "ups_pod_stop.txt") -Value "stop" -Force
        Start-Sleep -Seconds 2
        if (!$script:pythonProcess.HasExited) { $script:pythonProcess.Kill() }
    }
    $form.Close()
})
$form.Controls.Add($exitButton)

function Write-Log {
    param($Message)
    $logBox.AppendText($Message + "`r`n")
    $logBox.ScrollToCaret()
    $logBox.Refresh()
}

# =====================================================
# LETÖLTÉS INDÍTÁSA
# =====================================================
$startButton.Add_Click({
    $startButton.Enabled = $false
    $stopButton.Enabled = $true
    $script:stopRequested = $false

    $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
    if (Test-Path $stopFilePath) { Remove-Item $stopFilePath -Force }

    $url            = $urlBox.Text.Trim()
    $excelPath      = $excelBox.Text.Trim()
    $downloadFolder = $folderBox.Text.Trim()

    if (-not $url) {
        [System.Windows.Forms.MessageBox]::Show("Add meg az UPS URL-t!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }
    if (-not $excelPath -or -not (Test-Path $excelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes Excel fájlt kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }
    if (-not $downloadFolder -or -not (Test-Path $downloadFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes letöltési mappát kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }

    $portCheck = Test-NetConnection -ComputerName 127.0.0.1 -Port 9222 -WarningAction SilentlyContinue -InformationLevel Quiet
    if (-not $portCheck) {
        [System.Windows.Forms.MessageBox]::Show(
            "A POD Chrome nem fut!`n`nElőször kattints a 'POD Chrome indítása' gombra és jelentkezz be az UPS-be.",
            "POD Chrome nem fut", "OK", "Warning")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }

    Write-Log "==========================================="
    Write-Log "UPS POD Letöltő indítása"
    Write-Log "==========================================="
    Write-Log "Dátum: $(Get-Date)"
    Write-Log "Excel: $excelPath"
    Write-Log "Letöltési mappa: $downloadFolder"
    Write-Log ""

    $pythonScript = @'
import sys
import pandas as pd
import time
import os
import random
import base64
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, WebDriverException
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

STOP_FILE = os.path.join(os.environ['TEMP'], 'ups_pod_stop.txt')
GREEN_COLOR = '92D050'

def should_stop():
    return os.path.exists(STOP_FILE)

def log_message(msg):
    print(f"LOG: {msg}"); sys.stdout.flush()
def log_error(msg, details=""):
    print(f"LOG: [HIBA] {msg}")
    if details: print(f"LOG:   {details}")
    sys.stdout.flush()
def log_success(msg):
    print(f"LOG: [OK] {msg}"); sys.stdout.flush()
def log_step(step, msg):
    print(f"LOG:   [{step}] {msg}"); sys.stdout.flush()
def update_progress(current, total):
    print(f"PROGRESS: {current},{total}"); sys.stdout.flush()

def human_click(driver, element):
    actions = ActionChains(driver)
    actions.move_to_element(element)
    time.sleep(random.uniform(0.3, 0.8))
    actions.click()
    actions.perform()

def close_policy_popup(driver):
    try:
        if not driver.find_elements(By.CSS_SELECTOR, "#ups-updateProfile-popup-container"):
            return
        btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".ups-notNowButton")))
        human_click(driver, btn)
        log_success("Policy popup bezarva")
        time.sleep(1)
    except:
        pass

def close_chat_if_present(driver):
    try:
        if not driver.find_elements(By.CSS_SELECTOR, "div.WACBotContainer"):
            return
        btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACHeader__CloseAndRestartButton")))
        human_click(driver, btn)
        time.sleep(1)
        try:
            yes = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACConfirmModal__YesButton")))
            human_click(driver, yes)
        except:
            pass
        log_success("Chat bezarva")
        time.sleep(1)
    except:
        pass

def is_row_processed(ws, row_idx):
    for col in range(1, 6):
        cell = ws.cell(row=row_idx, column=col)
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
            if cell.fill.fgColor.rgb[-6:] == GREEN_COLOR:
                return True
    return False

def save_pod_pdf(driver, download_folder, new_name, tracking_window):
    try:
        windows_before = set(driver.window_handles)

        log_step("PDF", "Print this page gomb keresese...")
        print_btn = None
        for by, sel, desc in [
            (By.ID, "stApp_POD_btnPrint", "ID"),
            (By.LINK_TEXT, "Print this page", "Link szoveg"),
            (By.PARTIAL_LINK_TEXT, "Print", "Reszleges")
        ]:
            try:
                print_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, sel)))
                log_step("PDF", f"Print gomb talalva: {desc}")
                break
            except:
                continue

        if not print_btn:
            log_error("Print this page gomb nem talalhato")
            return False

        human_click(driver, print_btn)
        log_success("Print this page megnyomva - varunk az uj ablakra...")

        try:
            WebDriverWait(driver, 10).until(
                lambda d: len(d.window_handles) > len(windows_before)
            )
            new_wins = set(driver.window_handles) - windows_before
            if new_wins:
                print_window = new_wins.pop()
                driver.switch_to.window(print_window)
                log_success(f"Print preview ablakra valtva, URL: {driver.current_url}")
                time.sleep(3)
            else:
                log_step("PDF", "Nem nyilt uj ablak, maradunk")
        except TimeoutException:
            log_step("PDF", "Uj ablak nem nyilt 10mp alatt, folytatjuk CDP-vel")

        log_step("PDF", "CDP PDF mentes a print preview ablakbol...")
        try:
            pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "paperWidth": 8.27,
                "paperHeight": 11.69,
                "marginTop": 0.4,
                "marginBottom": 0.4,
                "marginLeft": 0.4,
                "marginRight": 0.4,
            })
            pdf_bytes = base64.b64decode(pdf_data['data'])
            output_path = os.path.join(download_folder, f"{new_name}.pdf")
            if os.path.exists(output_path):
                os.remove(output_path)
            with open(output_path, 'wb') as f:
                f.write(pdf_bytes)
            log_success(f"PDF mentve: {new_name}.pdf ({len(pdf_bytes)} bytes)")
            return True
        except Exception as e:
            log_error("CDP PDF mentes hiba", str(e))
            return False

    except Exception as e:
        log_error("PDF mentes hiba", str(e))
        return False

    finally:
        try:
            for handle in list(driver.window_handles):
                if handle != tracking_window:
                    driver.switch_to.window(handle)
                    driver.close()
                    log_step("Ablak", "Extra ablak bezarva")
        except Exception as e:
            log_step("Ablak", f"Bezarasi hiba: {str(e)}")
        try:
            driver.switch_to.window(tracking_window)
            log_step("Ablak", "Visszavaltas tracking ablakra")
        except:
            if driver.window_handles:
                driver.switch_to.window(driver.window_handles[0])

def main():
    if len(sys.argv) < 4:
        log_error("Hianyzo argumentumok"); return 1

    ups_url         = sys.argv[1]
    excel_path      = sys.argv[2]
    download_folder = sys.argv[3]

    log_message("="*60)
    log_message("UPS POD - debuggerAddress mod")
    log_message("="*60)

    log_message("[1/5] Excel beolvasasa...")
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        log_success(f"Excel beolvasva - {len(df)} sor")
    except Exception as e:
        log_error("Excel olvasasi hiba", str(e)); return 1

    required = ['Tracking Number', 'összefűz']
    missing = [c for c in required if c not in df.columns]
    if missing:
        log_error("Hianyzó oszlopok", str(missing)); return 1

    try:
        wb = load_workbook(excel_path)
        ws = wb.active
    except Exception as e:
        log_error("Excel megnyitas hiba", str(e)); return 1

    to_process = []
    for idx, row in df.iterrows():
        excel_row = idx + 2
        if is_row_processed(ws, excel_row):
            log_step("Szures", f"Sor {excel_row} mar zold, kihagyva")
            continue
        tracking = str(row['Tracking Number']).strip() if pd.notna(row['Tracking Number']) else ''
        new_name = str(row['összefűz']).strip() if pd.notna(row['összefűz']) else ''
        if not tracking or not new_name:
            continue
        to_process.append((idx, excel_row, tracking, new_name))

    total = len(to_process)
    if total == 0:
        log_message("Nincs feldolgozando sor."); return 0
    log_success(f"Feldolgozando sorok: {total}")
    update_progress(0, total)

    log_message("[2/5] Csatlakozas a POD Chrome-hoz (port 9222)...")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        service = Service(executable_path=r"C:\WebDriver\chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        log_success(f"Csatlakozva! Jelenlegi URL: {driver.current_url}")
    except Exception as e:
        log_error("Csatlakozasi hiba - fut-e a POD Chrome?", str(e))
        return 1

    try:
        # Bezarjuk az osszes meglevo tabot, csak egy maradjon
        all_handles = driver.window_handles
        if len(all_handles) > 1:
            log_step("Init", f"{len(all_handles)} tab talalhato, bezarjuk a feleslegeseket...")
            # Az elso tabon maradunk, a tobbi bezarasra kerul
            driver.switch_to.window(all_handles[0])
            for handle in all_handles[1:]:
                driver.switch_to.window(handle)
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
            log_success("Felesleges tabok bezarva, 1 tab maradt")

        log_step("Nav", f"Navigalas: {ups_url}")
        driver.get(ups_url)
        time.sleep(3)
        log_success("UPS tracking oldal betoltve")

        processed     = 0
        success_count = 0
        zold_fill = PatternFill(start_color=GREEN_COLOR, end_color=GREEN_COLOR, fill_type='solid')

        for idx, excel_row, tracking, new_name in to_process:
            if should_stop():
                log_message("Leallitasi keres eszlelve..."); break

            log_message("")
            log_message("-"*50)
            log_message(f"Feldolgozas: {tracking} -> {new_name} (sor: {excel_row})")
            log_message("-"*50)

            log_step("3a", "Tracking mezo keresese...")
            track_input = None
            for by, sel, desc in [
                (By.ID, "stApp_trackingNumber", "ID"),
                (By.CSS_SELECTOR, "textarea[formcontrolname='trackingNumber']", "Angular"),
                (By.CSS_SELECTOR, "textarea.ups-textbox_textarea", "Class"),
                (By.NAME, "trackingnumber", "NAME")
            ]:
                try:
                    track_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, sel)))
                    log_step("3a", f"Megtalalva: {desc}")
                    break
                except:
                    continue

            if not track_input:
                log_error("Tracking mezo nem talalhato"); continue

            human_click(driver, track_input)
            time.sleep(random.uniform(0.5, 1.0))
            track_input.clear()
            time.sleep(0.2)
            track_input.send_keys(Keys.CONTROL + "a")
            track_input.send_keys(Keys.DELETE)
            time.sleep(0.3)

            driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                track_input, tracking
            )
            log_step("3a", f"Beillesztve: '{tracking}'")
            time.sleep(random.uniform(0.5, 1.0))

            track_input.send_keys(Keys.TAB)
            time.sleep(random.uniform(1.0, 1.5))

            try:
                actual = track_input.get_attribute('value')
                log_step("3a", f"Mezo tartalma: '{actual}'")
                if actual.strip() != tracking.strip():
                    log_step("3a", "Ertek nem egyezik, ujra...")
                    human_click(driver, track_input)
                    track_input.clear()
                    time.sleep(0.5)
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                        track_input, tracking
                    )
                    time.sleep(0.8)
                    track_input.send_keys(Keys.TAB)
                    time.sleep(0.8)
            except:
                pass

            log_step("3b", "Track gomb keresese...")
            try:
                track_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "stApp_btnTrack"))
                )
                log_success("Track gomb megtalálva")
            except Exception as e:
                log_error("Track gomb hiba", str(e)); continue

            human_click(driver, track_btn)
            log_success("Track gomb megnyomva")

            # Dinamikus varakozas - max 8mp, de ha megjelenik a POD link azonnal tovabb
            log_step("Varas", "Tracking eredmenyre varunk (max 8mp)...")
            wait_start = time.time()
            while time.time() - wait_start < 8:
                if driver.find_elements(By.ID, "stApp_btnProofOfDeliveryonDetails"):
                    log_success("POD link megjelent, tovabb")
                    break
                time.sleep(0.5)

            close_policy_popup(driver)
            close_chat_if_present(driver)

            # Van POD link?
            pod_link = None
            used = ""
            for by, sel, desc in [
                (By.ID, "stApp_btnProofOfDeliveryonDetails", "ID"),
                (By.LINK_TEXT, "Proof of Delivery", "Link szoveg"),
                (By.PARTIAL_LINK_TEXT, "Proof", "Reszleges")
            ]:
                elems = driver.find_elements(by, sel)
                if elems:
                    pod_link = elems[0]
                    used = desc
                    log_success(f"POD link talalva: {desc}")
                    break

            if not pod_link:
                log_error("Nincs POD link - sor kihagyva, visszanavigalas...")
                driver.get(ups_url)
                time.sleep(random.uniform(3, 5))
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "stApp_trackingNumber"))
                    )
                except TimeoutException:
                    pass
                continue
            log_success("POD link megtalalhato - folytatjuk")

            tracking_window = driver.current_window_handle
            human_click(driver, pod_link)
            log_success(f"POD link megnyitva ({used})")

            try:
                WebDriverWait(driver, 8).until(lambda d: len(d.window_handles) > 1)
                for w in driver.window_handles:
                    if w != tracking_window:
                        driver.switch_to.window(w)
                        break
                log_success("POD ablakra valtva")
                time.sleep(3)
            except Exception as e:
                log_step("Ablak", f"POD ablak nem nyilt: {str(e)}")

            pdf_saved = save_pod_pdf(driver, download_folder, new_name, tracking_window)

            if pdf_saved:
                for col in range(1, 6):
                    ws.cell(row=excel_row, column=col).fill = zold_fill
                log_success(f"Sor {excel_row} zoldre szinezve")
                success_count += 1
            else:
                log_error("PDF mentes sikertelen")

            log_step("Nav", "Visszanavigalas...")
            driver.get(ups_url)
            time.sleep(random.uniform(3, 5))

            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "stApp_trackingNumber"))
                )
                log_success("Tracking oldal keszen all")
                time.sleep(random.uniform(1.5, 2.5))
            except TimeoutException:
                log_error("Tracking mezo nem jelent meg, folytatjuk...")

            processed += 1
            update_progress(processed, total)
            log_success(f"Feldolgozva: {processed}/{total}")

        # [4] Excel mentés - letöltési mappába
        log_message("")
        log_message("[4/5] Excel mentese...")
        excel_basename = os.path.basename(excel_path)
        if excel_basename.endswith('.xlsx'):
            excel_filename = excel_basename[:-5] + '_FELDOLGOZOTT.xlsx'
        else:
            excel_filename = excel_basename + '_FELDOLGOZOTT.xlsx'
        output_path = os.path.join(download_folder, excel_filename)
        try:
            wb.save(output_path)
            sys.stdout.write(f"LOG: [OK] Excel mentve: {output_path}\n")
            sys.stdout.write(f"LOG: [OK] Eredmeny: {success_count}/{total} sikeres\n")
            sys.stdout.write("LOG: [5/5] Kesz!\n")
            sys.stdout.flush()
        except Exception as e:
            log_error("Excel mentesi hiba", str(e)); return 1

        return 0

    except Exception as e:
        log_error("Varatlan hiba", str(e)); return 1
    finally:
        sys.stdout.write("LOG: A POD Chrome nyitva maradt.\n")
        sys.stdout.flush()
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)

if __name__ == "__main__":
    sys.exit(main())
'@

    $tempPython = [System.IO.Path]::GetTempFileName() + ".py"
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tempPython, $pythonScript, $utf8WithBom)

    Write-Log "Python script futtatasa..."
    Write-Log ""

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "py"
    $psi.Arguments = "`"$tempPython`" `"$url`" `"$excelPath`" `"$downloadFolder`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $script:pythonProcess = $process

    $outputEvent = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            if ($data.StartsWith("LOG: ")) {
                $message = $data.Substring(5)
                $form.BeginInvoke([Action]{ Write-Log $message })
            } elseif ($data.StartsWith("PROGRESS: ")) {
                $parts = $data.Substring(10).Split(',')
                if ($parts.Count -eq 2) {
                    $current = [int]$parts[0]; $total = [int]$parts[1]
                    $form.BeginInvoke([Action]{ $progressBar.Maximum = $total; $progressBar.Value = $current })
                }
            }
        }
    }

    $errorEvent = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            $form.BeginInvoke([Action]{ Write-Log "PYTHON HIBA: $data" })
            Add-Content -Path "C:\temp\python_hibak.log" -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $data`r`n" -ErrorAction SilentlyContinue
        }
    }

    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    $script:pythonProcess = $null

    Unregister-Event -SourceIdentifier $outputEvent.Name -Force -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $errorEvent.Name -Force -ErrorAction SilentlyContinue
    Remove-Item $tempPython -Force -ErrorAction SilentlyContinue

    Write-Log ""
    Write-Log "="*50
    if ($exitCode -eq 0) {
        Write-Log "SIKERESEN BEFEJEZODOTT"
        # Ablak előtérbe hozása
        $form.TopMost = $true
        $form.Activate()
        [System.Windows.Forms.MessageBox]::Show("A letöltés sikeresen befejeződött!", "Siker", "OK", "Information")
        $form.TopMost = $false
    } else {
        Write-Log "HIBA TORTENT (kód: $exitCode)"
        $form.TopMost = $true
        $form.Activate()
        [System.Windows.Forms.MessageBox]::Show("Hiba történt! Ellenőrizd a naplót.", "Hiba", "OK", "Error")
        $form.TopMost = $false
    }
    Write-Log "="*50

    $progressBar.Value = 0
    $startButton.Enabled = $true
    $stopButton.Enabled = $false
})

$form.ShowDialog()
