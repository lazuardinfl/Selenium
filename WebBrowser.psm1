using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Chrome
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace System

$Version = "4.21.0"
$BinPath = "Selenium\bin\$($Version)"

function Start-Browser {
    [OutputType([OpenQA.Selenium.Chromium.ChromiumDriver])]
    param (
        [Alias("BrowserName")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("UseUserProfile")] [string]$userProfile,
        [Alias("EnableLogging")] [switch]$log,
        [Alias("DisableReloadOnFail")] [switch]$noRepeat,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    switch ($type) {
        "Chrome" {
            $options = [ChromeOptions]::new()
            $userData = "$($env:LOCALAPPDATA)\Google\Chrome\User Data"
        }
        "Edge" {
            $options = [EdgeOptions]::new()
            $options.AddArgument("do-not-de-elevate") # prevent error when run as admin
            $options.AddArgument("enable-features=msEdgeTowerAutoHide")
            $options.AddUserProfilePreference("user_experience_metrics.personalization_data_consent_enabled", $true)
            $userData = "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data"
        }
        Default { if ($silent) { return $null } else { throw "Invalid browser type" } }
    }
    $options.AddArgument("ignore-ssl-errors")
    $options.AddArgument("ignore-certificate-errors")
    $options.AddArgument("remote-debugging-port=9222")
    $options.AddExcludedArgument("enable-automation")
    $options.AddUserProfilePreference("credentials_enable_service", $false)
    $options.AddUserProfilePreference("profile.password_manager_enabled", $false)
    if (!$log) { $options.AddExcludedArgument("enable-logging") }
    if ($userProfile) {
        if (Test-Path -Path "$($userData)\$($userProfile)") {
            $options.AddArgument("user-data-dir=$($userData)")
            $options.AddArgument("profile-directory=$($userProfile)")
        }
        elseif (Split-Path -Path $userProfile -IsAbsolute) {
            $options.AddArgument("user-data-dir=$($userProfile)")
        }
        elseif ($silent) { return $null }
        else { throw "Invalid user profile" }
    }
    $maxtry = 3
    for ($try = 1; $try -le $maxtry; $try++) {
        try {
            switch ($type) {
                "Edge" { $driver = [EdgeDriver]::new($options) }
                "Chrome" { $driver = [ChromeDriver]::new($options) }
            }
            $driver.Manage().Window.Maximize()
            return $driver
        }
        catch {
            if (($try -eq $maxtry) -or $noRepeat) {
                if ($silent) { return $null } else { throw }
            }
            else {
                Stop-Browser $driver -Force $type
                Remove-Variable driver -Scope Local
                Remove-Variable driver -Scope Script
                Start-Sleep -Seconds 5
            }
        }
    }
}

function Stop-Browser {
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("Force")] [ValidateSet("Chrome", "Edge")] [string]$type
    )
    try {
        $driver.Close()
        $driver.Quit()
    }
    catch {}
    if ($type) {
        $browser = @{
            Chrome = @{ Process = "chrome"; Driver = "chromedriver" }
            Edge = @{ Process = "msedge"; Driver = "msedgedriver" }
        }
        Get-Process $browser[$type].Process | ForEach-Object { $_.CloseMainWindow() } | Out-Null
        Wait-Process $browser[$type].Process -Timeout 10
        Get-Process $browser[$type].Process | Stop-Process -Force
        Wait-Process $browser[$type].Driver -Timeout 7
        Get-Process $browser[$type].Driver | Stop-Process -Force
    }
    Start-Sleep -Seconds $sleep
}

function Resume-Browser {
    [OutputType([OpenQA.Selenium.Chromium.ChromiumDriver])]
    param (
        [Alias("BrowserName")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    $debuggerAddress = "127.0.0.1:9222"
    try {
        switch ($type) {
            "Chrome" {
                $options = [ChromeOptions]::new()
                $options.DebuggerAddress = $debuggerAddress
                return [ChromeDriver]::new($options)
            }
            "Edge" {
                $options = [EdgeOptions]::new()
                $options.DebuggerAddress = $debuggerAddress
                return [EdgeDriver]::new($options)
            }
            Default { throw "Invalid browser type" }
        }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Get-WebDriverWait {
    [OutputType([OpenQA.Selenium.Support.UI.WebDriverWait])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return [WebDriverWait]::new($driver, (New-TimeSpan -Seconds 5)) }
    catch { if ($silent) { return $null } else { throw } }
}

function Import-SeleniumBinary {
    param (
        [Alias("LibraryPath")] [string]$path
    )
    $Script:BinPath = $env:JENKINS_URL ?
        "$([Environment]::GetEnvironmentVariable($env:BotAgent))\libraries\$($BinPath)" : "$($path)\$($BinPath)"
    if (!(Test-Path -Path $BinPath)) { New-Item -Path $BinPath -ItemType Directory -Force }
    Get-SeleniumBinary
    Import-Module "$($BinPath)\WebDriver.dll" -Global -Force -ErrorAction Stop
    Import-Module "$($BinPath)\WebDriver.Support.dll" -Global -Force -ErrorAction Stop
    Set-SeleniumEnvironment
}

function Get-SeleniumBinary {
    param (
        [Alias("DownloadUrl")] [string]$url
    )
    $url = $url ? $url : ($env:SeleniumUrl ? $env:SeleniumUrl : "https://globalcdn.nuget.org/packages")
    $tempPath = "$($BinPath)\temp"
    $binaries = @(
        @{
            File = "$($BinPath)\WebDriver.dll"; Temp = "$($tempPath)\selenium.webdriver.4.21.0.nupkg\lib\netstandard2.0\WebDriver.dll"
            Hash = "B8EB2044376281311020829A0E514BC18C20D4B03C3EF4131CD1C4DEC64D0813"; Package = "selenium.webdriver.4.21.0.nupkg"
        },
        @{
            File = "$($BinPath)\WebDriver.Support.dll"; Temp = "$($tempPath)\selenium.support.4.21.0.nupkg\lib\netstandard2.0\WebDriver.Support.dll"
            Hash = "711866886C2FA5395FCB7961E32A9B57EA89B1B479D8D5D1BC1D2D6178D96D7E"; Package = "selenium.support.4.21.0.nupkg"
        },
        @{
            File = "$($BinPath)\selenium-manager.exe"; Temp = "$($tempPath)\selenium.webdriver.4.21.0.nupkg\manager\windows\selenium-manager.exe"
            Hash = "B7B27C6DFE6F1D30BB63A3038C799E2C8E9E801C0AEE4528C7541D93F70DFDDB"; Package = "selenium.webdriver.4.21.0.nupkg"
        }
    )
    foreach ($binary in $binaries) {
        if ((Get-FileHash $binary.File).Hash -ne $binary.Hash) {
            if (!(Test-Path -Path $binary.Temp)) {
                New-Item -Path "$($tempPath)\$($binary.Package)" -ItemType Directory -Force
                Invoke-RestMethod -Uri "$($url)/$($binary.Package)" -OutFile "$($BinPath)\$($binary.Package)"
                Expand-Archive -Path "$($BinPath)\$($binary.Package)" -DestinationPath "$($tempPath)\$($binary.Package)" -Force
                Remove-Item -Path "$($BinPath)\$($binary.Package)" -Force
            }
            Copy-Item -Path $binary.Temp -Destination $binary.File -Force
        }
    }
    if (Test-Path -Path $tempPath) { Remove-Item -Path $tempPath -Recurse -Force }
}

function Set-SeleniumEnvironment {
    if (($env:Path -split ";") -notcontains $BinPath) { $env:Path += ";$($BinPath)" }
    $env:SE_MANAGER_PATH = "$($BinPath)\selenium-manager.exe"
    $env:SE_DRIVER_MIRROR_URL = $env:SeleniumUrl
    $env:SE_AVOID_BROWSER_DOWNLOAD = $true;
    $env:SE_AVOID_STATS = $true;
}