using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Chrome
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace System.Security.Principal

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
    if ([WindowsPrincipal]::new([WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]::Administrator)) { $options.AddArgument("do-not-de-elevate") }
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
        [Alias("Force")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("CurrentHandleOnly")] [switch]$current
    )
    try {
        if ($current) { $driver.Close() }
        else { $driver.Quit() }
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

function Invoke-BrowserNavigation {
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("NavigationMethod")] [ArgumentCompletions("protocol://url", "Back", "Forward", "Refresh",
            "FullScreen", "Maximize", "Minimize")] [string]$method,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($method) {
            "Back" { $driver.Navigate().Back() }
            "Forward" { $driver.Navigate().Forward() }
            "Refresh" { $driver.Navigate().Forward() }
            "FullScreen" { $driver.Manage().Window.FullScreen() }
            "Maximize" { $driver.Manage().Window.Maximize() }
            "Minimize" { $driver.Manage().Window.Minimize() }
            Default { $driver.Navigate().GoToUrl($method) }
        }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Get-DriverWait {
    [OutputType([OpenQA.Selenium.Support.UI.WebDriverWait])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try { return [WebDriverWait]::new($driver, (New-TimeSpan -Seconds 5)) }
    catch { if ($silent) { return $null } else { throw } }
}

function Set-PrivateEnvironment {
    param (
        [Alias("SeleniumUrl")] [string]$url
    )
    $env:SE_DRIVER_MIRROR_URL = $url ? $url : $env:SeleniumUrl
    $env:SE_AVOID_BROWSER_DOWNLOAD = $true;
    $env:SE_AVOID_STATS = $true;
}
