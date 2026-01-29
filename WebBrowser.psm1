using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Chrome
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace System.Management.Automation
using namespace System.Security.Principal

function Start-Browser {
    [OutputType([OpenQA.Selenium.Chromium.ChromiumDriver])]
    param (
        [Alias("BrowserName")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("Profile")] [ArgumentCompletions("Temporary", "Default", "Custom")] [string]$userProfile,
        [Alias("AddArguments")] [string[]]$arguments,
        [Alias("AddExcludedArguments")] [string[]]$exArguments,
        [Alias("HeadlessMode")] [switch]$headless,
        [Alias("EnableLogging")] [switch]$log,
        [Alias("DisableReloadOnFail")] [switch]$noRepeat,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($type) {
            "Chrome" {
                $options = [ChromeOptions]::new()
                $userData = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            }
            "Edge" {
                $options = [EdgeOptions]::new()
                $options.AddArgument("enable-features=msEdgeTowerAutoHide")
                $options.AddUserProfilePreference("user_experience_metrics.personalization_data_consent_enabled", $true)
                $userData = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
            }
            Default { throw [ValidationMetadataException] "Invalid browser type" }
        }
        $options.AddArgument("ignore-ssl-errors")
        $options.AddArgument("ignore-certificate-errors")
        $options.AddArgument("remote-debugging-port=9222")
        $options.AddExcludedArgument("enable-automation")
        $options.AddUserProfilePreference("credentials_enable_service", $false)
        $options.AddUserProfilePreference("profile.password_manager_enabled", $false)
        if ([WindowsPrincipal]::new([WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]::Administrator)) { $options.AddArgument("do-not-de-elevate") }
        if (!$log) { $options.AddExcludedArgument("enable-logging") }
        if ($headless) { $options.AddArgument("headless") }
        if ($userProfile -and ($userProfile -ne "Temporary")) {
            if (Test-Path -Path "$userData\$userProfile") {
                $options.AddArgument("user-data-dir=$userData")
                $options.AddArgument("profile-directory=$userProfile")
            }
            elseif (Split-Path -Path $userProfile -IsAbsolute) { $options.AddArgument("user-data-dir=$userProfile") }
            else { throw [ValidationMetadataException] "Invalid user profile" }
        }
        foreach ($arg in $arguments) { $options.AddArgument($arg) }
        foreach ($arg in $exArguments) { $options.AddExcludedArgument($arg) }
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
                if (($try -eq $maxtry) -or $noRepeat) { throw }
                else {
                    Stop-Browser $driver -Force $type -OnErrorContinue | Out-Null
                    Remove-Variable driver -Scope Local -ErrorAction SilentlyContinue
                    Remove-Variable driver -Scope Script -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 5
                }
            }
        }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Stop-Browser {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("Force")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("CurrentHandleOnly")] [switch]$current,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        if (!($driver -or $type)) { throw "Invalid parameter" }
        if ($driver) {
            if ($current) { $driver.Close() }
            else { $driver.Quit() }
        }
        if ($type) {
            $browser = @{
                Chrome = @{ Process = "chrome"; Driver = "chromedriver" }
                Edge = @{ Process = "msedge"; Driver = "msedgedriver" }
            }
            $procs = @($browser[$type].Process, $browser[$type].Driver)
            if (Get-Process $procs -ErrorAction SilentlyContinue) {
                Get-Process $browser[$type].Process -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() } | Out-Null
                for ($i = 0; $i -lt 7; $i++) {
                    Get-Process $procs -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    if (!(Get-Process $procs -ErrorAction SilentlyContinue)) { return $true }
                    Start-Sleep -Seconds 1
                }
                throw "Stop browser failed"
            }
        }
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
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
            "Refresh" { $driver.Navigate().Refresh() }
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
