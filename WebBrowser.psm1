using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Chrome
using namespace OpenQA.Selenium.Edge
using namespace OpenQA.Selenium.Support.UI
using namespace System

function Start-Browser {
    [OutputType([OpenQA.Selenium.Chromium.ChromiumDriver])]
    param (
        [Alias("BrowserName")] [ValidateSet("Chrome", "Edge")] [string]$type,
        [Alias("EnableLogging")] [switch]$log
    )
    switch ($type) {
        "Chrome" {
            $options = [ChromeOptions]::new()
        }
        "Edge" {
            $options = [EdgeOptions]::new()
            $options.AddArgument("do-not-de-elevate") # prevent error when run as admin
            $options.AddArgument("enable-features=msEdgeTowerAutoHide")
            $options.AddUserProfilePreference("user_experience_metrics.personalization_data_consent_enabled", $true)
        }
        Default {return $null}
    }
    $options.AddArgument("ignore-ssl-errors")
    $options.AddArgument("ignore-certificate-errors")
    $options.AddArgument("remote-debugging-port=9222")
    $options.AddExcludedArgument("enable-automation")
    $options.AddUserProfilePreference("credentials_enable_service", $false)
    $options.AddUserProfilePreference("profile.password_manager_enabled", $false)
    if (!$log) {$options.AddExcludedArgument("enable-logging")}
    $maxtry = 3
    for ($try = 0; $try -lt $maxtry; $try++) {
        try {
            switch ($type) {
                "Edge" {$driver = [EdgeDriver]::new($options)}
                "Chrome" {$driver = [ChromeDriver]::new($options)}
                Default {return $null}
            }
            $driver.Manage().Window.Maximize()
            return $driver
        }
        catch {
            Stop-Browser $driver $type
            Remove-Variable driver -Scope Local
            Remove-Variable driver -Scope Script
            Start-Sleep -s 5
        }
    }
    return $null
}

function Stop-Browser ($driver, $app, $waitafter) {
    try {
        $driver.Close()
        $driver.Quit()
    }
    catch {}
    switch ($app) {
        "IE" {
            Get-Process iexplore | ForEach-Object {$_.CloseMainWindow()}
            Wait-Process iexplore -Timeout 10
            Get-Process iexplore | Stop-Process -Force
            Wait-Process IEDriverServer -Timeout 7
            Get-Process IEDriverServer | Stop-Process -Force
        }
        "EdgeIE" {
            Get-Process iexplore, msedge | ForEach-Object {$_.CloseMainWindow()}
            Wait-Process iexplore, msedge -Timeout 10
            Get-Process iexplore, msedge | Stop-Process -Force
            Wait-Process IEDriverServer, msedgedriver -Timeout 7
            Get-Process IEDriverServer, msedgedriver | Stop-Process -Force
        }
        "Edge" {
            Get-Process msedge | ForEach-Object {$_.CloseMainWindow()}
            Wait-Process msedge -Timeout 10
            Get-Process msedge | Stop-Process -Force
            Wait-Process msedgedriver -Timeout 7
            Get-Process msedgedriver | Stop-Process -Force
        }
        "Chrome" {
            Get-Process chrome | ForEach-Object {$_.CloseMainWindow()}
            Wait-Process chrome -Timeout 10
            Get-Process chrome | Stop-Process -Force
            Wait-Process chromedriver -Timeout 7
            Get-Process chromedriver | Stop-Process -Force
        }
        Default {
            Get-Process iexplore, msedge, chrome | ForEach-Object {$_.CloseMainWindow()}
            Wait-Process iexplore, msedge, chrome -Timeout 10
            Get-Process iexplore, msedge, chrome | Stop-Process -Force
            Wait-Process IEDriverServer, msedgedriver, chromedriver -Timeout 7
            Get-Process IEDriverServer, msedgedriver, chromedriver | Stop-Process -Force
        }
    }
    if ($null -ne $waitafter) {Start-Sleep -s $waitafter}
}

function Resume-Browser {
    [OutputType([OpenQA.Selenium.Chromium.ChromiumDriver])]
    param (
        [Alias("BrowserName")] [ValidateSet("Chrome", "Edge")] [string]$type
    )
    $debuggerAddress = "127.0.0.1:9222"
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
        Default {return $null}
    }
}

function Get-WebDriverWait {
    [OutputType([OpenQA.Selenium.Support.UI.WebDriverWait])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.IWebDriver]$driver
    )
    return [WebDriverWait]::new($driver, (New-TimeSpan -Seconds 5))
}

function Import-SeleniumBinaries ($path) {
    $version = "Selenium\bin\4.21.0"
    $binPath = "$($path)\$($version)"
    if (!(Test-Path -Path $binPath)) {
        $binPath = "$([Environment]::GetEnvironmentVariable($env:BotAgent))\libraries\$($version)"
    }
    Import-Module "$($binPath)\WebDriver.dll" -Global -Force -ErrorAction Stop
    Import-Module "$($binPath)\WebDriver.Support.dll" -Global -Force -ErrorAction Stop
    if (($env:Path -split ";") -notcontains $binPath) {$env:Path += ";$($binPath)"}
    $env:SE_MANAGER_PATH = "$($binPath)\selenium-manager.exe"
    $env:SE_AVOID_BROWSER_DOWNLOAD = $true;
    $env:SE_AVOID_STATS = $true;
}