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
        [Alias("EnableLogging")] [switch]$log
    )
    switch ($type) {
        "Chrome" { $options = [ChromeOptions]::new() }
        "Edge" {
            $options = [EdgeOptions]::new()
            $options.AddArgument("do-not-de-elevate") # prevent error when run as admin
            $options.AddArgument("enable-features=msEdgeTowerAutoHide")
            $options.AddUserProfilePreference("user_experience_metrics.personalization_data_consent_enabled", $true)
        }
        Default { return $null }
    }
    $options.AddArgument("ignore-ssl-errors")
    $options.AddArgument("ignore-certificate-errors")
    $options.AddArgument("remote-debugging-port=9222")
    $options.AddExcludedArgument("enable-automation")
    $options.AddUserProfilePreference("credentials_enable_service", $false)
    $options.AddUserProfilePreference("profile.password_manager_enabled", $false)
    if (!$log) { $options.AddExcludedArgument("enable-logging") }
    $maxtry = 3
    for ($try = 0; $try -lt $maxtry; $try++) {
        try {
            switch ($type) {
                "Edge" { $driver = [EdgeDriver]::new($options) }
                "Chrome" { $driver = [ChromeDriver]::new($options) }
                Default { return $null }
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
        Default { return $null }
    }
}

function Get-WebDriverWait {
    [OutputType([OpenQA.Selenium.Support.UI.WebDriverWait])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.IWebDriver]$driver
    )
    return [WebDriverWait]::new($driver, (New-TimeSpan -Seconds 5))
}

function Import-SeleniumBinary ($path) {
    $Script:BinPath = $env:JENKINS_URL ?
        "$([Environment]::GetEnvironmentVariable($env:BotAgent))\libraries\$($BinPath)" : "$($path)\$($BinPath)"
    if (!(Test-Path -Path $BinPath)) { New-Item -Path $BinPath -ItemType Directory -Force }
    Get-SeleniumBinary
    Import-Module "$($BinPath)\WebDriver.dll" -Global -Force -ErrorAction Stop
    Import-Module "$($BinPath)\WebDriver.Support.dll" -Global -Force -ErrorAction Stop
    Set-SeleniumEnvironment
}

function Get-SeleniumBinary ($url) {
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