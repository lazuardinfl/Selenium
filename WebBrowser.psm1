using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Support.UI

function Start-Browser ($app, $headless) {
    switch ($app) {
        "EdgeIE" {
            $servicesEdgeIE = [IE.InternetExplorerDriverService]::CreateDefaultService()
            $servicesEdgeIE.Port = 9223
            $optionsEdgeIE = [IE.InternetExplorerOptions]::new()
            $optionsEdgeIE.AttachToEdgeChrome = $true
            $optionsEdgeIE.EdgeExecutablePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            $optionsEdgeIE.IntroduceInstabilityByIgnoringProtectedModeSettings = $true
            $optionsEdgeIE.IgnoreZoomLevel = $true
        }
        "Edge" {
            $optionsEdge = [Edge.EdgeOptions]::new()
            $optionsEdge.AddArgument("ignore-ssl-errors")
            $optionsEdge.AddArgument("ignore-certificate-errors")
            $optionsEdge.AddArgument("remote-debugging-port=9222") # prevent admin error and to reconnect session
            $optionsEdge.AddExcludedArgument("enable-logging") # hide console log
            # $optionsEdge.AddExcludedArgument("enable-automation") # failed at Edge > v109
            $optionsEdge.AddUserProfilePreference("credentials_enable_service", $false)
            $optionsEdge.AddUserProfilePreference("profile.password_manager_enabled", $false)
            $optionsEdge.AddUserProfilePreference("user_experience_metrics.personalization_data_consent_enabled", $true)
            if ($headless -eq "headless") {$optionsEdge.AddArgument("headless")} # without UI
        }
        "Chrome" {
            $optionsChrome = [Chrome.ChromeOptions]::new()
            $optionsChrome.AddArgument("ignore-ssl-errors")
            $optionsChrome.AddArgument("ignore-certificate-errors")
            $optionsChrome.AddArgument("remote-debugging-port=9222") # to reconnect session
            $optionsChrome.AddExcludedArgument("enable-automation")
            $optionsChrome.AddUserProfilePreference("credentials_enable_service", $false)
            $optionsChrome.AddUserProfilePreference("profile.password_manager_enabled", $false)
            if ($headless -eq "headless") {$optionsChrome.AddArgument("headless")} # without UI
        }
        Default {}
    }
    $maxtry = 3
    for ($try = 0; $try -lt $maxtry; $try++) {
        try {
            switch ($app) {
                "EdgeIE" {$driver = [IE.InternetExplorerDriver]::new($servicesEdgeIE, $optionsEdgeIE)}
                "Edge" {$driver = [Edge.EdgeDriver]::new($optionsEdge)}
                "Chrome" {$driver = [Chrome.ChromeDriver]::new($optionsChrome)}
                Default {}
            }
            Start-Sleep -s 2
            $driver.Manage().Window.Maximize()
            Start-Sleep -s 2
            return $driver
        }
        catch {
            Stop-Browser $driver $app
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

function Get-WebDriverWait ($driver) {
    return [WebDriverWait]::new($driver, (New-TimeSpan -Seconds 5))
}

function Import-SeleniumBinaries ($path) {
    $version = "Selenium\bin\4.21.0"
    $binPath = "$($path)\$($version)"
    if (!(Test-Path -Path $binPath)) {
        $binPath = "$([System.Environment]::GetEnvironmentVariable($env:BotAgent))\libraries\$($version)"
    }
    Import-Module "$($binPath)\WebDriver.dll" -Global -Force -ErrorAction Stop
    Import-Module "$($binPath)\WebDriver.Support.dll" -Global -Force -ErrorAction Stop
    if (($env:Path -split ";") -notcontains $binPath) {$env:Path += ";$($binPath)"}
    $env:SE_MANAGER_PATH = "$($binPath)\selenium-manager.exe"
    $env:SE_AVOID_BROWSER_DOWNLOAD = $true;
    $env:SE_AVOID_STATS = $true;
}