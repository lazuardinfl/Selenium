using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Interactions
using namespace OpenQA.Selenium.Support.UI
using namespace System
using namespace System.Management.Automation

function Find-Element {
    [OutputType([OpenQA.Selenium.WebElement])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($by) {
            "Id" { return $driver.FindElement([By]::Id($value)) }
            "XPath" { return $driver.FindElement([By]::XPath($value)) }
            Default { throw [ValidationMetadataException] "Invalid find by type" }
        }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Wait-Element {
    [OutputType([OpenQA.Selenium.WebElement])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WebDriverWait")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.Support.UI.WebDriverWait]$wait,
        [Alias("WaitMethod")] [ValidateSet("Appear", "Disappear")] [string]$method,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("TimeoutDuration")] [int]$timeout,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        $wait.Timeout = New-TimeSpan -Seconds $timeout
        $wait.Message = "Element $($by) '$($value)' not $($method.ToLower())"
        $found = switch ($method) {
            "Appear" {
                $wait.Until([Func[IWebDriver, WebElement]]{
                    try {
                        $element = Find-Element $driver $by $value
                        return $element.Displayed -and $element.Enabled ? $element : $null
                    }
                    catch [ValidationMetadataException] { throw }
                    catch { return $null }
                })
            }
            "Disappear" {
                $wait.Until([Func[IWebDriver, bool]]{
                    try {
                        $element = Find-Element $driver $by $value
                        return !$element.Displayed
                    }
                    catch [ValidationMetadataException] { throw }
                    catch { return $true }
                })
            }
            Default { throw "Invalid wait element method type" }
        }
        Start-Sleep -Seconds $sleep
        return $found
    }
    catch { if ($silent) { return $method -eq "Disappear" ? $false : $null } else { throw } }
}

function Invoke-Click {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        $element = Find-Element $driver $by $value
        $element.Click()
        Start-Sleep -Seconds $sleep
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Set-Text {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("TextInput")] [string]$text,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("EnterAfter")] [switch]$enter,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        $field = Find-Element $driver $by $value
        $field.Click()
        $field.Clear()
        $field.SendKeys($text)
        if ($enter) {
            Start-Sleep -Seconds 1
            $field.SendKeys([Keys]::Enter)
        }
        Start-Sleep -Seconds $sleep
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Get-Handle {
    [OutputType([string[]])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("HandleScope")] [ValidateSet("Current", "All")] [string]$scope,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($scope) {
            "Current" { return $driver.CurrentWindowHandle }
            "All" { return $driver.WindowHandles }
            Default { throw "Invalid handle scope" }
        }
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Switch-Handle {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("HandleType")] [ValidateSet("Alert", "Frame", "Tab", "Window")] [string]$handle,
        [Alias("HandleValue")] [ArgumentCompletions("AcceptAlert", "DismissAlert", "BaseFrame", "ParentFrame", "FrameValue",
            "NewTab", "NewWindow", "WindowValue")] $value,
        [Alias("WaitAfter")] [int]$sleep = 1,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($handle) {
            "Alert" {
                switch ($value) {
                    "AcceptAlert" { $driver.SwitchTo().Alert().Accept() }
                    "DismissAlert" { $driver.SwitchTo().Alert().Dismiss() }
                    Default { throw "Invalid alert method" }
                }
            }
            "Frame" {
                switch ($value) {
                    "BaseFrame" { $driver.SwitchTo().DefaultContent() | Out-Null }
                    "ParentFrame" { $driver.SwitchTo().ParentFrame() | Out-Null }
                    Default { $driver.SwitchTo().Frame($value) | Out-Null }
                }
            }
            { $handle -in @("Tab", "Window") } {
                switch ($value) {
                    "NewTab" { $driver.SwitchTo().NewWindow([WindowType]::Tab) | Out-Null }
                    "NewWindow" { $driver.SwitchTo().NewWindow([WindowType]::Window) | Out-Null }
                    Default { $driver.SwitchTo().Window($value) | Out-Null }
                }
            }
            Default { throw "Invalid handle type" }
        }
        Start-Sleep -Seconds $sleep
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Invoke-ScrollToElement {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [ValidateNotNullOrWhiteSpace()] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("ScrollType")] [ValidateSet("Native", "HTML")] [string]$type,
        [Alias("FindElementBy")] [ValidateSet("Id", "XPath")] [string]$elementBy,
        [Alias("Element")] [string]$elementValue,
        [Alias("FindScrollBy")] [ValidateSet("Id", "XPath")] [string]$scrollBy,
        [Alias("Scroll")] [string]$scrollValue,
        [Alias("TimeoutDuration")] [int]$timeout,
        [Alias("ScrollAgain")] [int]$scrollAfter,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        switch ($type) {
            "Native" { [Actions]::new($driver).ScrollToElement((Find-Element $driver $elementBy $elementValue)).Perform() }
            "HTML" {
                $wait = [WebDriverWait]::new($driver, (New-TimeSpan -Seconds $timeout))
                $wait.PollingInterval = New-TimeSpan -Milliseconds 1
                $wait.Message = "Scroll failed to element $($elementBy) '$($elementValue)'"
                $found = $wait.Until([Func[IWebDriver, bool]] {
                    try {
                        $element = Find-Element $driver $elementBy $elementValue
                        if ($element.Displayed -and $element.Enabled) { return $true }
                    }
                    catch [ValidationMetadataException] { throw }
                    catch {}
                    Invoke-Click $driver $scrollBy $scrollValue | Out-Null
                })
                if ($found) { for ($i = 0; $i -lt $scrollAfter; $i++) { Invoke-Click $driver $scrollBy $scrollValue | Out-Null } }
                return $found
            }
            Default { throw "Invalid scroll type" }
        }
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Update-Elements {
    [OutputType([hashtable])]
    param (
        [Alias("Hashtable")] [ValidateNotNullOrWhiteSpace()] [hashtable]$hash,
        [Alias("OldValue")] $old,
        [Alias("NewValue")] $new,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    try {
        @($hash.GetEnumerator() | Where-Object { $_.Value -match $old }) | ForEach-Object {
            $hash[$_.Key] = $hash[$_.Key] -replace $old, $new
        }
        return $hash
    }
    catch { if ($silent) { return $null } else { throw } }
}
