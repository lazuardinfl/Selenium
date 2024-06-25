using namespace OpenQA.Selenium
using namespace System
using namespace System.Management.Automation

function Find-Element {
    [OutputType([OpenQA.Selenium.WebElement])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
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

function Wait-Appear {
    [OutputType([OpenQA.Selenium.WebElement])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WebDriverWait")] [OpenQA.Selenium.Support.UI.WebDriverWait]$wait,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("WaitDuration")] [int]$duration,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $wait.Message = "Element $($by) '$($value)' not appear"
    try {
        $appear = $wait.Until([Func[IWebDriver, WebElement]] {
            try {
                $element = Find-Element $driver $by $value
                return $element.Displayed -and $element.Enabled ? $element : $null
            }
            catch [ValidationMetadataException] { throw }
            catch { return $null }
        })
        Start-Sleep -Seconds $sleep
        return $appear
    }
    catch { if ($silent) { return $null } else { throw } }
}

function Wait-Disappear {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WebDriverWait")] [OpenQA.Selenium.Support.UI.WebDriverWait]$wait,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("WaitDuration")] [int]$duration,
        [Alias("WaitAfter")] [int]$sleep,
        [Alias("OnErrorContinue")] [switch]$silent
    )
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $wait.Message = "Element $($by) '$($value)' not disappear"
    try {
        $disappear = $wait.Until([Func[IWebDriver, bool]] {
            try {
                $element = Find-Element $driver $by $value
                return !$element.Displayed
            }
            catch [ValidationMetadataException] { throw }
            catch { return $true }
        })
        Start-Sleep -Seconds $sleep
        return $disappear
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Invoke-Click {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
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
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
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

function Switch-Handle {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("HandleType")] [ValidateSet("Alert", "Frame", "Tab", "Window")] [string]$handle,
        [Alias("HandleValue")] [ArgumentCompletions("AcceptAlert", "DismissAlert", "BaseFrame", "ParentFrame")] $value,
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
            { $handle -in @("Tab", "Window") } { $driver.SwitchTo().Window($value) | Out-Null }
            Default { throw "Invalid handle type" }
        }
        Start-Sleep -Seconds $sleep
        return $true
    }
    catch { if ($silent) { return $false } else { throw } }
}

function Invoke-ScrollToElement ($driver, $byElement, $value, $byScroll, $scroll, $duration, $scrollafter, $fail) {
    Wait-Loop @{
        action = {
            try {
                switch ($byElement) {
                    "id" {$element = $driver.FindElement([By]::Id($value))}
                    "xpath" {$element = $driver.FindElement([By]::XPath($value))}
                    Default {}
                }
                if (($element.Displayed -and $element.Enabled) -eq $true) {
                    try {for ($i = 0; $i -lt $scrollafter; $i++) {Invoke-Click $driver $byScroll $scroll}} catch {}
                    return $true
                }
                else {throw}
            }
            catch {Invoke-Click $driver $byScroll $scroll}
        }
        timeout = {if ($null -ne $fail) {Stop-Error $fail}}
        duration = New-TimeSpan -Seconds $duration
    }
}

function Update-AutomationElements ($type, $old, $new) {
    switch ($type) {
        {$type -in @("id", "all")} {
            @($Id.GetEnumerator() | Where-Object {$_.Value -match $old}) | ForEach-Object {
                $Id[$_.Key] = $Id[$_.Key] -replace $old, $new
            }
        }
        {$type -in @("xpath", "all")} {
            @($XPath.GetEnumerator() | Where-Object {$_.Value -match $old}) | ForEach-Object {
                $XPath[$_.Key] = $XPath[$_.Key] -replace $old, $new
            }
        }
        Default {}
    }
}