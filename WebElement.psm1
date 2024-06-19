using namespace OpenQA.Selenium
using namespace System

function Wait-Appear {
    [OutputType([OpenQA.Selenium.WebElement])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WebDriverWait")] [OpenQA.Selenium.Support.UI.WebDriverWait]$wait,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("WaitDuration")] [int]$duration,
        [Alias("WaitAfter")] [int]$sleep
    )
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $appear = $wait.Until([Func[IWebDriver, WebElement]] {
        try {
            switch ($by) {
                "Id" { $element = $driver.FindElement([By]::Id($value)) }
                "XPath" { $element = $driver.FindElement([By]::XPath($value)) }
                Default { return $null }
            }
            return $element.Displayed -and $element.Enabled ? $element : $null
        }
        catch { return $null }
    })
    if ($appear) { Start-Sleep -Seconds $sleep }
    return $appear
}

function Wait-Disappear {
    [OutputType([bool])]
    param (
        [Alias("WebDriver")] [OpenQA.Selenium.WebDriver]$driver,
        [Alias("WebDriverWait")] [OpenQA.Selenium.Support.UI.WebDriverWait]$wait,
        [Alias("FindBy")] [ValidateSet("Id", "XPath")] [string]$by,
        [Alias("Element")] [string]$value,
        [Alias("WaitDuration")] [int]$duration,
        [Alias("WaitAfter")] [int]$sleep
    )
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $disappear = $wait.Until([Func[IWebDriver, bool]] {
        try {
            switch ($by) {
                "Id" { $element = $driver.FindElement([By]::Id($value)) }
                "XPath" { $element = $driver.FindElement([By]::XPath($value)) }
                Default { return $null }
            }
            return $element.Displayed ? $false : $true
        }
        catch { return $true }
    })
    if ($disappear) { Start-Sleep -Seconds $sleep }
    return $disappear ? $true : $false
}

function Invoke-Click ($driver, $by, $value) {
    switch ($by) {
        "id" {$driver.FindElement([By]::Id($value)).Click()}
        "xpath" {$driver.FindElement([By]::XPath($value)).Click()}
        Default {}
    }
}

function Set-Text ($driver, $by, $value, $text, $enter) {
    switch ($by) {
        "id" {$field = $driver.FindElement([By]::Id($value))}
        "xpath" {$field = $driver.FindElement([By]::XPath($value))}
        Default {}
    }
    $field.Click()
    $field.Clear()
    $field.SendKeys($text)
    if ($enter -eq "enter") {
        Start-Sleep -s 1
        $field.SendKeys([Keys]::Enter)
    }
}

function Switch-Handle ($driver, $wait, $by, $handle, $duration, $sleep) {
    if ($null -eq $duration) {$duration = 1}
    if ($null -eq $sleep) {$sleep = 1}
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $switched = $wait.Until([Func[IWebDriver, bool]]{
        try {
            switch ($by) {
                "alert" {
                    switch ($handle) {
                        "accept" {$driver.SwitchTo().Alert().Accept()}
                        "dismiss" {$driver.SwitchTo().Alert().Dismiss()}
                        Default {$driver.SwitchTo().Alert() | Out-Null}
                    }
                }
                "frame" {
                    switch ($handle) {
                        "BaseFrame" {$driver.SwitchTo().DefaultContent() | Out-Null}
                        "ParentFrame" {$driver.SwitchTo().ParentFrame() | Out-Null}
                        Default {$driver.SwitchTo().Frame($handle) | Out-Null}
                    }
                }
                {$by -in @("tab", "window")} {$driver.SwitchTo().Window($handle) | Out-Null}
                Default {throw}
            }
            return $true
        }
        catch {return $null}
    })
    if ($switched -eq $true) {Start-Sleep -s $sleep}
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