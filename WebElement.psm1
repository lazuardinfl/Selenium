using namespace OpenQA.Selenium
using namespace OpenQA.Selenium.Support.UI

function Wait-Appear ($driver, $wait, $by, $element, $duration, $waitafter) {
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $found = $wait.Until([System.Func[IWebDriver, bool]]{
        try {
            switch ($by) {
                "id" {$elm = $driver.FindElement([By]::Id($element))}
                "xpath" {$elm = $driver.FindElement([By]::XPath($element))}
                Default {}
            }
            return $elm.Displayed -and $elm.Enabled
        }
        catch {return $null}
    })
    if (($found -eq $true) -and ($null -ne $waitafter)) {Start-Sleep -s $waitafter}
}

function Wait-Disappear ($driver, $wait, $by, $element, $duration, $waitafter) {
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $disappear = $wait.Until([System.Func[IWebDriver, bool]]{
        try {
            switch ($by) {
                "id" {$elm = $driver.FindElement([By]::Id($element))}
                "xpath" {$elm = $driver.FindElement([By]::XPath($element))}
                Default {}
            }
            if ($elm.Displayed -eq $true) {return $false}
            else {return $true}
        }
        catch {return $true}
    })
    if (($disappear -eq $true) -and ($null -ne $waitafter)) {Start-Sleep -s $waitafter}
}

function Switch-Handle ($driver, $wait, $by, $handle, $duration, $waitafter) {
    if ($null -eq $duration) {$duration = 1}
    if ($null -eq $waitafter) {$waitafter = 1}
    $wait.Timeout = New-TimeSpan -Seconds $duration
    $switched = $wait.Until([System.Func[IWebDriver, bool]]{
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
    if ($switched -eq $true) {Start-Sleep -s $waitafter}
}

function Invoke-Click ($driver, $by, $element) {
    switch ($by) {
        "id" {$driver.FindElement([By]::Id($element)).Click()}
        "xpath" {$driver.FindElement([By]::XPath($element)).Click()}
        Default {}
    }
}

function Set-Text ($driver, $by, $element, $value, $enter) {
    switch ($by) {
        "id" {$field = $driver.FindElement([By]::Id($element))}
        "xpath" {$field = $driver.FindElement([By]::XPath($element))}
        Default {}
    }
    $field.Click()
    $field.Clear()
    $field.SendKeys($value)
    if ($enter -eq "enter") {
        Start-Sleep -s 1
        $field.SendKeys([Keys]::Enter)
    }
}

function Invoke-ScrollToElement ($driver, $byElement, $element, $byScroll, $scroll, $duration, $scrollafter, $fail) {
    Wait-Loop @{
        action = {
            try {
                switch ($byElement) {
                    "id" {$elm = $driver.FindElement([By]::Id($element))}
                    "xpath" {$elm = $driver.FindElement([By]::XPath($element))}
                    Default {}
                }
                if (($elm.Displayed -and $elm.Enabled) -eq $true) {
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