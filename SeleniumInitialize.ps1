$NuGetUrl = "https://api.nuget.org/v3/index.json"
$Resources = @(
    @{ Name = "Selenium.WebDriver"; Version = "4.21.0" },
    @{ Name = "Selenium.Support"; Version = "4.21.0" }
)
try {
    if (!(Get-PSResourceRepository | Where-Object { $_.Uri -eq $NuGetUrl })) {
        Register-PSResourceRepository -Name NuGetGallery -Uri $NuGetUrl -Priority 80 -Trusted -Force -ErrorAction Stop
    }
    foreach ($resource in $Resources) {
        if (!(Get-InstalledPSResource @resource -ErrorAction SilentlyContinue)) {
            Install-PSResource $resource.Name -Version "[$($resource.Version)]" -Scope CurrentUser -TrustRepository -AcceptLicense -SkipDependencyCheck -ErrorAction Stop
            #$location = (Get-InstalledPSResource Selenium.WebDriver -Version 4.21.0).InstalledLocation
            #Import-Module "$($location)\Selenium.WebDriver\4.21.0\lib\netstandard2.0\WebDriver.dll" -Force
        }
    }
}
catch {
    throw
}