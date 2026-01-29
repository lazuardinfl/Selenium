$NuGetUrl = "https://api.nuget.org/v3/index.json"
$Resources = @(
    @{ Name = "Selenium.WebDriver"; Version = "4.38.0"; Import = @("lib\net8.0\WebDriver.dll") },
    @{ Name = "Selenium.Support"; Version = "4.38.0"; Import = @("lib\netstandard2.0\WebDriver.Support.dll") }
)
try {
    # register nuget repo
    if (!(Get-PSResourceRepository | Where-Object { $_.Uri -eq $NuGetUrl })) {
        Register-PSResourceRepository -Name NuGetGallery -Uri $NuGetUrl -Priority 80 -Trusted -Force -ErrorAction Stop
    }
    # install
    foreach ($resource in $Resources) {
        if (!(Get-InstalledPSResource -Name $resource.Name -Version $resource.Version -ErrorAction SilentlyContinue)) {
            Install-PSResource $resource.Name -Version "[$($resource.Version)]" -Scope CurrentUser -TrustRepository -AcceptLicense -SkipDependencyCheck -ErrorAction Stop
        }
    }
    # import
    foreach ($resource in $Resources) {
        $path = (Get-InstalledPSResource -Name $resource.Name -Version $resource.Version -ErrorAction Stop).InstalledLocation
        foreach ($item in $resource.Import) {
            Import-Module "$path\$($resource.Name)\$($resource.Version)\$item" -Global -Force -ErrorAction Stop
        }
    }
    # manager
    $resource = $Resources | Where-Object { $_.Name -eq "Selenium.WebDriver" }
    $path = (Get-InstalledPSResource -Name $resource.Name -Version $resource.Version -ErrorAction Stop).InstalledLocation
    $env:SE_MANAGER_PATH = "$path\$($resource.Name)\$($resource.Version)\runtimes\win\native\selenium-manager.exe"
}
catch { throw }
