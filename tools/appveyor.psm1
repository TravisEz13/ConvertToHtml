[string] $moduleDir = Split-Path -Path $script:MyInvocation.MyCommand.Path –Parent

Set-StrictMode -Version latest
Function Invoke-AppveyorInstall
{
    Write-Info "Starting Install stage..."
    $nuGetPath = "$env:SystemDrive\nuget.exe"
    $webClient = New-Object 'System.Net.WebClient';
    $webClient.DownloadFile( 'https://oneget.org/nuget-anycpu-2.8.3.6.exe', $nuGetPath )  
    Write-Info "End Install stage."
}

Function Invoke-AppveyorBuild
{
    Write-Info "Starting Build stage..."

    $nuGetPath = "$env:SystemDrive\nuget.exe"
    mkdir -force .\out > $null
    mkdir -force .\nuget > $null
    mkdir -force .\examples > $null

    # Update version to current Build Version
    $versionParts = ($env:APPVEYOR_BUILD_VERSION).split('.')
    Import-Module .\ConvertToHtml
    $moduleInfo = Get-Module -Name ConvertToHtml
    $newVersion = New-Object -TypeName 'System.Version' -ArgumentList @($versionParts[0],$versionParts[1],$versionParts[2],$versionParts[3])
    $FunctionsToExport = @()
    foreach($key in $moduleInfo.ExportedFunctions.Keys)
    {
      $FunctionsToExport += $key
    }
    copy-item .\ConvertToHtml\ConvertTohtml.psd1 .\ConvertTohtmlOrigPsd1.ps1
    New-ModuleManifest -Path .\ConvertToHtml\ConvertTohtml.psd1 -Guid $moduleInfo.Guid -Author $moduleInfo.Author -CompanyName $moduleInfo.CompanyName `
    -Copyright $moduleInfo.Copyright -RootModule $moduleInfo.RootModule -ModuleVersion $newVersion -Description $moduleInfo.Description -FunctionsToExport $FunctionsToExport
    # Done Updating Version

    # Create Nuget package
    [xml]$xml = Get-Content -Raw .\ConvertToHtml\ConvertToHtml.nuspec
    $xml.package.metadata.version = $env:APPVEYOR_BUILD_VERSION
    $xml.OuterXml | out-file -FilePath .\ConvertToHtml\ConvertToHtml.nuspec
    &$nuGetPath pack .\ConvertToHtml\ConvertToHtml.nuspec -outputdirectory  .\nuget

    7z a -tzip .\out\ConvertToHtml.zip .\ConvertToHtml\*.*
    Write-Info "End Build Stage."
}

Function Invoke-AppveyorTest
{
    Write-Info "Starting Test stage..."
    # setup variables for the whole build process
    #
    $script:failedTestsCount = 0
    $webClient = New-Object 'System.Net.WebClient';
    #

 
    $coverage.CodeCoverage.MissedCommands | ConvertTo-FormattedHtml -OutClipboard
    $CodeCoverage = @('.\ConvertToHtml\exporttohtml.psm1')
    ".\tests" | %{ 
    $res = RunTest -filePath $_ -CodeCoverage @('.\ConvertToHtml\exporttohtml.psm1')
    $script:failedTestsCount += $res.FailedCount 
    $CodeCoverageTitle = "Code Coverage {0:F1}%"  -f (100 * ($res.CodeCoverage.NumberOfCommandsExecuted /$res.CodeCoverage.NumberOfCommandsAnalyzed))
    $res.CodeCoverage.MissedCommands | ConvertTo-FormattedHtml -title $CodeCoverageTitle | out-file .\examples\CodeCoverage.html
    }

    7z a -tzip .\out\examples.zip .\examples\*.html

    if ($script:failedTestsCount -gt 0) { throw "$($script:failedTestsCount) tests failed."} else {       ls .\nuget | % { Push-AppveyorArtifact $_.FullName }}
    Write-Info "End Test Stage."
}

function RunTest {
param
(
    [string]$filePath, [Object[]] $CodeCoverage
    )
    Write-Info "Running tests: $filePath"
    $testResultsFile = "TestsResults.xml"
    $res = Invoke-Pester -Path $filePath -OutputFormat NUnitXml -OutputFile $testResultsFile -PassThru -CodeCoverage $CodeCoverage
    $webClient.UploadFile("https://ci.appveyor.com/api/testresults/nunit/$($env:APPVEYOR_JOB_ID)", (Resolve-Path $testResultsFile))
    Write-Info "Done running tests."
    return $res
}

function Write-Info([string]$message) {
    Write-Host -ForegroundColor Yellow  "[APPVEYOR] [$([datetime]::UtcNow)] $message"
}