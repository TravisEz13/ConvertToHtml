[string] $moduleDir = Split-Path -Path $script:MyInvocation.MyCommand.Path -Parent

Set-StrictMode -Version latest
$webClient = New-Object 'System.Net.WebClient';
$repoName = ${env:APPVEYOR_REPO_NAME}
$pullRequestTitle = ${env:APPVEYOR_PULL_REQUEST_TITLE}
Function Invoke-AppveyorInstall
{
    Write-Info 'Starting Install stage...'
    Write-Info "Repo: $repoName"
    if($pullRequestTitle)
    {
        Write-Info "Pull Request:  $pullRequestTitle"    
    }

    $nuGetPath = "$env:SystemDrive\nuget.exe"
    $webClient.DownloadFile( 'https://oneget.org/nuget-anycpu-2.8.3.6.exe', $nuGetPath )  
    Write-Info 'End Install stage.'

}

Function Invoke-AppveyorBuild
{
    Write-Info 'Starting Build stage...'

    $nuGetPath = "$env:SystemDrive\nuget.exe"
    mkdir -force .\out > $null
    mkdir -force .\nuget > $null
    mkdir -force .\examples > $null

    Update-ModuleVersion

    
    if($repoName -ieq 'master' -and [string]::IsNullOrEmpty($pullRequestTitle))
    {
        $moduleName = 'ConvertToHtml'
    }
    else
    {
        $moduleName = "ConvertToHtml-$repoName"
    }
    Update-Nuspec -ModuleName $moduleName

    Write-Info 'Creating nuget package ...'
    &$nuGetPath pack .\ConvertToHtml\ConvertToHtml.nuspec -outputdirectory  .\nuget

    Write-Info 'Creating module zip ...'
    7z a -tzip .\out\ConvertToHtml.zip .\ConvertToHtml\*.*

    Write-Info 'End Build Stage.'
}

Function Invoke-AppveyorTest
{
    Write-Info 'Starting Test stage...'
    # setup variables for the whole build process
    #
    $script:failedTestsCount = 0
    #

    $CodeCoverage = @('.\ConvertToHtml\exporttohtml.psm1')
    '.\tests' | %{ 
    $res = Invoke-RunTest -filePath $_ -CodeCoverage @('.\ConvertToHtml\exporttohtml.psm1')
    $script:failedTestsCount += $res.FailedCount 
    $CodeCoverageTitle = 'Code Coverage {0:F1}%'  -f (100 * ($res.CodeCoverage.NumberOfCommandsExecuted /$res.CodeCoverage.NumberOfCommandsAnalyzed))
    $res.CodeCoverage.MissedCommands | ConvertTo-FormattedHtml -title $CodeCoverageTitle | out-file .\examples\CodeCoverage.html
    }

    Write-Info 'Creating Example zip ...'
    7z a -tzip .\out\examples.zip .\examples\*.html

    if ($script:failedTestsCount -gt 0) 
    { 
        throw "$($script:failedTestsCount) tests failed."
    } 
    else 
    {       
        Get-ChildItem .\nuget | % { 
                Push-AppveyorArtifact $_.FullName 
            }
    }
    Write-Info 'End Test Stage.'
}

function Invoke-RunTest {
    param
    (
        [string]
        $filePath, 
        
        [Object[]] 
        $CodeCoverage
    )
    Write-Info "Running tests: $filePath"
    $testResultsFile = 'TestsResults.xml'
    $res = Invoke-Pester -Path $filePath -OutputFormat NUnitXml -OutputFile $testResultsFile -PassThru -CodeCoverage $CodeCoverage
    $webClient.UploadFile("https://ci.appveyor.com/api/testresults/nunit/$($env:APPVEYOR_JOB_ID)", (Resolve-Path $testResultsFile))
    Write-Info 'Done running tests.'
    return $res
}

function Write-Info {
     param
     (
         [string]
         $message
     )

    Write-Host -ForegroundColor Yellow  "[APPVEYOR] [$([datetime]::UtcNow)] $message"
}

function Update-ModuleVersion
{
    Write-Info 'Updating Module version to: ${env:APPVEYOR_BUILD_VERSION}'
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

function Update-Nuspec
{
    param
    (
        [string]
        $ModuleName = 'ConvertToHtml'
    )

    Write-Info 'Updating Module version in nuspec...'
    [xml]$xml = Get-Content -Raw .\ConvertToHtml\ConvertToHtml.nuspec
    $xml.package.metadata.version = $env:APPVEYOR_BUILD_VERSION
    $xml.OuterXml | out-file -FilePath .\ConvertToHtml\ConvertToHtml.nuspec
}


}

