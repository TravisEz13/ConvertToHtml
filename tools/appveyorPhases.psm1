[string] $moduleDir = Split-Path -Path $script:MyInvocation.MyCommand.Path -Parent

Import-Module PoshBuildTools
Set-StrictMode -Version latest
$webClient = New-Object 'System.Net.WebClient';
$repoName = ${env:APPVEYOR_REPO_NAME}
$branchName = $env:APPVEYOR_REPO_BRANCH
$pullRequestTitle = ${env:APPVEYOR_PULL_REQUEST_TITLE}

$script:expectedModuleCount = 1
$moduleInfoList = @()
$moduleInfoList += New-BuildModuleInfo -ModuleName 'ConvertToHtml' -ModulePath '.\ConvertToHtml' -CodeCoverage @('.\ConvertToHtml\exporttohtml.psm1') -Tests @('.\tests')
$script:moduleBuildCount = 0
$script:failedTestsCount = 0
$script:passedTestsCount = 0

Function Invoke-AppveyorInstall
{
    Write-Info 'Starting Install stage...'
    Write-Info "Repo: $repoName"
    Write-Info "Branch: $branchName"
    if($pullRequestTitle)
    {
        Write-Info "Pull Request:  $pullRequestTitle"    
    }

    Install-NugetPackage -package pester

    Write-Info 'End Install stage.'
}

Function Invoke-AppveyorBuild
{
    Write-Info 'Starting Build stage...'
    mkdir -force .\out > $null
    mkdir -force .\nuget > $null
    mkdir -force .\examples > $null

    foreach($moduleInfo in $moduleInfoList)
    {

        $ModuleName = $moduleInfo.ModuleName
        $ModulePath = $moduleInfo.ModulePath
        if(test-path $modulePath)
        {
            Update-ModuleVersion -modulePath $ModulePath -moduleName $moduleName
            
            Update-Nuspec -modulePath $ModulePath -moduleName $ModuleName

            Write-Info 'Creating nuget package ...'
            nuget pack "$modulePath\${ModuleName}.nuspec" -outputdirectory  .\nuget

            Write-Info 'Creating module zip ...'
            7z a -tzip ".\out\$ModuleName.zip" ".\$ModuleName\*.*"

            $script:moduleBuildCount ++
        }
        else 
        {
            Write-Warning "Couldn't find module, $ModuleName at $ModulePath.."
        }
    }
    Write-Info 'End Build Stage.'
}

Function Invoke-AppveyorTest
{
    Write-Info 'Starting Test stage...'
    # setup variables for the whole build process
    #
    #
    Import-Module .\ConvertToHtml -force
    
    foreach($moduleInfo in $moduleInfoList)
    {
        $ModuleName = $moduleInfo.ModuleName
        $ModulePath = $moduleInfo.ModulePath
        $ModulePath = $moduleInfo.ModulePath
        if(test-path $modulePath)
        {
            $CodeCoverage = $moduleInfo.CodeCoverage
            $tests = $moduleInfo.Tests
            $tests | %{ 
                $res = Invoke-RunTest -filePath $_ -CodeCoverage $CodeCoverage
                $script:failedTestsCount += $res.FailedCount
                $script:passedTestsCount += $res.PassedCount 
                $CodeCoverageTitle = 'Code Coverage {0:F1}%'  -f (100 * ($res.CodeCoverage.NumberOfCommandsExecuted /$res.CodeCoverage.NumberOfCommandsAnalyzed))
                $res.CodeCoverage.MissedCommands | ConvertTo-FormattedHtml -title $CodeCoverageTitle | out-file .\out\CodeCoverage.html
            }
        }
    }

    Write-Info 'Creating Example zip ...'
    7z a -tzip .\out\examples.zip .\examples\*.html

    if ($script:failedTestsCount -gt 0) 
    { 
        throw "$($script:failedTestsCount) tests failed."
    } 
    elseif($script:passedTestsCount -eq 0)
    {
        throw "no tests passed"
    }
    elseif($script:moduleBuildCount -ne $script:expectedModuleCount)
    {
        throw "built ${script:moduleBuildCount} modules, but expected ${script:expectedModuleCount}"
    }
    else 
    {       
        if($branchName -ieq 'master' -and [string]::IsNullOrEmpty($pullRequestTitle))
        {
        Get-ChildItem .\nuget | % { 
                    Write-Info "Pushing nuget package $_.Name to Appveyor"
                    Push-AppveyorArtifact $_.FullName
            }
        }
        else 
        {
            Write-Info 'Skipping nuget package publishing because the build is not for the master branch or is a pull request.'
        }
    }
    Write-Info 'End Test Stage.'
}




