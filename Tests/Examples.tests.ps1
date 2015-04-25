<# 
.summary
    Test suite for common examples
#>
[CmdletBinding()]
param()

Import-Module $PSScriptRoot\..\ConvertToHtml\ConvertToHtml.psd1 -Force

$examplesPath = "$PSScriptRoot\..\examples"
$ErrorActionPreference = 'stop'
Set-StrictMode -Version latest

function Suite.BeforeAll {
    # Remove any leftovers from previous test runs
    Suite.AfterAll 

}

function Suite.AfterAll {
}

function Suite.BeforeEach {
    if(!(test-path $examplesPath))
    {
        md $examplesPath > $null
    }
}


try
{
    Describe 'Process example' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }

            It 'Should not throw' {
                $processReportHtml = get-process |select-object -first 10 | ConvertTo-FormattedHtml 
                $processReportHtml | out-file -filepath $examplesPath\Process.html
            }
    }

    Describe 'Dir example' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }

            It 'Should not throw' {
                $processReportHtml = dir| Select-Object Mode, lastwritetime, length, Name | ConvertTo-FormattedHtml 
                $processReportHtml | out-file -filepath $examplesPath\dir.html
            }
    }

}
finally
{
    Suite.AfterAll
}

