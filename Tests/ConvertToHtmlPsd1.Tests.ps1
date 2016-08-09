<# 
.summary
    Test suite for ConvertToHtmlPSD1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'stop'
Set-StrictMode -Version latest

function Suite.BeforeAll {
    # Remove any leftovers from previous test runs
    Suite.AfterAll 

}

function Suite.AfterAll {
}

function Suite.BeforeEach {
}

function Clear-TestDirectories {
}

try
{
    Suite.BeforeAll
    Describe 'Module Import' {
        It 'should not throw or have an error'{
            $error.Clear()
            {Import-Module "$PSScriptRoot\..\ConvertToHtml" -Force} | should not throw
            foreach($errorInstance in $error)
            {
                $errorInstance | Out-String | should be $null
            }
            $error.Count | should be 0
        }
        It 'should retain errors'{
            $error.Clear()
            Write-Error 'foo' -ErrorAction Continue 2> $null
            {Import-Module "$PSScriptRoot\..\ConvertToHtml" -Force} | should not throw
            $error[$error.Count -1].exception.message | should be 'foo'
        }
    }
}
finally
{
    Suite.AfterAll
}

