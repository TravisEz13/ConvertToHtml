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
            $error.Count | should be 0
        }
    }
}
finally
{
    Suite.AfterAll
}

