<# 
Copyright Microsoft 2015

.summary
    Test that describes code itself.
#>

[CmdletBinding()]
param()

if (!$PSScriptRoot) # $PSScriptRoot is not defined in 2.0
{
    $PSScriptRoot = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
}
$RepoRoot = "$PSScriptRoot\.."

# import common for $RepoRoot variable

$ErrorActionPreference = 'stop'
Set-StrictMode -Version latest

Describe 'Text files formatting' {
    
    $allTextFiles = Get-ChildItem -File -Recurse $RepoRoot | ? { @('.gitignore', '.gitattributes', '.ps1', '.psm1', '.psd1', '.json', '.xml', '.cmd', '.mof') -contains $_.Extension } 
    
    Context 'Files encoding' {

        It "Doesn't use Unicode encoding" {
            $allTextFiles | %{
                $path = $_.FullName
                $bytes = [System.IO.File]::ReadAllBytes($path)
                $zeroBytes = @($bytes -eq 0)
                if ($zeroBytes.Length) {
                    Write-Warning "File $($_.FullName) contains 0 bytes. It's probably uses Unicode and need to be converted to UTF-8"
                }
                $zeroBytes.Length | Should Be 0
            }
        }
    }

    Context 'Indentations' {

        It 'We are using spaces for indentaion, not tabs' {
            $totalTabsCount = 0
            $allTextFiles | %{
                $fileName = $_.FullName
                $tabStrings = (Get-Content $_.FullName -Raw) | Select-String "`t" | % {
                    Write-Warning "There are tab in $fileName"
                    $totalTabsCount++
                }
            }
            $totalTabsCount | Should Be 0
        }
    }
}