<# 
.summary
    Test suite for ExportToHtml.psm1
#>
[CmdletBinding()]
param()

Import-Module $PSScriptRoot\..\ConvertToHtml\ExportToHtml.psm1 -Force

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

    Describe 'Get-BackgroundColorStyle' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }


            It 'Should return null when property name doesnt exist in table' {
                [HashTable] $hashtable = @{foo={write-output 'aou'}}
                Get-BackgroundColorStyle -columnValue 'foo' -propertyName 'bar' -columnBackgroundColor $hashtable -this $null| should be $null
            }
            It 'Should return a static color' {
                [HashTable] $hashtable = @{foo={write-output 'testcolor'}}
                Get-BackgroundColorStyle -columnValue 'foo' -propertyName 'foo' -columnBackgroundColor $hashtable -this $null| should be "background-color:testcolor"
            }
            It 'Should return using this' {
                [HashTable] $hashtable = @{foo={write-output $this.bar}}
                [HashTable] $hashtable2 = @{bar='testcolor2'}
                Get-BackgroundColorStyle -columnValue 'foo' -propertyName 'foo' -columnBackgroundColor $hashtable -this $hashtable2| should be "background-color:testcolor2"
            }
            It 'Should return using columnValue' {
                [HashTable] $hashtable = @{foo={ if($columnValue -eq 'foo') {write-output 'testcolor3'} else {write-output 'fail'}}}
                Get-BackgroundColorStyle -columnValue 'foo' -propertyName 'foo' -columnBackgroundColor $hashtable -this $null| should be "background-color:testcolor3"
            }
    }


    Describe 'Get-HeadingName' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }


            It 'Should return propertyName when property name doesnt exist in table' {
                [HashTable] $hashtable = @{foo='bar'}
                Get-HeadingName -propertyName 'foo2' -ColumnHeadings $hashtable | should be 'foo2'
            }
            It 'Should return heading from table' {
                [HashTable] $hashtable = @{foo='bar'}
                Get-HeadingName -propertyName 'foo' -ColumnHeadings $hashtable | should be 'bar'
            }
    }

    Describe 'Format-Number' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }

            It 'value 1 length 3 should return 001' {
                Format-Number -value 1 -totalLength 3 | should be '001'
            }
            It 'value 28 length 2 should return 28' {
                Format-Number -value 28 -totalLength 2 | should be '28'
            }
            It 'value 28 length 1 should return 28' {
                Format-Number -value 28 -totalLength 1 | should be '28'
            }
    }

    Describe 'Get-CF_Html' {
        BeforeEach {
            Suite.BeforeEach
        }

        AfterEach {
        }

        $headerLength = 89

        $html = 'foo'
        $result = (Get-CF_Html -html $html)
        [Byte[]] $buffer = [System.Text.UnicodeEncoding]::Unicode.GetBytes($result)
        $memStream = New-Object -TypeName 'System.IO.MemoryStream' 
        $memStream.Write($buffer, 0, $buffer.length)
        $memStream.Position = 0
        $lines=@()
        try {
            $streamReader = New-Object -TypeName 'System.IO.StreamReader' -ArgumentList @($memStream, [System.Text.UnicodeEncoding]::Unicode)
            while(!$streamReader.EndOfStream)
            {
                $lines += $streamReader.ReadLine()
            }
        }
        finally
        {
            $memStream.Close()
        }

        $end = $headerLength + $html.length
        if($env:APPVEYOR -ne 'True')
        {
            # Don't run these on appveyor due to this issue:
            # https://github.com/TravisEz13/ConvertToHtml/issues/2

            It "Should not be longer than end length" {
                $result.length | should be $end
            }
            It "Should add header of length $headerLength"{
                ($result.length - $html.length) | should be $headerLength
            }
        }
        It 'First Header line should be version 0.9'{
            $lines[0] | should be "Version:0.9"
        }
        It "Second header line should be StartHTML:0000$headerLength"{
            $lines[1] | should be "StartHTML:0000$headerLength"
        }
        It "Third header line should be EndHTML:0000$end"{
            $lines[2] | should be "EndHTML:0000$end"
        }
        It "Forth Header line should be StartFragment:0000$headerLength"{
            $lines[3] | should be "StartFragment:0000$headerLength"
        }
        It "Fifth header line should be EndFragment:0000$end"{
            $lines[4] | should be "EndFragment:0000$end"
        }
        It "Should contain html fragment" {
            $result.EndsWith($html) | should be $true
        }
    }
}
finally
{
    Suite.AfterAll
}
