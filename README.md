# ConvertToHtml
PowerShell Module to convert PSObjects to formatted Html (targeted to Outlook)

I plan to publish to PowerShell Gallery, but currently to install:

 1. Clone the repo, 
 2. Copy the ConvertToHtml folder to C:\Program
    Files\WindowsPowerShell\Modules

[![Build status](https://ci.appveyor.com/api/projects/status/j1vu2x67hxjmbtes/branch/master?svg=true)](https://ci.appveyor.com/project/TravisEz13/converttohtml/branch/master)

Examples/Testing:
-----------------

If you make changes, please run the pester tests, and run these two tests:

    Get-Process | ConvertTo-FormattedHtml -OutClipboard

paste the results into outlook and verified they are formatted correctly

    dir| Select-Object Mode, lastwritetime, length, Name | ConvertTo-FormattedHtml -OutClipboard

paste the results into outlook and verified they are formatted correctly

Issues
------
For HTML to work in Outlook 2013 and older, only inline styles can be used.