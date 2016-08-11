# ConvertToHtml

PowerShell Module to convert PSObjects to formatted Html (targeted to Outlook)

[![Join the chat at https://gitter.im/TravisEz13/ConvertToHtml](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/TravisEz13/ConvertToHtml?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)
[![Build status](https://ci.appveyor.com/api/projects/status/j1vu2x67hxjmbtes/branch/master?svg=true)](https://ci.appveyor.com/project/TravisEz13/converttohtml/branch/master)
[![Stories in Ready](https://badge.waffle.io/TravisEz13/ConvertToHtml.png?label=ready&title=Ready)](https://waffle.io/TravisEz13/ConvertToHtml)
[![codecov.io](http://codecov.io/github/TravisEz13/ConvertToHtml/coverage.svg?branch=master)](http://codecov.io/github/TravisEz13/ConvertToHtml?branch=master)

[![codecov.io](http://codecov.io/github/TravisEz13/ConvertToHtml/branch.svg?branch=master)](http://codecov.io/github/TravisEz13/ConvertToHtml?branch=master)
WMF/PowerShell 5 Installation 
--------------------------------
From PowerShell run:

	Register-PSRepository -Name converttohtml -SourceLocation https://ci.appveyor.com/nuget/converttohtml
	Install-Module converttohtml -Scope CurrentUser

WMF/PowerShell 4 Installation
-----------------------------
 1. Download nuget.exe from [NuGet.org](https://nuget.org/nuget.exe) 
 2. &nuget.exe install ConvertToHtml -source https://ci.appveyor.com/nuget/converttohtml -outputDirectory "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\" -ExcludeVersion

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
