[string] $moduleDir = Split-Path -Path $script:MyInvocation.MyCommand.Path –Parent

Set-StrictMode -Version latest


<#
.Synopsis
    Converts the array of objects to a formatted HTML table.
#>
function ConvertTo-FormattedHtml
{
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [PSObject[]]$InputObject = $null,
        [string]$Title = $null,
        [switch]$OutClipboard = $false,
        [switch]$bodyOnly = $false,
        [string]$formatJson = $null
)
    Begin
    {
        [PSObject[]] $allInput = @();
    }

    process
    {
        [HashTable] $exportHtmlParams = @{bodyOnly = $bodyOnly}
        foreach($item in $InputObject)
        {
            $allInput += $item;
        }
    }

    end
    {

        if ($allInput[0] -is [PSObject])
        {
            [bool] $gotformatJson = $false
            if($formatJson.Length -gt 1)
            {
                [string] $formatObjectJson = $formatJson
                [bool] $gotformatJson = $true
            }
            else
            {
                [string] $formatObjectJson = Find-FormatJsonFromFile -allInput $allInput
                if($formatObjectJson)
                {
                    $gotformatJson = $true
                }
            }

            if($gotformatJson)
            {
                $formatJsonTables = Convert-FormatObjectJson -formatObjectJson $formatObjectJson

                [string] $result = Export-Html -InputObject $allInput -Property $formatJsonTables.Property -GroupBy $formatJsonTables.GroupBy -GroupByHeading $formatJsonTables.GroupByHeading -ColumnHeadings $formatJsonTables.columnHeadings -ColumnBackgroundColors $formatJsonTables.columnBackgroundColor @exportHtmlParams
                   
                if($OutClipboard)
                {
                    $result | Out-Clipboard
                }
                else
                {
                    $result
                }
            }
            else
            {
                Write-Verbose -Message 'Json missing, using properites...'

                [string] $result = Export-Html -InputObject $allInput -Property (Get-Properties -allInput $allInput) -Heading $title @exportHtmlParams
                if($OutClipboard)
                   {
                        $result | Out-Clipboard
                   }
                   else
                   {
                        $result
                   }
            }
        }
    }
}

<#
.Synopsis
    Converts a Format Json string to a PSObject
#>
function Convert-FormatObjectJson
{

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0, HelpMessage='Please add a help message here')]
        [Object]
        $formatObjectJson
    )    
    
    [PSObject] $formatObject = $formatObjectJson | ConvertFrom-Json
    
    [HashTable] $columnHeadings = @{}
    foreach($columnHeading in $formatObject.ColumnHeadings)
    {
        $columnHeading.PSObject.Members| ForEach-Object -Process {
            if($_.MemberType -eq 'NoteProperty') {
                $columnHeadings.Add($_.Name,$_.Value)
            }
        }
    }
    
    [HashTable] $columnBGColors = @{}
    foreach($columnHeading in $formatObject.ColumnBackgroundColor)
    {
        $columnHeading.PSObject.Members| ForEach-Object -Process {
            if($_.MemberType -eq 'NoteProperty') {
                $columnBGColors.Add($_.Name,$_.Value)
            }
        }
    }

    $returnValue = (New-Object -TypeName PSObject -Property @{
        ColumnHeadings = $columnHeadings
        ColumnBackgroundColor = $columnBGColors
        Property = $formatObject.Property
        GroupBy = $formatObject.GroupBy
        GroupByHeading = $formatObject.GroupByHeading

    })
    $returnValue.pstypenames.clear()
    $returnValue.pstypenames.add('ConvertToHtml.FormatTables')
    return $returnValue
}



<#
.Synopsis
    Gets the formatJson from a file
#>
function Find-FormatJsonFromFile
{
    param
    (
        [Parameter(Mandatory=$true)]
        [Object[]]
        $allInput
    )
    
    
    [int] $i = 0
    [string] $typenameFormatFilePath = $null
    do{
        $tempPath = (Join-Path -Path $moduleDir -ChildPath ('ExportHtml.' + $allInput[0].PSObject.TypeNames[$i] + '.json'))
        Write-Verbose -Message "looking for $tempPath"
        if (Test-Path -Path $tempPath)
        {
            $typenameFormatFilePath = $tempPath
            [string] $formatObjectJson = Get-Content -raw -Path $typenameFormatFilePath
            return $formatObjectJson
        }
        $i++
    }
    while($i -lt $allInput[0].PSObject.TypeNames.Count -and $null -eq $typenameFormatFilePath)
    return $null
}

<#
.Synopsis
    Creates a formatting Json for the array of objects.
#>
function New-FormattedHtmlJson
{
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [PSObject[]]$InputObject = $null
)
    Begin
    {
        [PSObject[]] $allInput = @();
    }

    process
    {
        foreach($item in $InputObject)
        {
            $allInput += $item;
        }
    }

    end
    {

        if ($allInput[0] -is [PSObject])
        {
                Write-Verbose -Message 'Creating Formatting Json'

                $formattingTables = Get-DefaultFormattingTables -allInput $allInput
                [HashTable] $columnHeadings = $formattingTables.ColumnHeadings
                [HashTable] $ColumnBackgroundColor= $formattingTables.ColumnBackgroundColor
                [string] $thisTypeName = $formattingTables.thisTypeName
                $properties = $formattingTables.Properties

                    Write-Verbose -Message 'Creating Json'
                    [HashTable] $jsonProperties = @{
                        'TypeName'=$thisTypeName
                        'Property'=$properties
                        'GroupBy'=$null;
                        'GroupByHeading'=$null;
                        'ColumnHeadings'= $columnHeadings;
                        'Heading' = $thisTypeName;
                        'ColumnBackgroundColor' = $ColumnBackgroundColor;
                    }
                    [PSObject] $jsonObject = New-Object –TypeName PSObject –Prop $jsonProperties 
                    $jsonObject.PSObject.TypeNames[0] = 'Export.Html.Format'
                    #[string] $filename="ExportHtml.$thisTypeName.Json"
                    $jsonObject | ConvertTo-Json | Write-Output
                    #Write-Warning -Message "Created missing Json: $filename"
            
        }
    }
}

<#
.Synopsis
    Creates the tables and arrays needed for various formatting functions, if we don't already have a formatting json
#>
function Get-DefaultFormattingTables
{
    param
    (
        [object[]] $allInput
    )

    [string[]] $properties = Get-Properties -allInput $allInput
    [HashTable] $columnHeadings = @{}
    foreach($property in $properties)
    {
        $columnHeadings.Add($property,$property)
    }
    
    [HashTable] $ColumnBackgroundColor=@{}
    $ColumnBackgroundColor.Add($properties[0], '#switch ($columnValue) { default { write-Output "#EE0000"} 0 { write-Output return}}  # you can also use $this, which is the current object');
    
    [string] $thisTypeName = $allInput[0].PSObject.TypeNames[0]

    return @{
            Properties = $properties
            ColumnHeadings = $columnHeadings
            ColumnBackgroundColor = $ColumnBackgroundColor
            thisTypeName = $thisTypeName
        }
}

<#
.Synopsis
    Get the properties for the first item out of the array of objects.
#>
function Get-Properties
{
    param
    (
        [object[]] $allInput
    )
    
    [string[]] $properties = @()
    $allInput[0] | Get-Member -membertype 'properties'| ForEach-Object -Process {$properties += $_.Name}
    return $properties
}

function Get-HtmlEncodedValue
{
    param(
        [string]$value
    )

    return $value.replace('<', '&lt;').replace('>','&gt;')
}


<#
.Synopsis
    Converts the array of objects to a formatted HTML table.
#>
function Export-Html
{
    param(
        [parameter(Mandatory=$true)]
        [Object[]]$InputObject, 
        [parameter(Mandatory=$true)]
        [Object[]] $Property,
        [Object] $GroupBy = $null,
        [Object] $GroupByHeading = $null,
        [String] $Heading = $null,
        [System.Collections.Hashtable] $ColumnHeadings = $null,
        [System.Collections.Hashtable] $ColumnBackgroundColors = $null,
        [switch] $bodyOnly
    )
    [string] $headingStyle='width:195.8pt;border-top:solid black 1.0pt;border-left:none;border-bottom:solid #4F81BD 1.5pt;border-right:none;background:#4F81BD;padding:0in 5.4pt 0in 5.4pt;height:20.25pt'
    [string] $greyStyle='width:195.8pt;border:none;border-bottom:solid #A7BFDE 1.5pt;background:#D9D9D9;padding:0in 5.4pt 0in 5.4pt;height:18.75pt'
    [string] $whiteStyle='width:195.8pt;border:none;border-bottom:solid #A7BFDE 1.5pt;padding:0in 5.4pt 0in 5.4pt;height:18.75pt'

    [string] $subject = $Heading
    
    [System.Text.StringBuilder] $sb = New-Object -TypeName 'System.Text.StringBuilder'
    if (-not $bodyOnly)
    {
        $sb.AppendLine("<!DOCTYPE html PUBLIC `"-//W3C//DTD XHTML 1.0 Transitional//EN`" `"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd`">")  > $null
        $sb.AppendLine('<html><head/><body>')  > $null
    }

    $sb.AppendFormat('<h1>{0}</h1>',$subject)  > $null
#    $trash=$sb.AppendLine("<h2>Summary</h2>");
#    $trash=$sb.AppendLine("[Add summary]");
#    $trash=$sb.AppendLine("<h2>Details</h2>");
    $sb.AppendLine('<table>')  > $null

    [string] $lastGroup = [string]::Empty

    [int] $rowCount = 0
    [bool] $firstGroup = $true;
    $InputObject | ForEach-Object -Process {
        [string] $color = '#C6EFCE' #lightGreen
        #if($_.State -eq 'Warning')
        #{$color='#FFEB9C'} #lightYellow

        #if($_.State -eq 'Error')
        #{$color='red'}

        if ($GroupBy)
        {
            if ($_.$GroupBy -ne $lastGroup)
            {
                if ($GroupByHeading)
                {
                    [string] $effectiveGroupByHeading = $GroupByHeading + ':  ' + $_.$GroupBy
                }
                else
                {
                    [string] $effectiveGroupByHeading = $GroupBy + ':  ' + $_.$GroupBy
                }

                [int] $columns = $Property.Count
                [int] $rowCount = 0
                if(-not $firstGroup)
                {
                    $sb.AppendLine("<tr><td colspan='$columns' >&nbsp</td></tr>")  > $null
                }

                $firstGroup = $false
                
                $sb.AppendFormat("<tr><th colspan='$columns' style='$headingStyle'>{0}</th></tr>", $effectiveGroupByHeading)  > $null
                $sb.AppendLine("<tr style='$greyStyle;font-size:13.0pt;color:#1F497D'>")  > $null
                foreach($propertyName in $Property)
                {
                    [string] $headingName = Get-HeadingName -PropertyName $propertyName -ColumnHeadings $ColumnHeadings
                    $sb.AppendLine("<th>$headingName</th>") > $null
                }
                $sb.AppendLine('</tr>')  > $null
                $lastGroup = $_.$GroupBy
            }
        }
        else
        {
            if($firstGroup)
            {
                $sb.AppendLine("<tr style='$greyStyle;font-size:13.0pt;color:#1F497D'>")  > $null
                foreach($propertyName in $Property)
                {
                    [string] $headingName = Get-HeadingName -PropertyName $propertyName -ColumnHeadings $ColumnHeadings
                    $sb.AppendLine("<th>$headingName</th>")  > $null
                }
                $sb.AppendLine('</tr>')  > $null
                $firstGroup = $false
            }
        }

        [string] $currentStyle = $greyStyle
        if(($rowCount %2) -eq 0)
        {
            $currentStyle=$whiteStyle
        }

        $sb.AppendLine("<tr style='$currentStyle'>")  > $null
        try
        {
            Set-StrictMode -Off
            foreach($propertyName in $Property)
            {
                [string] $columnValue = Get-HtmlEncodedValue -value $_.$propertyName
                [string] $style = Get-BackgroundColorStyle -ColumnValue $columnValue -PropertyName $propertyName -ColumnBackgroundColor $ColumnBackgroundColors -this $_
                $sb.AppendLine("<td style='$style'>$columnValue</td>")  > $null
            }
        }
        finally
        {
            Set-StrictMode -Version latest
        }
        $sb.AppendLine('</tr>')  > $null
        $sb.AppendLine([string]::Empty)  > $null
        $rowCount++
    }
    $sb.AppendLine('</table>')  > $null

    if (-not $bodyOnly)
    {
        $sb.AppendLine('</body>')  > $null
        $sb.AppendLine('</html>')  > $null
    }

    ($sb.ToString()) | Write-Output
    Write-Verbose -Message "Finished Export-Html for $rowCount rows."
} 


<#
.Synopsis
    Gets the background style for a propertyname and value combination
#>
function Get-BackgroundColorStyle
{
    param (
        $columnValue,
        
        [string] 
        $propertyName ,
        
        [System.Collections.Hashtable]
        $ColumnBackgroundColor,
        
        $this
    )
    if($ColumnBackgroundColor)
    {
        if ($ColumnBackgroundColor.ContainsKey($propertyName))
        {
            [scriptblock] $colorScriptBlock = $executioncontext.invokecommand.NewScriptBlock($ColumnBackgroundColor[$propertyName])
            Write-Verbose -Message "colorScriptBlock: $colorScriptBlock"
            [string] $color = Invoke-Command -ScriptBlock $colorScriptBlock
            if($color)
            {
                "background-color:$color" | Write-Output 
                return
            }
        }
    }

 }


<#
.Synopsis
    Gets the heading name for a property
#>
function Get-HeadingName
{
    param ( 
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $propertyName, 

            [parameter(Mandatory=$true)]
            [AllowNull()]
            [System.Collections.Hashtable]
            $ColumnHeadings
        )

    if($ColumnHeadings)
    {
        if ($ColumnHeadings.ContainsKey($propertyName))
        {
            $ColumnHeadings[$propertyName] | Write-Output 
            return
        }
    }

    Write-Verbose -Message "Didn't find Heading name for $propertyName, using propertyName"
    $propertyName | Write-Output    
}


<#
.Synopsis
    Puts HTML in the browser
#>
function Out-Browser
{
    param ( 
    [parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string] $InputObject)

    Begin
    {
        [System.Text.StringBuilder] $sb = New-Object -TypeName 'System.Text.StringBuilder';
    }
    Process
    {
        $sb.Append($InputObject) > $null
    }
    End
    {
        [string] $filename = [System.IO.Path]::GetRandomFileName() + '.html'
        Write-Verbose -Message "sending temp file $filename to browser."
        $filename = join-path -path $env:TEMP -ChildPath $filename
        ($sb.ToString()) | Out-File -FilePath $filename
        Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', 'start', $filename)
    }
}

Add-Type -Assembly @('PresentationCore')

<#
.Synopsis
    Puts HTML on the clipboard
#>
function Out-Clipboard
{
    param ( 
    [parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string] $InputObject,
    [Windows.TextDataFormat] $format=[Windows.TextDataFormat]::Html
    )

    Begin
    {
        [System.Text.StringBuilder] $sb = New-Object -TypeName 'System.Text.StringBuilder' 
    }
    Process
    {
        $sb.Append($InputObject) > $null
    }
    End
    {
        Write-Verbose -Message 'Sending to clipboard...'
        [string] $body = $sb.ToString()
        [string] $cfHtml=Get-CF_Html -html $body
        Clear-Clipboard

        $cfHtml|powershell.exe -NoProfile -STA -Command {
            Add-Type -Assembly PresentationCore

            $clipText = ($input | Out-String -Stream)
 
            ## And finally set the clipboard text
            #[Windows.Clipboard]::SetText($input,[Windows.TextDataFormat]::UnicodeText)
            #Write-Host $clipText
            [Windows.Clipboard]::SetText($clipText, [Windows.TextDataFormat]::Html)
        }
    }

}


<#
.Synopsis
    Converts HTML to the text expected on the clipboard
#>
function Get-CF_Html
{
    param([string]$html)

    #adding for script analyzer
    Write-Verbose -Message 'in get-cf_html'

    [string] $cfHtmlFormat = @"
Version:{0}
StartHTML:{1}
EndHTML:{2}
StartFragment:{3}
EndFragment:{4}
{5}
"@
    # Make sure header is 89 characters
    $totalLengthOfFields = 6

    $headerLength = 62 + 3 + ($totalLengthOfFields * 4)
    $startHtml = $headerLength 
    $startHtmlString = Format-Number -totalLength $totalLengthOfFields -value $startHtml

    $startFragment= Format-Number -totalLength $totalLengthOfFields -value $startHtml
    #$startSelection = Format-Number -totalLength $totalLengthOfFields -value $startHtml
    $endHtmlPosition = $startHtml + $html.Length
    $endHtml = Format-Number -totalLength $totalLengthOfFields -value $endHtmlPosition
    $endFragment = Format-Number -totalLength $totalLengthOfFields -value $endHtmlPosition
    #$endSelection = Format-Number -totalLength $totalLengthOfFields -value $endHtmlPosition

    $cfHtml=[String]::Format($cfHtmlFormat,'0.9',$startHtmlString,$endHtml,$startFragment,$endFragment,$html)

    Write-Debug $cfHtml
    $cfHtml | Write-Output 
}


<#
.Synopsis
    formats a number with leading zeros to a certain length
#>
function Format-Number
{
    param($value, 
    $totalLength)

    #TODO: make use string builder
        [string] $text = $value.ToString();
        while($text.Length -lt $totalLength)
        {
            $text = '0' + $text
        }
        $text | Write-Output 
}


<#
.Synopsis
    clears the clipboard
#>
function Clear-Clipboard
{
    param ( 
    )

    End
    {
        Write-Verbose -Message 'Clearing clipboard...'

        powershell.exe -NoProfile -STA -Command {
            Add-Type -Assembly PresentationCore
 
            [Windows.Clipboard]::Clear()
        }
    }

}