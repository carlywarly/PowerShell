Function Exportto-HTML {
    <#
    .SYNOPSIS
    Converts an object into a formatted HTML table 
    .DESCRIPTION
    Converts an object into a formatted HTML table. Alternate lines been white and grey and adds java to the html, so you can sort by clicking on the headers.   You can sort by string, date, IP and numbers.
    .PARAMETER Object
    The Object you what to export
    .PARAMETER Path
    Save the html as a file, otherwise return the html
    .PARAMETER NoSort
    Disable the sorting links, useful of pages with large tables
    .PARAMETER Title
    (Optional) Title of the page
    .PARAMETER Colour
    Colour of the table header, defaulted to Red
    .PARAMETER NoStyle
    (Optional) Whether to add sytle info, you only need one per webpage.
    .EXAMPLE
    # List services and save as a webpage
    Exportto-HTML (Get-Service | select DisplayName,Status,ServiceType) -Title "Services" -path .\services.html
    Start .\services.html
    .EXAMPLE
    # List services and hotfixes and just display html to screen   
    Get-Hotfix | select Description,HotFixID,InstalledOn | Exportto-HTML -Title "Hotfixes" 
    1.0 - Initial version
    #>
    [cmdletbinding()]
    param([parameter(ValueFromPipeline=$true,Mandatory=$false,Position=0)][psobject[]]$InObject,
        [parameter(ValueFromPipeline=$false,Mandatory=$False,Position=1)][string]$Path = "",
        [parameter(ValueFromPipeline=$false,Mandatory=$False,Position=2)][string]$Title = "",
        [parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$NoStyle = $False,
        [parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$NoSort = $False,
        [parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$List = $False
    )
    BEGIN {
           $Object = @()
    }
    Process {
           $Object += $InObject
    }
    End {
        Function FormatObject( $PSObject ) {
            If ($PSObject -ne $null) {
                $ret = @()
                Foreach ($line in $PSObject) {
                    $NewObject = New-Object –TypeName PSObject
                    foreach ($header in $line.psobject.properties) {
                        $Value = $header.value
                        if ($Value -ne $null) {
                            Switch ($Value.gettype().tostring()) {
                                "System.Boolean" { If ($Value) { $Value = "Yes" } else { $Value = "No" } }
                                "System.Array" { $Value = $Value -join ", " }
                                "System.Char[]" { $Value = $Value -join ", " }
                                "System.String[]" { $Value = $Value -join ", " }
                                "System.Int16" { $Value = "{0:N0}" -f $Value }
                                "System.Int16[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.UInt16" { $Value = "{0:N0}" -f $Value }
                                "System.UInt16[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.Int32" { $Value = "{0:N0}" -f $Value }
                                "System.Int32[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.UInt32" { $Value = "{0:N0}" -f $Value }
                                "System.UInt32[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.Int64" { $Value = "{0:N0}" -f $Value }
                                "System.Int64[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.UInt64" { $Value = "{0:N0}" -f $Value }
                                "System.UInt64[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
                                "System.Double" { $Value = "{0:N2}" -f $Value }
                                "System.Double[]" { $Value = ($Value | % { "{0:N2}" -f $_ }) -join ", " | Out-String }
                                "System.Single" { $Value = "{0:N2}" -f $Value }
                                "System.Single[]" { $Value = ($Value | % { "{0:N2}" -f $_ }) -join ", " | Out-String }                                
                                default { $Value = [string]$Value }
                            }
                        } else { $value = "" }
                        $HeaderName = (([regex]::replace(([regex]::replace($header.name,"[A-Z][a-z]+"," $&")),"[A-Z][A-Z]+"," $&")).replace("  "," ").trim())
                        $NewObject | Add-Member Noteproperty –Name $HeaderName –Value $Value
                    }
                    $ret += $NewObject 
                }
                $ret
            }
        }    

        If ($Object -ne $null) {
            $Object = FormatObject $Object
            
            [String]$HTML = ""
            If (!($NoStyle.IsPresent)) {
                $HTML += "<head><script type='text/javascript'>"
                $HTML += "`nvar Sort = 0;var LastCol = -1;"
                $HTML += "`nfunction sortTable(tableName,col) {"
                $HTML += "`n`tvar table, rows, switching, i, x, y, shouldSwitch;"
                $HTML += "`n`ttable = document.getElementById(tableName);"
                $HTML += "`n`tswitching = true;"
                $HTML += "`n`tif(LastCol == col) { Sort = Sort ^ 1; } else { LastCol = col; Sort = 0;}"
                $HTML += "`n`twhile (switching) {"
                $HTML += "`n`t`tswitching = false;"
                $HTML += "`n`t`trows = table.getElementsByTagName('TR');"        
                $HTML += "`n`t`tfor (i = 1;i < (rows.length - 1);i++) {"
                $HTML += "`n`t`t`tshouldSwitch = false;"
                $HTML += "`n`t`t`tx = rows[i].getElementsByTagName('TD')[col];"
                $HTML += "`n`t`t`ty = rows[i + 1].getElementsByTagName('TD')[col];"
                $HTML += "`n`t`t`txValue = (x.innerHTML.toLowerCase());"
                $HTML += "`n`t`t`tyValue = (y.innerHTML.toLowerCase());"
                $HTML += "`n`t`t`t if (((Sort == 0) && y.innerHTML.match('^[-+]?[0-9,]*\.?[0-9]+%?$')) || ((Sort == 1) && x.innerHTML.match('^[-+]?[0-9,]*\.?[0-9]+%?$'))) {"
                $HTML += "`n`t`t`t`txValue = parseInt(xValue.replace(',','').replace('%',''));"
                $HTML += "`n`t`t`t`tyValue = parseInt(yValue.replace(',','').replace('%',''));"
                $HTML += "`n`t`t`t}"
                $HTML += "`n`t`t`telse if ((x.innerHTML.match('^(0?[1-9]|[12][0-9]|3[01])[-\/.][0-9]{2}[-\/.][0-9]{2,4}( (0?[0-9]|1[0-9]|2[0-4]):[0-5][0-9]:[0-5][0-9])?$'))) {"
                $HTML += "`n`t`t`t`txValue = new Date(xValue.split('/')[1]+'-'+xValue.split('/')[0]+'-'+xValue.split('/')[2]);"
                $HTML += "`n`t`t`t`tyValue = new Date(yValue.split('/')[1]+'-'+yValue.split('/')[0]+'-'+yValue.split('/')[2]);"
                $HTML += "`n`t`t`t`tif (xValue == 'Invalid Date') { xValue = '' };"
                $HTML += "`n`t`t`t`tif (yValue == 'Invalid Date') { yValue = '' };"
                $HTML += "`n`t`t`t}"
                $HTML += "`n`t`t`telse if (x.innerHTML.match('^(?:(?:25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.){3}(?:25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])$')) {"
                $HTML += "`n`t`t`t`tvar IP1 = xValue.split('.');"
                $HTML += "`n`t`t`t`tvar IP2 = yValue.split('.');"
                $HTML += "`n`t`t`t`txValue = (IP1[0]*16777216) + (IP1[1]*65536) + (IP1[2]*256) + IP1[3]*1;"
                $HTML += "`n`t`t`t`tyValue = (IP2[0]*16777216) + (IP2[1]*65536) + (IP2[2]*256) + IP2[3]*1;"
                $HTML += "`n`t`t`t}"
                $HTML += "`n`t`t`tif(Sort == 0) {"
                $HTML += "`n`t`t`t`tif (xValue > yValue) {shouldSwitch= true;break;}"
                $HTML += "`n`t`t`t} else {"
                $HTML += "`n`t`t`t`tif (xValue < yValue) {shouldSwitch= true;break;}"
                $HTML += "`n`t`t`t}"
                $HTML += "`n`t`t}"
                $HTML += "`n`t`tif (shouldSwitch) {rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);switching = true;}"
                $HTML += "`n`t}`t"
                $HTML += "`n}"
                $HTML += "`n</script>"
                $HTML += "`n<style>"
                $HTML += "`ntable {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
                $HTML += "`nth {cursor: pointer;border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:$Colour;color:white;text-align:left;padding-left:8px;padding-right:8px;font-weight:bold}"
                $HTML += "`nth.sort-by {padding-right: 18px;position: relative;}"
                $HTML += "`nth.sort-by:before,th.sort-by:after {border: 4px solid transparent;content: '';display: block;height: 0;right: 5px;top: 50%;position: absolute;width: 0;}"
                $HTML += "`nth.sort-by:before {border-bottom-color: white;margin-top: -9px;}"
                $HTML += "`nth.sort-by:after {border-top-color: white;margin-top: 1px;}"
                $HTML += "`ntr:nth-child(even) {background-color: #dddddd;}"
                $HTML += "`ntd {border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align:left;padding-left:8px;padding-right:8px}"
                $HTML += "`n</style>`n"
            }            
            # Convert  to HTML
            If ($Title -ne $null) { $html += "<H2>$Title</H2><title>$Title</title>`n" } 
            If ($List.IsPresent) {
                # As List
                $HTML += $Object | ConvertTo-Html -Fragment -as List | out-string       
                $HTML = $HTML.Replace("&" + "lt;","<").Replace("&" + "gt;",">").Replace("&" + "quot;",'"').Replace("&" + "amp;",'&')
                $HTML = $HTML.replace("<tr><td>",'<tr><th>').replace(":</td><td>","</th><td>") 
                # Remove mouse pointer from style
                $HTML = $HTML.replace("cursor: pointer;","")
            } else {
                # As Table
               $HTML += $Object | ConvertTo-Html -Fragment -as Table | out-string       
                $HTML = $HTML.Replace("&" + "lt;","<").Replace("&" + "gt;",">").Replace("&" + "quot;",'"').Replace("&" + "amp;",'&')
                # Right justifiy any cells with numbers, can contain comma and/or %
                $HTML = [regex]::replace($HTML,">[-+]?[0-9,]*\.?[0-9]+%?<",' style="text-align:right"$&') 
                # Right justifiy any cells with binary prefixes 
                $HTML = [regex]::replace($HTML,">(?i:[0-9,\.]+\s?[KMGTPEZY]B)<",' style=" text-align:right"$&') 
                # Right justifiy n/a
                $HTML = [regex]::replace($HTML,">[nN]\/[aA]<\/td>",' style=" text-align:right"$&') 
                #Find first column header
                $pos = $HTML.IndexOf("<th>")
                $col = 0
                If ($NoSort.IsPresent) {
                    # Remove mouse pointer from style
                    $HTML = $HTML.replace("cursor: pointer;","")                    
                } else {
                    # Add hyperlinks to sort table
                    While($pos -ne -1) {
                        $HTML = $HTML.Insert($pos + 1,"#")
                        # Replace column header to point to sort function
                        $HTML = $HTML -replace "<#th>",('<th class="sort-by" onclick="sortTable(''' + $title + ''',' + $col + ')">')
                        $col = $col + 1
                        $pos = $HTML.IndexOf("<th>")
                    }
                }
            }  
            $HTML = $HTML.Replace("<table>",'<table id="' + $title + '">')
            $HTML = $HTML + "</body>"
        }
        If ($path -ne "") { 
            $HTML | Set-Content $path 
        } else {
            $HTML 
        }
    }
}