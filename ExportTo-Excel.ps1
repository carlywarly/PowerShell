Function ExportTo-Excel {
    <# 
    .SYNOPSIS 
    Exports an object into a excel spreadsheet and creates a formatted table in custom colours. 
    .DESCRIPTION 
    Exports an object into a excel spreadsheet and creates a formatted table in custom colours. 
    .PARAMETER Object 
    The object you what to convert 
    .PARAMETER path 
    Name of the excel sheet you wish to create 
    .PARAMETER NoCobber
    True or False, do not overwrite file 
    .PARAMETER Display 
    Open the spreadsheet when finished 
    .EXAMPLE
    # Get services, save them into an spreadsheet called .\services.xlsx and load the spreadsheet 
    Get-Service | select name,*st* | Export-Excel -Path .\service.xlsx -Display 
    .EXAMPLE
    # Save the PSDrives to as spreadsheet called PSDrives 
    Get-PSDrive | Select name,provider,Used,Free | Export-Excel -Path .\PSDrive.xlsx 
    #>
    
    Param(
            [parameter(Position=0, Mandatory=$false,ValueFromPipeline=$true)][Object]$InputObject,
            [parameter(Mandatory=$true,ParameterSetName="Path")][String]$Path = "",
            [Switch]$Display = $false,
            [Switch]$NoClobber = $false
        )
        BEGIN {
            $Object = @()
        }
        Process {
            $Object += $InputObject
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
                                    "System.Object[]" { $value = (FJConvertto-HTML $value)}
                                    default { $Value = [string]$Value }
                                }
                            } else { 
                                $value = "" 
                            }
                            $HeaderName = (([regex]::replace(([regex]::replace($header.name,"[A-Z][a-z]+"," $&")),"[A-Z][A-Z]+"," $&")).replace("  "," ").trim())
                            $NewObject | Add-Member Noteproperty –Name $HeaderName –Value $Value
                        }
                        $ret += $NewObject 
                    }
                    $ret
                }
            }    
            If (-not (Test-Path HKLM:SOFTWARE\Classes\Excel.Application)) {
                throw "Excel not installed"
                exit
            } else {
                Add-Type -AssemblyName Microsoft.Office.Interop.Excel    
            }    
            If ($Path -eq "") {
                throw "You must specify either the -Path parameter"
                exit
            }
            if ($NoClobber -and (Test-Path $Path)) {
                throw "Export-Excel : The file '$Path' already exists"
                exit
            }
            If ($InputObject -ne $null) {
                $Object = FormatObject $Object
                $xl = new-object -comobject excel.application
                $tempname = [System.IO.Path]::GetTempFileName().Replace(".tmp",".html")
                # As Table
                $HTML = $Object | ConvertTo-Html -Fragment -as Table | out-string       
                # Right justifiy any cells with numbers, can contain comma and/or %
                $HTML = [regex]::replace($HTML,">[-+]?[0-9,]*\.?[0-9]+%?<",' style="text-align:right"$&') 
                # Right justifiy any cells with binary prefixes 
                $HTML = [regex]::replace($HTML,">(?i:[0-9,\.]+\s?[KMGTPEZY]B)<",' style=" text-align:right"$&') 
                # Make alternative lines grey 
                $HTML = [regex]::replace($HTML,">[nN]\/[aA]<\/td>",' style=" text-align:right"$&') 
                $HTML | set-content $tempname
                $Workbook = $xl.workbooks.open($tempname)
                $Worksheet = $xl.ActiveSheet
                $MyTable = $xl.ActiveWorkbook.TableStyles.Add("MyBlue") 
                $MyTable.ShowAsAvailablePivotTableStyle = $False
                $MyTable.ShowAsAvailableTableStyle = $True | Out-Null
                #$MyTable.ShowAsAvailableSlicerStyle = $False 
                #$MyTable.ShowAsAvailableTimelineStyle = $False 
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideVertical).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlWholeTable).Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlInsideHorizontal).ColorIndex =  [Microsoft.Office.Interop.Excel.XlColorIndex]::xlAutomatic
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlHeaderRow).Font.FontStyle = "Bold"
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlHeaderRow).Font.TintAndShade = 0
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlHeaderRow).Font.ThemeColor = [Microsoft.Office.Interop.Excel.xlThemeColor]::xlThemeColorDark1
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlHeaderRow).Interior.Color = 12611584
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlHeaderRow).Interior.TintAndShade = 0
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlRowStripe2).Interior.ThemeColor = [Microsoft.Office.Interop.Excel.xlThemeColor]::xlThemeColorDark1
                $MyTable.TableStyleElements.Item([Microsoft.Office.Interop.Excel.XlTableStyleElementType]::xlRowStripe2).Interior.TintAndShade = -0.249946592608417                   
                $Worksheet.Name = "Sheet1"
                # Create Range for table
                $Worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $WorkSheet.UsedRange,$null , 1).Name = "Table1"
                # Set table format to Fujitsu
                $Worksheet.ListObjects.Item("Table1").TableStyle = "MyBlue"
                # Justify the columns and Row
                Foreach($column in $WorkSheet.UsedRange.Columns) { $column.columnwidth = $column.columnwidth * 1.25 }
                $Worksheet.Columns.AutoFit() | Out-Null
                $Worksheet.Rows.AutoFit() | Out-Null
                # Left justify headers
                $Worksheet.ListObjects.Item("Table1").HeaderRowRange.HorizontalAlignment = [Microsoft.Office.Interop.Excel.Constants]::xlLeft
                If ($Path.substring(0,2) -eq ".\") {
                      $Path = $Path.replace(".\",($PWD.ToString() + "\"))
                }
                $workbook.saveas($path, 51)
                If ($Display -eq $true) {
                     $xl.visible = $true
                } else {
                     $xl.quit() | Out-Null 
                     [gc]::collect()
                     [gc]::WaitForPendingFinalizers()
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | out-null
                Remove-Variable xl
            }
        }
    }

    Get-Service | select name,*st* | Exportto-Excel -Path .\service.xlsx -Display