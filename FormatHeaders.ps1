
Function FormatHeaders( $PSObject ) {
    <# 
    .SYNOPSIS 
    Formats the headers of an object
    .PARAMETER Object 
    The object you what to convert 
    .PARAMETER path 
    Name of the excel sheet you wish to create 
    .EXAMPLE
    RenameHeaders $(gwmi win32_service | select -first 5 serv*,start*) | ft
    #>
    
    If ($PSObject -ne $null) {
        $ret = @()

        Foreach ($line in $PSObject) {
            $NewObject = New-Object -TypeName PSObject
            foreach ($header in $line.psobject.properties) {
                $NewObject | Add-Member Noteproperty -Name $(([regex]::replace(([regex]::replace($header.name,"[A-Z][a-z]+"," $&")),"[A-Z][A-Z]+"," $&")).replace("  "," ").trim() ) -Value $header.value
            }
            $ret += $NewObject 
        }
        $ret
    }
}