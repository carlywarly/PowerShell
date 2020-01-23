Function MyConvertto-HTML {
	<#
	.SYNOPSIS
	Converts an object into a formatted HTML table
	.DESCRIPTION
	Converts an object into a formatted HTML table. The header being corporate colours, with alternate lines been white and grey. So that it can be e-mailed or save to a html doc (or both).
	in green.
	.AUTHOR
	Carl Armstrong
	.PARAMETER Object
	The Object you what to convert
	.PARAMETER aa
	(Optional) Header for table
	.PARAMETER NoStyle
	(Optional) Whether to add sytle info, you only need one per webpage.
	.PARAMETER NoHeader
	(Optional) Whether to include table headers
	.EXAMPLE
	# List services and save as a webpage
	ASConvertto-HTML (Get-Service | select DisplayName,Status,ServiceType) -Header "Services" | Set-Content .\services.html
	Start .\services.html
	.EXAMPLE
	# List services and hotfixes and e-mail as html report
	$Email = ASConvertto-HTML (Get-Service | select DisplayName,Status,ServiceType) -Header "Services"
	$Email += Get-Hotfix | select Description,HotFixID,InstalledOn | ASConvertto-HTML -Header "Hotfixes" -NoStyle
	Send-MailMessage -From "noreply@na.com" -To "someone@b.com" -subject "Services" -SmtpServer webmail.c.com -Body $Email -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
	.Version
	2.5 - Added ability to use pipeline as an input
	2.6 - Added right justify for IP Addresses
	.NOTES
	We can make alternative lines grey via the style sheets using the below but this does not work
	when e-mailing the HTML.
	$HTML += "TR:nth-child(even) {background-color: #dddddd;}"
	#>
	[cmdletbinding()]
	param([parameter(ValueFromPipeline=$true,Mandatory=$false,Position=0)][psobject[]]$InObject,
		[parameter(ValueFromPipeline=$false,Mandatory=$False,Position=1)][string]$Header = "",
		[parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$NoHeader = $False,
		[parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$NoStyle = $False,
		[parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$List = $False,
		[parameter(ValueFromPipeline=$false,Mandatory=$False)][Switch]$Table = $False
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
								"System.Boolean" {
									If ($Value) { $Value = "Yes" }
									else { $Value = "No" } }
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
									"System.Double[]" { $Value = ($Value | % { "{0:N0}" -f $_ }) -join ", " | Out-String }
									"System.Single" { $Value = "{0:N2}" -f $Value }
									"System.Single[]" { $Value = ($Value | % { "{0:N2}" -f $_ }) -join ", " | Out-String }
									"System.Object[]" {
										$value = (FJConvertto-HTML $value)
									}
									default { $Value = [string]$Value }
								}
							} else { $value = "" }
							$HeaderName = (([regex]::replace(([regex]::replace($header.name,"[A-Z][a-z]+"," $&")),"[A-Z][A-Z]+"," $&")).replace(" "," ").replace(" ( "," (").trim())
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
					$HTML += "<style>TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
					$HTML += "TH{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#4da6ff;color:black;text-align:left;padding-left:8px;padding-right:8px;font-weight:bold}"
					$HTML += "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align:left;padding-left:8px;padding-right:8px}"
					$HTML += "</style>"
				}
				# Convert to HTML
				If ($Header -ne $null) { $html += "<H2>$Header</H2>" }
				if ($List.IsPresent) {
					# As List
					$HTML += $Object | ConvertTo-Html -Fragment -as List | out-string
					$HTML = $HTML.Replace("&" + "lt;","<").Replace("&" + "gt;",">").Replace("&" + "quot;",'"').Replace("&" + "amp;",'&')
					$HTML = $HTML.replace("<tr><td>",'<tr><th>').replace(":</td><td>","</th><td>")
				} else {
					# As Table
					$HTML += $Object | ConvertTo-Html -Fragment -as Table | out-string
					$HTML = $HTML.Replace("&" + "lt;","<").Replace("&" + "gt;",">").Replace("&" + "quot;",'"').Replace("&" + "amp;",'&')
					# Right justifiy any cells with numbers, can contain comma and/or %
					$HTML = [regex]::replace($HTML,">[-+]?[0-9,]*\.?[0-9]+?<",' style="text-align:right"$&')
					# Right justifiy any cells with binary prefixes
					$HTML = [regex]::replace($HTML,">(?i:[0-9,\.]+\s?[KMGTPEZY]B)<",' style="text-align:right"$&')
					# Right justifiy any cells with IP addresses
					$HTML = [regex]::replace($HTML,">(?:(?:25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)<",' style=" text-align:right"$&')
					# Right justifiy any cells with hr(s)
					$HTML = [regex]::replace($HTML,">\d* hr\(s\)<",' style=" text-align:right"$&')
					# Right justify N/A
					$HTML = [regex]::replace($HTML,">[nN]\/[aA]<\/td>",' style=" text-align:right"$&')
					# Make alternative lines grey
					$HTML = [regex]::replace($HTML,"<tr><td>.+\n<tr",'$& style="background:#C0C0C0"')
					# Make alternative lines grey
					$HTML = [regex]::replace($HTML,">[0-9]%<",' style="text-align:right;font-weight:bold;background:red;color:yellow"$&')
					$HTML = [regex]::replace($HTML,">1[0-4]%<",' style="text-align:right;font-weight:bold;background:yellow;color:red"$&')
					$HTML = [regex]::replace($HTML,">1[5-9]%<",' style="text-align:right"$&')
					$HTML = [regex]::replace($HTML,">[2-9][0-9]%<",' style="text-align:right"$&')
					$HTML = [regex]::replace($HTML,">100%<",' style="text-align:right"$&')
				}
				If ($NoHeader.ispresent) {
					$HTML = [regex]::replace($HTML,"<tr><th>.+</th></tr>",'')
				}
				$HTML = $HTML.Replace("<table>",'<table id="Table">')
				$HTML
			}
	    }
	}
