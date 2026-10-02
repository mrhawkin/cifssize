function DisplayInBytes($num) 
{
    $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
    $index = 0
    while ($num -gt 1kb) 
    {
        $num = $num / 1kb
        $index++
    } 

    "{0:N1} {1}" -f $num, $suffix[$index]
}

# Init table
$output_data = @()
$excel_data  = @()

# Get the list of shares
$UNCPaths = Get-Content junction_points.txt | Select-Object

foreach ($path in $UNCPaths) {

	# Map share to drive
        $nwobj=new-object -comobject WScript.Network
        $status=$nwobj.mapnetworkdrive("O:",$path)
        $drive=get-psdrive O
	
	# Remove the first common part of the path
	$pathshort = $path.Substring(35,$path.Length-35)

	# Find stats about the share
        $bytesfree=($drive.free)
	$bytesused=($drive.used)
	$bytestotal=($bytesfree+$bytesused)

	# Convert to Terabytes for excelfile
	$TBfree=([math]::Round($bytesfree/1TB,1))
	$TBused=([math]::Round($bytesused/1TB,1))
	$TBtotal=([math]::Round($bytestotal/1TB,1))

	# Convert dynamically for display
        $freehuman=DisplayInBytes($bytesfree)
        $usedhuman=DisplayInBytes($bytesused)
        $totalhuman=DisplayInBytes($bytestotal)

        $freeinPercentage=($bytesfree/$bytestotal*100)
        $freewithoutdecimal=([math]::Round($freeinPercentage))
        
	# Make stats row
	$output_line = @(
		[pscustomobject]@{ Path = $pathshort; "Free(%)" = $freewithoutdecimal; Used = $usedhuman; Total = $totalhuman}
	)
	$excel_line = @(
		[pscustomobject]@{ Path = $pathshort; "Free(%)" = $freewithoutdecimal; Used = $TBused; Total = $TBtotal}
	)

	# Add row to table
        $output_data += $output_line
	$excel_data += $excel_line
        $status=$nwobj.removenetworkdrive("O:")
}

$output_data | ForEach {[PSCustomObject]$_} | Format-Table -AutoSize

# Grupper etter første del av Path.
# Eksempel: "gruppe1\område\share" -> "gruppe1"
$gruppe_data = $excel_data |
    Group-Object {
        ($_.Path.TrimStart('\', '/') -split '[\\/]')[0]
    } |
    ForEach-Object {
        [pscustomobject]@{
            Gruppe         = $_.Name
            'Brukt TB'     = [math]::Round(
                ($_.Group | Measure-Object -Property Used -Sum).Sum, 1
            )
            'Totalt TB'    = [math]::Round(
                ($_.Group | Measure-Object -Property Total -Sum).Sum, 1
            )
        }
    } |
    Sort-Object Gruppe

$gruppe_data = @($gruppe_data)

$gruppe_data += [pscustomobject]@{
    Gruppe      = 'TOTALT'
    'Brukt TB'  = [math]::Round(
        ($gruppe_data | Measure-Object -Property 'Brukt TB' -Sum).Sum, 1
    )
    'Totalt TB' = [math]::Round(
        ($gruppe_data | Measure-Object -Property 'Totalt TB' -Sum).Sum, 1
    )
}

$filename = "usagedata_$(Get-Date -Format 'yyyy-MM').xlsx"
Write-Host "Eksporterer til: $filename"

$excel_data | Export-Excel `
    -Path $filename `
    -WorksheetName 'Detaljer' `
    -AutoSize `
    -TableName 'UsageDetails' `
    -TableStyle Medium6 `
    -FreezeTopRow `
    -ClearSheet

$gruppe_data | Export-Excel `
    -Path $filename `
    -WorksheetName 'Per gruppe' `
    -AutoSize `
    -TableName 'UsageGroups' `
    -TableStyle Medium6 `
    -FreezeTopRow `
    -ClearSheet
