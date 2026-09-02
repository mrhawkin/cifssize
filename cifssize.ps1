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
$data = @()

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

# Write to CSV file
# $excel_data | Export-Csv -Path "usagedata.csv" -NoTypeInformation
$excel_data | Export-Excel -Path "usagedata.xlsx" -AutoSize -TableName table -TableStyle Medium6 -FreezeTopRow