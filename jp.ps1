# Define the directory path to search
$directoryPath = "z:\"

# Recursive function to find and traverse junction points
function Get-JunctionPoints {
    param (
        [string]$Path,
        [int]$Depth = 0,
        [int]$MaxDepth = 2
    )

    # Return if the current depth exceeds the maximum depth
    if ($Depth -gt $MaxDepth) {
        return
    }

    # Get all directories in the current path
    Get-ChildItem -Path $Path -Directory | ForEach-Object {
        # Check if the directory is a junction point
        if ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            # Output the full path of the junction point
            [PSCustomObject]@{
                FullPath = $_.FullName
                # Use fsutil to get target information if needed
                Target = Get-JunctionTarget -Path $_.FullName
            }
        }

        # Recursively process subdirectories, including junction points
        Get-JunctionPoints -Path $_.FullName -Depth ($Depth + 1) -MaxDepth $MaxDepth
    }
}

# Helper function to get the target of a junction point using fsutil
function Get-JunctionTarget {
    param (
        [string]$Path
    )
    try {
        $output = & fsutil reparsepoint query "$Path" 2>&1
        if ($output -match "Substitute Name:\s+(.*)") {
            return $matches[1].Trim()
        } else {
            return "Unknown"
        }
    } catch {
        return "Error retrieving target"
    }
}

# Execute the function and collect results
$result = Get-JunctionPoints -Path $directoryPath -MaxDepth 2

# Output the full paths of all found junction points
$result | ForEach-Object { $_.FullPath }

# Define output file path
$outputFile = "junction_points.txt"

# Output the full paths of all found junction points with path replacement
$result | ForEach-Object {
	$_.FullPath -replace [regex]::Escape($directoryPath), "\\forskning.it.ntnu.no\ntnu\mh\kin\"
} | Out-File -FilePath $outputFile -Encoding UTF8

# Write-Output "Traversal completed. Found $($result.Count) junction points in $directoryPath (up to 2 levels)."
