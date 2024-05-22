param (
    # Parameter help description
    [Parameter(AttributeValues)]
    [ParameterType]
    $ParameterName
)

# START Using ArrayList to build data collections
$ReportDate = Get-Date -Format FileDate
$ScriptName = $MyInvocation.MyCommand.Name
$FilePath = "$PSSCRIPTROOT\output\$ScriptName-$ReportDate"
$report = [System.Collections.ArrayList][PSCustomObject]::new()
$i = 0

# Collect data into variable
$collection # = Run-Command

foreach ($item in $collection) {
    $i++
    Write-Progress -Activity "Processing $($item)" -PercentComplete (($i/$collection.Length)*100)

    # Some Process of $item

    # Build the data object
    $ReportItem = [PSCustomObject][ordered]@{
        Property1 = $item.Property1
        Property2 = $item.Property2
        Property3 = $item.Property3
    }

    # Add data object to report collection
    $report.add($ReportItem) > $null
}

# Report can be exported in multiple ways
$Report | Export-CSV -Path "$FilePath.csv"
$Report | Export-Excel -Path "$FilePath.xlsx"
$Report | ConvertTo-Json > "$FilePath.json"

Write-Host "File created: $FilePath"

# END Using ArrayList to build data collections

# START Transcription Block
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
$Date = Get-Date -Format "yyyyddMM-HHmm"
$ScriptName = $MyInvocation.MyCommand.Name
$LogPath = "$PSScriptRoot\output\$ScriptName-$Date.txt"
Start-Transcript -path $LogPath | Out-Null

# do some stuff

Stop-Transcript | Out-Null
# END Transcription Block
