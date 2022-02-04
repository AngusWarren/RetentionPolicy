$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

class RetentionPolicy {
    [Int] $Monthly = 99999
    [Int] $Weekly = 45
    [Int] $Daily = 21
    [Int] $IntraDaily = 3
}

function New-RetentionPolicy {
    <#
        .SYNOPSIS
            Sets up the retention policy used by Start-RetentionPolicyCleanup.
        .EXAMPLE
            $policy = New-RetentionPolicy
        .EXAMPLE
            $policy = New-RetentionPolicy -Weekly 45 -Daily 14 -IntraDaily 3
        .EXAMPLE
            $policy = New-RetentionPolicy -Monthly 24 -Weekly 45 -Daily 14 -IntraDaily 3
    #>
    [CmdletBinding()]
    param (
        # Number of days to retain monthly files
        [Int]
        $Monthly = 99999,

        # Number of days to retain weekly files
        [Int]
        $Weekly = 45,

        # Number of days to retain daily files
        [Int]
        $Daily = 21,

        # Number of days to retain intradaily files
        [Int]
        $IntraDaily = 21
    )

    process {
        return [RetentionPolicy]@{
            Monthly    = $Monthly
            Weekly     = $Weekly
            Daily      = $Daily
            IntraDaily = $IntraDaily
        }
    }
}


function Initialize-RetentionPolicy {
    <#
        .SYNOPSIS
            Checks an input array against a retention policy and adds properties to the output
        .EXAMPLE
            Initialize-RetentionPolicy
    #>
    [CmdletBinding()]
    param (
        # Retention policy to use. Created with New-RetentionPolicy
        [Parameter(Mandatory)]
        [RetentionPolicy]
        $Policy,

        # Objects to process from the pipeline
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject[]]
        $InputObject,

        # Property to use for date matching and sorting.
        [Parameter(Mandatory)]
        [String]
        $DateProperty,

        # Use this if you prefer to keep the newest file from each time window.
        [Switch]
        $PreferNewest
    )

    begin {
        $objects = New-Object System.Collections.Generic.List[PSObject]
    }

    process {
        foreach ($object in $InputObject) {
            if ($object.$DateProperty -isnot [DateTime]) {
                throw "$DateProperty is not a valid [DateTime] for $object"
            }
            $objects.Add($object)
        }
    }

    end {
        $objects = $objects | Sort-Object $DateProperty -Descending:$PreferNewest

        $now = Get-Date
        $found = New-Object System.Collections.Generic.List[String]

        foreach ($object in $objects) {
            $date = $object.$DateProperty
            $isoDate = ConvertTo-IsoDate -InputObject $date

            $yearMonthDay = $isoDate.IsoDateFormat
            $yearMonth = "$( $date.year )-$( $date.month )"
            $yearWeek = "$( $isoDate.IsoYear )-$( $isoDate.IsoWeek )"

            $retentionReason = @()
            if ($date -gt $now.AddDays(-$Policy.Monthly) -and $yearMonth -notin $found) {
                $found.Add($yearMonth)
                $retentionReason += 'Monthly'
            }
            if ($date -gt $now.AddDays(-$Policy.Weekly) -and $yearWeek -notin $found) {
                $found.Add($yearWeek)
                $retentionReason += 'Weekly'
            }
            if ($date -gt $now.AddDays(-$Policy.Daily) -and $yearMonthDay -notin $found) {
                $found.Add($yearMonthDay)
                $retentionReason += 'Daily'
            }
            if ($date -gt $now.AddDays(-$Policy.IntraDaily)) {
                $retentionReason += 'IntraDaily'
            }
            $object | Add-Member -MemberType NoteProperty -Name 'Retain' -Value ($retentionReason.Count -gt 0)
            $object | Add-Member -MemberType NoteProperty -Name 'RetentionReason' -Value $retentionReason
            $object
        }
    }
}

function Start-RetentionPolicyCleanup {
    <#
        .SYNOPSIS
            Removes old files from a directory except for those matching a retention policy.
        .EXAMPLE
            $Params = @{
                Policy = New-RetentionPolicy -Weekly 45 -Daily 14 -IntraDaily 3
                Source = 'C:\Backups'
                FileNamePattern = 'prefix_(\d{4}-\d\d-\d\d)_.*\.tgz'
                DateProperty = 'FileNamePattern'
                MinimumSize = 3500MB
            }
            Start-RetentionPolicyCleanup @Params -WhatIf
        .EXAMPLE
            $policy = New-RetentionPolicy -Weekly 45 -Daily 14 -IntraDaily 3
            $src = 'C:\Backups'
            Start-RetentionPolicyCleanup -Policy $Policy -Source $src -FileNamePattern 'prefix_.*' -DateProperty LastWriteTime
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        # Retention policy to use. Created with New-RetentionPolicy
        [Parameter(Mandatory)]
        [RetentionPolicy]
        $Policy,

        # Directory to search for files.
        [Parameter(Mandatory)]
        [String]
        $Source,

        # Destination to move found backups. If a relative path is given, it's joined with the $Source param.
        [String]
        $Destination = 'CleanedFiles',

        # Only files matching this regular expression will be included in the job. Please make it very specific.
        [Parameter(Mandatory)]
        [String]
        $FileNamePattern,

        # Property to use for date matching and sorting.
        # If the FileNamePattern is used, it will use the first capture group from the FileNamePattern paramater.
        # TODO: This should have auto-complete suggestions but still allow free text fields.
        [Parameter(Mandatory)]
        [ValidateSet('FileNamePattern', 'CreationTime', 'LastWriteTime')]
        [String]
        $DateProperty,

        # Only work on files bigger than this. PowerShell allows units to be included, eg. 3500MB
        # This is used to prevent 0 length files taking up the place of a legitmate file.
        [Int64]
        $MinimumSize = 10MB,

        # Delete files rather than moving them.
        [Switch]
        $Delete
    )

    process {
        if (-not [System.IO.Path]::IsPathRooted($Destination)) {
            $Destination = Join-Path $Source $Destination
        }
        $destinationExists = Test-Path -PathType Container -Path $Destination
        if ($destinationExists -eq $false -and $PSCmdlet.ShouldProcess($Destination, 'Create Destination')) {
            $null = New-Item -Path $Destination -ItemType Directory -Force
        }

        $allFiles = Get-ChildItem -File $Source
        #this was used for debuging
        #$allFiles = Import-Clixml oldfilelist.xml
        $matchingFiles = $allFiles | Where-Object {
            $_.Length -gt $MinimumSize -and $_.Name -match $FileNamePattern
        }

        if ($DateProperty -eq 'FileNamePattern') {
            foreach ($file in $matchingFiles) {
                $fileNameDate = $file.Name -replace $FileNamePattern, '$1' | Get-Date
                $file | Add-Member -NotePropertyName 'FileNamePattern' -NotePropertyValue $fileNameDate
            }
        }
        $files = $matchingFiles | Initialize-RetentionPolicy -Policy $Policy -DateProperty $DateProperty

        $filesToRemove = $files | Where-Object { $_.Retain -eq $false }
        if ($VerbosePreference -eq 'Continue') {
            $filesToKeep = $files | Where-Object { $_.Retain -eq $true }
            Write-Verbose ('Cleaning {0} files and retaining {1}.' -f $filesToRemove.Count, $filesToKeep.Count)
        }

        foreach ($file in $filesToRemove) {
            if ($Delete) {
                if ($PSCmdlet.ShouldProcess($File.FullName, 'Delete')) {
                    Remove-Item -Path $File.FullName
                }
            } else {
                if ($PSCmdlet.ShouldProcess($File.FullName, "Move to $Destination")) {
                    Move-Item -Path $File.FullName -Destination $Destination
                }
            }
        }
    }
}

function ConvertTo-IsoDate {
    <#
        .SYNOPSIS
            Converts a DateTime to an object containing ISO date values.
        .DESCRIPTION
            Required because PowerShell 5.1 returns false ISO Year and ISO Week values.
        .EXAMPLE
            $date | ConverTo-IsoDate
        .EXAMPLE
            ConvertTo-IsoDate $date
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline, Mandatory, Position = 0)]
        [DateTime]
        $InputObject
    )

    process {
        $DayofWeek = [int]$DateTime.DayOfWeek
        if ($DayofWeek -eq 0) {
            $DayofWeek = 7
        }
        $Thursday = $DateTime.AddDays(4 - $DayofWeek)

        [PSCustomObject]@{
            IsoDateFormat = (Get-Date $DateTime -Format 'yyyy-MM-dd')
            IsoYear       = $Thursday.Year
            IsoWeek       = 1 + [Math]::Floor(($Thursday.DayOfYear - 1) / 7)
        }
    }
}

Export-ModuleMember -Function New-RetentionPolicy
Export-ModuleMember -Function Initialize-RetentionPolicy
Export-ModuleMember -Function Start-RetentionPolicyCleanup
