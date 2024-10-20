param (
    [string]$srcroot,
    [string]$dest,
    [switch]$dryrun = $false,
    [switch]$verbose = $false
)

$shell = New-Object -COMObject Shell.Application

function Move-Image([System.IO.FileInfo]$file, [string]$dest) {
    $parent = $file.Directory
    $filename = $file.Name
    $fullname = $file.FullName

    # Note that `$folder` is a COM object unlike `$parent`.
    $folder = $shell.Namespace($parent.FullName)
    $index = 12

    $propertyName = $folder.GetDetailsOf($null, $index)
    if ($propertyName -ne '撮影日時') {
        Write-Output "Unexpected property name of index 12: $tagName"
        return
    }

    $folderItem = $folder.ParseName($filename)
    if ($null -eq $folderItem) {
        $parentDirName = $parent.FullName
        Write-Output "Failed to obtain $filename in $parentDirName."
        return
    }

    $datestr = $folder.GetDetailsOf($folderItem, $index) -replace "[\p{Cc}\p{Cf}]", ""
    if (!($datestr -match "^([0-9]+)/([0-9]+)/([0-9]+) ")) {
        Write-Output "Failed to obtain the shooting date information in $fullname."
        return
    }

    [int]$nyear = $Matches.1
    [int]$nmonth = $Matches.2
    [int]$nday = $Matches.3
    [string]$syear = $Matches.1
    [string]$smonth = $Matches.2
    [string]$sday = $Matches.3

    if (($nyear -lt 1970) -or ($nmonth -lt 1) -or ($nmonth -gt 12) -or ($nday -lt 1) -or ($nday -gt 31)) {
        Write-Output "Invalid shooting date information: `$datestr = $datestr"
        return
    }

    $destdir = "$dest/$syear/$smonth"

    if ($verbose) {
        Write-Output "  `$year = $syear, `$month = $smonth, `$day = $sday, `$dest = $dest"
    }

    if (!(Test-Path $destdir)) {
        if ($verbose) {
            Write-Output "  Creating $destdir."
        }
        if (!$dryrun) {
            New-Item $destdir -ItemType directory > $null
        }
    }

    $dest = "$destdir/$filename"
    $filenameBody = $filename -replace "\.[a-zA-Z0-9]+$", ""
    $ext = $file.Extension
    [int] $counter = 0
    while (Test-Path $dest) {
        $counter = $counter + 1
        $dest = "$destdir/$filenameBody" + "_" + $counter + $ext
    }
    
    if ($verbose) {
        $path = $fullname
        Write-Output "  Moving $path into $dest."
    }
    if ($dryrun) {
        return
    }

    Move-Item $fullname -destination $dest

}


if ($verbose) {
    Write-Output "`$srcroot = $srcroot"
    Write-Output "`$dest = $dest"
    Write-Output "`$dryrun = $dryrun"
    Write-Output "`$verbose = $verbose"
}

if ($srcroot -eq "") {
    Write-Output "The srcroot parameter is required."
    Break
}

if ($dest -eq "") {
    Write-Output "The dest parameter is required."
    Break
}

Get-ChildItem -Path $srcroot -Recurse | ForEach-Object {
    $item = $_
    $filename = $item.Name
    $fullname = $item.FullName
    $ext = $item.Extension

    if ($item.PsIsContainer) {
        # Ignore directories.
    } elseif (($filename -eq "desktop.ini") -or ($filename -eq "Thumbs.db") -or ($filename -eq ".picasa.ini")) {
        # Ignore these files.
    } elseif (($ext -eq ".jpg") -or ($ext -eq ".jpeg")) {
        if ($verbose) {
            Write-Output "$fullname is a JPEG file."
        }
        Move-Image $item $dest/JPEG

    } elseif (($ext -eq ".CR2") -or ($ext -eq ".DNG") -or ($ext -eq ".RW2")) {
        if ($verbose) {
            Write-Output "$fullname is a raw file."
        }
        Move-Image $item.FullName $dest/RAW
    } else {
        Write-Output "$fullname is an unknown file."
    }
}