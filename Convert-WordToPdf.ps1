<#
.SYNOPSIS
Converts Microsoft Word documents to PDF.

.DESCRIPTION
Converts .doc and .docx files from a specified directory to PDF using Microsoft Word.

By default:
- Only files in the specified directory are processed.
- Existing PDF files are skipped.
- Original Word documents are preserved.

Use -Recurse to include subdirectories.
Use -Force to overwrite existing PDF files.

Microsoft Word must be installed on the computer.

.PARAMETER Path
The directory containing the Word documents to convert.

.PARAMETER Recurse
Includes files located in subdirectories.

.PARAMETER Force
Overwrites existing PDF files.

.EXAMPLE
.\Convert-WordToPdf.ps1 -Path "C:\Documents"

Converts all .doc and .docx files directly inside C:\Documents.

.EXAMPLE
.\Convert-WordToPdf.ps1 -Path "C:\Documents" -Recurse

Converts all .doc and .docx files inside C:\Documents and its subdirectories.

.EXAMPLE
.\Convert-WordToPdf.ps1 -Path "C:\Documents" -Force

Converts the documents and overwrites existing PDF files.

.EXAMPLE
.\Convert-WordToPdf.ps1 -Path "C:\Documents" -Recurse -Force

Processes the entire directory tree and overwrites existing PDF files.

.NOTES
Version: 1.0.0
Requires Microsoft Word on Windows.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [switch]$Recurse,

    [switch]$Force
)

# Validate the directory.
if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Error "The specified directory does not exist: $Path"
    return
}

# Resolve the directory to an absolute path.
$resolvedPath = (Resolve-Path -LiteralPath $Path).Path

# Find Word documents.
if ($Recurse) {
    $files = @(
        Get-ChildItem -LiteralPath $resolvedPath -File -Recurse | Where-Object { $_.Extension -in ".doc", ".docx" }
    )
}
else {
    $files = @(
        Get-ChildItem -LiteralPath $resolvedPath -File | Where-Object { $_.Extension -in ".doc", ".docx" }
    )
}

if ($files.Count -eq 0) {
    Write-Host "No Word documents found."
    return
}

# Start Microsoft Word.
try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
}
catch {
    Write-Error "Microsoft Word could not be started. Make sure it is installed."
    return
}

$word.Visible = $false
$word.DisplayAlerts = 0

$converted = 0
$skipped = 0
$failed = 0

try {
    foreach ($file in $files) {
        $docPath = $file.FullName
        $pdfPath = [System.IO.Path]::ChangeExtension($docPath, ".pdf")

        if ((Test-Path -LiteralPath $pdfPath) -and -not $Force) {
            Write-Host "Skipped: $pdfPath already exists."
            $skipped++
            continue
        }

        Write-Host "Converting: $docPath"

        $document = $null

        try {
            $document = $word.Documents.Open($docPath)

            # 17 = wdFormatPDF
            $document.SaveAs([ref]$pdfPath, [ref]17)

            Write-Host "Created: $pdfPath"
            $converted++
        }
        catch {
            Write-Warning "Failed to convert '$docPath': $($_.Exception.Message)"
            $failed++
        }
        finally {
            if ($null -ne $document) {
                $document.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
            }
        }
    }
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Conversion complete."
Write-Host "Converted: $converted"
Write-Host "Skipped:   $skipped"
Write-Host "Failed:    $failed"
