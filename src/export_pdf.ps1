param(
    [string]$Docx = "paper/Random_Forest_Digit_Reproduction_Paper.docx"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$DocxPath = [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $Docx))
$PdfPath = [System.IO.Path]::ChangeExtension($DocxPath, ".pdf")
$Word = $null
$Document = $null

try {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0
    $Document = $Word.Documents.Open($DocxPath, $false, $true)
    $Document.ExportAsFixedFormat(
        $PdfPath,
        17,
        $false,
        0,
        0,
        0,
        0,
        0,
        $true,
        $true,
        0,
        $true,
        $true,
        $false
    )
    Write-Output "Exported $PdfPath"
}
finally {
    if ($null -ne $Document) { $Document.Close($false) }
    if ($null -ne $Word) { $Word.Quit() }
    if ($null -ne $Document) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
    }
    if ($null -ne $Word) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null
    }
}
