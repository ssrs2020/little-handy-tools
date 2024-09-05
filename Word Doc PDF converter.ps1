# Function to convert PDF to Word
function Convert-PDFToWord {
    param (
        [string]$pdfPath,
        [string]$wordOutputPath
    )

    # Create Word COM object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    try {
        # Open the PDF file
        $document = $word.Documents.Open($pdfPath, [ref]$false, [ref]$true, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false)

        # Save the document as a Word file
        $document.SaveAs([ref]$wordOutputPath, [ref]16)  # 16 represents wdFormatDocument (Word format)
        Write-Host "Conversion successful: PDF to Word"
    } catch {
        Write-Error "Error during conversion: $_"
    } finally {
        # Close the document and quit Word
        $document.Close()
        $word.Quit()
    }
}

# Function to convert Word to PDF
function Convert-WordToPDF {
    param (
        [string]$wordPath,
        [string]$pdfOutputPath
    )

    # Create Word COM object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    try {
        # Open the Word file
        $document = $word.Documents.Open($wordPath)

        # Save the document as a PDF file
        $document.SaveAs([ref]$pdfOutputPath, [ref]17)  # 17 represents wdFormatPDF
        Write-Host "Conversion successful: Word to PDF"
    } catch {
        Write-Error "Error during conversion: $_"
    } finally {
        # Close the document and quit Word
        $document.Close()
        $word.Quit()
    }
}

# Main logic
$conversionType = Read-Host "Choose conversion type: 'PDFToWord' or 'WordToPDF'"

if ($conversionType -eq 'PDFToWord') {
    $pdfPath = Read-Host "Enter the path of the PDF file"
    $wordOutputPath = Read-Host "Enter the output path for the Word document"
    Convert-PDFToWord -pdfPath $pdfPath -wordOutputPath $wordOutputPath
} elseif ($conversionType -eq 'WordToPDF') {
    $wordPath = Read-Host "Enter the path of the Word file"
    $pdfOutputPath = Read-Host "Enter the output path for the PDF document"
    Convert-WordToPDF -wordPath $wordPath -pdfOutputPath $pdfOutputPath
} else {
    Write-Host "Invalid conversion type. Please choose 'PDFToWord' or 'WordToPDF'."
}
