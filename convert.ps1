# Requirements to run this script:
# 1. LibreOffice Impress installed and added to the system PATH
# 2. ImageMagick installed and added to the system PATH
#   2.1. Ghostscript is required for ImageMagick to convert PDF to TIFF

# Define the directories
$inputDirectory = "input"
$cacheDirectory = "cache"
$outputDirectory = "output"
$libreOfficeImpressPath = "simpress"
$imageMagickPath = "magick"

# Function to convert PDF to TIFF
function Convert-PdfToTiff {
    param (
        [string]$pdfFilePath,
        [string]$folderName
    )

    # Create a folder for each slide
    $slideFolder = Join-Path -Path $outputDirectory -ChildPath $folderName
    if (!(Test-Path $slideFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $slideFolder | Out-Null
    }
    
    # Extract PDF file name without extension
    $pdfFileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfFilePath)
    
    # Construct TIFF file path with folder name
    $tiffFilePath = Join-Path -Path $slideFolder -ChildPath "$pdfFileName.tiff"

    # Execute ImageMagick command to convert PDF to TIFF
    & $imageMagickPath "$pdfFilePath" +adjoin "$tiffFilePath"
}

# Check if the input directory exists
if (Test-Path $inputDirectory) {
    Write-Host "Directory '$inputDirectory' exists. Searching for PPTX files..."

    # Use Get-ChildItem to recursively find PPTX files in the directory
    $pptxFiles = Get-ChildItem -Path $inputDirectory -Recurse -Filter *.pptx

    # Check if any PPTX files were found
    if ($pptxFiles.Count -gt 0) {
        Write-Host "Found $($pptxFiles.Count) PPTX files. Converting to PDF..."

        # Execute command to convert all PPTX files to PDF with wait for completion
        Start-Process -FilePath $libreOfficeImpressPath -ArgumentList "--headless --convert-to pdf --outdir $cacheDirectory $inputDirectory\*.pptx" -Wait

        # Check the exit code of the last command
        if ($LastExitCode -eq 0) {
            Write-Host "Conversion from PPTX to PDF completed successfully."

            # Get all PDF files in the cache directory
            $pdfFiles = Get-ChildItem -Path $cacheDirectory -Filter *.pdf

            # Check if any PDF files were found
            if ($pdfFiles.Count -gt 0) {
                Write-Host "Found $($pdfFiles.Count) PDF files. Converting to TIFF..."

                # Loop through each PDF file and convert to TIFF
                foreach ($pdfFile in $pdfFiles) {
                    # Determine folder name from PDF file name
                    $folderName = $pdfFile.BaseName

                    # Call function to convert PDF to TIFF with folder name
                    Convert-PdfToTiff -pdfFilePath $pdfFile.FullName -folderName $folderName
                }

                Write-Host "Conversion from PDF to TIFF completed."
            } else {
                Write-Host "No PDF files found in the directory '$cacheDirectory'."
            }
        } else {
            Write-Host "Conversion from PPTX to PDF failed with exit code: $LastExitCode"
        }
    } else {
        Write-Host "No PPTX files found in the directory '$inputDirectory'."
    }
} else {
    Write-Host "Directory '$inputDirectory' does not exist."
}
