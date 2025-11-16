[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host ""
Write-Host "=== Docx Replace V2 - Actual Replace Mode ==="
Write-Host ""

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml

$dictPath = ".\dict.csv"
if (-not (Test-Path $dictPath))
{
    Write-Host "Error: Cannot find dict.csv" -ForegroundColor Red
    pause
    exit
}

$rules = Import-Csv $dictPath
Write-Host "Loaded dict.csv: $($rules.Count) rules"
Write-Host ""

$docs = Get-ChildItem -Filter *.docx | Where-Object { $_.Name -notlike "*_fixed.docx" }
if ($docs.Count -eq 0)
{
    Write-Host "Error: No docx files found" -ForegroundColor Red
    pause
    exit
}

foreach ($file in $docs)
{
    Write-Host "================================================"
    Write-Host "Processing: $($file.Name)"
    Write-Host "================================================"
    
    # Create temp folder for extraction
    $tempFolder = ".\temp_$($file.BaseName)_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    
    try
    {
        # Extract docx (it's a zip file)
        [System.IO.Compression.ZipFile]::ExtractToDirectory($file.FullName, $tempFolder)
        Write-Host "Extracted docx to temp folder"
        
        # Read document.xml
        $docXmlPath = Join-Path $tempFolder "word\document.xml"
        if (-not (Test-Path $docXmlPath))
        {
            Write-Host "Error: Cannot find word/document.xml" -ForegroundColor Red
            Remove-Item $tempFolder -Recurse -Force
            continue
        }
        
        $xmlContent = [System.IO.File]::ReadAllText($docXmlPath, [System.Text.Encoding]::UTF8)
        Write-Host "Original document.xml length: $($xmlContent.Length) chars"
        
        # Parse XML
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.PreserveWhitespace = $true
        try
        {
            $xmlDoc.LoadXml($xmlContent)
        }
        catch
        {
            Write-Host "Error: XML parsing failed" -ForegroundColor Red
            Remove-Item $tempFolder -Recurse -Force
            continue
        }
        
        # Get all text nodes
        $ns = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $textNodes = $xmlDoc.SelectNodes("//w:t", $ns)
        
        Write-Host "Found $($textNodes.Count) text nodes"
        
        # Perform replacements
        $totalReplacements = 0
        
        foreach ($rule in $rules)
        {
            $wrong = $rule.wrong
            $right = $rule.right
            $countThisRule = 0
            
            foreach ($node in $textNodes)
            {
                if ($node.InnerText.Contains($wrong))
                {
                    $oldText = $node.InnerText
                    $newText = $oldText.Replace($wrong, $right)
                    $node.InnerText = $newText
                    $countThisRule++
                    $totalReplacements++
                }
            }
            
            if ($countThisRule -gt 0)
            {
                Write-Host "  [$countThisRule] '$wrong' -> '$right'" -ForegroundColor Green
            }
        }
        
        Write-Host ""
        Write-Host "Total replacements made: $totalReplacements" -ForegroundColor Cyan
        
        if ($totalReplacements -eq 0)
        {
            Write-Host "No replacements needed for this file" -ForegroundColor Yellow
            Remove-Item $tempFolder -Recurse -Force
            continue
        }
        
        # Save modified XML
        $xmlDoc.Save($docXmlPath)
        Write-Host "Saved modified document.xml"
        
        # Create output filename
        $outputPath = Join-Path $file.DirectoryName "$($file.BaseName)_fixed.docx"
        
        # Remove old output if exists
        if (Test-Path $outputPath)
        {
            Remove-Item $outputPath -Force
            Write-Host "Removed old _fixed.docx file"
        }
        
        # Compress back to docx
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempFolder, $outputPath)
        Write-Host "Created new file: $outputPath" -ForegroundColor Green
        
        # Clean up temp folder
        Remove-Item $tempFolder -Recurse -Force
        Write-Host "Cleaned up temp files"
        Write-Host ""
    }
    catch
    {
        Write-Host "Error processing file: $_" -ForegroundColor Red
        if (Test-Path $tempFolder)
        {
            Remove-Item $tempFolder -Recurse -Force
        }
    }
}

Write-Host ""
Write-Host "=== Processing Complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Modified files have '_fixed.docx' suffix"
Write-Host ""
pause