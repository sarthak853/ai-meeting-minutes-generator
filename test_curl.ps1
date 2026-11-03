$ProgressPreference = 'SilentlyContinue'
# Launch server in background
$serverProc = Start-Process -FilePath ".\venv\Scripts\python.exe" -ArgumentList "app.py" -PassThru
Start-Sleep -Seconds 3

try {
    # 1. Test GET /
    Write-Host "Testing GET /..."
    $indexRes = Invoke-WebRequest -Uri "http://127.0.0.1:5000/" -UseBasicParsing
    Write-Host "GET / Status: $($indexRes.StatusCode)"

    # 2. Test POST /generate via curl
    Write-Host "`nTesting POST /generate with sample_transcript.txt..."
    $body = @{ transcript = (Get-Content "sample_transcript.txt" -Raw) } | ConvertTo-Json
    $generateRes = Invoke-RestMethod -Uri "http://127.0.0.1:5000/generate" -Method POST -Body $body -ContentType "application/json"
    
    Write-Host "Title: $($generateRes.title)"
    Write-Host "Decisions:"
    $generateRes.decisions | ForEach-Object { Write-Host " - [$($_.status)] $($_.description)" }
    Write-Host "Action Items:"
    $generateRes.action_items | ForEach-Object { Write-Host " - [$($_.owner)] $($_.task) (Deadline: $($_.deadline))" }

    # 3. Test POST /export/pdf
    Write-Host "`nTesting POST /export/pdf..."
    $pdfPayload = $generateRes | ConvertTo-Json -Depth 5
    $pdfFile = "test_output.pdf"
    Invoke-WebRequest -Uri "http://127.0.0.1:5000/export/pdf" -Method POST -Body $pdfPayload -ContentType "application/json" -OutFile $pdfFile
    $pdfBytes = [System.IO.File]::ReadAllBytes($pdfFile)
    $pdfHeader = [System.Text.Encoding]::ASCII.GetString($pdfBytes[0..4])
    Write-Host "PDF Size: $($pdfBytes.Length) bytes, Header: $pdfHeader"

    # 4. Test POST /export/docx
    Write-Host "`nTesting POST /export/docx..."
    $docxFile = "test_output.docx"
    Invoke-WebRequest -Uri "http://127.0.0.1:5000/export/docx" -Method POST -Body $pdfPayload -ContentType "application/json" -OutFile $docxFile
    $docxBytes = [System.IO.File]::ReadAllBytes($docxFile)
    $docxHeader = [System.Text.Encoding]::ASCII.GetString($docxBytes[0..3])
    Write-Host "DOCX Size: $($docxBytes.Length) bytes, Header: $docxHeader"

    # Assertions
    if ($pdfHeader -eq "%PDF-") {
        Write-Host "SUCCESS: Valid %PDF- signature verified!" -ForegroundColor Green
    } else {
        Write-Host "FAILED: PDF header mismatch" -ForegroundColor Red
    }

    if ($docxHeader -eq "PK`0`0" -or $docxHeader.StartsWith("PK")) {
        Write-Host "SUCCESS: Valid PK.. DOCX zip signature verified!" -ForegroundColor Green
    } else {
        Write-Host "FAILED: DOCX header mismatch" -ForegroundColor Red
    }

} finally {
    Stop-Process -Id $serverProc.Id -Force -ErrorAction SilentlyContinue
    Remove-Item "test_output.pdf", "test_output.docx" -ErrorAction SilentlyContinue
}
