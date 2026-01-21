function Send-DocDigitizerDocument {
    <#
    .SYNOPSIS
        Sends a document to the DocDigitizer API for processing.

    .DESCRIPTION
        Uploads a PDF or image file to the DocDigitizer pipeline for OCR, classification,
        and data extraction. Returns JSON results from the processing pipeline.

        Supported file types:
        - PDF files (.pdf)
        - Image files (.jpg, .jpeg, .png, .tiff, .tif, .bmp, .webp, .gif)

        For OCR-only processing (without data extraction), use -Pipeline "OCROnlyPipeline".

    .PARAMETER FilePath
        Path to the file to process. Accepts PDF and image files. Accepts pipeline input.

    .PARAMETER DocumentId
        Unique identifier for the document (GUID). If not provided, a new GUID is generated.

    .PARAMETER ContextId
        Context identifier for grouping related documents (GUID). If not provided, a new GUID is generated.

    .PARAMETER Pipeline
        Name of the pipeline to execute. Overrides the default pipeline.
        Examples: MainPipelineWithOCR, MainPipelineWithFile, OCROnlyPipeline

        OCROnlyPipeline: Returns OCR text, word positions, and line structure without
        data extraction. Output is in output.documents[] instead of output.extractions[].

    .PARAMETER LogLevel
        Response verbosity level. Valid values: Minimal, Medium, Full.
        - Minimal: Only essential data (default)
        - Medium: Include metadata and basic execution info
        - Full: Complete execution details including all plugin outputs

    .PARAMETER BaseUrl
        Base URL of the DocDigitizer API. Defaults to environment variable
        DOCDIGITIZER_URL or the built-in default (https://apix.docdigitizer.com/sync).

    .PARAMETER TimeoutSec
        Request timeout in seconds. Defaults to 300 (5 minutes).

    .PARAMETER SaveExtraction
        Automatically saves the JSON result to a file named {originalFilename}_extraction.json
        in the same directory as the input file.

    .PARAMETER OutputPath
        Custom path to save the JSON result. Overrides -SaveExtraction location.

    .PARAMETER Depth
        JSON serialization depth. Defaults to 20 for deep nested objects.

    .PARAMETER IncludeRawData
        When using OCROnlyPipeline, includes detailed raw data in the output:
        words with bounding boxes, token mappings, raw Google Vision response, and page images.
        By default, only essential fields (fullText, wordCount, lines, metadata) are returned.

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "invoice.pdf"

        Processes invoice.pdf and returns JSON result.

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "invoice.pdf" -SaveExtraction

        Processes invoice.pdf and saves result to invoice_extraction.json

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "invoice.pdf" -OutputPath "C:\results\invoice_result.json"

        Processes invoice.pdf and saves result to specified path.

    .EXAMPLE
        Get-ChildItem *.pdf | Send-DocDigitizerDocument -SaveExtraction

        Processes all PDFs and saves each result alongside the original file.

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "document.png" -Pipeline "OCROnlyPipeline"

        Performs OCR-only processing on an image file, returning text and word positions.

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "scan.jpg" -Pipeline "OCROnlyPipeline" -SaveExtraction

        Performs OCR on a JPEG image and saves the result to scan_extraction.json.

    .EXAMPLE
        Send-DocDigitizerDocument -FilePath "scan.png" -Pipeline "OCROnlyPipeline" -IncludeRawData

        Performs OCR and includes all raw data (words, bounding boxes, token mappings, page images).

    .OUTPUTS
        JSON string containing the extraction results and metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [ValidateScript({
            if (-not (Test-Path $_)) { throw "File not found: $_" }
            $validExtensions = @('.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.webp', '.gif')
            $ext = [System.IO.Path]::GetExtension($_).ToLower()
            if ($ext -notin $validExtensions) { throw "Unsupported file type: $_. Supported: PDF, JPEG, PNG, TIFF, BMP, WebP, GIF" }
            $true
        })]
        [string]$FilePath,

        [Parameter()]
        [guid]$DocumentId = [guid]::NewGuid(),

        [Parameter()]
        [guid]$ContextId = [guid]::NewGuid(),

        [Parameter()]
        [string]$Pipeline,

        [Parameter()]
        [ValidateSet('Minimal', 'Medium', 'Full')]
        [string]$LogLevel,

        [Parameter()]
        [string]$BaseUrl,

        [Parameter()]
        [int]$TimeoutSec,

        [Parameter()]
        [switch]$SaveExtraction,

        [Parameter()]
        [string]$OutputPath,

        [Parameter()]
        [int]$Depth = 20,

        [Parameter()]
        [switch]$IncludeRawData
    )

    begin {
        $config = Get-ModuleConfig

        # Use provided values or fall back to config
        if (-not $BaseUrl) { $BaseUrl = $config.BaseUrl }
        if (-not $Pipeline -and $config.Pipeline) { $Pipeline = $config.Pipeline }
        if (-not $LogLevel -and $config.LogLevel) { $LogLevel = $config.LogLevel }
        if (-not $TimeoutSec) { $TimeoutSec = $config.Timeout }
    }

    process {
        $fullPath = Resolve-Path $FilePath -ErrorAction Stop
        $fileName = Split-Path $fullPath -Leaf

        Write-CommandLog -CommandName 'Send-DocDigitizerDocument' -Message "Processing '$fileName'"
        Write-Verbose "Processing: $fullPath"
        Write-Verbose "DocumentId: $DocumentId"
        Write-Verbose "ContextId: $ContextId"
        Write-Verbose "Pipeline: $(if ($Pipeline) { $Pipeline } else { '(default)' })"
        Write-Verbose "BaseUrl: $BaseUrl"

        # Build headers
        $headers = @{}
        if ($config.ApiKey) {
            $headers['x-api-key'] = $config.ApiKey
        }
        if ($Pipeline) {
            $headers['X-DD-Pipeline'] = $Pipeline
            Write-Verbose "Header X-DD-Pipeline: $Pipeline"
        }
        if ($LogLevel) {
            $headers['X-DD-LogLevel'] = $LogLevel
            Write-Verbose "Header X-DD-LogLevel: $LogLevel"
        }

        try {
            # PowerShell 7+ supports -Form parameter for multipart
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $form = @{
                    files     = Get-Item $fullPath
                    id        = $DocumentId.ToString()
                    contextID = $ContextId.ToString()
                }

                $response = Invoke-RestMethod -Uri $BaseUrl -Method Post -Form $form -Headers $headers -TimeoutSec $TimeoutSec
            }
            else {
                # Fallback for PowerShell 5.1
                $multipart = Build-MultipartForm -FilePath $fullPath -DocumentId $DocumentId -ContextId $ContextId
                $response = Invoke-RestMethod -Uri $BaseUrl -Method Post -Body $multipart.Body -ContentType $multipart.ContentType -Headers $headers -TimeoutSec $TimeoutSec
            }

            # Check for error state
            $stateText = if ($response.stateText) { $response.stateText } else { $response.StateText }
            if ($stateText -eq 'ERROR') {
                $errorMsg = if ($response.messages) { $response.messages -join '; ' } elseif ($response.Messages) { $response.Messages -join '; ' } else { 'Unknown error' }
                $traceId = if ($response.traceId) { $response.traceId } else { $response.TraceId }
                Write-Error "DocDigitizer processing failed: $errorMsg (TraceId: $traceId)"
                return
            }

            # Build result object with metadata
            $traceId = if ($response.traceId) { $response.traceId } else { $response.TraceId }
            $pipelineName = if ($response.pipeline) { $response.pipeline } else { $response.Pipeline }
            $pageCount = if ($null -ne $response.numberPages) { $response.numberPages } else { $response.NumberPages }
            $output = if ($response.output) { $response.output } else { $response.Output }
            $timers = if ($response.timers) { $response.timers } else { $response.Timers }

            # Filter out raw data from OCROnlyPipeline output unless explicitly requested
            if (-not $IncludeRawData -and $output.documents) {
                foreach ($doc in $output.documents) {
                    # Remove pageImageBase64
                    if ($doc.PSObject.Properties['pageImageBase64']) {
                        $doc.PSObject.Properties.Remove('pageImageBase64')
                    }
                    # Remove rawOCRResponse
                    if ($doc.PSObject.Properties['rawOCRResponse']) {
                        $doc.PSObject.Properties.Remove('rawOCRResponse')
                    }
                    # Remove words array from ocr (keep fullText and wordCount)
                    if ($doc.ocr -and $doc.ocr.PSObject.Properties['words']) {
                        $doc.ocr.PSObject.Properties.Remove('words')
                    }
                    # Remove tokensByLine and linesByToken from textLines (keep lines)
                    if ($doc.textLines) {
                        if ($doc.textLines.PSObject.Properties['tokensByLine']) {
                            $doc.textLines.PSObject.Properties.Remove('tokensByLine')
                        }
                        if ($doc.textLines.PSObject.Properties['linesByToken']) {
                            $doc.textLines.PSObject.Properties.Remove('linesByToken')
                        }
                    }
                }
            }
            $pluginExecutions = if ($response.plugin_executions) { $response.plugin_executions } else { $response.PluginExecutions }

            $result = [ordered]@{
                traceId          = $traceId
                state            = $stateText
                pipeline         = $pipelineName
                pageCount        = $pageCount
                filePath         = $fullPath.Path
                documentId       = $DocumentId.ToString()
                contextId        = $ContextId.ToString()
                output           = $output
                timers           = $timers
            }

            # Include plugin executions if available (when LogLevel is set)
            if ($pluginExecutions) {
                $result.pluginExecutions = $pluginExecutions
            }

            # Convert to JSON and format nicely
            $json = $result | ConvertTo-Json -Depth $Depth -Compress | Format-Json

            # Determine output file path
            $savePath = $null
            if ($OutputPath) {
                $savePath = $OutputPath
            }
            elseif ($SaveExtraction) {
                $directory = Split-Path $fullPath.Path -Parent
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath.Path)
                $savePath = Join-Path $directory "${baseName}_extraction.json"
            }

            # Save to file if requested
            if ($savePath) {
                $json | Out-File -FilePath $savePath -Encoding UTF8 -Force
                Write-Verbose "Saved extraction to: $savePath"
                Write-Host "Saved: $savePath" -ForegroundColor Green
            }

            Write-CommandLog -CommandName 'Send-DocDigitizerDocument' -Message "Completed '$fileName' (TraceId: $traceId)" -Status Complete

            # Return JSON
            return $json
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = $_.Exception.Response.StatusCode.value__
            }
            $errorMessage = "Failed to process document '$fullPath': $($_.Exception.Message)"
            if ($statusCode) {
                $errorMessage += " (HTTP $statusCode)"
            }
            Write-Error $errorMessage
        }
    }
}
