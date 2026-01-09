# DocDigitizer PowerShell Module

A PowerShell module for interacting with the DocDigitizer document processing API.

## Installation

### Option 1: Clone and Import

```powershell
# Clone the repository
git clone https://github.com/DocDigitizer/dd-cmdlets.git

# Import the module
Import-Module ".\dd-cmdlets\DocDigitizer.psd1"
```

### Option 2: Install to PowerShell Modules folder

```powershell
# Clone directly to user modules folder
git clone https://github.com/DocDigitizer/dd-cmdlets.git "$env:USERPROFILE\Documents\PowerShell\Modules\DocDigitizer"

# Then import (no path needed)
Import-Module DocDigitizer
```

### Option 3: Download ZIP

1. Download from https://github.com/DocDigitizer/dd-cmdlets/archive/refs/heads/master.zip
2. Extract to a folder of your choice
3. Import the module:

```powershell
Import-Module "C:\path\to\dd-cmdlets\DocDigitizer.psd1"
```

## Quick Start

```powershell
# Import the module
Import-Module .\dd-cmdlets\DocDigitizer.psd1

# Test connection
Test-DocDigitizerConnection

# Process a document (returns JSON)
Send-DocDigitizerDocument -FilePath "invoice.pdf"

# Process and auto-save to invoice_extraction.json
Send-DocDigitizerDocument -FilePath "invoice.pdf" -SaveExtraction
```

## Available Commands

| Command                       | Description                        |
| ----------------------------- | ---------------------------------- |
| `Send-DocDigitizerDocument`   | Send a PDF document for processing |
| `Test-DocDigitizerConnection` | Test API connectivity              |
| `Get-DocDigitizerConfig`      | View current configuration         |
| `Set-DocDigitizerConfig`      | Set default configuration values   |
| `Get-DocDigitizerHelp`        | Show help and usage examples       |

## Usage Examples

### Basic Document Processing

```powershell
# Process a single document (returns JSON)
Send-DocDigitizerDocument -FilePath "invoice.pdf"

# Process with specific IDs
Send-DocDigitizerDocument -FilePath "invoice.pdf" -DocumentId $docId -ContextId $contextId
```

### Saving Results

```powershell
# Auto-save to {filename}_extraction.json in same directory
Send-DocDigitizerDocument -FilePath "invoice.pdf" -SaveExtraction
# Creates: invoice_extraction.json

# Save to custom path
Send-DocDigitizerDocument -FilePath "invoice.pdf" -OutputPath "C:\results\my_result.json"

# Batch process and save all
Get-ChildItem *.pdf | Send-DocDigitizerDocument -SaveExtraction
```

### Pipeline Selection

```powershell
# Use a specific pipeline
Send-DocDigitizerDocument -FilePath "invoice.pdf" -Pipeline "MainPipelineWithOCR"

# Available pipelines (check your server):
# - MainPipelineWithOCR
# - MainPipelineWithFile
# - SingleDocPipelineWithOCR
# - SingleDocPipelineWithFile
```

### Response Verbosity

```powershell
# Minimal response (default) - just essential data
Send-DocDigitizerDocument -FilePath "invoice.pdf"

# Medium response - includes metadata
Send-DocDigitizerDocument -FilePath "invoice.pdf" -LogLevel Medium

# Full response - complete execution details
Send-DocDigitizerDocument -FilePath "invoice.pdf" -LogLevel Full

# Get raw API response
Send-DocDigitizerDocument -FilePath "invoice.pdf" -PassThru
```

### Batch Processing

```powershell
# Process all PDFs in a folder with same context
$contextId = [guid]::NewGuid()
Get-ChildItem *.pdf | Send-DocDigitizerDocument -ContextId $contextId

# Process and collect results
$results = Get-ChildItem *.pdf | Send-DocDigitizerDocument -ContextId $contextId
$results | Export-Csv -Path "results.csv"
```

### Health Check

```powershell
# Detailed connection test
Test-DocDigitizerConnection

# Output:
# Url       : http://localhost:5000
# Connected : True
# Response  : I am alive
# Error     :
# Latency   : 45

# Simple boolean check
if (Test-DocDigitizerConnection -Quiet) {
    Send-DocDigitizerDocument -FilePath "invoice.pdf"
}
```

## Configuration

### Environment Variables

| Variable                | Description               | Default                 |
| ----------------------- | ------------------------- | ----------------------- |
| `DOCDIGITIZER_URL`      | API base URL              | `https://apix.docdigitizer.com/sync` |
| `DOCDIGITIZER_APIKEY`   | API key for authentication | (built-in default)     |
| `DOCDIGITIZER_PIPELINE` | Default pipeline name     | (server default)        |
| `DOCDIGITIZER_LOGLEVEL` | Default log level         | `Minimal`               |
| `DOCDIGITIZER_TIMEOUT`  | Request timeout (seconds) | `300`                   |

### Setting Configuration

```powershell
# Set for current session
Set-DocDigitizerConfig -BaseUrl "https://apix.docdigitizer.com/sync" -Pipeline "MainPipelineWithOCR"

# Persist to PowerShell profile (survives session restarts)
Set-DocDigitizerConfig -BaseUrl "https://apix.docdigitizer.com/sync" -Persist

# View current configuration
Get-DocDigitizerConfig
```

### Using Different Environments

```powershell
# Development
Set-DocDigitizerConfig -BaseUrl "http://localhost:5000"

# Production
Set-DocDigitizerConfig -BaseUrl "https://docingester-prod.example.com"

# One-off call to different server
Send-DocDigitizerDocument -FilePath "doc.pdf" -BaseUrl "https://other-server.com"
```

## Output Format

The `Send-DocDigitizerDocument` command returns a **JSON string**:

```json
{
  "traceId": "ABC1234",
  "state": "PROCESSINGX",
  "pipeline": "MainPipelineWithOCR",
  "pageCount": 3,
  "filePath": "C:\\docs\\invoice.pdf",
  "documentId": "a1b2c3d4-...",
  "contextId": "e5f6g7h8-...",
  "output": {
    "extractions": [...]
  },
  "timers": {
    "DocIngester_Total": 1234.56,
    "DocWorker_Total": 1200.00
  }
}
```

### Working with JSON output

```powershell
# Parse JSON to object if needed
$json = Send-DocDigitizerDocument -FilePath "invoice.pdf"
$result = $json | ConvertFrom-Json

# Access specific fields
$result.output.extractions
$result.timers

# Save to file automatically
Send-DocDigitizerDocument -FilePath "invoice.pdf" -SaveExtraction
```

## Error Handling

```powershell
# Errors are written to the error stream
$result = Send-DocDigitizerDocument -FilePath "invalid.pdf" -ErrorAction SilentlyContinue
if (-not $result) {
    Write-Host "Processing failed"
}

# Or use try/catch with -ErrorAction Stop
try {
    $result = Send-DocDigitizerDocument -FilePath "doc.pdf" -ErrorAction Stop
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
```

## Requirements

- PowerShell 5.1 or later (PowerShell 7+ recommended for better performance)
- Network access to DocIngester API

## Troubleshooting

### Connection Issues

```powershell
# Test connectivity
Test-DocDigitizerConnection -Verbose

# Check configuration
Get-DocDigitizerConfig
```

### Timeout Issues

```powershell
# Increase timeout for large documents
Send-DocDigitizerDocument -FilePath "large.pdf" -TimeoutSec 600

# Or set as default
Set-DocDigitizerConfig -Timeout 600
```

### Debug Mode

```powershell
# Enable verbose output
Send-DocDigitizerDocument -FilePath "doc.pdf" -Verbose
```

## License

MIT License - See LICENSE file for details.
