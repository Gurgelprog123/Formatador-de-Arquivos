
# Caminho da pasta de origem 
$src_path = Join-Path (Split-Path -parent $MyInvocation.MyCommand.Path) "arquivos_pptx"

# Caminho da pasta de destino 
$dst_path = Join-Path (Split-Path -parent $MyInvocation.MyCommand.Path) "arquivos_pdfs"

# Cria a pasta de destino caso não exista
if (!(Test-Path $dst_path)) {
    New-Item -ItemType Directory -Path $dst_path | Out-Null
}

# Cria objeto PowerPoint
$ppt_app = New-Object -ComObject PowerPoint.Application

# Lista para acumular os arquivos com erro
$erro_files = @()

# Procura todos os PPTX na pasta
Get-ChildItem -Path $src_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processando" $_.FullName "..."
    if (Test-Path $_.FullName) {
        $success = $false
        $attempts = 0

        while (-not $success -and $attempts -lt 2) {
            try {
                $attempts++
                $document = $ppt_app.Presentations.Open($_.FullName, $false, $false, $false)
                if ($document -ne $null) {
                    $pdf_filename = "$($dst_path)\$($_.BaseName).pdf"
                    $opt = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
                    $document.SaveAs($pdf_filename, $opt)
                    $document.Close()
                    $success = $true
                }
            }
            catch {
                if ($attempts -eq 2) {
                    $erro_files += $_.FullName
                }
            }
        }
    } else {
        $erro_files += $_.FullName
    }
}

# Salva os arquivos errados
if ($erro_files.Count -gt 0) {
    $error_log = Join-Path $dst_path "erros.txt"
    $erro_files | Out-File -FilePath $error_log -Encoding UTF8
}

# Finaliza PowerPoint
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
