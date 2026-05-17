# ─────────────────────────────────────────────
#  Ferramenta Word — Instalador para Windows
# ─────────────────────────────────────────────

$MANIFEST_URL = "https://leocorrea99.github.io/ferramenta-word/manifest.xml"
$CATALOG_DIR  = "$env:APPDATA\FerramantaWord"
$MANIFEST_PATH = "$CATALOG_DIR\ferramenta-word.xml"

Clear-Host
Write-Host ""
Write-Host "  ┌──────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │      Ferramenta Word             │" -ForegroundColor Cyan
Write-Host "  │      Instalador — Windows        │" -ForegroundColor Cyan
Write-Host "  └──────────────────────────────────┘" -ForegroundColor Cyan
Write-Host ""

# Fecha o Word se estiver aberto
$wordProc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProc) {
    Write-Host "  → Fechando o Microsoft Word..." -ForegroundColor Yellow
    $wordProc | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# Cria a pasta do catálogo
New-Item -ItemType Directory -Force -Path $CATALOG_DIR | Out-Null

# Baixa o manifest
Write-Host "  → Baixando manifest..."
try {
    Invoke-WebRequest -Uri $MANIFEST_URL -OutFile $MANIFEST_PATH -UseBasicParsing
} catch {
    Write-Host ""
    Write-Host "  ✗  Erro ao baixar. Verifique sua internet." -ForegroundColor Red
    Read-Host "  Pressione Enter para fechar"
    exit 1
}

# Registra a pasta como catálogo confiável no Word (via Registry)
$regBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$guid = [System.Guid]::NewGuid().ToString("B").ToUpper()

try {
    New-Item -Path "$regBase\$guid" -Force | Out-Null
    Set-ItemProperty -Path "$regBase\$guid" -Name "Id"    -Value $guid
    Set-ItemProperty -Path "$regBase\$guid" -Name "Url"   -Value $CATALOG_DIR
    Set-ItemProperty -Path "$regBase\$guid" -Name "Flags" -Value 1 -Type DWord
} catch {
    Write-Host "  Aviso: não foi possível registrar automaticamente." -ForegroundColor Yellow
    Write-Host "  Adicione manualmente em: Word → Opções → Central de Confiabilidade"
    Write-Host "  → Catálogos de Suplementos Confiáveis → URL: $CATALOG_DIR"
}

Write-Host ""
Write-Host "  ✓  Instalado com sucesso!" -ForegroundColor Green
Write-Host ""
Write-Host "  Como abrir:"
Write-Host "  Word → Inserir → Suplementos"
Write-Host "  Aba ""Pasta Compartilhada"" → Ferramenta Word"
Write-Host ""

$abrir = Read-Host "  Abrir o Word agora? (s/n)"
if ($abrir -match "^[Ss]$") {
    Start-Process "WINWORD.EXE"
}

Write-Host ""
Read-Host "  Pressione Enter para fechar"
