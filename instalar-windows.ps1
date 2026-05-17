# Ferramenta Word - Instalador para Windows

$MANIFEST_URL  = "https://leocorrea99.github.io/ferramenta-word/manifest.xml"
$MANIFEST_DIR  = "$env:APPDATA\FerramantaWord"
$MANIFEST_PATH = "$MANIFEST_DIR\ferramenta-word.xml"

Clear-Host
Write-Host ""
Write-Host "  +----------------------------------+" -ForegroundColor Cyan
Write-Host "  |      Ferramenta Word             |" -ForegroundColor Cyan
Write-Host "  |      Instalador - Windows        |" -ForegroundColor Cyan
Write-Host "  +----------------------------------+" -ForegroundColor Cyan
Write-Host ""

# Fecha o Word se estiver aberto
$wordProc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProc) {
    Write-Host "  -> Fechando o Microsoft Word..." -ForegroundColor Yellow
    $wordProc | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# Cria a pasta e baixa o manifest
Write-Host "  -> Baixando manifest..."
New-Item -ItemType Directory -Force -Path $MANIFEST_DIR | Out-Null

try {
    Invoke-WebRequest -Uri $MANIFEST_URL -OutFile $MANIFEST_PATH -UseBasicParsing
} catch {
    Write-Host ""
    Write-Host "  ERRO: Nao foi possivel baixar. Verifique sua internet." -ForegroundColor Red
    Read-Host "  Pressione Enter para fechar"
    exit 1
}

# Registra como Suplemento do Desenvolvedor (equivalente a pasta wef do Mac)
Write-Host "  -> Registrando suplemento..."

$devKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
try {
    New-Item -Path $devKey -Force | Out-Null
    New-ItemProperty -Path $devKey -Name "FerramontaWord" -Value $MANIFEST_PATH -PropertyType String -Force | Out-Null
} catch {
    Write-Host "  Aviso: registro principal falhou, usando metodo alternativo..." -ForegroundColor Yellow
}

# Metodo alternativo: catalogo confiavel (fallback)
$regBase = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$guid = [System.Guid]::NewGuid().ToString("B").ToUpper()
try {
    New-Item -Path "$regBase\$guid" -Force | Out-Null
    Set-ItemProperty -Path "$regBase\$guid" -Name "Id"    -Value $guid
    Set-ItemProperty -Path "$regBase\$guid" -Name "Url"   -Value $MANIFEST_DIR
    Set-ItemProperty -Path "$regBase\$guid" -Name "Flags" -Value 1 -Type DWord
} catch {}

Write-Host ""
Write-Host "  OK! Instalado com sucesso!" -ForegroundColor Green
Write-Host ""
Write-Host "  Para abrir no Word:" -ForegroundColor White
Write-Host "  Inserir -> Suplementos -> Suplementos do Desenvolvedor" -ForegroundColor Gray
Write-Host ""

$abrir = Read-Host "  Abrir o Word agora? (s/n)"
if ($abrir -match "^[Ss]$") {
    Start-Process "WINWORD.EXE"
    Write-Host ""
    Write-Host "  Aguarde o Word abrir e va em:" -ForegroundColor Cyan
    Write-Host "  Inserir -> Suplementos -> Suplementos do Desenvolvedor" -ForegroundColor Cyan
}

Write-Host ""
Read-Host "  Pressione Enter para fechar"
