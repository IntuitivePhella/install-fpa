# Salva um print da TELA PRIMARIA inteira em PNG, em resolucao real (o script se declara DPI-aware, entao
# escala de 125%/150% nao corta nem encolhe a imagem). Usado como prova de Localidade/filtros.
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\capturar_tela.ps1 -Destino C:\pasta\cp.png
# Para capturar todos os monitores (area virtual), acrescente -TodosMonitores.
# O destino deve ficar em PAPAKI_DADOS_DIR (fora do repositorio): o print pode mostrar dados de outros aplicativos.
# Nao imprime nada do conteudo da tela; so grava o arquivo e devolve erro (codigo 1) se algo falhar.
param(
    [Parameter(Mandatory = $true)][string]$Destino,
    [switch]$TodosMonitores
)
$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    Add-Type -Namespace Papaki -Name Dpi -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool SetProcessDPIAware();'
    [void][Papaki.Dpi]::SetProcessDPIAware()   # antes de ler qualquer medida de tela
    $pasta = Split-Path -Parent ([System.IO.Path]::GetFullPath($Destino))
    if ($pasta -and -not (Test-Path -LiteralPath $pasta)) { throw "pasta de destino nao existe" }
    if ($TodosMonitores) {
        $tela = [System.Windows.Forms.SystemInformation]::VirtualScreen
    } else {
        $tela = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    }
    $bmp = New-Object System.Drawing.Bitmap $tela.Width, $tela.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($tela.Left, $tela.Top, 0, 0, $bmp.Size)
    $bmp.Save($Destino, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
} catch {
    [Console]::Error.WriteLine("ERRO: nao foi possivel salvar o print ($($_.Exception.GetType().Name))")
    exit 1
}
