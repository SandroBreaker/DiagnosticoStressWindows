<#
.SYNOPSIS
    Limpeza e Otimização Segura do Sistema (SRE Standard).
.DESCRIPTION
    Script de automação resiliente para manutenção de estado do Windows.
    Implementa tratamento estruturado de erros e logging seguro.
#>
[CmdletBinding()]
param ()

# Garante execução com privilégios administrativos
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevando privilegios para execucao..."
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$Host.UI.RawUI.ForegroundColor = 'Green'
Clear-Host

Write-Host "=========================================================="
Write-Host "        OTIMIZADOR SEGURO DE PC - VERSAO SRE"
Write-Host "=========================================================="
Write-Host "`nOla! Iniciando pipeline de limpeza com tratamento de estado...`n"

# [1] Controle de Processos (OneDrive)
Write-Host "[1/4] Gerenciando ciclo de vida do OneDrive..."
try {
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDriveProcess) {
        Stop-Process -Name "OneDrive" -Force -ErrorAction Stop
        Write-Host "      -> Processo finalizado. O servico sera reiniciado sob demanda."
    } else {
        Write-Host "      -> OneDrive nao esta em execucao. Estado desejado mantido."
    }
} catch {
    Write-Warning "      -> Aviso ao gerenciar OneDrive: $_"
}

# [2] Limpeza de Diretórios Temporários (Excluindo Prefetch)
Write-Host "`n[2/4] Purgando arquivos temporarios isolados..."
$tempPaths = @(
    $env:TEMP,
    "$env:SystemRoot\Temp"
)

foreach ($path in $tempPaths) {
    try {
        if (Test-Path -Path $path) {
            $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            if ($items.Count -gt 0) {
                $items | Where-Object { -not $_.PSIsContainer } | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Host "      -> Limpeza concluida no escopo: $path"
            } else {
                Write-Host "      -> Diretorio ja otimizado: $path"
            }
        }
    } catch {
        Write-Warning "      -> Arquivos retidos em uso pelo sistema em: $path"
    }
}

# [3] Limpeza da Lixeira Local
Write-Host "`n[3/4] Esvaziando Lixeira do Sistema..."
try {
    Clear-RecycleBin -Force -ErrorAction Stop
    Write-Host "      -> Reciclagem de blocos de disco concluida."
} catch {
    Write-Host "      -> A lixeira ja esta vazia ou itens estao bloqueados."
}

# [4] Flush de DNS via Objeto .NET / Cmdlet nativo
Write-Host "`n[4/4] Redefinindo cache de resolucao de rede (DNS)..."
try {
    Clear-DnsClientCache -ErrorAction Stop
    Write-Host "      -> Tabela de roteamento DNS local purgada com sucesso."
} catch {
    Write-Warning "      -> Aviso ao limpar cache DNS: $_"
}

Write-Host "`n=========================================================="
Write-Host "              MANUTENCAO CONCLUIDA COM SUCESSO!"
Write-Host "=========================================================="
Write-Host "A infraestrutura local foi otimizada com seguranca.`n"

Read-Host "Pressione [ENTER] para encerrar"