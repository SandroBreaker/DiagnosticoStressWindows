# 🚀 Diagnóstico Robusto de Stress - Windows (v7)

Este é um utilitário avançado em PowerShell para diagnóstico de performance e análise de "stress" em sistemas Windows. Ele coleta métricas profundas de hardware, sistema operacional e processos em tempo real, gerando relatórios detalhados para identificação de gargalos (bottlenecks).

## ✨ Principais Funcionalidades

- **Métricas de Sistema (Snapshot):** Informações detalhadas da CPU, RAM, Discos (incluindo tipo de mídia SSD/HDD), Versão do Windows, Uptime e Plano de Energia.
- **Análise de Processos (Delta CPU):** Diferencia o consumo real de CPU durante o período de amostragem, permitindo ver quem realmente está "pesando" no sistema.
- **Monitoramento de Navegadores:** Captura específica do consumo de recursos (RAM/CPU) de instâncias do Edge, Chrome, Firefox, Brave, Opera, etc.
- **Timeline de Performance:** Gera dados de série temporal (timeseries) para CPU, RAM e Disco, permitindo analisar picos durante o teste.
- **Mapeamento de Serviços:** Identifica quais serviços do Windows estão rodando dentro de processos compartilhados (ex: `svchost.exe`).
- **Relatório Visual (Viewer):** Gera um painel interativo em HTML para visualização rica das métricas coletadas.
- **Exportação:** Suporte a CSV (para timeseries) e JSON (relatório completo).
- **Gerenciamento de Logs:** Opções de retenção automática (arquivamento ou limpeza de execuções antigas).

## 🚀 Como Utilizar

Abra o PowerShell como **Administrador** no diretório do script e execute:

### Execução Padrão (2 minutos de amostragem)
```powershell
.\DiagnosticoStressWindows.v7.ps1
```

### Modo Rápido (Amostra de 2 segundos)
Ideal para verificações instantâneas.
```powershell
.\DiagnosticoStressWindows.v7.ps1 -Quick
```

### Incluindo Detalhes de Serviços
Lista quais serviços estão atrelados aos processos que mais consomem CPU.
```powershell
.\DiagnosticoStressWindows.v7.ps1 -IncludeServices
```

### Amostragem Customizada
Define o tempo e a quantidade de processos no topo.
```powershell
.\DiagnosticoStressWindows.v7.ps1 -SampleSeconds 60 -Top 15
```

### Manutenção de Logs
Para manter apenas as execuções dos últimos 5 dias e limpar o resto:
```powershell
.\DiagnosticoStressWindows.v7.ps1 -PruneOldRuns -RetentionDays 5
```

## 📂 Estrutura de Saída

Cada execução cria uma pasta no formato `run-YYYYMMDD-HHMMSS` contendo:
- `system-snapshot.json`: Dados técnicos do hardware e OS.
- `system-timeseries.csv`: Histórico de uso de recursos do sistema.
- `top-process-timeseries.csv`: Histórico de uso dos processos mais pesados.
- `browser-processes.csv`: Detalhamento específico de processos de navegadores.

## 🎨 Visualizador (Viewer)

O projeto inclui o **DiagnosticoStressWindows_Viewer**. Ao final de cada execução, o script perguntará se você deseja gerar o relatório visual.
- O relatório será gerado em: `DiagnosticoStressWindows_Viewer\viewer-report.html`.
- Você também pode usar o `Gerar_Viewer_Diagnostico.cmd` para atualizar o viewer manualmente.

## 🛠️ Integração com Menu de Contexto

Você pode adicionar o Diagnóstico ao menu de contexto do Windows (clicar com o botão direito em uma pasta) usando o arquivo:
`Atualizar_MenuContexto_DiagnosticoStressWindows_RaizFixa.reg`

---
*Desenvolvido para análise técnica profunda e suporte a hardware/software Windows.*
