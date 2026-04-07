# Diagnostico Stress Windows Viewer v4

## O que mudou

- interface refeita com hierarquia visual decente
- resumo executivo por run
- stress score com contexto
- gráficos explicados
- tabelas com barras de intensidade
- comparativo entre runs em cards
- sem CDN e sem servidor

## Como usar

1. extraia esta pasta em `C:\dev\DiagnosticoStressWindows\DiagnosticoStressWindows_Viewer`
2. rode `Gerar_Viewer_Diagnostico.cmd`
3. o script vai ler sempre a raiz fixa `C:\dev\DiagnosticoStressWindows`
4. o arquivo `viewer-report.html` será gerado/atualizado na própria pasta do viewer

## O que ele lê

Cada run precisa ter:

- `system-snapshot.json`
- `system-timeseries.csv`
- `top-process-timeseries.csv`
- `browser-processes.csv` (opcional)

## Observação

Se o `DiagnosticoStressWindows.v7.ps1` perguntar no fim da coleta se deseja gerar o viewer, responder `S` já atualiza esse mesmo HTML.
