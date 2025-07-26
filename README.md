# Automação de Geração de Grade de Carregamento

Este projeto automatiza a geração de planilhas filtradas por transportadora a partir de um Excel de cargas.

## 🔧 O que o script faz

- Verifica se o arquivo Excel está em uso
- Aguarda até ser liberado
- Identifica o dia seguinte útil (pula domingos)
- Copia a aba do dia e filtra apenas as linhas da transportadora com status válido
- Mantém toda a formatação (fontes, cores, largura de colunas)
- Salva a nova planilha em uma pasta monitorada
- O Power Automate detecta e envia o e-mail automaticamente com o anexo

## 📁 Tecnologias

- Python
- openpyxl
- win32com
- Power Automate

## 💡 Benefícios

- Redução de trabalho manual
- Aumento de confiabilidade
- Gatilho automatizado de envio para transportadora
