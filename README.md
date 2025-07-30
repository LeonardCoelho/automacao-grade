# 🚛 Automação de Geração de Grade de Carregamento

Este projeto automatiza a criação de uma planilha filtrada para uma transportadora específica (MVC), baseada em um arquivo Excel de ordens de coleta, com formatação mantida, validações aplicadas e integração com Power Automate.

---

## 🔧 O que o script faz

- Verifica se o arquivo Excel está em uso
- Aguarda até o arquivo ser liberado
- Identifica o próximo dia útil (pula domingos)
- Copia a aba referente ao dia
- Filtra apenas as linhas da transportadora "MVC" com status válido
- Mantém toda a formatação original (fontes, cores, largura de colunas)
- Salva a nova planilha em uma pasta monitorada
- O Power Automate detecta e envia o e-mail automaticamente com o anexo

---

## 📁 Estrutura do Projeto

```
automacao-grade/
├── data/
│   └── ORDEM_DE_COLETA.xlsx
├── images/
│   └── Print planilha final.jpg
├── src/
│   └── atualizar_grade_MVC.py
├── README.md
└── requirements.txt
```

---

## ⚙️ Tecnologias Utilizadas

- Python 3.10+
- `openpyxl` para manipulação do Excel com formatação
- `win32com` para detectar se o Excel está aberto
- `Power Automate` para envio automático

---

## 💡 Benefícios

- Redução de trabalho manual
- Aumento da confiabilidade no envio
- Gatilho automatizado de e-mail
- Garante que só registros válidos sejam processados

---

## 📝 Observação

A pasta de destino é monitorada pelo Power Automate no OneDrive. Após a geração da planilha filtrada, o fluxo dispara o envio automático para a transportadora correta.

---

## ✉️ Contato

Desenvolvido por [Leonardo Coelho](https://github.com/LeonardCoelho) 🚀
