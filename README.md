# 🚛 Automação de Geração de Grade de Carregamento

Automação em Python que gera uma planilha filtrada para uma transportadora específica, com base em um arquivo de ordens de coleta. Mantém toda a **formatação original**, aplica **validações** e **integra com Power Automate** para envio automático por e-mail.

## 🧠 Objetivo

Reduzir retrabalho manual e garantir que **apenas os dados válidos** de carregamento sejam enviados à transportadora, de forma **rápida, confiável e automatizada**.

## 🔧 O que o script faz

1. Verifica se o Excel está aberto (e aguarda liberar).  
2. Identifica o próximo dia útil (pula domingos).  
3. Copia a aba correspondente ao próximo dia.  
4. Filtra apenas os registros da transportadora + status válido.  
5. Mantém toda a **formatação original** (fontes, cores, larguras...).  
6. Salva a nova planilha numa pasta monitorada.  
7. O **Power Automate** detecta e envia o e-mail com o anexo.

## 📂 Estrutura do Projeto

```
automacao-grade/
├── data/
│   └── ORDEM_DE_COLETA.xlsx           # Planilha base
├── images/
│   └── print-planilha-final.jpg       # Exemplo do resultado
├── src/
│   └── atualizar_grade.py         # Script principal
├── README.md
└── requirements.txt                   # Bibliotecas necessárias
```

## ⚙️ Tecnologias e Bibliotecas

- 🐍 Python 3.10+  
- 📊 openpyxl – manipulação de Excel com formatação  
- 🖥️ pywin32 – detecção de planilhas abertas  
- ☁️ Power Automate (OneDrive) – envio automático

## ✅ Benefícios

- 🚫 Elimina etapas manuais repetitivas  
- 📈 Aumenta a confiabilidade da grade  
- 📬 Envio automático por gatilho  
- 🔒 Garante que só dados válidos sejam usados

## 📸 Exemplo da planilha gerada

![Print planilha final](images/Print%20planilha%20final.jpg)

## ▶️ Como Usar

1. Clone o repositório:  
```bash
git clone https://github.com/LeonardCoelho/automacao-grade.git
```

2. Instale as dependências:  
```bash
pip install -r requirements.txt
```

3. Coloque sua planilha base em `data/ORDEM_DE_COLETA.xlsx`.

4. Execute o script manualmente ou agende com o Task Scheduler:  
```bash
python src/atualizar_grade.py
```

## 🔄 Integração com Power Automate

A pasta de destino é monitorada pelo **Power Automate** (via OneDrive). Assim que a nova planilha é gerada, o fluxo é disparado e o e-mail com o anexo é enviado automaticamente para a transportadora correta.  
> ✅ **O script já está em produção com agendamento automático a cada 1 hora via Task Scheduler do Windows.**

## 🙋‍♂️ Autor

Desenvolvido por **Leonardo Coelho**  
📫 [linkedin.com/in/leonardocoelho](https://www.linkedin.com/in/leonardocoelho)

## 🏁 Próximos passos (ideias)

- [x] Deploy com agendamento automático a cada 1h  
- [ ] Parametrizar transportadora e dias via `.env` ou interface  
- [ ] Adicionar logs automáticos
