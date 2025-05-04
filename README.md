
# 🤖 Automação de Relatório Diário (Excel + E-mail)

Este projeto RPA realiza a automação completa de um processo comum em empresas: **ler uma planilha de dados**, **consolidar informações** (totais por vendedor, região, etc.), **gerar gráficos** e **enviar um e-mail automático com os relatórios** em anexo.

---

## 📌 Funcionalidades

- 📥 Leitura de planilha Excel com registros de vendas.
- 📊 Consolidação de:
  - Total geral de vendas
  - Total por vendedor
  - Total por região
- 📈 Geração de gráfico de barras
- 📤 Geração de relatório em Excel e PDF
- ✉️ Envio automático por e-mail com os arquivos anexados

---

## 🧰 Tecnologias Utilizadas

- `Python`
  - `pandas`
  - `openpyxl`
  - `matplotlib`
  - `yagmail` ou `smtplib` para envio de e-mail
  - `fpdf` ou `xlsx2pdf` para exportação em PDF (opcional)
- Compatível com UiPath (versão futura possível com atividades `Read Range`, `Group By`, `Send Outlook Mail`)

---

## 📁 Estrutura do Projeto

```
relatorio-diario-automacao/
├── README.md
├── data/
│   ├── relatorio_base.xlsx         # Planilha original com dados brutos
│   └── relatorio_consolidado.xlsx # Planilha gerada com totais e resumo
├── output/
│   └── relatorio.pdf               # PDF final com gráficos
├── scripts/
│   └── main.py                     # Script Python principal
├── config/
│   └── settings.json               # Configuração de caminhos e e-mail
├── requirements.txt                # Lista de dependências
└── .gitignore
```

---

## 📦 Requisitos

```bash
pip install -r requirements.txt
```

Conteúdo sugerido para o `requirements.txt`:
```txt
pandas
openpyxl
matplotlib
yagmail
fpdf
```

---

## ⚙️ Como Usar

1. Edite o arquivo `config/settings.json` com seus dados:
   - Caminhos de entrada e saída
   - Dados do SMTP (e-mail e senha de aplicativo)
2. Coloque sua planilha em `data/relatorio_base.xlsx`
3. Execute o script:

```bash
python scripts/main.py
```

4. Verifique:
   - Arquivo consolidado em `data/relatorio_consolidado.xlsx`
   - PDF gerado em `output/relatorio.pdf`
   - E-mail enviado ao destinatário

---

## 🧪 Exemplo de Dados

Um exemplo de planilha com dados fictícios pode ser encontrado em `data/relatorio_base.xlsx`, com colunas:
- `Data`
- `Vendedor`
- `Região`
- `Valor da Venda (R$)`

---

## 📮 Contato

Este projeto faz parte do meu portfólio como desenvolvedor RPA.  
Dúvidas ou sugestões? Fique à vontade para abrir uma issue ou me contactar!

---