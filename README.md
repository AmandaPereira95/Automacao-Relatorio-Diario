
# 🤖 Relatório Diário de Vendas - RPA com Python

Este projeto realiza uma automação completa para geração e envio diário de relatórios de vendas, utilizando Python com foco em boas práticas e organização orientada a objetos.

---

## 📌 Funcionalidades

- Leitura de planilhas Excel com dados brutos
- Consolidação de dados por vendedor e por região
- Geração de estatísticas e indicadores-chave (total, média, maior/menor venda)
- Criação de gráficos (.png)
- Exportação do relatório em:
  - Excel consolidado (.xlsx)
  - Relatório em PDF (.pdf)
- Envio automático de e-mail com os arquivos em anexo

---

## 🛠️ Tecnologias Utilizadas

- `Python 3.10+`
- `pandas`, `openpyxl`, `matplotlib`, `numpy`
- `yagmail` para envio de e-mail
- `fpdf` para criação do PDF
- `python-dotenv` para segurança das credenciais

---

## 📁 Estrutura

```
.
├── main.py
├── config/
│   └── settings.json
├── output/
│   ├── vendas_por_vendedor.png
│   ├── vendas_por_regiao.png
│   ├── relatorio.pdf
├── data/
│   └── relatorio_base.xlsx
├── .env
└── requirements.txt
```

---

## ⚙️ Como usar

1. Clone o repositório:
```bash
git clone https://github.com/seunome/relatorio-rpa.git
cd relatorio-rpa
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Configure:
   - Crie um arquivo `.env` com:
     ```
     EMAIL=seuemail@gmail.com
     EMAIL_SENHA=sua_senha_de_app
     ```
   - Edite `config/settings.json` com os caminhos e destinatário

4. Execute:
```bash
python main.py
```

---

## 📊 Exemplo de estatísticas geradas

- Total de Vendas: R$ 23.400,00
- Média por Venda: R$ 1.950,00
- Maior Venda: R$ 4.900,00
- Menor Venda: R$ 620,00
- Quantidade de Vendas: 12

---

## 📬 Exemplo de E-mail

**Assunto:** Relatório de Vendas - Diário  
**Corpo:**
```
Olá,

Segue em anexo o relatório diário de vendas gerado automaticamente pelo robô.

📊 Total de Vendas: R$ 23.400,00

Atenciosamente,
Bot RPA de Relatórios
```

---

## 📮 Contato

Este projeto foi desenvolvido por Daniel Freire da Costa como parte do portfólio de automação RPA com Python.