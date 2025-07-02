# LexBot

LexBot é um robô em Python desenvolvido por **Carlos Alexandre** para auxiliar advogados na extração de dados de documentos digitalizados do tipo "inicial jurídica" (como CCBs).

Ele automatiza o processo de leitura de PDFs escaneados, identificando e extraindo informações como:

- Nome/Razão Social
- CNPJ ou CPF
- Endereço e Número
- Bairro
- Município
- UF
- CEP

> Caso o CEP esteja mal posicionado ou em formatos como `13.188-181`, `18.284402` ou `13188181`, o LexBot reconhece, formata corretamente e consulta a [API ViaCEP](https://viacep.com.br/) para completar os dados.

---

## 📦 Funcionalidades

- 🧠 Leitura inteligente do bloco **QUADRO III – EMITENTE** do PDF
- 📄 OCR com **Tesseract** e **pdf2image**
- 🔍 Busca e correção automática de CEPs misturados em outras colunas
- 🌐 Consulta automática à **API ViaCEP**
- 📊 Exportação em planilha Excel (`.xlsx`) limpa e estruturada

---

## 📂 Estrutura esperada

```
C:\
 └── import\
      ├── clientes\
      │     ├── CARLOS ALEXANDRE\
      │     │     └── CCB - CARLOS ALEXANDRE.pdf
      │     └── OUTRO CLIENTE\
      │           └── CCB - EXEMPLO.pdf
      └── dados_extraidos.xlsx
```

---

## ⚙️ Requisitos

- Python 3.9+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) instalado
- [Poppler](http://blog.alivate.com.au/poppler-windows/) (para `pdf2image`)
- Bibliotecas Python:

```bash
pip install pytesseract pdf2image pandas requests openpyxl
```

---

## 🚀 Como executar

```bash
python lexbot.py
```

---

## 🛠️ Desenvolvido por

Carlos Alexandre • [github.com/carloscamposmiranda](https://github.com/seu-usuario)

Este projeto é de uso **livre e educativo**, especialmente útil para escritórios jurídicos que desejam automatizar o preenchimento de planilhas a partir de documentos digitalizados.