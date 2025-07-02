import os
import re
import pytesseract
import pandas as pd
import requests
from pdf2image import convert_from_path
from pathlib import Path

# Caminhos fixos
CAMINHO_TESSERACT = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
CAMINHO_PDF_BASE = r'C:\import\clientes'
ARQUIVO_SAIDA = r'C:\import\dados_extraidos.xlsx'

# Configuração
pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT
dados = []

def extrair_info(texto):
    nome = cnpj = cpf = endereco = bairro = municipio = uf = cep = numero = ''

    # Isola o trecho entre QUADRO III – EMITENTE e QUADRO IV
    match = re.search(r'QUADRO III.*?EMITENTE(.*?QUADRO IV)', texto, re.DOTALL | re.IGNORECASE)
    bloco = match.group(1) if match else texto

    linhas = [l.strip() for l in bloco.split('\n') if l.strip()]
    linhas_iter = iter(linhas)

    for linha in linhas_iter:
        linha_upper = linha.upper()

        if 'RAZÃO SOCIAL' in linha_upper or 'NOME/RAZÃO SOCIAL' in linha_upper:
            linha_nome = next(linhas_iter, '').strip()
            cnpj_match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', linha_nome)
            cpf_match = re.search(r'\d{3}\.\d{3}\.\d{3}-\d{2}', linha_nome)
            if cnpj_match:
                cnpj = cnpj_match.group()
                nome = linha_nome.replace(cnpj, '')
            elif cpf_match:
                cpf = cpf_match.group()
                nome = linha_nome.replace(cpf, '')
            else:
                nome = linha_nome
            nome = re.sub(r'[^A-Za-zÀ-ÿ\s]', '', nome).strip()

        elif 'CNPJ' in linha_upper and not cnpj:
            cnpj = next(linhas_iter, '').strip()

        elif 'ENDEREÇO' in linha_upper:
            endereco_linha = next(linhas_iter, '').strip()
            endereco = re.sub(r',?\s*\d{1,5}\b', '', endereco_linha).strip()
            num_match = re.search(r'\b\d{1,5}\b', endereco_linha)
            if num_match:
                numero = num_match.group()

        elif 'BAIRRO' in linha_upper:
            bairro = next(linhas_iter, '').strip()

        elif 'MUNICÍPIO' in linha_upper:
            municipio = next(linhas_iter, '').strip()

        elif linha_upper == 'UF':
            uf = next(linhas_iter, '').strip()

        elif 'CEP' in linha_upper and not cep:
            cep_inline = re.search(r'\d{2}\.\d{3}-\d{3}|\d{5}-\d{3}', linha)
            if cep_inline:
                cep_raw = re.sub(r'\D', '', cep_inline.group())
                if len(cep_raw) == 8:
                    cep = cep_raw[:5] + '-' + cep_raw[5:]
            else:
                prox = next(linhas_iter, '').strip()
                cep_next = re.search(r'\d{2}\.\d{3}-\d{3}|\d{5}-\d{3}', prox)
                if cep_next:
                    cep_raw = re.sub(r'\D', '', cep_next.group())
                    if len(cep_raw) == 8:
                        cep = cep_raw[:5] + '-' + cep_raw[5:]

    # Consulta ViaCEP
    if re.match(r'^\d{5}-\d{3}$', cep):
        try:
            response = requests.get(f'https://viacep.com.br/ws/{cep.replace("-", "")}/json/', timeout=5)
            if response.status_code == 200:
                data = response.json()
                if 'erro' not in data:
                    endereco = data.get('logradouro', endereco) or endereco
                    bairro = data.get('bairro', bairro) or bairro
                    municipio = data.get('localidade', municipio) or municipio
                    uf = data.get('uf', uf) or uf
        except Exception as e:
            print(f'⚠️ Erro ao consultar ViaCEP ({cep}): {e}')

    return nome, endereco, numero, bairro, municipio, uf, cep, cnpj or cpf

# def extrair_dif(df):
#     cep_regex = r'\d{2}\.\d{3}-\d{3}|\d{5}-\d{3}'
#
#     for idx, row in df.iterrows():
#         for col in ['Bairro', 'Município', 'UF']:
#             texto = str(row[col])
#             cep_match = re.search(cep_regex, texto)
#             if cep_match:
#                 cep_raw = re.sub(r'\D', '', cep_match.group())
#                 if len(cep_raw) == 8:
#                     cep_formatado = cep_raw[:5] + '-' + cep_raw[5:]
#                     df.at[idx, 'CEP'] = cep_formatado
#                     df.at[idx, col] = texto.replace(cep_match.group(), '').strip(' ,|-')
#
#                     try:
#                         r = requests.get(f'https://viacep.com.br/ws/{cep_raw}/json/', timeout=5)
#                         if r.status_code == 200:
#                             data = r.json()
#                             if 'erro' not in data:
#                                 if data.get('bairro'):
#                                     df.at[idx, 'Bairro'] = data['bairro']
#                                 if data.get('localidade'):
#                                     df.at[idx, 'Município'] = data['localidade']
#                                 if data.get('uf'):
#                                     df.at[idx, 'UF'] = data['uf']
#
#                                 # Se logradouro vier vazio, mantém o extraído via OCR
#                                 if data.get('logradouro'):
#                                     df.at[idx, 'Endereço'] = data['logradouro']
#                                 else:
#                                     df.at[idx, 'Endereço'] = row['Endereço']
#                     except Exception as e:
#                         print(f"⚠️ Erro consultando ViaCEP para linha {idx}: {e}")
#     return df
def extrair_dif(df):
    # Regex cobre:
    # - 12345678
    # - 12.345-678
    # - 12345-678
    # - 12.345678 ✅
    cep_regex = r'(\d{5}-\d{3}|\d{2}\.\d{3}-\d{3}|\d{8}|\d{2}\.\d{6})'

    for idx, row in df.iterrows():
        # Se nome estiver vazio, use o nome da pasta
        if not str(row['Nome']).strip():
            df.at[idx, 'Nome'] = row['Cliente (pasta)'].strip()

        # Verifica CEP em colunas: Bairro, Município, UF
        for col in ['Bairro', 'Município', 'UF']:
            texto = str(row[col])
            cep_match = re.search(cep_regex, texto)
            if cep_match:
                cep_raw = re.sub(r'\D', '', cep_match.group())
                if len(cep_raw) == 8:
                    cep_formatado = cep_raw[:5] + '-' + cep_raw[5:]
                    df.at[idx, 'CEP'] = cep_formatado
                    df.at[idx, col] = texto.replace(cep_match.group(), '').strip(' ,|-')

                    try:
                        r = requests.get(f'https://viacep.com.br/ws/{cep_raw}/json/', timeout=5)
                        if r.status_code == 200:
                            data = r.json()
                            if 'erro' not in data:
                                if data.get('bairro'):
                                    df.at[idx, 'Bairro'] = data['bairro']
                                if data.get('localidade'):
                                    df.at[idx, 'Município'] = data['localidade']
                                if data.get('uf'):
                                    df.at[idx, 'UF'] = data['uf']
                                if data.get('logradouro'):
                                    df.at[idx, 'Endereço'] = data['logradouro']
                                else:
                                    df.at[idx, 'Endereço'] = row['Endereço']
                    except Exception as e:
                        print(f"⚠️ Erro consultando ViaCEP para linha {idx}: {e}")
    return df


# Processa os arquivos PDF nas pastas
total_pastas = sum(1 for _ in Path(CAMINHO_PDF_BASE).iterdir() if _.is_dir())
pasta_atual = 1

for pasta_cliente in Path(CAMINHO_PDF_BASE).iterdir():
    if pasta_cliente.is_dir():
        print(f"\n📁 Pasta ({pasta_atual}/{total_pastas}): {pasta_cliente.name}")
        pasta_atual += 1
        for arquivo in pasta_cliente.iterdir():
            if arquivo.name.upper().startswith("CCB") and arquivo.suffix.lower() == ".pdf":
                print(f"  ⏳ Processando arquivo: {arquivo.name}")
                try:
                    imagens = convert_from_path(str(arquivo), dpi=300, first_page=1, last_page=1)
                    texto_total = pytesseract.image_to_string(imagens[0], lang='por')

                    nome, endereco, numero, bairro, municipio, uf, cep, cnpj = extrair_info(texto_total)

                    dados.append({
                        'Cliente (pasta)': pasta_cliente.name,
                        'Arquivo': arquivo.name,
                        'Nome': nome,
                        'Endereço': endereco,
                        'Número': numero,
                        'Bairro': bairro,
                        'Município': municipio,
                        'UF': uf,
                        'CEP': cep,
                        'CNPJ/CPF': cnpj
                    })

                    print(f"  ✅ Extraído com sucesso.")

                except Exception as e:
                    print(f"  ❌ ERRO ao processar {arquivo.name}: {e}")

# Salva a planilha com pós-processamento
df = pd.DataFrame(dados)
df = extrair_dif(df)
df.to_excel(ARQUIVO_SAIDA, index=False)
print(f"\n📄 Extração finalizada. Arquivo salvo em: {ARQUIVO_SAIDA}")