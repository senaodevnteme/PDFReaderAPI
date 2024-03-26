import azure.functions as func
import re
import pandas as pd
import io
from io import BytesIO
import pdfplumber

app = func.FunctionApp()

def extrair_tabela(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        for page in pdf_file.pages:
            tabela = page.extract_table()

            if tabela:
                cabecalhos = tabela[0]

                if 'Codigo' in cabecalhos and 'Descricao' in cabecalhos:
                    descricao_idx = cabecalhos.index('Descricao')

                    for linha in tabela[1:]:
                        
                        # Verifica se o código NCM está presente
                        if linha[cabecalhos.index('Class.\nFiscal')]:
                            item_type = "1"
                        else:
                            item_type = ""
                        cst_valor = linha[cabecalhos.index('CST')]

                        dados_linha = {
                            "Mfr Part # (*)": linha[cabecalhos.index('Codigo')],
                            "Description(*)": linha[descricao_idx],
                            "Item Type": item_type,
                            "Quantity(*)": linha[cabecalhos.index('Qtd.')],
                            "Internal Comments": f"CST {cst_valor}",
                        }

                        list_price = linha[cabecalhos.index('Preco Total')]
                        unit_price = linha[cabecalhos.index('Preco Unit')]
                        our_cost = linha[cabecalhos.index('Valor IPI')]

                        # Verifica se todos os dados estão presentes
                        if None not in (list_price, unit_price, our_cost):
                            # Remove o R$ e substitui , por .
                            list_price = list_price.replace('R$', '').replace('.', '').replace(',', '.')
                            unit_price = unit_price.replace('R$', '').replace('.', '').replace(',', '.')
                            our_cost = our_cost.replace('R$', '').replace('.', '').replace(',', '.')

                            # Calcula o total
                            quantidade = int(linha[cabecalhos.index('Qtd.')])
                            if quantidade > 1:
                                # Ajusta o cálculo do preço total
                                total_cost = (float(list_price) + float(our_cost)) / quantidade
                            else:
                                total_cost = float(list_price) + float(our_cost)

                            dados_linha.update({
                                "List Price": total_cost,
                                "Unit Price(*)": total_cost,
                                "Our Cost(*)": total_cost,
                            })
                            

                            dados_linha.update({
                                "Manufacturer": "Dell",
                                "Brazil NCM Code": linha[cabecalhos.index('Class.\nFiscal')].replace('.', ''),
                                "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                                "Cost Factor 1": "0",
                                "Cost Factor 2": "0",
                                "Currency": "BRL",
                                "Preferred Supplier": "Others",
                                "Cost Factor 3": "0",
                                "Cost Factor 4": "0",
                                "Cost Factor 5": "0.5",
                            })

                            yield dados_linha

def extrair_informacoes_texto(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        informacoes = []

        for page in pdf_file.pages:
            texto = page.extract_text()
            linhas = texto.split('\n')

            for i in range(len(linhas)-1):
                # Modificação na expressão regular para identificar o início do part number
                match_part_number = re.search(r'\b[78]\d{2}-\d{4}\b', linhas[i])
                
                if match_part_number:
                    mfr_part = match_part_number.group(0).strip()

                    descricao = linhas[i].split(mfr_part)[1].replace('-', '').strip() + ' ' + ' '.join(linhas[i+1:i+4])
                    descricao_limpa = descricao.split('(', 1)[0].strip()
                    
                    concatLinhas = linhas[i] + linhas[i+1]

                    # Modificação para buscar a quantidade de forma mais flexível
                    quantidade_match = re.search(r'QTD:\s*(\d+)', concatLinhas)
                    
                    if quantidade_match:
                        preco_match = re.search(r'UNIT:\s*R\$\s*(\d+\,\d+)', concatLinhas)
                        preco = preco_match.group(1).replace('VL UNIT:', '').replace('R$', '').replace(',', '.').strip() if preco_match else "N/A"

                        informacoes.append({
                            "Manufacturer": "Dell",
                            "Mfr Part # (*)": mfr_part,
                            "Description(*)": descricao_limpa,
                            "Item Type": "7",
                            "Quantity(*)": quantidade_match.group(1),
                            "List Price": preco,
                            "Unit Price(*)": preco,
                            "Our Cost(*)": preco,
                            "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                            "Cost Factor 1": "0",
                            "Cost Factor 2": "0",
                            "Currency": "BRL",
                            "Preferred Supplier": "Others",
                            "Cost Factor 3": "0",
                            "Cost Factor 4": "0",
                            "Cost Factor 5": "0.5",
                        })

        return informacoes

def extrair_informacoes_software(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        informacoes = []

        for page in pdf_file.pages:
            texto = page.extract_text()
            linhas = texto.split('\n')

            for i in range(len(linhas)-1):
                software_code_match = re.search(r'\d{3}-[A-Z]+', linhas[i])

                if software_code_match:
                    mfr_part = software_code_match.group(0).strip()

                    descricao = linhas[i].split(mfr_part)[1].replace('-', '').strip() + ' ' + ' '.join(linhas[i+1:i+4])
                    descricao_limpa = descricao.split('(', 1)[0].strip()
                    
                    concatLinhas = linhas[i] + linhas[i+1]

                    quantidade_match = re.search(r'QTD:\s*(\d+)', concatLinhas)
                    
                    if quantidade_match:
                        preco_match = re.search(r'UNIT:\s*R\$\s*([\d\.,]+)', concatLinhas)
                        preco = preco_match.group(1).replace('R$', '').replace(',', '.').strip() if preco_match else "N/A"

                        informacoes.append({
                            "Manufacturer": "Dell",
                            "Mfr Part # (*)": mfr_part,
                            "Description(*)": descricao_limpa,
                            "Item Type": "5",
                            "Quantity(*)": quantidade_match.group(1),
                            "List Price": preco,
                            "Unit Price(*)": preco,
                            "Our Cost(*)": preco,
                            "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                            "Cost Factor 1": "0",
                            "Cost Factor 2": "0",
                            "Currency": "BRL",
                            "Preferred Supplier": "Others",
                            "Cost Factor 3": "0",
                            "Cost Factor 4": "0",
                            "Cost Factor 5": "0.5",
                        })

        return informacoes

def extrair_informacoes_infoblox(pdf_bytes):
    informacoes = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        for page in pdf_file.pages:
            # Extrair todas as tabelas da página
            tables = page.extract_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})

            # Flag para indicar que encontramos o cabeçalho da tabela
            dentro_da_tabela = False

            for table in tables:
                for linha in table:
                    # Verificar se a linha é não nula e começa com um número
                    if linha and isinstance(linha[0], str) and linha[0].isdigit():
                        dentro_da_tabela = True
                        # Lógica para descrição
                        descricao_parts = [linha[3]]
                        for i in range(4, len(linha)):
                            if i < len(linha) - 1 and isinstance(linha[i], str) and linha[i + 1].isdigit():
                                break
                            descricao_parts.append(linha[i])
                        descricao = ' '.join(descricao_parts).replace('\n',' ')
                        # Lógica para classificação do item type
                        item_type = "N/A"
                        if any(palavra in descricao.lower() for palavra in ["Software", "License", "Subscription"]):
                            item_type = "5"
                        elif "Services" in descricao.lower():
                            item_type = "7"
                        # Lógica para currency ser correta
                        net_price = str(linha[-1]).strip()
                        currency = "USD"
                        if net_price.startswith("R$"):
                            currency = "BRL"
                        # Extrair as informações da linha
                        informacao = {
                            "Manufacturer": "Infoblox",
                            "Mfr Part # (*)": linha[2],
                            "Description(*)": descricao,
                            "Item Type": item_type,
                            "Quantity(*)": int(linha[5]),
                            "List Price": linha[6].replace('$', ''),
                            "Unit Price(*)": linha[6].replace('$', ''),
                            "Our Cost(*)": linha[6].replace('$', ''),
                            "UNSPSC": "",
                            "External Comments": "",
                            "Internal Comments": "",
                            "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                            "Cost Factor 1": "0",
                            "Cost Factor 2": "0",
                            "Surcharges": "",
                            "Vendor Maintenance": "",
                            "Local Maintenance": "",
                            "Currency": currency,
                            "Required Section": "",
                            "Solution Type": "",
                            "Preferred Supplier": "Infoblox",
                            "UOM": "",
                            "Brazil NCM Code": "",
                            "Cost Factor 3": "0",
                            "Cost Factor 4": "0",
                            "Cost Factor 5": "0.5",
                            "DealType": "",
                            "DealValue": "",
                            "SubItemType": "",
                            "CategoryCode": "",
                            "VendorQuoteNumber": "",
                        }

                        informacoes.append(informacao)
                    elif dentro_da_tabela and linha and isinstance(linha[0], str) and linha[0].strip() == '':
                        dentro_da_tabela = False

    return informacoes

def extrair_informacoes_f5_1(pdf_bytes):
    informacoes = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        start_processing = False
        part_number = None
        description_parts = []  
        qty = 0
        net_price = 0

        for page in pdf_file.pages:
            texto_pagina = page.extract_text()

            linhas = [linha.strip() for linha in texto_pagina.split('\n') if linha.strip()]

            informacao_item = {}  

            for i, linha in enumerate(linhas):
                if "Line Part # Description Qty Net Price Ext. Sale Price" in linha:
                    start_processing = True
                    continue
                if linha.startswith('Total:'):
                    start_processing = False
                    break

                if start_processing and any(linha.startswith(f"{j}.") for j in range(1, 10)) and len(linha.split()) > 4:
                    partes = linha.split()
                    part_number = partes[1]

                    item_type = "7" if "SVC" in part_number else "5"

                    description_parts = partes[2:-2]  # Ajuste do índice para excluir apenas a parte final da linha

                    qty_index = -3 if partes[-3].replace('$', '').replace(',', '').replace('.', '').isdigit() else -2
                    qty_str = ''.join(filter(str.isdigit, partes[qty_index]))  # Extrai apenas os dígitos
                    qty = int(qty_str)

                    net_price_index = -2 if partes[-2].replace('$', '').replace(',', '').replace('.', '').isdigit() else -1
                    net_price = float(partes[net_price_index].replace('$', '').replace(',', ''))

                    # Verifica se a última parte da descrição contém apenas dígitos
                    last_description_part = description_parts[-1]
                    if last_description_part.isdigit():
                        description_parts = description_parts[:-1]

                    # Alteração aqui para unir todas as partes da descrição
                    next_line = linhas[i + 1] if i + 1 < len(linhas) else ""
                    description = ' '.join(description_parts + [next_line])

                    # Adiciona as informações do item ao dicionário
                    informacao_item = {
                        "Manufacturer": "F5 Networks",
                        "Mfr Part # (*)": part_number,
                        "Description(*)": description,
                        "Item Type": item_type,
                        "Quantity(*)": qty,
                        "List Price": net_price,
                        "Unit Price(*)": net_price,
                        "Our Cost(*)": net_price,
                        "UNSPSC": "",
                        "External Comments": "",
                        "Internal Comments": "",
                        "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                        "Cost Factor 1": "0",
                        "Cost Factor 2": "0",
                        "Surcharges": "",
                        "Vendor Maintenance": "",
                        "Local Maintenance": "",
                        "Currency": "USD",
                        "Required Section": "",
                        "Solution Type": "",
                        "Preferred Supplier": "F5 networks",
                        "UOM": "",
                        "Brazil NCM Code": "",
                        "Cost Factor 3": "0",
                        "Cost Factor 4": "0",
                        "Cost Factor 5": "0.5",
                        "DealType": "",
                        "DealValue": "",
                        "SubItemType": "",
                        "CategoryCode": "",
                        "VendorQuoteNumber": "",
                    }

                    informacoes.append(informacao_item)

    return informacoes

def extrair_informacoes_f5_2(pdf_bytes):
    informacoes = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        for page in pdf_file.pages:
            table = page.extract_tables()

            if not table:
                continue

            df = pd.DataFrame(table[1][1:], columns=table[1][0])

            for _, row in df.iterrows():
                descricao = row["Description"].replace('\n', ' ')

                informacao_item = {
                    "Manufacturer": "F5 Networks",
                    "Mfr Part # (*)": row["Part#"],
                    "Description(*)": descricao,
                    "Item Type": "7" if "SVC" in row["Part#"] else "5",
                    "Quantity(*)": int(float(row["Extended Price"].replace('$', '').replace(',', '')) / float(row["Unit Price"].replace('$', '').replace(',', ''))),
                    "List Price": float(row["Extended Price"].replace('$', '').replace(',', '')),
                    "Unit Price(*)": float(row["Unit Price"].replace('$', '').replace(',', '')),
                    "Our Cost(*)": float(row["Net Price"].replace('$', '').replace(',', '')),
                    "UNSPSC": "",
                    "External Comments": "",
                    "Internal Comments": "",
                    "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                    "Cost Factor 1": "0",
                    "Cost Factor 2": "0",
                    "Surcharges": "",
                    "Vendor Maintenance": "",
                    "Local Maintenance": "",
                    "Currency": "USD",
                    "Required Section": "",
                    "Solution Type": "",
                    "Preferred Supplier": "F5 networks",
                    "UOM": "",
                    "Brazil NCM Code": "",
                    "Cost Factor 3": "0",
                    "Cost Factor 4": "0",
                    "Cost Factor 5": "0.5",
                    "DealType": "",
                    "DealValue": "",
                    "SubItemType": "",
                    "CategoryCode": "",
                    "VendorQuoteNumber": "",
                }

                informacoes.append(informacao_item)
                

    return informacoes

def extrair_descrição(partes, start_index, end_index):
    descrição = ' '.join(partes[start_index:end_index])
    return descrição

def extrair_informacoes_palo_alto(pdf_bytes):
    informacoes_item = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_file:
        start_processing = False  

        for page in pdf_file.pages:
            texto_pagina = page.extract_text()
            linhas = [linha.strip() for linha in texto_pagina.split('\n') if linha.strip()]

            for linha in linhas:
                linha = re.sub(r'(\d+),(\d{2})(\d{2}),', r'\1,\2 \3,', linha)
                linha = re.sub(r'\s{2,}', ' ', linha)
                partes = re.split(r'\s+', linha)

                header_items = ["Qtde", "Estoque", "Origem", "Part", "Number", "NCM", "SKU", "Descrição", "ECCNMoeda", "Valor", "Unitário", "Valor", "Total", "sem", "ST/Difal", "ICMS", "ST", "/", "Difal", "Valor", "Total", "Comissão", "Liquida", "IPI", "ICMSISSObservações"]
                if all(item in linha for item in header_items):
                    start_processing = True
                    continue
                if linha.startswith('Total Grupo USC:'):
                    start_processing = False
                    break

                if start_processing:
                    if len(partes) >= 15:
                        print(partes)
                        if partes[6] == "Software":
                            item_type = "5"
                        elif partes[6] == "Service":
                            item_type = "7"
                        else: 
                            item_type = "1"
                            ncm = partes[6]

                        try:
                            end_of_description_index = partes.index('USC') - 1
                        except ValueError:
                            end_of_description_index = len(partes)  
                        
                        start_of_description_index = 7  
                        for i, parte in enumerate(partes[start_of_description_index:], start=start_of_description_index):
                            if not re.fullmatch(r'EX\.\d+', parte):
                                start_of_description_index = i
                                break
                            
                        qtde = partes[0]
                        part_number = partes[5]
                        descricao = extrair_descrição(partes, start_of_description_index, end_of_description_index)
                        valor_unitario = partes[end_of_description_index + 2].replace('.', '').replace(',', '.')
                        moeda = "USD" if "USC" in partes else partes[-11]

                        informacao_item = {
                            "Manufacturer": "Palo Alto",
                            "Mfr Part # (*)": part_number,
                            "Description(*)": descricao,
                            "Item Type": item_type,
                            "Quantity(*)": qtde,
                            "List Price": valor_unitario,
                            "Unit Price(*)": valor_unitario,
                            "Our Cost(*)": valor_unitario,
                            "UNSPSC": "",
                            "External Comments": "",
                            "Internal Comments": "",
                            "Price Rule('Fixed' , 'Margin', 'Discount')": "Fixed",
                            "Cost Factor 1": "0",
                            "Cost Factor 2": "0",
                            "Surcharges": "",
                            "Vendor Maintenance": "",
                            "Local Maintenance": "",
                            "Currency": moeda,
                            "Required Section": "",
                            "Solution Type": "",
                            "Preferred Supplier": "Others",
                            "UOM": "",
                            "Brazil NCM Code": ncm if item_type == "1" else "",
                            "Cost Factor 3": "0",
                            "Cost Factor 4": "0",
                            "Cost Factor 5": "0.5",
                            "DealType": "",
                            "DealValue": "",
                            "SubItemType": "",
                            "CategoryCode": "",
                            "VendorQuoteNumber": "",
                        }

                        informacoes_item.append(informacao_item)

    return informacoes_item

@app.route(route="MyFunction", methods=['GET'])
def main(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(open('index.html', 'r').read(), mimetype='text/html')

@app.route(route="MyFunction", methods=['POST'])
def MyFunction(req: func.HttpRequest) -> func.HttpResponse:
       
    manufacturer = req.form.get('manufacturer')
    pdf_bytes = req.get_body()
    
    if pdf_bytes is None:
        return func.HttpResponse("Nenhum arquivo recebido")
    
    colunas_padrao = [
        "Manufacturer",
        "Mfr Part # (*)",
        "Description(*)",
        "Item Type",
        "Quantity(*)",
        "List Price",
        "Unit Price(*)",
        "Our Cost(*)",
        "UNSPSC",
        "External Comments",
        "Internal Comments",
        "Price Rule('Fixed' , 'Margin', 'Discount')",
        "Cost Factor 1",
        "Cost Factor 2",
        "Surcharges",
        "Vendor Maintenance",
        "Local Maintenance",
        "Currency",
        "Required Section",
        "Solution Type",
        "Preferred Supplier",
        "UOM",
        "Brazil NCM Code",
        "Cost Factor 3",
        "Cost Factor 4",
        "Cost Factor 5",
        "DealType",
        "DealValue",
        "SubItemType",
        "CategoryCode",
        "VendorQuoteNumber"
    ]
    
    df = pd.DataFrame()
    
    if manufacturer == 'Dell':
        resultados_tabela = [resultado for resultado in extrair_tabela(pdf_bytes)]
        df_tabela = pd.DataFrame(resultados_tabela)

        resultados_texto = extrair_informacoes_texto(pdf_bytes)
        df_texto = pd.DataFrame(resultados_texto)

        resultados_software = extrair_informacoes_software(pdf_bytes)
        df_software = pd.DataFrame(resultados_software)

        # Adiciona colunas ausentes com valores padrão
        for coluna in colunas_padrao:
            if coluna not in df_tabela.columns:
                df_tabela[coluna] = ""

            if coluna not in df_texto.columns:
                df_texto[coluna] = ""

            if coluna not in df_software.columns:
                df_software[coluna] = ""

        # Reorganiza as colunas na ordem desejada
        df_tabela = df_tabela[colunas_padrao]
        df_texto = df_texto[colunas_padrao]
        df_software = df_software[colunas_padrao]

        # Concatena os DataFrames
        df = pd.concat([df_tabela, df_texto, df_software], ignore_index=True)
        
    elif manufacturer == 'Infoblox':
        resultados_infoblox = extrair_informacoes_infoblox(pdf_bytes)
        df = pd.DataFrame(resultados_infoblox)
        
    elif manufacturer == 'F5 - 1':
        resultados_f5_1 = extrair_informacoes_f5_1(pdf_bytes) 
        df = pd.DataFrame(resultados_f5_1)
    
    elif manufacturer == 'F5 - 2':
        resultados_f5_2 = extrair_informacoes_f5_2(pdf_bytes)
        df = pd.DataFrame(resultados_f5_2)
    
    elif manufacturer == 'Palo Alto - Ingram':
        resultado_palo_alto = extrair_informacoes_palo_alto(pdf_bytes)
        df = pd.DataFrame(resultado_palo_alto)
        
    # Adiciona valores padrão nas colunas extras
    for coluna in colunas_padrao:
        if coluna not in df:
            df[coluna] = ""

    if manufacturer == "Dell":
        df["Preferred Supplier"] = "DELL"
        
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return func.HttpResponse(body=output.getvalue(), status_code=200, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


