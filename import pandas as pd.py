import pandas as pd
from barcode import Code128  # Importa o gerador de código de barras do tipo Code128
from barcode.writer import ImageWriter  # Usado para salvar o código de barras como imagem
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime

def carregar_dados():
    caminho_arquivo = r'C:\Users\ter07068\OneDrive - M DIAS BRANCO S.A. INDUSTRIA E COMERCIO DE ALIMENTOS\Área de Trabalho\Python\Log demanda\Materiais.xlsx'
    df = pd.read_excel(caminho_arquivo)
    return df

def formatar_data(data):
    if pd.isna(data):
        return 'N/A'
    return datetime.strftime(pd.to_datetime(data), "%d/%m/%Y")

def remover_prefixo_e_zeros(codigo_completo):
    codigo_limpo = codigo_completo.lstrip("EMBZMP0")
    return codigo_limpo

def gerar_codigo_controle(codigo_completo, lote):
    codigo_limpo = remover_prefixo_e_zeros(codigo_completo)
    ultimos_3_codigo = codigo_limpo[-3:]
    ultimos_3_lote = lote[-3:]
    prefixo = "EMB" if codigo_completo.startswith("EMB") or codigo_completo.startswith("ZEMB") else "MP"
    return f"{prefixo}{ultimos_3_codigo}{ultimos_3_lote}"

def criar_codigo_barras(informacoes, caminho_arquivo):
    # Gera o código de barras do tipo Code128 com base nas informações fornecidas
    codigo_barras = Code128(informacoes, writer=ImageWriter(), add_checksum=False)
    codigo_barras.save(caminho_arquivo)
    
    # Abre a imagem gerada do código de barras
    imagem_barras = Image.open(f"{caminho_arquivo}.png")
    return imagem_barras

def criar_imagem_a4(codigo, descricao, lote, validade, recebimento, quantidade, nota_fiscal, fornecedor, codigo_barras_img):
    largura_a4, altura_a4 = (2480, 3508)
    imagem = Image.new("RGB", (largura_a4, altura_a4), "white")
    draw = ImageDraw.Draw(imagem)
    
    try:
        fonte_codigo_enorme = ImageFont.truetype("arialbd.ttf", 450)  # Fonte negrito
        fonte_codigo_grande = ImageFont.truetype("arialbd.ttf", 300)  # Fonte negrito
        fonte_codigo_media = ImageFont.truetype("arial.ttf", 150)
        fonte_outros = ImageFont.truetype("arial.ttf", 100)
        fonte_pequena = ImageFont.truetype("arial.ttf", 50)
    except IOError:
        fonte_codigo_grande = fonte_codigo_media = fonte_outros = fonte_pequena = ImageFont.load_default()

    margem = 50
    espaçamento = 50  # Ajustar espaçamento para 70
    espaçamento_codigo = 50
    espaçamento_quantidade = 35
    y_offset = margem

    # Função para desenhar caixinhas
    def desenhar_caixa(texto, y, preenchimento="lightgray", fonte=fonte_outros, altura_extra=0):
        largura_caixa = largura_a4 - 2 * margem
        bbox = draw.textbbox((0, 0), texto, font=fonte)
        altura_caixa = bbox[3] - bbox[1] + 2 * margem // 2 + altura_extra
        draw.rectangle([margem, y, largura_caixa + margem, y + altura_caixa], outline="black", fill=preenchimento)
        draw.text((margem + 10, y + margem // 4), texto, font=fonte, fill="black")
        return altura_caixa

    # Código
    y_offset += desenhar_caixa("Código", y_offset) + espaçamento_codigo
    draw.text((margem + 10, y_offset + 20), codigo, font=fonte_codigo_grande, fill="black", underline=True)  # Sublinha
    y_offset += 300 + espaçamento  # Altura fixa para o código

    # Descrição
    y_offset += desenhar_caixa(f"Descrição: {descricao}", y_offset) + espaçamento

    # Lote
    y_offset += desenhar_caixa(f"Lote: {lote}", y_offset) + espaçamento

    # Data de Validade
    y_offset += desenhar_caixa(f"Data de Validade: {validade}", y_offset) + espaçamento

    # Data de Recebimento
    y_offset += desenhar_caixa(f"Data de Fabricação: {recebimento}", y_offset) + espaçamento

    # Nota Fiscal
    y_offset += desenhar_caixa(f"Nota Fiscal: {nota_fiscal}", y_offset) + espaçamento

    # Fornecedor
    y_offset += desenhar_caixa(f"Fornecedor: {fornecedor}", y_offset) + espaçamento

    # Quantidade
    y_offset += desenhar_caixa("Quantidade", y_offset) + espaçamento_quantidade
    draw.text((largura_a4 // 2, y_offset), quantidade, font=fonte_codigo_enorme, fill="black", underline=True)  # Sublinha centralizada
    y_offset += 400 + espaçamento  # Altura fixa para a quantidade

    # Linha separadora
    linha_y = y_offset  # Ajuste a altura da linha separadora
    draw.line([margem, linha_y, largura_a4 - margem, linha_y], fill="black", width=3)
    y_offset += espaçamento

    # Ajuste do Código de Barras
    barras_largura, barras_altura = codigo_barras_img.size
    novo_tamanho_barras = (barras_largura // 2, barras_altura // 2)
    codigo_barras_img = codigo_barras_img.resize(novo_tamanho_barras, Image.Resampling.LANCZOS)

    # Posição do Código de Barras
    pos_barras_x = (largura_a4 - novo_tamanho_barras[0]) // 2 - margem
    pos_barras_y = altura_a4 - novo_tamanho_barras[1] - margem
    imagem.paste(codigo_barras_img, (pos_barras_x, pos_barras_y))

    return imagem

def editar_informacoes(data):
    while True:
        print("\nInformações atuais:")
        for i, (key, value) in enumerate(data.items(), 1):
            print(f"{i}. {key}: {value}")
        
        editar = input("\nDeseja editar alguma informação? (s/n): ").strip().lower()
        
        if editar == 's':
            opcao = int(input("Digite o número da informação que deseja editar: ").strip())
            chave_selecionada = list(data.keys())[opcao - 1]
            
            nova_informacao = input(f"Digite a nova informação para {chave_selecionada}: ").strip()
            confirmacao = input(f"Tem certeza que deseja alterar '{data[chave_selecionada]}' para '{nova_informacao}'? (s/n): ").strip().lower()
            
            if confirmacao == 's':
                data[chave_selecionada] = nova_informacao
                print(f"{chave_selecionada} atualizado com sucesso!")
            else:
                print("Alteração cancelada.")
        else:
            break

def main():
    df = carregar_dados()
    
    while True:
        meses_visualizar = int(input("Quantos meses você deseja visualizar os lotes? ").strip())
        data_limite = datetime.now() - pd.DateOffset(months=meses_visualizar - 1)
        
        codigo_produto = input("Digite o código do produto: ").strip()

        df['Material'] = df['Material'].astype(str)
        df['Lote'] = df['Lote'].astype(str)
        df['Lote'] = df['Lote'].str.strip().str.upper()

        df_codigo = df[df['Material'].str.endswith(codigo_produto.zfill(9)) | df['Material'].str.endswith(codigo_produto.zfill(10))]

        if not df_codigo.empty:
            df_codigo.loc[:, 'Data de produção'] = pd.to_datetime(df_codigo['Data de produção'], errors='coerce')
            df_codigo = df_codigo[df_codigo['Data de produção'] >= data_limite]
            
            if df_codigo.empty:
                print(f"Nenhum lote encontrado para o código {codigo_produto} nos últimos {meses_visualizar} meses.")
                continue

            lotes = df_codigo['Lote'].unique()
            if len(lotes) > 1:
                print("\nLotes disponíveis:")
                for i, lote in enumerate(lotes, 1):
                    print(f"{i}. {lote}")
                indice_lote = int(input("\nEscolha o número do lote desejado: ").strip())
                lote_escolhido = lotes[indice_lote - 1]
            else:
                lote_escolhido = lotes[0]

            df_lote = df_codigo[df_codigo['Lote'] == lote_escolhido].iloc[0]

            descricao = df_lote['Descrição de material']
            lote = df_lote['Lote']
            validade = formatar_data(df_lote['Data de vencimento'])
            recebimento = formatar_data(df_lote['Data de produção'])
            nota_fiscal = input("Digite a nota fiscal: ").strip()
            fornecedor = df_lote['Fornecedor']
            quantidade = input("Digite a quantidade: ").strip()

            dados = {
                "Código": codigo_produto,
                "Descrição": descricao,
                "Lote": lote,
                "Validade": validade,
                "Data de Fabricação": recebimento,
                "Quantidade": quantidade,
                "Nota Fiscal": nota_fiscal,
                "Fornecedor": fornecedor
            }

            editar_informacoes(dados)

            # Gera o código de barras
            caminho_arquivo_barras = os.path.join(os.getcwd(), f"codigo_barras_{codigo_produto}")
            codigo_barras_img = criar_codigo_barras(codigo_produto, caminho_arquivo_barras)

            # Cria a imagem A4 com o código de barras
            imagem_final = criar_imagem_a4(
                dados["Código"], dados["Descrição"], dados["Lote"], dados["Validade"],
                dados["Data de Fabricação"], dados["Quantidade"], dados["Nota Fiscal"],
                dados["Fornecedor"], codigo_barras_img
            )

            caminho_imagem = os.path.join(os.getcwd(), f"CodigoBarras_{codigo_produto}.png")
            imagem_final.save(caminho_imagem)
            print(f"Código de barras gerado e salvo em: {caminho_imagem}")

            continuar = input("\nDeseja gerar outro código de barras? (s/n): ").strip().lower()
            if continuar != 's':
                break
        else:
            print(f"Nenhum lote encontrado para o código {codigo_produto}.")
