import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime

# Função para carregar os dados do arquivo Excel
def carregar_dados():
    caminho_arquivo = r'C:\Users\ter07068\OneDrive - M DIAS BRANCO S.A. INDUSTRIA E COMERCIO DE ALIMENTOS\Área de Trabalho\Python\Log demanda\Materiais.xlsx'
    df = pd.read_excel(caminho_arquivo)
    return df

# Função para formatar datas
def formatar_data(data):
    if pd.isna(data):
        return 'N/A'
    return datetime.strftime(pd.to_datetime(data), "%d/%m/%Y")

# Função para criar o QR Code a partir de uma string de informações
def criar_qrcode(informacoes):
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=35,
        border=4,
    )
    qr.add_data(informacoes)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")


# Função para criar uma imagem A4 com os dados e QR Code
def criar_imagem_a4(codigo, descricao, lote, validade, recebimento, quantidade, nota_fiscal, fornecedor, qrcode_img):
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
    y_offset += desenhar_caixa(f"Data de Recebimento: {recebimento}", y_offset) + espaçamento

    # Nota Fiscal
    y_offset += desenhar_caixa(f"Nota Fiscal: {nota_fiscal}", y_offset) + espaçamento

    # Fornecedor
    y_offset += desenhar_caixa(f"Fornecedor: {fornecedor}", y_offset) + espaçamento

       # Quantidade
    y_offset += desenhar_caixa("Quantidade", y_offset) + espaçamento_quantidade
    draw.text((largura_a4 // 2, y_offset), quantidade, font=fonte_codigo_enorme, fill="black", underline=True)  # Sublinha centralizada
    y_offset += 400 + espaçamento  # Altura fixa para a quantidade

    # Linha separadora
    linha_y = y_offset # Ajuste a altura da linha separadora
    draw.line([margem, linha_y, largura_a4 - margem, linha_y], fill="black", width=3)
    y_offset += espaçamento

    # Ajuste do QR Code
    qr_largura, qr_altura = qrcode_img.size
    novo_tamanho_qr = (qr_largura // 2, qr_altura // 2)
    qrcode_img = qrcode_img.resize(novo_tamanho_qr, Image.Resampling.LANCZOS)

    # Posição do QR Code
    pos_qr_x = (largura_a4 - novo_tamanho_qr[0]) // 2 - margem
    pos_qr_y = altura_a4 - novo_tamanho_qr[1] - margem
    imagem.paste(qrcode_img, (pos_qr_x, pos_qr_y))

    return imagem




# Função para editar informações inseridas pelo usuário
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

# Função principal
def main():
    df = carregar_dados()
    
    while True:
        meses_visualizar = int(input("Quantos meses você deseja visualizar os lotes? ").strip())
        data_limite = datetime.now() - pd.DateOffset(months=meses_visualizar-1)
        
        codigo_produto = input("Digite o código do produto: ").strip()
        
        df_codigo = df[df['Material'].str.endswith(codigo_produto.zfill(9)) | df['Material'].str.endswith(codigo_produto.zfill(10))]

        if not df_codigo.empty:
            df_codigo.loc[:, 'Data de produção'] = pd.to_datetime(df_codigo['Data de produção'], errors='coerce')
            df_codigo = df_codigo[df_codigo['Data de produção'] >= data_limite]
            
            if df_codigo.empty:
                print(f"Nenhum lote encontrado para o código {codigo_produto} nos últimos {meses_visualizar} meses.")
                continue

            lotes = df_codigo['Lote'].unique()
            if len(lotes) > 1:
                print(f"Existem múltiplos lotes para o código {codigo_produto}.")
                for i, lote in enumerate(lotes):
                    print(f"{i + 1}: Lote {lote}")
                escolha = int(input("Selecione o número do lote desejado: ")) - 1
                lote_selecionado = lotes[escolha]
                df_lote = df_codigo[df_codigo['Lote'] == lote_selecionado]
            else:
                df_lote = df_codigo
            
            descricao = df_lote['Descrição de material'].values[0]
            lote = df_lote['Lote'].values[0]
            validade = formatar_data(df_lote['Data de vencimento'].values[0])
            recebimento = formatar_data(df_lote['Data de produção'].values[0])
            quantidade = input("Digite a quantidade do produto: ").strip()
            nota_fiscal = input("Digite a Nota Fiscal: ").strip()
            fornecedor = input("Digite o Fornecedor: ").strip()

            # Preparando as informações para o QR Code
            informacoes_qr = f"Código: {codigo_produto}\nDescrição: {descricao}\nLote: {lote}\nValidade: {validade}\nRecebimento: {recebimento}\nQuantidade: {quantidade}\nNota Fiscal: {nota_fiscal}\nFornecedor: {fornecedor}"

            # Gerando o QR Code
            qrcode_img = criar_qrcode(informacoes_qr)

            # Armazenando os dados para possível edição
            dados_inseridos = {
                "Código": codigo_produto,
                "Descrição": descricao,
                "Lote": lote,
                "Validade": validade,
                "Recebimento": recebimento,
                "Quantidade": quantidade,
                "Nota Fiscal": nota_fiscal,
                "Fornecedor": fornecedor
            }

            # Permitir edição das informações
            editar_informacoes(dados_inseridos)

            # Criando a imagem A4
            imagem = criar_imagem_a4(
                dados_inseridos['Código'],
                dados_inseridos['Descrição'],
                dados_inseridos['Lote'],
                dados_inseridos['Validade'],
                dados_inseridos['Recebimento'],
                dados_inseridos['Quantidade'],
                dados_inseridos['Nota Fiscal'],
                dados_inseridos['Fornecedor'],
                qrcode_img
            )

            # Salvando a imagem
            caminho_imagem = rf"C:\Users\ter07068\OneDrive - M DIAS BRANCO S.A. INDUSTRIA E COMERCIO DE ALIMENTOS\Área de Trabalho\Python\Log demanda\QRCode\{dados_inseridos['Código']}_{dados_inseridos['Lote']}.png"
            imagem.save(caminho_imagem)
            print(f"Imagem salva como: {caminho_imagem}")

        else:
            print(f"Nenhum produto encontrado para o código {codigo_produto}.")
        
        continuar = input("Deseja continuar? (s/n): ").strip().lower()
        if continuar != 's':
            break

if __name__ == "__main__":
    main()
