import pandas as pd
import qrcode
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

def remover_prefixo_e_zeros(codigo_completo):
    codigo_limpo = codigo_completo.lstrip("EMBZMP0")
    return codigo_limpo

def gerar_codigo_controle(codigo_completo, lote):
    codigo_limpo = remover_prefixo_e_zeros(codigo_completo)
    ultimos_3_codigo = codigo_limpo[-3:]
    ultimos_3_lote = lote[-3:]
    prefixo = "EMB" if codigo_completo.startswith("EMB") or codigo_completo.startswith("ZEMB") else "MP"
    return f"{prefixo}{ultimos_3_codigo}{ultimos_3_lote}"

def criar_imagem_a4(codigo, descricao, lote, validade, recebimento, quantidade, nota_fiscal, fornecedor, qrcode_img, codigo_controle):
    largura_a4, altura_a4 = (2480, 3508)
    imagem = Image.new("RGB", (largura_a4, altura_a4), "white")
    draw = ImageDraw.Draw(imagem)
    
    try:
        fonte_codigo_enorme = ImageFont.truetype("arial.ttf", 450)
        fonte_codigo_grande = ImageFont.truetype("arial.ttf", 300)  # Para o código e a quantidade em grande destaque
        fonte_codigo_media = ImageFont.truetype("arial.ttf", 150)   # Para o texto da quantidade
        fonte_outros = ImageFont.truetype("arial.ttf", 100)         # Para os demais textos
        fonte_pequena = ImageFont.truetype("arial.ttf", 50)         # Para o código de controle
    except IOError:
        fonte_codigo_grande = fonte_codigo_media = fonte_outros = fonte_pequena = ImageFont.load_default()
    
    margem = 50
    y_offset = margem

    # Código
    draw.text((margem, y_offset), "Código", font=fonte_codigo_media, fill="black")
    y_offset += fonte_codigo_media.getbbox("Código")[3] + margem

    draw.text((margem, y_offset), codigo, font=fonte_codigo_grande, fill="black")
    y_offset += fonte_codigo_grande.getbbox(codigo)[3] + 2 * margem

    # Descrição
    draw.text((margem, y_offset), f"Descrição: {descricao}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Descrição: {descricao}")[3] + margem

    # Lote
    draw.text((margem, y_offset), f"Lote: {lote}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Lote: {lote}")[3] + margem

    # Data de Validade
    draw.text((margem, y_offset), f"Data de Validade: {validade}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Data de Validade: {validade}")[3] + margem

    # Data de Recebimento
    draw.text((margem, y_offset), f"Data de Recebimento: {recebimento}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Data de Recebimento: {recebimento}")[3] + margem

    # Nota Fiscal
    draw.text((margem, y_offset), f"Nota Fiscal: {nota_fiscal}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Nota Fiscal: {nota_fiscal}")[3] + margem
    
    # Fornecedor
    draw.text((margem, y_offset), f"Fornecedor: {fornecedor}", font=fonte_outros, fill="black")
    y_offset += fonte_outros.getbbox(f"Fornecedor: {fornecedor}")[3] + 2 * margem

    # Quantidade
    draw.text((margem, y_offset), "Quantidade", font=fonte_codigo_media, fill="black")
    y_offset += fonte_codigo_media.getbbox("Quantidade")[3] + margem

    # Quantidade em números (grande)
    draw.text((margem, y_offset), quantidade, font=fonte_codigo_enorme, fill="black")
    y_offset += fonte_codigo_enorme.getbbox(quantidade)[3] + 2 * margem

    # Ajuste do QR Code
    qr_largura, qr_altura = qrcode_img.size
    novo_tamanho_qr = (qr_largura // 2, qr_altura // 2)
    qrcode_img = qrcode_img.resize(novo_tamanho_qr, Image.Resampling.LANCZOS)

    # Calcula a posição do QR Code no canto inferior direito
    pos_qr_x = (largura_a4 - novo_tamanho_qr[0]) // 2 - margem
    pos_qr_y = altura_a4 - novo_tamanho_qr[1] - margem

    # Coloca o QR Code na imagem na posição calculada
    imagem.paste(qrcode_img, (pos_qr_x, pos_qr_y))

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
            
            quantidade = input("Digite a quantidade no pallet: ").strip()
            nota_fiscal = input("Digite a Nota Fiscal: ").strip()
            fornecedor = input("Digite o Fornecedor: ").strip()
            
            codigo_completo = df_lote['Material'].values[0]
            codigo_controle = gerar_codigo_controle(codigo_completo, lote)
            
            data = {
                'Código': codigo_completo,
                'Descrição': descricao,
                'Lote': lote,
                'Data de Fabricação': recebimento,
                'Data de vencimento': validade,
                'Quantidade': quantidade,
                'Nota Fiscal': nota_fiscal,
                'Fornecedor': fornecedor
            }
            
            # Etapa de edição
            editar_informacoes(data)
            
            data_str = "\n".join(f"{key}: {value}" for key, value in data.items())
            qr_img = criar_qrcode(data_str)
            
            a4_img = criar_imagem_a4(
                codigo_completo, 
                descricao, 
                lote, 
                validade, 
                recebimento, 
                quantidade, 
                nota_fiscal, 
                fornecedor, 
                qr_img, 
                codigo_controle
            )
            
            diretorio_qrcode = r'C:\Users\ter07068\OneDrive - M DIAS BRANCO S.A. INDUSTRIA E COMERCIO DE ALIMENTOS\Área de Trabalho\Python\Log demanda'
            nome_arquivo = f"{codigo_completo}_qrcode.png"
            caminho_completo = os.path.join(diretorio_qrcode, nome_arquivo)
            a4_img.save(caminho_completo)
            
            print(f"Imagem com QR Code salva em: {caminho_completo}")
            
            continuar = input("Deseja gerar outro QR Code? (s/n): ").strip().lower()
            if continuar != 's':
                break
        else:
            print(f"Código {codigo_produto} não encontrado.")

if __name__ == "__main__":
    main()
