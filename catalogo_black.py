import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os
import re
from datetime import date
import math

# === CONFIGURAÇÕES GERAIS ===
excel_path = "produtos.xlsx"
pdf_path = "catalogo_amaisciclo_bike_friday.pdf"
logo_path = "logo_amaisciclo.png"
img_desconto = "10porcem.jpg"
img_dir = "img_produtos"

# Cores
COR_AZUL_CODIGO = colors.Color(red=0.8, green=0.0, blue=0.0) # Vermelho Forte para o Código
COR_SOMBRA = colors.Color(0.8, 0.8, 0.8) # Sombra mais clara
COR_FUNDO_CARD = colors.white
COR_FUNDO_ESCURO = colors.Color(red=0.0, green=0.0, blue=0.0) # Preto Absoluto para o fundo da capa
COR_FAIXA_MEIO = colors.Color(red=0.1, green=0.1, blue=0.1) # Cinza Quase Preto para a faixa da capa/índice
COR_TEXTO_CLARO = colors.white
COR_FUNDO_CLARO = colors.Color(0.95, 0.97, 1.0)
COR_LINK_AZUL = colors.Color(red=0.9, green=0.4, blue=0.1) # Laranja Forte para Destaque/Links

# === FUNÇÃO DE NORMALIZAÇÃO DO CÓDIGO ===
def normalize_code(code_str):
    """Garante que o código exato seja a primeira tentativa de nome de arquivo."""
    code_str = str(code_str).strip()
    tentativas = [code_str]
    cleaned_code = re.sub(r'[.,]', '', code_str)
    if cleaned_code != code_str:
        tentativas.append(cleaned_code)
    underscore_code = code_str.replace('.', '_')
    if underscore_code != code_str and underscore_code != cleaned_code:
        tentativas.append(underscore_code)
    return list(set(tentativas))

# === FUNÇÕES DE LAYOUT (CABECALHO/RODAPE/CAPA/INDICE) ===

def cabecalho(c, largura, altura, pagina, categoria_atual=""):
    c.setFillColorRGB(0.95, 0.95, 0.95)
    ALTURA_CABECALHO = 1.5 * cm
    c.rect(0, altura - ALTURA_CABECALHO, largura, ALTURA_CABECALHO, fill=True, stroke=0)
    try:
        c.drawImage(logo_path, 2 * cm, altura - ALTURA_CABECALHO + 0.3 * cm, width=3.0 * cm, preserveAspectRatio=True, mask='auto')
    except:
        pass
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(7 * cm, altura - ALTURA_CABECALHO + 0.5 * cm, "BIKE FRIDAY A+Ciclo")

    
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.setLineWidth(1)
    c.line(1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm, largura - 1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm)

def rodape(c, largura, altura, pagina):
    ALTURA_RODAPE = 1.5 * cm
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, largura, ALTURA_RODAPE, fill=True, stroke=0)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 8)
    c.drawString(2 * cm, 0.5 * cm, "Acesse: b2b.amaisciclo.com.br")
    c.drawRightString(largura - 2 * cm, 0.5 * cm, f"Página {pagina}")

def criar_capa(c, largura, altura, logo_path, tipo_ordenacao):
    """Desenha a página de capa do catálogo, incluindo a data de geração."""
    data_geracao = date.today().strftime("%d/%m/%Y")
    
    c.setFillColor(COR_FUNDO_ESCURO)
    c.rect(0, 0, largura, altura, fill=1, stroke=0)

    faixa_altura = altura * 0.60
    faixa_y = altura * 0.60
    c.setFillColor(COR_FAIXA_MEIO)
    c.rect(0, faixa_y, largura, faixa_altura, fill=1, stroke=0)
    
    try:
        logo_capa_width = 13 * cm
        logo_topo_y = altura * 0.65
        c.drawImage(logo_path, (largura - logo_capa_width) / 2, logo_topo_y, 
                     width=logo_capa_width, preserveAspectRatio=True, mask='auto')
    except Exception as e:
        c.setFillColor(COR_TEXTO_CLARO)
        c.setFont("Helvetica-Bold", 40)
        c.drawCentredString(largura / 2, altura * 0.82, "A+CICLO")

    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica-Bold", 48)
    c.drawCentredString(largura / 2, altura * 0.50, "BIKE FRIDAY")
    c.setFillColor(colors.red)
    c.setFont("Helvetica", 22)
    c.drawCentredString(largura / 2, altura * 0.45, "PRODUTOS COM ATÉ 44% DE DESCONTO")
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 14)
    c.drawCentredString(largura / 2, altura * 0.38, "Peças e Acessórios para Ciclismo")

    box_largura = 12 * cm
    box_altura = 1.2 * cm
    box_x = (largura - box_largura) / 2
    box_y = altura * 0.20
    
    c.setFillColor(COR_FUNDO_ESCURO)
    c.roundRect(box_x, box_y, box_largura, box_altura, 0.5 * cm, fill=1, stroke=0)
    
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 10)
    
    texto_data = f" Válido de {data_geracao} até 29/11/2025!"
        
    c.drawCentredString(largura / 2, box_y + 0.4 * cm, texto_data)

    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 8)
    c.drawCentredString(largura / 2, 1 * cm, "Catálogo Digital - Versão 2.0")

    c.showPage()


# === MODO DE GERAÇÃO FIXO ===
TIPO_ORDENACAO = 'A'
print("Catálogo configurado para ordenação Alfabética (Geral).")

# === LEITURA E PRÉ-PROCESSAMENTO DA PLANILHA ===
try:
    df = pd.read_excel(excel_path, dtype={'Código do Produto': str})
    df['Categoria'] = df['Categoria'].fillna('Diversos').astype(str).str.strip()
    
    # Ordenação base: SEMPRE ALFABÉTICA
    # Ordenar todos os produtos por Descrição e, secundariamente, por Código do Produto
    df = df.sort_values(by=['Descrição', 'Código do Produto'])
    
    # O iterador será a lista de todas as linhas do DataFrame (simulando um único grupo)
    produtos_iteracao = [("ALFABÉTICA GERAL", df.iterrows())] 
    
except FileNotFoundError:
    print(f"ERRO: Arquivo Excel não encontrado em: {excel_path}")
    exit()
except Exception as e:
    print(f"ERRO: Falha ao ler o arquivo Excel: {e}")
    exit()

# === CRIAÇÃO DO PDF ===
c = canvas.Canvas(pdf_path, pagesize=A4)
largura, altura = A4

# Estilos para o Paragraph (descrição)
styles = getSampleStyleSheet()
styleN = styles['Normal']
styleN.fontSize = 5.5 # Reduzido de 6 para 5.5
styleN.leading = 6.5  # Reduzido de 8 para 6.5
styleN.alignment = 1
styleN.fontName = 'Helvetica'
styleN.textColor = colors.black

# === CONFIGURAÇÕES DE LAYOUT DO PRODUTO (BLOCO MENOR) ===
ALTURA_RODAPE = 1.5 * cm
ALTURA_CABECALHO = 1.5 * cm
MARGEM_SUPERIOR = ALTURA_CABECALHO + 0.5 * cm

produtos_por_linha = 3
espacamento_horizontal = 1 * cm
largura_produto_bloco = (largura - 3 * cm - 2 * espacamento_horizontal) / produtos_por_linha
altura_produto_bloco = 5.5 * cm 
espacamento_vertical = altura_produto_bloco + 0.3 * cm 
y_inicio_produtos = altura - MARGEM_SUPERIOR
x_inicio = 1.5 * cm

# --- INÍCIO DA GERAÇÃO DO PDF ---
print("Iniciando geração da Capa...")

# 1. Gerar a Capa (passando o tipo de ordenação para o texto)
criar_capa(c, largura, altura, logo_path, TIPO_ORDENACAO)

pagina = 1 # A primeira página de conteúdo

# 3. Loop Final para Conteúdo
y = y_inicio_produtos
erros_imagem = 0
produto_index_na_pagina = 0

print(f"Iniciando conteúdo do catálogo (a partir da Página {pagina})...")

# Itera sobre o grupo único "ALFABÉTICA GERAL"
for grupo_key, grupo_data in produtos_iteracao:
    
    categoria_atual = grupo_key # Será "ALFABÉTICA GERAL"
    
    # Itera sobre os produtos do grupo
    it_produtos = grupo_data

    # Desenha cabeçalho na primeira página
    cabecalho(c, largura, altura, pagina, categoria_atual)
    
    for i, row in it_produtos:
        col = produto_index_na_pagina % produtos_por_linha
        x_bloco = x_inicio + col * (largura_produto_bloco + espacamento_horizontal)

        codigo_produto = str(row.get("Código do Produto", "")).strip()
        descricao = str(row.get("Descrição", "")).strip()

        y_bloco_topo = y
        x_bloco_centro = x_bloco + largura_produto_bloco / 2

        # 1. Cartão, Sombra, Imagem, Botão, Descrição (Lógica de desenho mantida)
        sombra_offset = 0.05 * cm
        c.setFillColor(COR_SOMBRA)
        c.roundRect(x_bloco + sombra_offset, y_bloco_topo - altura_produto_bloco + sombra_offset, largura_produto_bloco, altura_produto_bloco, 0.2 * cm, fill=1, stroke=0)
        c.setFillColor(COR_FUNDO_CARD)
        c.setStrokeColor(COR_SOMBRA)
        c.setLineWidth(0.5)
        c.roundRect(x_bloco, y_bloco_topo - altura_produto_bloco, largura_produto_bloco, altura_produto_bloco, 0.2 * cm, fill=1, stroke=1)
        
        # --- ÁREA DA IMAGEM ---
        max_altura_img_area = 3.5 * cm 
        y_img_area_topo = y_bloco_topo - 0.3 * cm
        y_img_area_fundo = y_img_area_topo - max_altura_img_area 
        largura_img_area = largura_produto_bloco * 0.8
        
        image_loaded = False
        caminho_imagem = None
        if codigo_produto:
            for cod in normalize_code(codigo_produto):
                for ext in [".jpg", ".jpeg", ".png"]:
                    tentativa = os.path.join(img_dir, f"{cod}{ext}")
                    if os.path.exists(tentativa):
                        caminho_imagem = tentativa
                        image_loaded = True
                        break
                if image_loaded: break
        
        if image_loaded:
            try:
                img = ImageReader(caminho_imagem)
                img_largura, img_altura = img.getSize()
                proporcao = img_largura / img_altura
                largura_final = largura_img_area
                altura_final = largura_final / proporcao
                if altura_final > max_altura_img_area:
                    altura_final = max_altura_img_area
                    largura_final = altura_final * proporcao 
                x_img = x_bloco_centro - largura_final / 2
                c.drawImage(img, x_img, y_img_area_fundo + (max_altura_img_area - altura_final)/2, 
                             width=largura_final, height=altura_final, preserveAspectRatio=True, mask='auto')
            except Exception as e:
                erros_imagem += 1
        else:
            erros_imagem += 1
            c.setFillColor(colors.lightgrey)
            c.rect(x_bloco_centro - largura_img_area/2, y_img_area_fundo, largura_img_area, max_altura_img_area, fill=1, stroke=0)
            c.setFillColor(colors.darkgrey)
            c.setFont("Helvetica-Oblique", 8)
            c.drawCentredString(x_bloco_centro, y_img_area_fundo + max_altura_img_area / 2, "Sem imagem")
            
        # --- POSICIONAMENTO DINÂMICO DE PREÇOS E CÓDIGO ---
        
        # Inicia a posição abaixo da área da imagem
        y_current = y_img_area_fundo - 0.2 * cm
        precos_existentes = False
        
        preco_antigo = row.get("Preço Antigo", "")
        preco_promocional = row.get("Preço Promoção", "")

        # 1. PREÇOS
        if preco_promocional:
            precos_existentes = True
            
            # Preço antigo (cinza e riscado)
            if preco_antigo:
                preco_antigo_txt = f"R$ {preco_antigo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                c.setFont("Helvetica", 7)
                c.setFillColor(colors.grey)
                c.drawCentredString(x_bloco_centro, y_current, preco_antigo_txt)

                # Linha de risco sobre o preço antigo
                text_width = c.stringWidth(preco_antigo_txt, "Helvetica", 7)
                c.setStrokeColor(colors.grey)
                c.setLineWidth(0.5)
                c.line(x_bloco_centro - text_width / 2, y_current + 1, x_bloco_centro + text_width / 2, y_current + 1)
                
                y_current -= 0.35 * cm  # Move para baixo (espaço entre preços)

            # Preço promocional (vermelho e maior)
            preco_promo_txt = f"R$ {preco_promocional:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.red)
            c.drawCentredString(x_bloco_centro, y_current, preco_promo_txt)
            
            y_current -= 0.6 * cm # Adiciona espaço abaixo do preço promocional para o código

        # 2. CÓDIGO DO PRODUTO (sempre abaixo do último preço ou abaixo da imagem se sem preço)
        
        largura_cod_btn = largura_produto_bloco * 0.2
        altura_cod_btn = 0.4 * cm
        x_cod_btn = x_bloco_centro - largura_cod_btn / 2

        if precos_existentes:
            # Posição calculada após os preços
            y_cod_btn = y_current 
        else:
            # Posição padrão se não houver preços
            y_cod_btn = y_img_area_fundo - 0.9 * cm 

        # Desenho do botão do código
        c.setFillColor(COR_AZUL_CODIGO)
        c.setFont("Helvetica-Bold", 6.5) 
        c.roundRect(x_cod_btn, y_cod_btn, largura_cod_btn, altura_cod_btn, 0.15 * cm, fill=1, stroke=0)
        c.setFillColor(colors.white) 
        c.drawCentredString(x_bloco_centro, y_cod_btn + 0.10 * cm, codigo_produto) 

        # 3. DESCRIÇÃO (Fundo do Card)
        c.setFillColor(colors.black)
        desc_limpa = " ".join(descricao.split())
        p = Paragraph(desc_limpa, styleN)
        largura_desc_area = largura_produto_bloco * 0.9
        y_desc_base = y_bloco_topo - altura_produto_bloco + 0.2 * cm 
        # Área de wrap reduzida para 0.6 * cm para limitar a altura da descrição
        p_width, p_height = p.wrapOn(c, largura_desc_area, 0.5 * cm) 
        c.saveState()
        c.translate(x_bloco_centro - p_width / 2, y_desc_base)
        p.drawOn(c, 0, 0)
        c.restoreState()


        # === PRÓXIMO BLOCO / QUEBRA DE PÁGINA ===
        if col == produtos_por_linha - 1:
            y -= espacamento_vertical
            produto_index_na_pagina = 0
            
            if y - altura_produto_bloco < ALTURA_RODAPE + 0.5 * cm:
                rodape(c, largura, altura, pagina)
                c.showPage()
                pagina += 1
                cabecalho(c, largura, altura, pagina, categoria_atual)
                y = y_inicio_produtos 
        else:
            produto_index_na_pagina += 1


# === FINALIZA ===
# Garante que o rodapé da última página seja desenhado
if y != y_inicio_produtos or produto_index_na_pagina != 0:
    rodape(c, largura, altura, pagina)
    
c.save()

print("\n--- Geração Concluída ---")
print(f"✅ Catálogo gerado com sucesso: {pdf_path}")
print(f"Total de páginas: {pagina}")
if erros_imagem > 0:
    print(f"⚠️ {erros_imagem} imagem(ns) não encontrada(s) ou falhou no carregamento.")