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
pdf_path = "catalogo_amaisciclo_compacto.pdf"
logo_path = "logo_bikeline.png"
img_dir = "img_produtos"

# Cores
COR_AZUL_CODIGO = colors.Color(0.25, 0.45, 0.85)
COR_SOMBRA = colors.Color(0.9, 0.9, 0.9)
COR_FUNDO_CARD = colors.white
COR_FUNDO_ESCURO = colors.Color(red=28/255, green=35/255, blue=46/255) 
COR_FAIXA_MEIO = colors.Color(red=39/255, green=47/255, blue=58/255)
COR_TEXTO_CLARO = colors.white
COR_FUNDO_CLARO = colors.Color(0.95, 0.97, 1.0)
COR_LINK_AZUL = colors.Color(0.1, 0.3, 0.8)


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
    c.drawString(6 * cm, altura - ALTURA_CABECALHO + 0.7 * cm, "Catálogo A+Ciclo")
    
    # Adiciona a categoria atual no cabeçalho
    if categoria_atual:
        c.setFont("Helvetica", 10)
        c.drawRightString(largura - 2 * cm, altura - ALTURA_CABECALHO + 1.2 * cm, f"Categoria: {categoria_atual.upper()}")
    
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.setLineWidth(1)
    c.line(1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm, largura - 1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm)

def rodape(c, largura, altura, pagina):
    ALTURA_RODAPE = 1.5 * cm
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, largura, ALTURA_RODAPE, fill=True, stroke=0)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 8)
    c.drawString(2 * cm, 0.5 * cm, "www.bikeline.com.br")
    c.drawRightString(largura - 2 * cm, 0.5 * cm, f"Página {pagina}")

def criar_capa(c, largura, altura, logo_path, tipo_ordenacao):
    """Desenha a página de capa do catálogo, incluindo a data de geração."""
    data_geracao = date.today().strftime("%d/%m/%Y")
    
    c.setFillColor(COR_FUNDO_ESCURO)
    c.rect(0, 0, largura, altura, fill=1, stroke=0)

    faixa_altura = altura * 0.25 
    faixa_y = altura * 0.50
    c.setFillColor(COR_FAIXA_MEIO)
    c.rect(0, faixa_y, largura, faixa_altura, fill=1, stroke=0)
    
    try:
        logo_capa_width = 7 * cm
        logo_topo_y = altura * 0.82
        c.drawImage(logo_path, (largura - logo_capa_width) / 2, logo_topo_y, 
                    width=logo_capa_width, preserveAspectRatio=True, mask='auto')
    except Exception as e:
        c.setFillColor(COR_TEXTO_CLARO)
        c.setFont("Helvetica-Bold", 40)
        c.drawCentredString(largura / 2, altura * 0.82, "A+CICLO")

    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica-Bold", 32)
    c.drawCentredString(largura / 2, altura * 0.45, "CATÁLOGO")
    c.setFont("Helvetica", 28)
    c.drawCentredString(largura / 2, altura * 0.40, "DE PRODUTOS")
    c.setFont("Helvetica", 14)
    c.drawCentredString(largura / 2, altura * 0.35, "Peças e Acessórios para Ciclismo")

    box_largura = 12 * cm
    box_altura = 1.2 * cm
    box_x = (largura - box_largura) / 2
    box_y = altura * 0.20
    
    c.setFillColor(COR_FUNDO_ESCURO)
    c.roundRect(box_x, box_y, box_largura, box_altura, 0.5 * cm, fill=1, stroke=0)
    
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 10)
    if tipo_ordenacao == 'C':
        ordem_texto = "Organizado por Categoria"
    else:
        ordem_texto = "Ordenados Alfabeticamente"
        
    texto_data = f"{ordem_texto} - Gerado em {data_geracao}"
    c.drawCentredString(largura / 2, box_y + 0.4 * cm, texto_data)

    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 8)
    c.drawCentredString(largura / 2, 1 * cm, "Catálogo Digital - Versão 2.0")

    c.showPage()

def criar_indice(c, largura, altura, categorias_map):
    """Desenha a página de índice com links para as categorias."""
    
    c.setFillColor(COR_FAIXA_MEIO)
    c.rect(0, altura * 0.75, largura, altura * 0.25, fill=1, stroke=0)
    
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(largura / 2, altura * 0.88, "ÍNDICE DE CATEGORIAS")
    c.setFont("Helvetica", 14)
    c.drawCentredString(largura / 2, altura * 0.83, "Navegue pelas principais categorias de produtos")
    
    categorias_ordenadas = sorted(categorias_map.keys())
    num_categorias = len(categorias_ordenadas)
    meio = math.ceil(num_categorias / 2)

    x_col1 = 2 * cm
    x_col2 = largura / 2 + 0.5 * cm
    y_start = altura * 0.70
    y_step = 1.5 * cm
    
    def desenhar_item_indice(c, num, nome, pagina, x, y):
        c.setStrokeColor(colors.lightgrey)
        c.setLineWidth(0.5)
        c.rect(x - 0.2 * cm, y - 1.2 * cm, largura/2 - 2*cm, 1.2 * cm, fill=0, stroke=1)
        
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 12)
        c.drawString(x, y - 0.7 * cm, f"{num}. {nome.upper()}")
        
        c.setFillColor(COR_LINK_AZUL)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x_col1 + largura/2 - 4.5 * cm, y - 0.7 * cm, f"Pág. {pagina}")
        

    for i in range(meio):
        nome = categorias_ordenadas[i]
        pagina = categorias_map[nome]
        desenhar_item_indice(c, i + 1, nome, pagina, x_col1, y_start - i * y_step)

    for j in range(num_categorias - meio):
        nome = categorias_ordenadas[meio + j]
        pagina = categorias_map[nome]
        desenhar_item_indice(c, meio + j + 1, nome, pagina, x_col2, y_start - j * y_step)

    box_largura_dica = largura - 4 * cm
    box_altura_dica = 3 * cm
    box_x_dica = 2 * cm
    box_y_dica = 3 * cm
    
    c.setFillColor(COR_FUNDO_CLARO)
    c.setStrokeColor(COR_LINK_AZUL)
    c.setLineWidth(1)
    c.roundRect(box_x_dica, box_y_dica, box_largura_dica, box_altura_dica, 0.5 * cm, fill=1, stroke=1)
    
    c.setFillColor(COR_LINK_AZUL)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(box_x_dica + 0.5 * cm, box_y_dica + box_altura_dica - 0.5 * cm, "Dica de Navegação")

    c.setFillColor(colors.darkgrey)
    c.setFont("Helvetica", 8)
    c.drawString(box_x_dica + 0.5 * cm, box_y_dica + box_altura_dica - 1.5 * cm, "Os produtos estão organizados por categorias para facilitar sua busca.")
    c.drawString(box_x_dica + 0.5 * cm, box_y_dica + box_altura_dica - 2.2 * cm, "Use este índice como referência rápida para encontrar o que procura.")
    
    c.showPage()

# === SELEÇÃO DO MODO DE GERAÇÃO (NOVO) ===
while True:
    print("\n------------------------------------------------------")
    print("Selecione o modo de organização do Catálogo:")
    print("  [C] - Por Categoria (Com Índice de Categorias)")
    print("  [A] - Alfabética (Geral, sem Índice)")
    escolha = input("Digite C ou A: ").strip().upper()
    print("------------------------------------------------------")
    if escolha in ['C', 'A']:
        TIPO_ORDENACAO = escolha
        break
    else:
        print("Opção inválida. Por favor, digite C para Categoria ou A para Alfabética.")

# === LEITURA E PRÉ-PROCESSAMENTO DA PLANILHA ===
try:
    df = pd.read_excel(excel_path, dtype={'Código do Produto': str})
    df['Categoria'] = df['Categoria'].fillna('Diversos').astype(str).str.strip()
    
    # Ordenação base
    if TIPO_ORDENACAO == 'C':
        # Ordenar por Categoria e, dentro dela, por Código do Produto
        df = df.sort_values(by=['Categoria', 'Código do Produto'])
        
        # Preparação para o loop de categorias
        produtos_iteracao = df.groupby('Categoria', sort=True)
    else: # Ordem Alfabética Geral
        # Ordenar todos os produtos por Descrição ou Código do Produto
        df = df.sort_values(by=['Descrição', 'Código do Produto'])
        
        # O iterador será a lista de todas as linhas do DataFrame
        # Criamos um iterador que simula o groupby para manter a estrutura do loop
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
styleN.fontSize = 6
styleN.leading = 8 
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

# 2. Gerar o Índice (APENAS SE FOR ORDENADO POR CATEGORIA)
if TIPO_ORDENACAO == 'C':
    
    # 2a. Pré-mapeamento das Páginas de Categoria (necessário para o índice)
    current_page_map = pagina + 1 # Começa após a capa e índice
    categorias_paginas = {}
    
    # Usando uma estimativa simplificada para evitar complexidade excessiva
    produtos_por_pagina = math.floor((altura - MARGEM_SUPERIOR - ALTURA_RODAPE) / espacamento_vertical) * produtos_por_linha
    
    for categoria_nome, grupo_df in produtos_iteracao:
        categorias_paginas[categoria_nome] = current_page_map
        
        # Calcula as páginas para esta categoria e avança o contador
        num_produtos = len(grupo_df)
        paginas_categoria = math.ceil(num_produtos / produtos_por_pagina)
        current_page_map += max(1, paginas_categoria) # Garante que a próxima categoria inicie na página correta

    print(f"Gerando Índice (Página {pagina})...")
    criar_indice(c, largura, altura, categorias_paginas)
    pagina += 1 # Avança para a primeira página de conteúdo (após a capa e índice)

# 3. Loop Final para Conteúdo
y = y_inicio_produtos
erros_imagem = 0
produto_index_na_pagina = 0

print(f"Iniciando conteúdo do catálogo (a partir da Página {pagina})...")

# Itera sobre os grupos (categorias ou o grupo único "ALFABÉTICA GERAL")
for grupo_key, grupo_data in produtos_iteracao:
    
    categoria_atual = grupo_key if TIPO_ORDENACAO == 'C' else ""

    # FORÇAR NOVA PÁGINA PARA CADA CATEGORIA (se não for a primeira página de conteúdo)
    if TIPO_ORDENACAO == 'C' and produto_index_na_pagina != 0:
        # Finaliza a página anterior
        rodape(c, largura, altura, pagina)
        c.showPage()
        pagina += 1
        y = y_inicio_produtos # Reinicia Y no topo da nova página
    
    # Itera sobre os produtos do grupo/categoria
    # Se for Categoria, grupo_data é um DataFrame. Se for Alfabética, é um iterrows.
    if TIPO_ORDENACAO == 'C':
        it_produtos = grupo_data.iterrows()
    else:
        it_produtos = grupo_data

    # Desenha cabeçalho na primeira página deste grupo
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
        
        max_altura_img_area = 3.5 * cm 
        y_img_area_topo = y_bloco_topo - 0.3 * cm
        y_img_area_fundo = y_img_area_topo - max_altura_img_area 
        largura_img_area = largura_produto_bloco * 0.95
        
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
            
        c.setFillColor(COR_AZUL_CODIGO)
        c.setFont("Helvetica-Bold", 6.5) 
        largura_cod_btn = largura_produto_bloco * 0.25
        altura_cod_btn = 0.4 * cm
        x_cod_btn = x_bloco_centro - largura_cod_btn / 2
        y_cod_btn = y_img_area_fundo - altura_cod_btn - 0.1 * cm 
        c.roundRect(x_cod_btn, y_cod_btn, largura_cod_btn, altura_cod_btn, 0.15 * cm, fill=1, stroke=0)
        c.setFillColor(colors.white) 
        c.drawCentredString(x_bloco_centro, y_cod_btn + 0.15 * cm, codigo_produto) 
        
        c.setFillColor(colors.black)
        desc_limpa = " ".join(descricao.split())
        p = Paragraph(desc_limpa, styleN)
        largura_desc_area = largura_produto_bloco * 0.9
        y_desc_base = y_bloco_topo - altura_produto_bloco + 0.2 * cm 
        p_width, p_height = p.wrapOn(c, largura_desc_area, 0.8 * cm)
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
rodape(c, largura, altura, pagina)
c.save()

print("\n--- Geração Concluída ---")
print(f"✅ Catálogo gerado com sucesso: {pdf_path}")
print(f"Total de páginas: {pagina}")
if erros_imagem > 0:
    print(f"⚠️ {erros_imagem} imagem(ns) não encontrada(s) ou falhou no carregamento.")