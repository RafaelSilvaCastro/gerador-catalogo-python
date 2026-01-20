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

# === CONFIGURAÇÕES GERAIS ===
excel_path = "produtos.xlsx"
pdf_path = "pdfs/catalogo_amaisciclo_bicicletas.pdf"
logo_path = "img/logo_amaisciclo.png"
fundo = "img/fundo.jpg"
img_desconto = "10porcem.jpg"
img_dir = "img/img_produtos"

# Cores
COR_DESTAQUE_CODIGO = colors.Color(red=0.8, green=0.0, blue=0.0) # Vermelho Forte para o Código
COR_SOMBRA = colors.Color(0.8, 0.8, 0.8) # Sombra mais clara
COR_FUNDO_CARD = colors.white
COR_FUNDO_ESCURO = colors.Color(red=0.0, green=0.0, blue=0.0) # Preto Absoluto para o fundo da capa
COR_FAIXA_MEIO = colors.Color(red=0.1, green=0.1, blue=0.1) # Cinza Quase Preto para a faixa da capa/índice
COR_TEXTO_CLARO = colors.white
COR_FUNDO_CLARO = colors.Color(0.95, 0.97, 1.0)
COR_LINK_AZUL = colors.Color(red=0.9, green=0.4, blue=0.1) # Laranja Forte para Destaque/Links

# Variáveis de Layout
ALTURA_RODAPE = 1.5 * cm
ALTURA_CABECALHO = 1.5 * cm
MARGEM_SUPERIOR = ALTURA_CABECALHO + 0.5 * cm

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

def cabecalho(c, largura, altura, pagina, titulo_grupo=""):
    """
    Desenha o cabeçalho, agora usando titulo_grupo para exibir a Categoria e o Tamanho.
    """
    c.setFillColorRGB(0.95, 0.95, 0.95)
    ALTURA_CABECALHO = 1.5 * cm
    c.rect(0, altura - ALTURA_CABECALHO, largura, ALTURA_CABECALHO, fill=True, stroke=0)
    try:
        c.drawImage(logo_path, 2 * cm, altura - ALTURA_CABECALHO + 0.3 * cm, width=3.0 * cm, preserveAspectRatio=True, mask='auto')
    except:
        pass
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 18)
    # Exibe o título do grupo (ex: Bicicletas - Tamanho 17)
    c.drawString(1.6 * cm, altura - ALTURA_CABECALHO + 0.5 * cm, f"CATÁLOGO: {titulo_grupo}")
    
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
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(largura / 2, altura * 0.50, "CATÁLOGO DE BICICLETAS")
    c.setFillColor(colors.red)
    c.setFont("Helvetica", 20)
    c.drawCentredString(largura / 2, altura * 0.45, "EDIÇÃO - 1/1")
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 14)
    c.drawCentredString(largura / 2, altura * 0.38, "")

    box_largura = 12 * cm
    box_altura = 1.2 * cm
    box_x = (largura - box_largura) / 2
    box_y = altura * 0.04
    
    c.setFillColor(COR_FUNDO_ESCURO)
    c.roundRect(box_x, box_y, box_largura, box_altura, 0.5 * cm, fill=1, stroke=0)
    
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 10)
    
    texto_data = f" Emissão {data_geracao}"
        
    c.drawCentredString(largura / 2, box_y + 0.4 * cm, texto_data)

    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 8)
    c.drawCentredString(largura / 2, 1 * cm, "Catálogo Digital - Versão Bicicletas")

    c.showPage()
    
# === PÁGINA DE ESPECIFICAÇÕES (AGORA COM 2 COLUNAS E FLUXO DE PÁGINAS) ===
def criar_pagina_especificacao(c, largura, altura, pagina_num):
    """
    Desenha a página de especificações e condições comerciais com layout de duas colunas,
    incluindo uma imagem de fundo (watermark) e garantindo a quebra de página.
    Os componentes de cada bicicleta são ordenados alfabeticamente.
    """
    
    # Função auxiliar para desenhar Fundo, Cabeçalho e Rodapé em qualquer página de especificação
    def draw_spec_page_elements(c_local, largura_local, altura_local, pagina_num_local, categoria_atual_local):
        # 1. Desenha o Fundo (Watermark)
        try:
            c_local.saveState()
            
            # Usando o fundo como imagem de fundo (watermark)
            bg_img = ImageReader(fundo) 
            img_w, img_h = bg_img.getSize()
            
            # Define a opacidade (muito baixa para não atrapalhar a leitura)
            c_local.setFillAlpha(0.6)
            
            # Calcula a escala para caber na página (manter proporção e centralizar)
            # Aumenta o tamanho (x1.5) para garantir que cubra a página completamente
            scale = min(largura_local / img_w, altura_local / img_h) * 1.5 
            final_w = img_w * scale
            final_h = img_h * scale
            
            x_bg = (largura_local - final_w) / 2
            y_bg = (altura_local - final_h) / 2
            
            # Desenha a imagem (o mask='auto' ajuda a lidar com fundos transparentes)
            c_local.drawImage(bg_img, x_bg, y_bg, width=final_w, height=final_h, 
                              preserveAspectRatio=True, mask='auto')
            
            c_local.restoreState() # Restaura o estado original de opacidade (1.0)
            
        except Exception as e:
            # Em caso de falha no carregamento da imagem de fundo, não trava o processo.
            pass

        # 2. Desenha Cabeçalho e Rodapé
        cabecalho(c_local, largura_local, altura_local, pagina_num_local, categoria_atual_local)
        rodape(c_local, largura_local, altura_local, pagina_num_local)

    texto_especificacao_raw = """
    <h1><strong>DESCRIÇÃO BICICLETA 26 VIKING</strong></h1>
    <p>QUADRO 26 AL VIKING DIRT J.</p>
    <p>CAMARAS 26 BUTIL</p>
    <p>CARRINHO DE SELIM IMP PRETO</p>
    <p>FREIO DISCO DT/TR (PINCA/ROTOR)</p>
    <p>GARFO 29 S. 28.6 AH SET PRETO</p>
    <p>MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO</p>
    <p>PEDIV TRIPLO ENC 24/34/42 PRETO</p>
    <p>PNEUS 26X1.95 KENDA K-90 SLICK PTO</p>
    <p>ARO 26 VMAXX SL VZAN 36F PTO DISC S/ILHOS</p>
    <p>RAIO 251X2.0MM IMP ZINC</p>
    <p>CUBO AL DT/TR ESFERADO PRETO C/ BLOCAGEM</p>
    <p>SELIM MTB PRETO C/CARRINHO</p>
    <p>CANOTE AÇO 27.2X350MM PRETO</p>
    <p>GUIDAO DH 31.8 ACO 700MM PRETO</p>
    <p>KIT TRANSMISSÃO 21V</p>
    <p>PEDAL 9/16 PLATAFORMA NYLON PRETO</p>
    <p>CORRENTE FINA 7/8V MODELO DG51 116 INDEX</p>
    <p>ABRAC. SELIM 31.8 AL PRETO</p>
    <p>MOV DIRECAO M. OVER PRETO</p>
    <p>RODA LIVRE 7V INDEX 14/28D</p>
    <p>MANOPLA MTB PRETO</p>
    <p>SUP MTB 31.8 (60MM) OU SIMILAR PTO</p>
    
    <h1><strong>DESCRIÇÃO BICICLETA 29 FIRST 21V</strong></h1>
    <p>QUADRO 29 ALUMÍNIO FIRST</p>
    <p>CÂMARAS BUTIL 48MM</p>
    <p>CARRINHO DE SELIM IMP PRETO</p>
    <p>FREIO DISCO DT/TR (PINCA/ROTOR)</p>
    <p>GARFO 29 S. 28.6 AH SET PRETO, CANELAS DE 38MM E CURSO DE 80MM</p>
    <p>MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO</p>
    <p>PEDIV TRIPLO ENC 24/34/42 PRETO</p>
    <p>PNEUS 29X2.10 SRI PTO</p>
    <p>CUBO AL DT/TR 36 FUROS ESFERADO C/ BLOCAGEM PRETO</p>
    <p>AROS 29/36 FUROS PRETO DISC</p>
    <p>SELIM MTB PRETO C/CARRINHO</p>
    <p>CANOTE AÇO 27.2X350MM PRETO</p>
    <p>GUIDAO DH 31.8 ACO 700MM PRETO</p>
    <p>KIT TRANSMISSÃO 21V</p>
    <p>PEDAL 9/16 PLATAFORMA NYLON PRETO</p>
    <p>CORRENTE FINA 7/8V MODELO DG51 116 INDEX</p>
    <p>ABRAC. SELIM 31.8 AL PRETO</p>
    <p>MOV DIRECAO M. OVER PRETO</p>
    <p>RODA LIVRE 7V INDEX 14/28D</p>
    <p>MANOPLA MTB PRETO</p>
    <p>SUPORTE MTB 31.8 PRETO</p>
    
    <h1><strong>DESCRIÇÃO BICICLETA 29 FIRST 24V</strong></h1>
    <p>QUADRO 29 ALUMÍNIO FIRST</p>
    <p>CÂMARAS BUTIL 48MM</p>
    <p>FREIO HIDRÁULICO </p>
    <p>GARFO 29 S. 28.6 AH SET PRETO, CURSO DE 100MM E TRAVA NO OMBRO</p>
    <p>MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO</p>
    <p>PEDIV TRIPLO ENC 24/34/42 PRETO</p>
    <p>PNEUS 29X2.10 SRI PTO</p>
    <p>CUBO AL K7 36 FUROS  C/ BLOCAGEM E ROLAMENTOS</p>
    <p>AROS 29/36 FUROS PRETO DISC</p>
    <p>SELIM MTB PRETO C/CARRINHO</p>
    <p>CANOTE ALUMÍNIO 27.2X350MM PRETO COM CARRINHO</p>
    <p>GUIDAO ALUMÍNIO 31.8 CURVO 20MM PRETO</p>
    <p>KIT TRANSMISSÃO 3X8V MICROSHIFT CASSETE 11/36D</p>
    <p>PEDAL 9/16 PLAT S/ ESFERA NYLON PRETO</p>
    <p>ABRAC. SELIM 31.8 ALUMÍNIO PRETO</p>
    <p>MOV DIRECAO M. OVER PRETO</p>
    <p>MANOPLA MTB PRETO</p>
    <p>SUPORTE MTB 31.8 PRETO</p>

    """
    
    # 1. Preparação dos Estilos
    styles = getSampleStyleSheet()
    
    # Estilo para os Títulos H1 (Descrição da Bicicleta) - Mantendo o tamanho 16 do usuário
    styleH1 = styles['Heading1']
    styleH1.fontSize = 16
    styleH1.leading = 20
    styleH1.textColor = COR_FAIXA_MEIO 
    styleH1.spaceBefore = 0.5 * cm
    styleH1.spaceAfter = 0.1 * cm

    # Estilo para o Conteúdo (Lista de Componentes) - Mantendo o tamanho 8 e leading 11 do usuário
    styleContent = styles['Normal']
    styleContent.fontSize = 8
    styleContent.leading = 11 
    styleContent.spaceBefore = 0
    styleContent.spaceAfter = 0
    
    # 2. Definição da Área de Desenho e Colunas
    x_margin = 2.5 * cm
    full_width = largura - 5 * cm
    col_spacing = 0.8 * cm # Espaçamento entre as colunas
    col_width = (full_width - col_spacing) / 2 # Largura de cada coluna
    
    y_start_pos = altura - MARGEM_SUPERIOR # Posição Y inicial abaixo do cabeçalho
    y_current = y_start_pos
    
    # 3. Processamento e Separação do Conteúdo em Blocos
    
    # Remove tags <br/> soltas para evitar erros de parsing e duplicação
    texto_especificacao_limpo = re.sub(r'</p>\s*<br\s*/>|<br\s*/>', '</p>', texto_especificacao_raw, flags=re.IGNORECASE)

    # Divide o texto em blocos (H1 + P's)
    blocos_texto = re.split(r'(<h1>.*?<\/strong><\/h1>)', texto_especificacao_limpo, flags=re.IGNORECASE | re.DOTALL)
    
    # Estrutura final: lista de tuplas [(H1_texto, [P1_texto, P2_texto, ...])]
    conteudo_formatado = []
    current_h1 = ""
    for bloco in blocos_texto:
        bloco = bloco.strip()
        if not bloco:
            continue
        
        # Se for um título H1
        if bloco.startswith('<h1'):
            current_h1 = bloco
            
        # Se for um bloco de parágrafos (componentes)
        elif current_h1:
            lista_itens = re.findall(r'<p>(.*?)<\/p>', bloco, re.DOTALL)
            
            # Extrai e ordena os itens da lista de componentes em ordem alfabética.
            componentes_ordenados = sorted([item.strip() for item in lista_itens if item.strip()])
            
            # Armazena o título e a lista de itens ordenados
            conteudo_formatado.append((current_h1, componentes_ordenados))
            current_h1 = "" # Reseta o título
            
    # 4. Loop de Desenho com Controle de Quebra de Página
    
    # Desenha o primeiro cabeçalho, rodapé E FUNDO
    draw_spec_page_elements(c, largura, altura, pagina_num, "ESPECIFICAÇÕES")

    for h1_text, components_list in conteudo_formatado:
        
        # --- A. Desenha o Título H1 (Spanning Full Width) ---
        p_h1 = Paragraph(h1_text, styleH1)
        p_h1_width, p_h1_height = p_h1.wrapOn(c, full_width, altura) 
        
        # Checagem de quebra de página para o título
        if y_current - p_h1_height < ALTURA_RODAPE + 0.5 * cm:
            c.showPage()
            pagina_num += 1
            draw_spec_page_elements(c, largura, altura, pagina_num, "ESPECIFICAÇÕES TÉCNICAS (Cont.)")
            y_current = y_start_pos
            
        # Desenha o H1
        y_current -= p_h1_height
        c.saveState()
        c.translate(x_margin, y_current)
        p_h1.drawOn(c, 0, 0)
        c.restoreState()
        
        y_current -= 0.1 * cm 
        
        # --- B. Desenha a Lista de Componentes em Duas Colunas ---
        
        total_items = len(components_list)
        if total_items == 0:
            continue

        # Divide a lista em duas metades (Left e Right)
        half_point = (total_items + 1) // 2
        left_half = components_list[:half_point]
        right_half = components_list[half_point:]
        
        # Converte as listas de volta para HTML, usando <br/> para forçar a quebra de linha
        # NOTA: O ReportLab usa o <br/> de forma mais robusta do que simplesmente juntar os parágrafos
        left_content_html = "<br/>".join(left_half)
        right_content_html = "<br/>".join(right_half) if right_half else ""

        p_left = Paragraph(left_content_html, styleContent)
        p_right = Paragraph(right_content_html, styleContent)
        
        # Pre-calcula a altura que as colunas irão ocupar
        p_left_width, p_left_height = p_left.wrapOn(c, col_width, altura)
        p_right_width, p_right_height = p_right.wrapOn(c, col_width, altura)
        
        # A altura que o bloco todo vai ocupar é a maior das duas colunas
        block_height = max(p_left_height, p_right_height)
        
        # Verifica se o bloco de 2 colunas cabe na página
        if y_current - block_height < ALTURA_RODAPE + 0.5 * cm:
            # Não cabe: Força quebra de página
            c.showPage()
            pagina_num += 1
            draw_spec_page_elements(c, largura, altura, pagina_num, "ESPECIFICAÇÕES TÉCNICAS (Cont.)")
            y_current = y_start_pos # Reseta a posição Y
            
            # Recalcula o wrap na nova altura (em caso de multi-página, para garantir a precisão)
            p_left.wrapOn(c, col_width, altura)
            p_right.wrapOn(c, col_width, altura)
            # A altura é recalculada, mas o wrapOn já define p_left_height e p_right_height

        # Atualiza a posição Y para o topo do bloco de duas colunas
        y_current -= block_height
        
        # --- Desenha Coluna Esquerda ---
        c.saveState()
        c.translate(x_margin, y_current) 
        p_left.drawOn(c, 0, 0)
        c.restoreState()
        
        # --- Desenha Coluna Direita ---
        if right_half:
            c.saveState()
            x_right_col = x_margin + col_width + col_spacing
            c.translate(x_right_col, y_current)
            p_right.drawOn(c, 0, 0)
            c.restoreState()
            
        # Adiciona o espaço total ocupado pelo bloco + margem de separação
        y_current -= 0.5 * cm 
        
    # Quebra de Página para o próximo conteúdo (produtos)
    c.showPage()
    
    # Retorna o novo número da página
    return pagina_num + 1

# === MODO DE GERAÇÃO FIXO (Mantido) ===
TIPO_ORDENACAO = 'C'
print("Catálogo configurado para ordenação por Categoria e Tamanho.")

# === LEITURA E PRÉ-PROCESSAMENTO DA PLANILHA (AJUSTADO PARA NOVO AGRUPAMENTO) ===
try:
    # Apenas para o ambiente de execução, simulamos a criação de um DataFrame se o arquivo não existir
    if not os.path.exists(excel_path):
        print(f"ATENÇÃO: Arquivo Excel não encontrado em: {excel_path}. Criando DataFrame de simulação.")
        data = {
            'Código do Produto': ['1001', '1002', '2001', '2002', '1003', '1004', '1005', '1006'],
            'Descrição': ['Bicicleta Aro 29 Pro', 'Bicicleta Aro 26 Infantil', 'Peça de Freio Disco', 'Peça de Selim Conforto', 'Bicicleta Urbano Light', 'Bicicleta Aro 29 Sport', 'Bicicleta Aro 29 Elite', 'Bicicleta Aro 26 Confort'],
            'Categoria': ['Bicicletas', 'Bicicletas', 'Componentes', 'Componentes', 'Bicicletas', 'Bicicletas', 'Bicicletas', 'Bicicletas'],
            # NOVA COLUNA TAMANHO DA BICICLETA
            'Tamanho da Bicicleta': ['17', '13', 'N/A', 'N/A', '15', '17', '19', '13'], 
            'Preço Antigo': [2500.00, 1200.00, None, None, 1800.00, 2100.00, 3500.00, 1300.00],
            'Preço Promoção': [2350.00, 1050.00, 45.00, 60.00, 1699.00, 1999.00, 3200.00, 1150.00]
        }
        df = pd.DataFrame(data)
    else:
        df = pd.read_excel(excel_path, dtype={'Código do Produto': str})
    
    # Assegura que todas as colunas de agrupamento são strings e trata NaN
    df['Categoria'] = df['Categoria'].fillna('Diversos').astype(str).str.strip()
    df['Tamanho da Bicicleta'] = df['Tamanho da Bicicleta'].fillna('N/A').astype(str).str.strip() # Tratamento da nova coluna
    
    # 1. ORDENAÇÃO: Primeiro por Categoria, depois por Tamanho da Bicicleta, e então por Código do Produto
    df = df.sort_values(by=['Categoria', 'Tamanho da Bicicleta', 'Código do Produto'])
    
    # 2. AGRUPAMENTO: Agrupa por Categoria E Tamanho da Bicicleta
    # O iterador será uma lista de tuplas (Nome da Categoria, DataFrame do Grupo)
    produtos_iteracao = df.groupby(['Categoria', 'Tamanho da Bicicleta'])
    
except Exception as e:
    print(f"ERRO: Falha ao ler o arquivo Excel ou criar o DataFrame: {e}")
    exit()

# === CRIAÇÃO DO PDF ===
c = canvas.Canvas(pdf_path, pagesize=A4)
largura, altura = A4

# Estilos para o Paragraph (descrição dos produtos)
styles = getSampleStyleSheet()
styleN = styles['Normal']
styleN.fontSize = 5 # Reduzido de 6 para 5.5
styleN.leading = 6 # Reduzido de 8 para 6.5
styleN.alignment = 1
styleN.fontName = 'Helvetica'
styleN.textColor = colors.black

# === CONFIGURAÇÕES DE LAYOUT DO PRODUTO (BLOCO MENOR) ===
produtos_por_linha = 3
espacamento_horizontal = 1 * cm
largura_produto_bloco = (largura - 3 * cm - 2 * espacamento_horizontal) / produtos_por_linha
altura_produto_bloco = 5.5 * cm 
espacamento_vertical = altura_produto_bloco + 0.3 * cm 
y_inicio_produtos = altura - MARGEM_SUPERIOR

# --- INÍCIO DA GERAÇÃO DO PDF ---
print("Iniciando geração da Capa...")

# 1. Gerar a Capa (Página 1)
criar_capa(c, largura, altura, logo_path, TIPO_ORDENACAO)
pagina = 2 # A próxima página a ser desenhada é a 2

# ** NOVO: 2. Gerar a Página de Especificação **
pagina = criar_pagina_especificacao(c, largura, altura, pagina) 

erros_imagem = 0
primeiro_grupo = True # Flag para tratar a primeira página de conteúdo

print(f"Iniciando conteúdo do catálogo (a partir da Página {pagina})...")

# 3. Loop Final para Conteúdo
# Itera sobre os grupos de categorias (agora a chave é uma tupla: (Categoria, Tamanho))
for (categoria_atual, tamanho_atual), df_grupo in produtos_iteracao:
    
    # Define o título do grupo (ex: Bicicletas - Tamanho 17)
    if tamanho_atual == 'N/A':
        titulo_grupo = f"{categoria_atual}" # Se não houver tamanho, exibe apenas a categoria
    else:
        titulo_grupo = f"{categoria_atual} - TAMANHO {tamanho_atual}"

    # Se não for o primeiro grupo E já houver conteúdo na página, força a quebra.
    if not primeiro_grupo:
        # 1. Garante que o rodapé e a quebra de página ocorram antes do novo grupo
        if produto_index_na_pagina != 0 or y != y_inicio_produtos:
             rodape(c, largura, altura, pagina)
             c.showPage()
             pagina += 1
    
    print(f"Processando Grupo: {titulo_grupo}")

    # Reconfigura a posição Y e o índice na página para o novo grupo
    y = y_inicio_produtos
    produto_index_na_pagina = 0
    primeiro_grupo = False
    
    # Desenha cabeçalho da nova página/grupo
    cabecalho(c, largura, altura, pagina, titulo_grupo)
    
    # Itera sobre os produtos do grupo (Iterrows retorna o índice e a série)
    for i, row in df_grupo.iterrows():
        col = produto_index_na_pagina % produtos_por_linha
        x_inicio = 1.5 * cm
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
            
        # --- POSICIONAMENTO DINÂMICO DE CÓDIGO ---
        
        largura_cod_btn = largura_produto_bloco * 0.4
        altura_cod_btn = 0.35 * cm
        x_cod_btn = x_bloco_centro - largura_cod_btn / 2
        y_cod_btn = y_img_area_fundo - 0.8 * cm 

        # Desenho do botão do código
        c.setFillColor(COR_DESTAQUE_CODIGO) # Usando a cor vermelha de destaque
        c.setFont("Helvetica-Bold", 6) 
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


        # === PRÓXIMO BLOCO / QUEBRA DE PÁGINA (DENTRO DO GRUPO) ===
        if col == produtos_por_linha - 1:
            y -= espacamento_vertical
            produto_index_na_pagina = 0
            
            if y - altura_produto_bloco < ALTURA_RODAPE + 0.5 * cm:
                # Quebra de página dentro do mesmo Grupo (Categoria + Tamanho)
                rodape(c, largura, altura, pagina)
                c.showPage()
                pagina += 1
                cabecalho(c, largura, altura, pagina, titulo_grupo) # Novo cabeçalho na nova página
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
if erros_imagem > 0:
    print(f"⚠️ {erros_imagem} imagem(ns) não encontrada(s) ou falhou no carregamento.")