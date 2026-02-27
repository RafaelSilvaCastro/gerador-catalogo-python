import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame, KeepInFrame, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER # Importação para alinhamento
import os
import re
from datetime import date

# === CONFIGURAÇÕES GERAIS ===
excel_path = "produtos.xlsx"
pdf_path = "pdfs/catalogo_amaisciclo_bicicletas_epecificações.pdf"
logo_path = "img/logo_amaisciclo.png"
img_dir = "img/img"

# Cores
COR_DESTAQUE_CODIGO = colors.Color(red=0.8, green=0.0, blue=0.0) 
COR_SOMBRA = colors.Color(0.8, 0.8, 0.8) 
COR_FUNDO_CARD = colors.white
COR_FUNDO_ESCURO = colors.black 
COR_FAIXA_MEIO = colors.Color(red=0.1, green=0.1, blue=0.1) 
COR_TEXTO_CLARO = colors.white

# Variáveis de Layout
ALTURA_RODAPE = 1.5 * cm
ALTURA_CABECALHO = 1.5 * cm
MARGEM_SUPERIOR = ALTURA_CABECALHO + 0.5 * cm

# === TEXTOS DAS DIVISÓRIAS ===
TEXTOS_DIVISORIAS = {
    "DESCRIÇÃO BICICLETA 26 VIKING": (
    "CAMARAS 26 BUTIL<br/>"
    "CARRINHO DE SELIM IMP PRETO<br/>"
    "FREIO DISCO DT/TR (PINCA/ROTOR)<br/>"
    "GARFO 29 S. 28.6 AH SET PRETO<br/>"
    "MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO<br/>"
    "PEDIV TRIPLO ENC 24/34/42 PRETO<br/>"
    "PNEUS 26X1.95 KENDA K-90 SLICK PTO<br/>"
    "ARO 26 VMAXX SL VZAN 36F PTO DISC S/ILHOS<br/>"
    "RAIO 251X2.0MM IMP ZINC<br/>"
    "CUBO AL DT/TR ESFERADO PRETO C/ BLOCAGEM<br/>"
    "SELIM MTB PRETO C/CARRINHO<br/>"
    "CANOTE AÇO 27.2X350MM PRETO<br/>"
    "GUIDAO DH 31.8 ACO 700MM PRETO<br/>"
    "KIT TRANSMISSÃO 21V<br/>"
    "PEDAL 9/16 PLATAFORMA NYLON PRETO<br/>"
    "CORRENTE FINA 7/8V MODELO DG51 116 INDEX<br/>"
    "ABRAC. SELIM 31.8 AL PRETO<br/>"
    "MOV DIRECAO M. OVER PRETO<br/>"
    "RODA LIVRE 7V INDEX 14/28D<br/>"
    "MANOPLA MTB PRETO<br/>"
    "SUP MTB 31.8 (60MM) OU SIMILAR PTO."
    ),
    
    "DESCRIÇÃO BICICLETA 29 FIRST 21V": (
    "QUADRO 29 ALUMÍNIO FIRST<br/>"
    "CÂMARAS BUTIL 48MM<br/>"
    "CARRINHO DE SELIM IMP PRETO<br/>"
    "FREIO DISCO DT/TR (PINCA/ROTOR)<br/>"
    "GARFO 29 S. 28.6 AH SET PRETO, CANELAS DE 38MM E CURSO DE 80MM<br/>"
    "MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO<br/>"
    "PEDIV TRIPLO ENC 24/34/42 PRETO<br/>"
    "PNEUS 29X2.10 SRI PTO<br/>"
    "CUBO AL DT/TR 36 FUROS ESFERADO C/ BLOCAGEM PRETO<br/>"
    "AROS 29/36 FUROS PRETO DISC<br/>"
    "SELIM MTB PRETO C/CARRINHO<br/>"
    "CANOTE AÇO 27.2X350MM PRETO<br/>"
    "GUIDAO DH 31.8 ACO 700MM PRETO<br/>"
    "KIT TRANSMISSÃO 21V<br/>"
    "PEDAL 9/16 PLATAFORMA NYLON PRETO<br/>"
    "CORRENTE FINA 7/8V MODELO DG51 116 INDEX<br/>"
    "ABRAC. SELIM 31.8 AL PRETO<br/>"
    "MOV DIRECAO M. OVER PRETO<br/>"
    "RODA LIVRE 7V INDEX 14/28D<br/>"
    "MANOPLA MTB PRETO<br/>"
    "SUPORTE MTB 31.8 PRETO."
    ),
    
    "DESCRIÇÃO BICICLETA 29 FIRST 24V": (
    "QUADRO 29 ALUMÍNIO FIRST<br/>"
    "CÂMARAS BUTIL 48MM<br/>"
    "FREIO HIDRÁULICO <br/>"
    "GARFO 29 S. 28.6 AH SET PRETO, CURSO DE 100MM E TRAVA NO OMBRO<br/>"
    "MOV CENTRAL 34.7/122MM C/ ROLAMENTO SELADO<br/>"
    "PEDIV TRIPLO ENC 24/34/42 PRETO<br/>"
    "PNEUS 29X2.10 SRI PTO<br/>"
    "CUBO AL K7 36 FUROS  C/ BLOCAGEM E ROLAMENTOS<br/>"
    "AROS 29/36 FUROS PRETO DISC<br/>"
    "SELIM MTB PRETO C/CARRINHO<br/>"
    "CANOTE ALUMÍNIO 27.2X350MM PRETO COM CARRINHO<br/>"
    "GUIDAO ALUMÍNIO 31.8 CURVO 20MM PRETO<br/>"
    "KIT TRANSMISSÃO 3X8V MICROSHIFT CASSETE 11/36D<br/>"
    "PEDAL 9/16 PLAT S/ ESFERA NYLON PRETO<br/>"
    "ABRAC. SELIM 31.8 ALUMÍNIO PRETO<br/>"
    "MOV DIRECAO M. OVER PRETO<br/>"
    "MANOPLA MTB PRETO<br/>"
    "SUPORTE MTB 31.8 PRETO."
    )
}

# === FUNÇÕES DE SUPORTE ===

def normalize_code(code_str):
    code_str = str(code_str).strip()
    tentativas = [code_str]
    cleaned_code = re.sub(r'[.,]', '', code_str)
    if cleaned_code != code_str: tentativas.append(cleaned_code)
    underscore_code = code_str.replace('.', '_')
    if underscore_code != code_str and underscore_code != cleaned_code: tentativas.append(underscore_code)
    return list(set(tentativas))

def cabecalho(c, largura, altura, pagina, titulo_grupo=""):
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, altura - ALTURA_CABECALHO, largura, ALTURA_CABECALHO, fill=True, stroke=0)
    try:
        c.drawImage(logo_path, 2 * cm, altura - ALTURA_CABECALHO + 0.3 * cm, width=3.0 * cm, preserveAspectRatio=True, mask='auto')
    except: pass
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(1.5 * cm, altura - ALTURA_CABECALHO + 0.5 * cm, f"CATÁLOGO: {titulo_bike}")
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.line(1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm, largura - 1.5 * cm, altura - ALTURA_CABECALHO - 0.1 * cm)

def rodape(c, largura, altura, pagina):
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, largura, ALTURA_RODAPE, fill=True, stroke=0)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 8)
    c.drawString(2 * cm, 0.5 * cm, "Acesse: b2b.amaisciclo.com.br")
    c.drawRightString(largura - 2 * cm, 0.5 * cm, f"Página {pagina}")

def criar_capa(c, largura, altura, logo_path):
    c.setFillColor(COR_FUNDO_ESCURO)
    c.rect(0, 0, largura, altura, fill=1, stroke=0)
    try:
        larg_logo = 13 * cm
        c.drawImage(logo_path, (largura - larg_logo)/2, altura * 0.65, width=larg_logo, preserveAspectRatio=True, mask='auto')
    except: pass
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(largura / 2, altura * 0.45, "CATÁLOGO DE BICICLETAS")
    c.setFillColor(colors.red)
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(largura / 2, altura * 0.40, "EDIÇÃO - 2026")
    c.setFillColor(COR_TEXTO_CLARO)
    c.setFont("Helvetica", 10)
    data_geracao = date.today().strftime("%d/%m/%Y")
    c.drawCentredString(largura / 2, 2 * cm, f"Emissão {data_geracao}")
    c.showPage()

def criar_pagina_divisoria(c, largura, altura, tipo_nome):
    c.setFillColor(COR_FUNDO_ESCURO)
    c.rect(0, 0, largura, altura, fill=1, stroke=0)
    
    # Faixa Cinza
    c.setFillColor(COR_FAIXA_MEIO)
    c.rect(0, altura * 0.15, largura, altura * 0.70, fill=1, stroke=0)
    
    estilos = getSampleStyleSheet()
    
    # Estilo do Título (Centralizado)
    estilo_titulo = ParagraphStyle(
        'DivTitle',
        parent=estilos['Normal'],
        fontName='Helvetica-Bold',
        fontSize=32,
        textColor=colors.red,
        alignment=TA_CENTER, # Centralizado
        leading=36,
        spaceAfter=20
    )
    
    # Estilo do Corpo (Justificado)
    estilo_corpo = ParagraphStyle(
        'DivBody',
        parent=estilos['Normal'],
        fontName='Helvetica',
        fontSize=14,
        textColor=colors.white,
        alignment=TA_JUSTIFY, # JUSTIFICADO
        leading=18,
        leftIndent=20,  # Recuo para não colar na borda
        rightIndent=20
    )
    
    conteudo = []
    conteudo.append(Paragraph(tipo_nome.upper(), estilo_titulo))
    
    texto_desc = TEXTOS_DIVISORIAS.get(tipo_nome, "Descrição técnica em breve.")
    conteudo.append(Paragraph(texto_desc, estilo_corpo))
    
    larg_f = largura - 4*cm
    alt_f = altura * 0.60
    # Centralizado verticalmente na página
    f = Frame(2*cm, altura * 0.20, larg_f, alt_f, showBoundary=0)
    
    f.addFromList([KeepInFrame(larg_f, alt_f, conteudo)], c)
    c.showPage()

# === PROCESSAMENTO DOS DADOS ===
try:
    df = pd.read_excel(excel_path, dtype={'Código do Produto': str})
    df['Categoria'] = df['Categoria'].fillna('Diversos').astype(str).str.strip()
    df['Tamanho da Bicicleta'] = df['Tamanho da Bicicleta'].fillna('N/A').astype(str).str.strip()
    df['Tipo'] = df['Tipo'].fillna('OUTROS').astype(str).str.strip()
    df = df.sort_values(by=['Tipo', 'Categoria', 'Tamanho da Bicicleta', 'Código do Produto'])
    produtos_iteracao = df.groupby(['Tipo', 'Categoria', 'Tamanho da Bicicleta'])
except Exception as e:
    print(f"Erro: {e}"); exit()

# === GERAÇÃO DO PDF ===
if not os.path.exists('pdfs'): os.makedirs('pdfs')
c = canvas.Canvas(pdf_path, pagesize=A4)
largura, altura = A4

criar_capa(c, largura, altura, logo_path)
pagina = 2
tipo_atual_rastreio = None

# Layout Grid
prod_por_linha = 3
espac_h = 1 * cm
larg_prod = (largura - 3 * cm - 2 * espac_h) / prod_por_linha
alt_prod = 5.5 * cm 
y_inicio = altura - MARGEM_SUPERIOR

styleN = getSampleStyleSheet()['Normal']
styleN.fontSize = 5
styleN.leading = 6
styleN.alignment = TA_CENTER

for (tipo, categoria, tamanho), df_grupo in produtos_iteracao:
    if tipo != tipo_atual_rastreio:
        criar_pagina_divisoria(c, largura, altura, tipo)
        tipo_atual_rastreio = tipo
        pagina += 1

    titulo_bike = f"{categoria} - TAM {tamanho}" if tamanho != 'N/A' else f"{categoria}"
    titulo_grupo = f"{categoria} {tipo} - TAM {tamanho}" if tamanho != 'N/A' else f"{categoria} {tipo}"
    y = y_inicio
    idx = 0
    cabecalho(c, largura, altura, pagina, titulo_grupo)

    for i, row in df_grupo.iterrows():
        col = idx % prod_por_linha
        x_bloco = 1.5 * cm + col * (larg_prod + espac_h)
        x_centro = x_bloco + larg_prod / 2

        # Card
        c.setFillColor(COR_SOMBRA)
        c.roundRect(x_bloco + 0.05*cm, y - alt_prod + 0.05*cm, larg_prod, alt_prod, 0.2*cm, fill=1, stroke=0)
        c.setFillColor(COR_FUNDO_CARD)
        c.setStrokeColor(COR_SOMBRA)
        c.roundRect(x_bloco, y - alt_prod, larg_prod, alt_prod, 0.2*cm, fill=1, stroke=1)
        
        # Imagem
        max_h_img = 3.5 * cm 
        y_img_fundo = y - 0.3 * cm - max_h_img
        codigo_produto = str(row.get("Código do Produto", "")).strip()
        image_loaded = False
        
        for cod in normalize_code(codigo_produto):
            for ext in [".jpg", ".jpeg", ".png"]:
                tentativa = os.path.join(img_dir, f"{cod}{ext}")
                if os.path.exists(tentativa):
                    try:
                        c.drawImage(ImageReader(tentativa), x_bloco + 0.1*cm, y_img_fundo, width=larg_prod-0.2*cm, height=max_h_img, preserveAspectRatio=True, mask='auto')
                        image_loaded = True
                        break
                    except: pass
            if image_loaded: break
        
        if not image_loaded:
            c.setFillColor(colors.lightgrey)
            c.drawCentredString(x_centro, y_img_fundo + max_h_img/2, "Sem imagem")

        # Código
        y_cod = y_img_fundo - 0.6 * cm 
        c.setFillColor(COR_DESTAQUE_CODIGO)
        c.roundRect(x_centro - 1*cm, y_cod, 2*cm, 0.35 * cm, 0.15 * cm, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 6)
        c.drawCentredString(x_centro, y_cod + 0.10 * cm, codigo_produto)

        # Descrição
        c.setFillColor(colors.black)
        desc = " ".join(str(row.get("Descrição", "")).split())
        p = Paragraph(desc, styleN)
        pw, ph = p.wrap(larg_prod * 0.9, 0.8 * cm)
        p.drawOn(c, x_centro - pw / 2, y - alt_prod + 0.2 * cm)

        if col == prod_por_linha - 1:
            y -= alt_prod + 0.3 * cm
            idx = 0
            if y - alt_prod < ALTURA_RODAPE + 0.5 * cm:
                rodape(c, largura, altura, pagina)
                c.showPage()
                pagina += 1
                cabecalho(c, largura, altura, pagina, titulo_grupo)
                y = y_inicio
        else: idx += 1

    rodape(c, largura, altura, pagina)
    c.showPage()
    pagina += 1

c.save()